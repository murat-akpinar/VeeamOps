# Veeam Backup Job ve VM Listesi - Excel Export

# Veeam PowerShell modülünü yükle
try {
    Add-PSSnapin -Name VeeamPSSnapIn -ErrorAction Stop
} catch {
    try {
        Import-Module Veeam.Backup.PowerShell -ErrorAction Stop
    } catch {
        Write-Error "Veeam PowerShell modülü yüklenemedi. Script Veeam Backup Server üzerinde çalıştırılmalidir."
        exit 1
    }
}

# ─── YARDIMCI FONKSİYONLAR ───────────────────────────────────────────────────

$gunTR = @{
    "Sunday"    = "Pazar"
    "Monday"    = "Pazartesi"
    "Tuesday"   = "Sali"
    "Wednesday" = "Carsamba"
    "Thursday"  = "Persembe"
    "Friday"    = "Cuma"
    "Saturday"  = "Cumartesi"
}

function Get-ScheduleGunler {
    param($job)
    try {
        $sched = $job.ScheduleOptions

        # Job sonrasi mi?
        try { if ($sched.OptionsScheduleAfterJob.IsEnabled -eq $true) { return "Job Sonrasi" } } catch {}

        # Periyodik mi?
        try { if ($sched.OptionsPeriodically.Enabled) { return "Surekli/Periyodik" } } catch {}

        # Aylik mi?
        try {
            if ($sched.OptionsMonthly.Enabled) {
                $m = $sched.OptionsMonthly
                return "Aylik ($($m.DayNumberInMonth) $($m.DayOfWeek))"
            }
        } catch {}

        # Gunluk — once kind kontrolü (Everyday/WeekDays/Weekend/SpecificDays)
        try {
            $daily = $sched.OptionsDaily
            if ($daily -ne $null) {
                $kind = $null
                try { $kind = $daily.DayNumberInMonth.ToString() } catch {}
                if (-not $kind) { try { $kind = $daily.Kind.ToString() } catch {} }

                if ($kind -match "Everyday|Every") { return "Her Gun" }
                if ($kind -match "WeekDay")         { return "Is Gunleri (Pzt-Cum)" }
                if ($kind -match "Weekend")         { return "Hafta Sonu (Cmt-Paz)" }

                # Spesifik gunler — DaysSrv gercek property adi, Days eski/alternatif
                $days = @()
                try { $days = @($daily.DaysSrv | Where-Object { $_ -ne $null }) } catch {}
                if ($days.Count -eq 0) { try { $days = @($daily.Days | Where-Object { $_ -ne $null }) } catch {} }
                if ($days.Count -gt 0) {
                    if ($days.Count -ge 7) { return "Her Gun" }
                    $isimler = $days | Where-Object { $_ -ne $null } | ForEach-Object {
                        $key = $_.ToString().Trim()
                        if ($gunTR.ContainsKey($key)) { $gunTR[$key] } else { $key }
                    }
                    return ($isimler -join ", ")
                }
            }
        } catch {}

        return "N/A"
    } catch { return "N/A" }
}

function Get-ScheduleBaslamaSaati {
    param($job)
    try {
        $sched = $job.ScheduleOptions

        # Dogru property adi: StartDateTimeLocal (StartDateTime degil)
        try {
            $dt = $sched.StartDateTimeLocal
            if ($dt -and $dt -is [datetime] -and $dt.Year -gt 1970) {
                return $dt.ToString("HH:mm")
            }
        } catch {}

        # Aylik schedule'in kendi saati
        try {
            if ($sched.OptionsMonthly.Enabled -eq $true) {
                $t = $sched.OptionsMonthly.Time
                if ($t -and $t -is [datetime] -and $t.Year -gt 1) {
                    return $t.ToString("HH:mm")
                }
            }
        } catch {}

        return "N/A"
    } catch { return "N/A" }
}

function Get-RetentionPolicy {
    param($job)
    try {
        $bso     = $job.BackupStorageOptions
        $retType = try { $bso.RetentionType.ToString() } catch { "Cycles" }
        if ($retType -match "Day") {
            $days = try { $bso.RetainDaysToKeep } catch { try { $bso.RetainDays } catch { 0 } }
            return "$days gun"
        } else {
            $cycles = try { $bso.RetainCycles } catch { 0 }
            return "$cycles rp"
        }
    } catch { return "N/A" }
}

function Get-SyntheticFullGunu {
    param($job)
    # NOT: Veeam XML'de "Syntethic" (yazim hatasi) kullaniliyor — "Synthetic" degil
    try {
        $xmlDoc = $job.Options.Options.Document
        if ($null -eq $xmlDoc) { return "Yok" }

        $sfNode = $xmlDoc.SelectSingleNode("//TransformFullToSyntethic")
        $sfEnabled = ($sfNode -ne $null -and $sfNode.InnerText -eq "True")

        $sfDayNode = $xmlDoc.SelectSingleNode("//TransformToSyntethicDays/DayOfWeek")
        if ($sfEnabled -and $sfDayNode -ne $null -and $sfDayNode.InnerText) {
            $d = $sfDayNode.InnerText.Trim()
            if ($gunTR.ContainsKey($d)) { return $gunTR[$d] } else { return $d }
        }

        $afNode    = $xmlDoc.SelectSingleNode("//EnableFullBackup")
        $afEnabled = ($afNode -ne $null -and $afNode.InnerText -eq "True")

        $afDayNode = $xmlDoc.SelectSingleNode("//FullBackupDays/DayOfWeek")
        if ($afEnabled -and $afDayNode -ne $null -and $afDayNode.InnerText) {
            $d = $afDayNode.InnerText.Trim()
            if ($gunTR.ContainsKey($d)) { return $gunTR[$d] } else { return $d }
        }

        return "Yok"
    } catch { return "N/A" }
}

function Get-FullBackupGunu {
    param($job)
    try {
        $parts = [System.Collections.Generic.List[string]]::new()

        $desiredTR = @{
            "First"     = "Ilk";    "Second" = "Ikinci"; "Third" = "Ucuncu"
            "Fourth"    = "Dorduncu"; "Last"  = "Son"
            "January"   = "Ocak";   "February" = "Subat";    "March"     = "Mart"
            "April"     = "Nisan";  "May"      = "Mayis";    "June"      = "Haziran"
            "July"      = "Temmuz"; "August"   = "Agustos";  "September" = "Eylul"
            "October"   = "Ekim";   "November" = "Kasim";    "December"  = "Aralik"
        }

        # ── GfsPolicy (UI'daki "Configure GFS" diyalogu) ─────────────────────
        try {
            $gfs = $job.Options.GfsPolicy
            if ($gfs -ne $null -and $gfs.IsEnabled -eq $true) {

                # Haftalik
                try {
                    $w = $gfs.Weekly
                    if ($w -ne $null -and $w.IsEnabled -eq $true) {
                        $keep = $w.KeepBackupsForNumberOfWeeks
                        $g    = try { $w.DesiredTime.ToString() } catch { "" }
                        $gun  = if ($gunTR.ContainsKey($g)) { $gunTR[$g] } else { $g }
                        $parts.Add("Haftalik:$gun ($keep hafta)")
                    }
                } catch {}

                # Aylik
                try {
                    $m = $gfs.Monthly
                    if ($m -ne $null -and $m.IsEnabled -eq $true) {
                        $keep    = $m.KeepBackupsForNumberOfMonths
                        $raw     = try { $m.DesiredTime.ToString() } catch { $null }
                        $desired = if ($raw -and $desiredTR.ContainsKey($raw)) { $desiredTR[$raw] } else { $raw }
                        if ($desired) { $parts.Add("Aylik:$desired ($keep ay)") }
                        else          { $parts.Add("Aylik ($keep ay)") }
                    }
                } catch {}

                # Yillik
                try {
                    $y = $gfs.Yearly
                    if ($y -ne $null -and $y.IsEnabled -eq $true) {
                        $keep    = $y.KeepBackupsForNumberOfYears
                        $raw     = try { $y.DesiredTime.ToString() } catch { $null }
                        $desired = if ($raw -and $desiredTR.ContainsKey($raw)) { $desiredTR[$raw] } else { $raw }
                        if ($desired) { $parts.Add("Yillik:$desired ($keep yil)") }
                        else          { $parts.Add("Yillik ($keep yil)") }
                    }
                } catch {}

                if ($parts.Count -gt 0) { return $parts -join " | " }
                return "GFS Aktif"
            }
        } catch {}

        # ── Eski GenerationPolicy fallback ────────────────────────────────────
        try {
            $gp = $job.Options.GenerationPolicy
            if ($gp.GFSWeeklyBackupsEnabled -eq $true) {
                $g   = $gp.WeeklyBackupDayOfWeek.ToString()
                $gun = if ($gunTR.ContainsKey($g)) { $gunTR[$g] } else { $g }
                $parts.Add("Haftalik:$gun ($($gp.GFSWeeklyBackups) hafta)")
            }
            if ($gp.GFSMonthlyBackupsEnabled -eq $true) { $parts.Add("Aylik ($($gp.GFSMonthlyBackups) ay)") }
            if ($gp.GFSYearlyBackupsEnabled  -eq $true) { $parts.Add("Yillik ($($gp.GFSYearlyBackups) yil)") }
            if ($parts.Count -gt 0) { return $parts -join " | " }
        } catch {}

        # ── Aktif Full Backup ─────────────────────────────────────────────────
        if ($job.BackupStorageOptions.EnableFullBackup -eq $true) {
            return "Aktif Full Backup"
        }

        return "GFS Yok"
    } catch { return "N/A" }
}

# ─── VERİ TOPLAMA ─────────────────────────────────────────────────────────────

Write-Host "`n=== VEEAM BACKUP JOB LİSTESİ ===" -ForegroundColor Cyan
Write-Host "Tarih: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host ("=" * 60) -ForegroundColor Cyan

$jobs = Get-VBRJob | Where-Object { $_.JobType -eq "Backup" } | Sort-Object Name

if ($jobs.Count -eq 0) {
    Write-Host "Hic backup job bulunamadi." -ForegroundColor Yellow
    exit 0
}

Write-Host "Toplam Job Sayisi: $($jobs.Count) — veriler toplanıyor...`n" -ForegroundColor Green

$jobVeriler = foreach ($job in $jobs) {
    Write-Host "  Isleniyor: $($job.Name)" -ForegroundColor Gray

    $objects      = @($job.GetObjectsInJob() | Sort-Object Name)
    $gunler       = Get-ScheduleGunler       -job $job
    $saat         = Get-ScheduleBaslamaSaati -job $job
    $synthFullGun = Get-SyntheticFullGunu    -job $job
    $fullGun      = Get-FullBackupGunu       -job $job
    $retPolicy    = Get-RetentionPolicy      -job $job

    [PSCustomObject]@{
        Job          = $job
        Objects      = $objects
        Durum        = if ($job.IsScheduleEnabled -eq $true) { "Aktif" } else { "Pasif" }
        SonCalisma   = if ($job.LatestRunLocal) { $job.LatestRunLocal.ToString("dd.MM.yyyy HH:mm") } else { "-" }
        Gunler       = $gunler
        Saat         = $saat
        RetPolicy    = $retPolicy
        SynthFullGun = $synthFullGun
        FullGun      = $fullGun
    }
}

# Ekran çıktısı
Write-Host ""
foreach ($v in $jobVeriler) {
    Write-Host "JOB: $($v.Job.Name)" -ForegroundColor Yellow
    Write-Host "  Durum           : $($v.Durum)"
    Write-Host "  Son Calisma     : $($v.SonCalisma)"
    Write-Host "  Calisma Gunleri : $($v.Gunler)"
    Write-Host "  Baslama Saati   : $($v.Saat)"
    Write-Host "  Retention Policy: $($v.RetPolicy)"
    Write-Host "  Full Backup     : $($v.SynthFullGun)"
    Write-Host "  GFS             : $($v.FullGun)"
    Write-Host "  VM Sayisi       : $($v.Objects.Count)"
    foreach ($obj in $v.Objects) { Write-Host "    - $($obj.Name)" -ForegroundColor Gray }
    Write-Host ""
}

# ─── EXCEL EXPORT ─────────────────────────────────────────────────────────────

$outputPath = "C:\scripts\Veeam_Backup_Jobs_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
if (-not (Test-Path "C:\scripts")) { New-Item -ItemType Directory -Path "C:\scripts" | Out-Null }

Write-Host "Excel verisi hazirlaniyor..." -ForegroundColor Cyan

# ── 1. Tüm satırları önce PowerShell listesinde oluştur ──────────────────────
#    (COM'a tek seferde yazacağız — GC sorununu önler)

$COLS = 10
$genislikler = @(5, 44, 11, 20, 11, 36, 14, 18, 22, 40)
$basliklar   = @("No","Job Adi / VM Adi","Durum","Son Calisma","VM Sayisi",
                 "Calisma Gunleri","Baslama Saati","Retention Policy","Full Backup","GFS")

# rowData  : her eleman bir object[] (satır değerleri)
# rowMeta  : her eleman string (header / job:Aktif / job:Pasif / vm:çift / vm:tek)
$rowData = [System.Collections.Generic.List[object[]]]::new()
$rowMeta = [System.Collections.Generic.List[string]]::new()

$rowData.Add([object[]]$basliklar)
$rowMeta.Add("header")

$jobIdx = 1
foreach ($v in $jobVeriler) {
    $rowData.Add([object[]]@(
        [object]$jobIdx,
        [object]$v.Job.Name,
        [object]$v.Durum,
        [object]$v.SonCalisma,
        [object]$v.Objects.Count,
        [object]$v.Gunler,
        [object]$v.Saat,
        [object]$v.RetPolicy,
        [object]$v.SynthFullGun,
        [object]$v.FullGun
    ))
    $rowMeta.Add("job:$($v.Durum)")

    $vmIdx = 1
    foreach ($obj in $v.Objects) {
        $rowData.Add([object[]]@(
            [object]"",
            [object]("     $([char]0x21B3) VM $vmIdx  -  $($obj.Name)"),
            [object]"", [object]"", [object]"",
            [object]"", [object]"", [object]"", [object]"", [object]""
        ))
        $rowMeta.Add("vm:$(($vmIdx % 2))")
        $vmIdx++
    }
    $jobIdx++
}

$totalRows = $rowData.Count

Write-Host "xlsx olusturuluyor ($totalRows satir, COM kullanilmiyor)..." -ForegroundColor Cyan

# ── Yardimci fonksiyonlar ─────────────────────────────────────────────────────
function XE  { param([string]$s) [System.Security.SecurityElement]::Escape($s) }
function CL  { param([int]$n)    [string][char](64 + $n) }   # 1=A, 2=B ... 9=I

# ── OOXML sabit dosyalari ─────────────────────────────────────────────────────
$ctXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"           ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"  ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml"             ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'

$relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'

$wbRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"    Target="styles.xml"/>
</Relationships>'

$wbXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Veeam Jobs" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>'

# Stil indeksleri: 0=Varsayilan, 1=Baslik, 2=Job, 3=VM-tek, 4=VM-cift, 5=Job-Pasif
$stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="10"/><name val="Calibri"/><color rgb="FF000000"/></font>
    <font><b/><sz val="10"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>
    <font><sz val="10"/><name val="Calibri"/><color rgb="FF000000"/></font>
  </fonts>
  <fills count="7">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF2F5496"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDEEAF1"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFBDD7EE"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF7B0000"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left   style="thin"><color rgb="FFBBBBBB"/></left>
      <right  style="thin"><color rgb="FFBBBBBB"/></right>
      <top    style="thin"><color rgb="FFBBBBBB"/></top>
      <bottom style="thin"><color rgb="FFBBBBBB"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="2" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'

# ── Sheet XML: tüm satırları StringBuilder ile oluştur ───────────────────────
$sb = [System.Text.StringBuilder]::new(1MB)
$null = $sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
$null = $sb.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
$null = $sb.AppendLine('  <sheetViews><sheetView workbookViewId="0">')
$null = $sb.AppendLine('    <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>')
$null = $sb.AppendLine('  </sheetView></sheetViews>')
$null = $sb.AppendLine('  <cols>')
for ($c = 1; $c -le $COLS; $c++) {
    $null = $sb.AppendLine("    <col min=""$c"" max=""$c"" width=""$($genislikler[$c-1])"" customWidth=""1""/>")
}
$null = $sb.AppendLine('  </cols>')
$null = $sb.AppendLine('  <sheetData>')

for ($r = 0; $r -lt $totalRows; $r++) {
    $er   = $r + 1
    $meta = $rowMeta[$r]
    $ht   = switch -Wildcard ($meta) { "header" {22} "job:*" {20} default {17} }

    $null = $sb.AppendLine("    <row r=""$er"" ht=""$ht"" customHeight=""1"">")

    for ($c = 1; $c -le $COLS; $c++) {
        $cref = "$(CL $c)$er"
        $val  = $rowData[$r][$c - 1]

        $si = switch -Wildcard ($meta) {
            "header"    { 1 }
            "job:Pasif" { 5 }
            "job:*"     { 2 }
            "vm:1"      { 3 }
            "vm:0"      { 4 }
            default     { 0 }
        }

        if ($null -eq $val -or "$val" -eq "") {
            $null = $sb.AppendLine("      <c r=""$cref"" s=""$si""/>")
        } elseif ($val -is [int] -or $val -is [long] -or $val -is [double]) {
            $null = $sb.AppendLine("      <c r=""$cref"" s=""$si""><v>$val</v></c>")
        } else {
            $null = $sb.AppendLine("      <c r=""$cref"" s=""$si"" t=""inlineStr""><is><t>$(XE "$val")</t></is></c>")
        }
    }

    $null = $sb.AppendLine("    </row>")
}

$null = $sb.AppendLine('  </sheetData>')
$null = $sb.AppendLine('</worksheet>')
$sheetXml = $sb.ToString()

# ── Temp dizinine XML dosyalarini yaz ─────────────────────────────────────────
$tmp = "$env:TEMP\vxlsx_$(Get-Random)"
$null = New-Item -ItemType Directory "$tmp"
$null = New-Item -ItemType Directory "$tmp\_rels"
$null = New-Item -ItemType Directory "$tmp\xl"
$null = New-Item -ItemType Directory "$tmp\xl\_rels"
$null = New-Item -ItemType Directory "$tmp\xl\worksheets"

$enc = [System.Text.UTF8Encoding]::new($false)   # UTF-8 BOM'suz
[System.IO.File]::WriteAllText("$tmp\[Content_Types].xml",         $ctXml,     $enc)
[System.IO.File]::WriteAllText("$tmp\_rels\.rels",                 $relsXml,   $enc)
[System.IO.File]::WriteAllText("$tmp\xl\_rels\workbook.xml.rels",  $wbRelsXml, $enc)
[System.IO.File]::WriteAllText("$tmp\xl\workbook.xml",             $wbXml,     $enc)
[System.IO.File]::WriteAllText("$tmp\xl\styles.xml",               $stylesXml, $enc)
[System.IO.File]::WriteAllText("$tmp\xl\worksheets\sheet1.xml",    $sheetXml,  $enc)

# ── ZIP → xlsx ────────────────────────────────────────────────────────────────
Add-Type -AssemblyName System.IO.Compression.FileSystem
if (Test-Path $outputPath) { Remove-Item $outputPath -Force }
[System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $outputPath)
Remove-Item -Recurse -Force $tmp

Write-Host "Excel dosyasi kaydedildi: $outputPath" -ForegroundColor Green
Write-Host "=== BITTI ===" -ForegroundColor Cyan

if (Test-Path $outputPath) { Start-Process $outputPath }
