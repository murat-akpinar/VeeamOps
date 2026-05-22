# Veeam - Job durum tablosu: Full Backup ve GFS analizi

try {
    Add-PSSnapin -Name VeeamPSSnapIn -ErrorAction Stop
} catch {
    try {
        Import-Module Veeam.Backup.PowerShell -ErrorAction Stop
    } catch {
        Write-Error "Veeam PowerShell modulu yuklenemedi. Script Veeam Backup Server uzerinde calistirilmalidir."
        exit 1
    }
}

# ─── YARDIMCI ─────────────────────────────────────────────────────────────────

$gunTR = @{
    Sunday    = "Pazar"; Monday    = "Pazartesi"; Tuesday   = "Sali"
    Wednesday = "Carsamba"; Thursday  = "Persembe"; Friday    = "Cuma"
    Saturday  = "Cumartesi"
}

$desiredTR = @{
    First = "Ilk"; Second = "Ikinci"; Third = "Ucuncu"; Fourth = "Dorduncu"; Last = "Son"
    January = "Oca"; February = "Sub"; March = "Mar"; April = "Nis"
    May = "May"; June = "Haz"; July = "Tem"; August = "Agu"
    September = "Eyl"; October = "Eki"; November = "Kas"; December = "Ara"
}

function TR { param($tbl, $k) if ($tbl.ContainsKey($k)) { $tbl[$k] } else { $k } }

function Get-FullBackupBilgi {
    param($job)
    try {
        $xmlDoc = $job.Options.Options.Document
        if ($null -eq $xmlDoc) { return "Yok" }

        $sfNode = $xmlDoc.SelectSingleNode("//TransformFullToSyntethic")
        if ($sfNode -ne $null -and $sfNode.InnerText -eq "True") {
            $dn = $xmlDoc.SelectSingleNode("//TransformToSyntethicDays/DayOfWeek")
            $gun = if ($dn -ne $null -and $dn.InnerText) { TR $gunTR $dn.InnerText.Trim() } else { "?" }
            return "Synth:$gun"
        }

        $afNode = $xmlDoc.SelectSingleNode("//EnableFullBackup")
        if ($afNode -ne $null -and $afNode.InnerText -eq "True") {
            $dn = $xmlDoc.SelectSingleNode("//FullBackupDays/DayOfWeek")
            $gun = if ($dn -ne $null -and $dn.InnerText) { TR $gunTR $dn.InnerText.Trim() } else { "?" }
            return "Aktif:$gun"
        }

        return "Yok"
    } catch { return "N/A" }
}

function Get-GFSBilgi {
    param($job)
    try {
        $parts = [System.Collections.Generic.List[string]]::new()

        try {
            $gfs = $job.Options.GfsPolicy
            if ($gfs -ne $null -and $gfs.IsEnabled -eq $true) {
                try {
                    $w = $gfs.Weekly
                    if ($w -ne $null -and $w.IsEnabled -eq $true) {
                        $keep = $w.KeepBackupsForNumberOfWeeks
                        $g    = try { $w.DesiredTime.ToString() } catch { "" }
                        $gun  = TR $gunTR $g
                        $parts.Add("H:$gun(${keep}h)")
                    }
                } catch {}

                try {
                    $m = $gfs.Monthly
                    if ($m -ne $null -and $m.IsEnabled -eq $true) {
                        $keep = $m.KeepBackupsForNumberOfMonths
                        $raw  = try { $m.DesiredTime.ToString() } catch { $null }
                        $d    = if ($raw) { TR $desiredTR $raw } else { $null }
                        $parts.Add($(if ($d) { "A:$d(${keep}a)" } else { "A:(${keep}a)" }))
                    }
                } catch {}

                try {
                    $y = $gfs.Yearly
                    if ($y -ne $null -and $y.IsEnabled -eq $true) {
                        $keep = $y.KeepBackupsForNumberOfYears
                        $raw  = try { $y.DesiredTime.ToString() } catch { $null }
                        $d    = if ($raw) { TR $desiredTR $raw } else { $null }
                        $parts.Add($(if ($d) { "Y:$d(${keep}y)" } else { "Y:(${keep}y)" }))
                    }
                } catch {}

                if ($parts.Count -gt 0) { return ($parts -join "  ") }
                return "GFS Aktif"
            }
        } catch {}

        try {
            $gp = $job.Options.GenerationPolicy
            if ($gp.GFSWeeklyBackupsEnabled  -eq $true) {
                $g = TR $gunTR $gp.WeeklyBackupDayOfWeek.ToString()
                $parts.Add("H:$g($($gp.GFSWeeklyBackups)h)")
            }
            if ($gp.GFSMonthlyBackupsEnabled -eq $true) { $parts.Add("A:($($gp.GFSMonthlyBackups)a)") }
            if ($gp.GFSYearlyBackupsEnabled  -eq $true) { $parts.Add("Y:($($gp.GFSYearlyBackups)y)") }
            if ($parts.Count -gt 0) { return ($parts -join "  ") }
        } catch {}

        return "GFS Yok"
    } catch { return "N/A" }
}

# ─── VERİ TOPLAMA ─────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "Veeam job verileri toplanıyor..." -ForegroundColor Gray

$jobs = Get-VBRJob | Where-Object { $_.JobType -eq "Backup" } | Sort-Object Name

if ($jobs.Count -eq 0) {
    Write-Host "Hic backup job bulunamadi." -ForegroundColor Yellow
    exit 0
}

$rows = foreach ($job in $jobs) {
    [PSCustomObject]@{
        Ad         = $job.Name
        Durum      = if ($job.IsScheduleEnabled -eq $true) { "Aktif" } else { "Pasif" }
        VM         = @($job.GetObjectsInJob()).Count
        FullBackup = Get-FullBackupBilgi -job $job
        GFS        = Get-GFSBilgi        -job $job
    }
}

# ─── TABLO ────────────────────────────────────────────────────────────────────

$w = @{ No = 4; Ad = 42; Durum = 7; VM = 5; Full = 16; GFS = 42 }
$totalW = $w.No + $w.Ad + $w.Durum + $w.VM + $w.Full + $w.GFS + 17

function P { param([string]$s, [int]$n)
    if ($s.Length -gt $n) { $s.Substring(0, $n - 1) + "~" }
    else                   { $s.PadRight($n) }
}

$sep = "-" * $totalW
$hdr = "$(P 'No'          $w.No) | $(P 'Job Adi' $w.Ad) | $(P 'Durum' $w.Durum) | $(P 'VM' $w.VM) | $(P 'Full Backup' $w.Full) | $(P 'GFS' $w.GFS)"

Write-Host ""
Write-Host "=== VEEAM JOB DURUM TABLOSU ===" -ForegroundColor Cyan
Write-Host "Tarih : $(Get-Date -Format 'yyyy-MM-dd HH:mm')  |  Toplam : $($rows.Count) job" -ForegroundColor DarkGray
Write-Host $sep         -ForegroundColor DarkGray
Write-Host $hdr         -ForegroundColor Cyan
Write-Host $sep         -ForegroundColor DarkGray

$idx = 1
foreach ($r in $rows) {
    $fullEksik = ($r.FullBackup -eq "Yok" -or $r.FullBackup -eq "N/A")
    $gfsEksik  = ($r.GFS       -eq "GFS Yok" -or $r.GFS -eq "N/A")

    $color = if    ($fullEksik -and $gfsEksik)  { "Red"    }
             elseif($fullEksik -or  $gfsEksik)  { "Yellow" }
             else                               { "White"  }

    $line = "$(P "$idx" $w.No) | $(P $r.Ad $w.Ad) | $(P $r.Durum $w.Durum) | $(P "$($r.VM)" $w.VM) | $(P $r.FullBackup $w.Full) | $(P $r.GFS $w.GFS)"
    Write-Host $line -ForegroundColor $color
    $idx++
}

Write-Host $sep -ForegroundColor DarkGray

# ─── ÖZET ─────────────────────────────────────────────────────────────────────

$listeTamam   = @($rows | Where-Object { $_.FullBackup -notin @("Yok","N/A") -and $_.GFS -notin @("GFS Yok","N/A") })
$listeEksik   = @($rows | Where-Object { $_.FullBackup -in  @("Yok","N/A") -and $_.GFS -in  @("GFS Yok","N/A") })
$listeKismi   = @($rows | Where-Object {
    ($_.FullBackup -in @("Yok","N/A") -xor $_.GFS -in @("GFS Yok","N/A"))
})

Write-Host ""
Write-Host "OZET" -ForegroundColor Cyan
Write-Host "  Tamam  (Full + GFS var)    : $($listeTamam.Count) job"   -ForegroundColor White
Write-Host "  Kismi  (Biri eksik)        : $($listeKismi.Count) job"   -ForegroundColor Yellow
Write-Host "  Eksik  (Her ikisi de yok)  : $($listeEksik.Count) job"   -ForegroundColor Red
Write-Host ""

# Kirmizi joblar varsa listele
if ($listeEksik.Count -gt 0) {
    Write-Host "DIKKAT — Full Backup ve GFS tanimlanmamis joblar:" -ForegroundColor Red
    $listeEksik | ForEach-Object { Write-Host "  - $($_.Ad)" -ForegroundColor Red }
    Write-Host ""
}

# Sari joblar varsa listele
if ($listeKismi.Count -gt 0) {
    Write-Host "UYARI — Eksik konfigurasyonlu joblar:" -ForegroundColor Yellow
    $listeKismi | ForEach-Object {
        $ne = @()
        if ($_.FullBackup -in @("Yok","N/A")) { $ne += "Full Backup" }
        if ($_.GFS        -in @("GFS Yok","N/A")) { $ne += "GFS" }
        Write-Host "  - $($_.Ad)  [$($ne -join ', ') eksik]" -ForegroundColor Yellow
    }
    Write-Host ""
}
