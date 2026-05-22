# ─── VEEAM JOB TEST SCRIPT ───────────────────────────────────────────────────
# Amac: Ana scripte gitmeden önce tüm property'leri ekranda gör
# Kullanim: script'i çalistirmadan önce $TEST_JOB'u değiştir

$TEST_JOB = "Web-Servers-DijitalYayinlar-2_STO"   # <-- buraya test edilecek job adini yaz

# Veeam modülü
try { Add-PSSnapin -Name VeeamPSSnapIn -ErrorAction Stop } catch {
    try { Import-Module Veeam.Backup.PowerShell -ErrorAction Stop } catch {
        Write-Error "Veeam modülü yüklenemedi."; exit 1
    }
}

$job = Get-VBRJob | Where-Object { $_.Name -eq $TEST_JOB }
if (-not $job) { Write-Error "Job bulunamadı: $TEST_JOB"; exit 1 }

function Sec { param($title) Write-Host "`n$("=" * 60)" -ForegroundColor Cyan; Write-Host "  $title" -ForegroundColor Cyan; Write-Host ("=" * 60) -ForegroundColor Cyan }
function Row { param($k, $v) Write-Host ("  {0,-35} : {1}" -f $k, $v) }
function Try-Get { param($label, [scriptblock]$sb)
    try { $v = & $sb; Row $label $(if ($null -eq $v) {"<null>"} else {$v}) }
    catch { Row $label "<HATA: $_>" }
}

# ─── 1. TEMEL JOB BİLGİLERİ ──────────────────────────────────────────────────
Sec "1. TEMEL JOB BİLGİLERİ"
Row "Name"               $job.Name
Row "JobType"            $job.JobType
Try-Get "IsEnabled"            { $job.IsEnabled }
Try-Get "IsScheduleEnabled"    { $job.IsScheduleEnabled }
Try-Get "IsRunning"            { $job.IsRunning }
Try-Get "LatestRunLocal"       { $job.LatestRunLocal }
Try-Get "NextRun"              { $job.NextRun }
Try-Get "LastResult"           { $job.LastResult }
Try-Get "State"                { $job.State }
Try-Get "JobEnabled (alias?)"  { $job.PSObject.Properties['JobEnabled'].Value }

# ─── 2. TÜM TOP-LEVEL PROPERTY'LER ──────────────────────────────────────────
Sec "2. JOB NESNESININ TÜM PROPERTY'LERI"
$job.PSObject.Properties | Sort-Object Name | ForEach-Object {
    try { Row $_.Name $_.Value } catch { Row $_.Name "<okunamadi>" }
}

# ─── 3. SCHEDULE OPTIONS ─────────────────────────────────────────────────────
Sec "3. SCHEDULE OPTIONS"
$sched = $job.ScheduleOptions
Try-Get "RunManually"                       { $sched.RunManually }
Try-Get "StartDateTime"                     { $sched.StartDateTime }
Try-Get "StartDateTime.Year"               { $sched.StartDateTime.Year }

Try-Get "OptionScheduleAfterJob.IsEnabled" { $sched.OptionScheduleAfterJob.IsEnabled }
Try-Get "OptionPeriodically.IsEnabled"     { $sched.OptionPeriodically.IsEnabled }

Try-Get "OptionScheduleDaily (tip)"        { $sched.OptionScheduleDaily.GetType().Name }
Try-Get "OptionScheduleDaily.DaysSrv"      { $sched.OptionScheduleDaily.DaysSrv -join ", " }
Try-Get "OptionScheduleDaily.StartTimeLocal" { $sched.OptionScheduleDaily.StartTimeLocal }
Try-Get "OptionScheduleDaily.StartDateTime"  { $sched.OptionScheduleDaily.StartDateTime }
Try-Get "OptionScheduleDaily.Kind"           { $sched.OptionScheduleDaily.Kind }
Try-Get "OptionScheduleDaily.Period"         { $sched.OptionScheduleDaily.Period }

Write-Host "`n  --- OptionsDaily tüm property'ler ---" -ForegroundColor Yellow
try { $sched.OptionsDaily.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value } } catch { Write-Host "  <erisim hatasi>" }

Write-Host "`n  --- ScheduleOptions tüm property'ler ---" -ForegroundColor Yellow
$sched.PSObject.Properties | Sort-Object Name | ForEach-Object {
    try { Row $_.Name $_.Value } catch { Row $_.Name "<okunamadi>" }
}

# ─── 4. FULL BACKUP SCHEDULE ─────────────────────────────────────────────────
Sec "4. FULL BACKUP SCHEDULE (Synthetic + Active)"

# BackupStorageOptions.EnableFullBackup — aktif full backup acik mi?
Try-Get "BackupStorageOptions.EnableFullBackup"  { $job.BackupStorageOptions.EnableFullBackup }

Write-Host "`n  --- OptionSyntheticFull tüm property'ler ---" -ForegroundColor Yellow
try { $sched.OptionSyntheticFull.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value } } catch { Write-Host "  <erisim hatasi>" }

Write-Host "`n  --- OptionPeriodicalFull tüm property'ler ---" -ForegroundColor Yellow
try { $sched.OptionPeriodicalFull.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value } } catch { Write-Host "  <erisim hatasi>" }

# ─── 5. GFS (GRANDFATHER-FATHER-SON) RETENSİYON ─────────────────────────────
Sec "5. GFS RETANSIYON (Keep weekly/monthly/yearly)"
Try-Get "BackupStorageOptions.EnableFullBackup"  { $job.BackupStorageOptions.EnableFullBackup }
Try-Get "BackupStorageOptions.RetentionType"     { $job.BackupStorageOptions.RetentionType }

Write-Host "`n  --- BackupStorageOptions tüm property'ler ---" -ForegroundColor Yellow
try { $job.BackupStorageOptions.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value } } catch { Write-Host "  <erisim hatasi>" }

# job.Options (CJobOptions) — GFS buraya gömülü olabilir
Write-Host "`n  --- job.Options tüm property'ler ---" -ForegroundColor Yellow
try { $job.Options.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value } } catch { Write-Host "  <erisim hatasi>" }

# GenerationPolicy içinde GFS olabilir
Write-Host "`n  --- job.Options.GenerationPolicy ---" -ForegroundColor Yellow
try { $job.Options.GenerationPolicy.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value } } catch { Write-Host "  <erisim hatasi>" }

# GfsPolicy — UI'daki Configure GFS diyalogunun verisi burada
Write-Host "`n  --- job.Options.GfsPolicy ---" -ForegroundColor Yellow
try { $job.Options.GfsPolicy.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value } } catch { Write-Host "  <erisim hatasi>" }

Write-Host "`n  --- GfsPolicy.Weekly tüm property'ler ---" -ForegroundColor Yellow
try { $job.Options.GfsPolicy.Weekly.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value } } catch { Write-Host "  <erisim hatasi>" }

Write-Host "`n  --- GfsPolicy.Monthly tüm property'ler ---" -ForegroundColor Yellow
try { $job.Options.GfsPolicy.Monthly.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value } } catch { Write-Host "  <erisim hatasi>" }

Write-Host "`n  --- GfsPolicy.Yearly tüm property'ler ---" -ForegroundColor Yellow
try { $job.Options.GfsPolicy.Yearly.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value } } catch { Write-Host "  <erisim hatasi>" }

# GenerationPolicy alt objeleri (MonthlyBackup, YearlyBackup detayi)
Write-Host "`n  --- GenerationPolicy.MonthlyBackup ---" -ForegroundColor Yellow
try { $job.Options.GenerationPolicy.MonthlyBackup.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value } } catch { Write-Host "  <erisim hatasi>" }

Write-Host "`n  --- GenerationPolicy.YearlyBackup ---" -ForegroundColor Yellow
try { $job.Options.GenerationPolicy.YearlyBackup.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value } } catch { Write-Host "  <erisim hatasi>" }

# Ozet: GFS durumu
Write-Host "`n  --- GFS OZET ---" -ForegroundColor Green
$gp = $job.Options.GenerationPolicy
Row "GFSWeeklyBackupsEnabled"  $gp.GFSWeeklyBackupsEnabled
Row "GFSWeeklyBackups"         $gp.GFSWeeklyBackups
Row "WeeklyBackupDayOfWeek"    $gp.WeeklyBackupDayOfWeek
Row "GFSMonthlyBackupsEnabled" $gp.GFSMonthlyBackupsEnabled
Row "GFSMonthlyBackups"        $gp.GFSMonthlyBackups
Row "GFSYearlyBackupsEnabled"  $gp.GFSYearlyBackupsEnabled
Row "GFSYearlyBackups"         $gp.GFSYearlyBackups
Row "KeepGfsBackup"            $gp.KeepGfsBackup
Row "EnableFullBackup (BSO)"   $job.BackupStorageOptions.EnableFullBackup

# ─── 6. VM LİSTESİ ───────────────────────────────────────────────────────────
Sec "6. JOB'A DAHIL VM'LER"
$objects = $job.GetObjectsInJob()
Row "VM Sayisi" $objects.Count
$objects | Sort-Object Name | ForEach-Object { Write-Host "    - $($_.Name)" -ForegroundColor Gray }

# ─── 7. RESTORE POINTS ───────────────────────────────────────────────────────
Sec "7. RESTORE POINTS (Toplam Backup Sayisi)"
try {
    $backuplar = Get-VBRBackup | Where-Object { $_.JobId -eq $job.Id }
    Row "Backup nesnesi sayisi" (@($backuplar).Count)
    $toplam = 0
    foreach ($bk in @($backuplar)) {
        $pts = Get-VBRRestorePoint -Backup $bk -ErrorAction SilentlyContinue
        $toplam += @($pts).Count
    }
    Row "Toplam restore point" $toplam
} catch { Row "HATA" $_ }

# ─── 8. RETENSİYON POLİSİ ────────────────────────────────────────────────────
Sec "8. RETENSİYON POLİSİ (BackupStorageOptions)"
Try-Get "RetentionType"                   { $job.BackupStorageOptions.RetentionType }
Try-Get "RetainCycles"                    { $job.BackupStorageOptions.RetainCycles }
Try-Get "RetainDays"                      { $job.BackupStorageOptions.RetainDays }

Write-Host "`n  --- GfsPolicy.Monthly ek detaylar ---" -ForegroundColor Yellow
Try-Get "GfsPolicy.Monthly.DesiredTime"   { $job.Options.GfsPolicy.Monthly.DesiredTime }
Try-Get "GfsPolicy.Monthly.DayOfWeek"     { $job.Options.GfsPolicy.Monthly.DayOfWeek }
Try-Get "GfsPolicy.Monthly.WeekOfMonth"   { $job.Options.GfsPolicy.Monthly.WeekOfMonth }

Write-Host "`n  --- GfsPolicy.Yearly ek detaylar ---" -ForegroundColor Yellow
Try-Get "GfsPolicy.Yearly.DesiredTime"    { $job.Options.GfsPolicy.Yearly.DesiredTime }
Try-Get "GfsPolicy.Yearly.MonthOfYear"    { $job.Options.GfsPolicy.Yearly.MonthOfYear }

Write-Host "`n  --- OptionSyntheticFull (Advanced > Create Synthetic full backups periodically on) ---" -ForegroundColor Yellow
Try-Get "OptionSyntheticFull.Enabled"     { $sched.OptionSyntheticFull.Enabled }
Try-Get "OptionSyntheticFull.DaysSrv"     { $sched.OptionSyntheticFull.DaysSrv -join ", " }
Try-Get "OptionSyntheticFull.Days"        { $sched.OptionSyntheticFull.Days -join ", " }
Try-Get "OptionPeriodicalFull.Enabled"    { $sched.OptionPeriodicalFull.Enabled }
Try-Get "OptionPeriodicalFull.DaysSrv"    { $sched.OptionPeriodicalFull.DaysSrv -join ", " }

Write-Host "`n  --- OptionSyntheticFull tum property'ler (tekrar) ---" -ForegroundColor Yellow
try { $sched.OptionSyntheticFull.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value } } catch { Write-Host "  <erisim hatasi>" }


# ─── 9. ANA SCRIPT FORMATI ONIZLEME ─────────────────────────────────────────
Sec "9. ANA SCRIPT FORMATI ONIZLEME"

$_gunTR = @{
    "Sunday"    = "Pazar"
    "Monday"    = "Pazartesi"
    "Tuesday"   = "Sali"
    "Wednesday" = "Carsamba"
    "Thursday"  = "Persembe"
    "Friday"    = "Cuma"
    "Saturday"  = "Cumartesi"
}

$_desiredTR = @{
    "First"     = "Ilk"
    "Second"    = "Ikinci"
    "Third"     = "Ucuncu"
    "Fourth"    = "Dorduncu"
    "Last"      = "Son"
    "January"   = "Ocak"
    "February"  = "Subat"
    "March"     = "Mart"
    "April"     = "Nisan"
    "May"       = "Mayis"
    "June"      = "Haziran"
    "July"      = "Temmuz"
    "August"    = "Agustos"
    "September" = "Eylul"
    "October"   = "Ekim"
    "November"  = "Kasim"
    "December"  = "Aralik"
}

# ── Retention Policy ─────────────────────────────────────────────────────────
$_retPolicy = "N/A"
try {
    $bso     = $job.BackupStorageOptions
    $retType = try { $bso.RetentionType.ToString() } catch { "Cycles" }
    if ($retType -match "Day") {
        $d = try { $bso.RetainDaysToKeep } catch { try { $bso.RetainDays } catch { 0 } }
        $_retPolicy = "$d gun"
    } else {
        $_retPolicy = "$(try{$bso.RetainCycles}catch{0}) rp"
    }
} catch {}

# ── Synthetic Full ve Active Full günü (CDomContainer XML) ───────────────────
# NOT: Veeam XML'de "Syntethic" (yazim hatasi) kullaniliyor — "Synthetic" degil
$_synthEnabled   = $false
$_synthDay       = $null
$_actFullEnabled = $false
$_actFullDay     = $null
try {
    $xmlDoc = $job.Options.Options.Document
    if ($xmlDoc -ne $null) {
        # Synthetic full enabled?
        $n = $xmlDoc.SelectSingleNode("//TransformFullToSyntethic")
        if ($n) { $_synthEnabled = ($n.InnerText -eq "True") }
        Write-Host "  [DEBUG] TransformFullToSyntethic = $(if($n){$n.InnerText}else{'<node yok>'})" -ForegroundColor Magenta

        # Synthetic full günü
        $n = $xmlDoc.SelectSingleNode("//TransformToSyntethicDays/DayOfWeek")
        if ($n -and $n.InnerText) {
            $d = $n.InnerText.Trim()
            $_synthDay = if ($_gunTR.ContainsKey($d)) { $_gunTR[$d] } else { $d }
        }
        Write-Host "  [DEBUG] TransformToSyntethicDays/DayOfWeek = $(if($n){$n.InnerText}else{'<node yok>'}) -> $_synthDay" -ForegroundColor Magenta

        # Active full backup enabled?
        $n = $xmlDoc.SelectSingleNode("//EnableFullBackup")
        if ($n) { $_actFullEnabled = ($n.InnerText -eq "True") }
        Write-Host "  [DEBUG] EnableFullBackup = $(if($n){$n.InnerText}else{'<node yok>'})" -ForegroundColor Magenta

        # Active full günü
        $n = $xmlDoc.SelectSingleNode("//FullBackupDays/DayOfWeek")
        if ($n -and $n.InnerText) {
            $d = $n.InnerText.Trim()
            $_actFullDay = if ($_gunTR.ContainsKey($d)) { $_gunTR[$d] } else { $d }
        }
        Write-Host "  [DEBUG] FullBackupDays/DayOfWeek = $(if($n){$n.InnerText}else{'<node yok>'}) -> $_actFullDay" -ForegroundColor Magenta
    }
} catch { Write-Host "  [DEBUG SynFull] HATA: $_" -ForegroundColor Red }

# ── GFS bilgileri ─────────────────────────────────────────────────────────────
$_parts = [System.Collections.Generic.List[string]]::new()
try {
    $gfs = $job.Options.GfsPolicy
    if ($gfs -ne $null -and $gfs.IsEnabled -eq $true) {

        # Haftalik — gun: GFS Weekly DesiredTime
        try {
            $w = $gfs.Weekly
            Write-Host "  [DEBUG] GfsPolicy.Weekly = $($w | Out-String -Width 200)" -ForegroundColor Magenta
            Write-Host "  [DEBUG] Weekly.IsEnabled = $($w.IsEnabled)" -ForegroundColor Magenta
            if ($w -ne $null -and $w.IsEnabled -eq $true) {
                $keep   = $w.KeepBackupsForNumberOfWeeks
                $g      = try { $w.DesiredTime.ToString() } catch { "" }
                Write-Host "  [DEBUG] Weekly.DesiredTime = '$g'" -ForegroundColor Magenta
                $gunAdi = if ($_gunTR.ContainsKey($g)) { $_gunTR[$g] } else { $g }
                $_parts.Add("Haftalik:$gunAdi ($keep hafta)")
            }
        } catch { Write-Host "  [DEBUG] Weekly HATA: $_" -ForegroundColor Red }

        # Aylik — DesiredTime: First/Last vb.
        try {
            $m = $gfs.Monthly
            if ($m -ne $null -and $m.IsEnabled -eq $true) {
                $keep    = $m.KeepBackupsForNumberOfMonths
                $raw     = try { $m.DesiredTime.ToString() } catch { $null }
                $tr      = if ($raw -and $_desiredTR.ContainsKey($raw)) { $_desiredTR[$raw] } else { $raw }
                if ($tr) { $_parts.Add("Aylik:$tr ($keep ay)") }
                else     { $_parts.Add("Aylik ($keep ay)") }
            }
        } catch {}

        # Yillik — DesiredTime: December vb.
        try {
            $y = $gfs.Yearly
            if ($y -ne $null -and $y.IsEnabled -eq $true) {
                $keep    = $y.KeepBackupsForNumberOfYears
                $raw     = try { $y.DesiredTime.ToString() } catch { $null }
                $tr      = if ($raw -and $_desiredTR.ContainsKey($raw)) { $_desiredTR[$raw] } else { $raw }
                if ($tr) { $_parts.Add("Yillik:$tr ($keep yil)") }
                else     { $_parts.Add("Yillik ($keep yil)") }
            }
        } catch {}
    }
} catch {}

# GenerationPolicy fallback (GfsPolicy yoksa)
if ($_parts.Count -eq 0) {
    try {
        $gp = $job.Options.GenerationPolicy
        if ($gp.GFSWeeklyBackupsEnabled -eq $true) {
            $g      = $gp.WeeklyBackupDayOfWeek.ToString()
            $gunAdi = if ($_gunTR.ContainsKey($g)) { $_gunTR[$g] } else { $g }
            $_parts.Add("Haftalik:$gunAdi ($($gp.GFSWeeklyBackups) hafta)")
        }
        if ($gp.GFSMonthlyBackupsEnabled -eq $true) { $_parts.Add("Aylik ($($gp.GFSMonthlyBackups) ay)") }
        if ($gp.GFSYearlyBackupsEnabled  -eq $true) { $_parts.Add("Yillik ($($gp.GFSYearlyBackups) yil)") }
    } catch {}
}

# ── Schedule bilgileri ────────────────────────────────────────────────────────
$_durum      = if ($job.IsScheduleEnabled -eq $true) { "Aktif" } else { "Pasif" }
$_sonCalisma = if ($job.LatestRunLocal) { $job.LatestRunLocal.ToString("dd.MM.yyyy HH:mm") } else { "-" }

$_calismaGun = "N/A"
try {
    $daily9 = $sched.OptionsDaily
    if ($daily9 -ne $null) {
        $k9 = try { $daily9.Kind.ToString() } catch { "" }
        if     ($k9 -match "Everyday|Every") { $_calismaGun = "Her Gun" }
        elseif ($k9 -match "WeekDay")        { $_calismaGun = "Is Gunleri (Pzt-Cum)" }
        elseif ($k9 -match "Weekend")        { $_calismaGun = "Hafta Sonu (Cmt-Paz)" }
        else {
            $dd9 = @(); try { $dd9 = @($daily9.DaysSrv | Where-Object { $_ -ne $null }) } catch {}
            if   ($dd9.Count -ge 7)   { $_calismaGun = "Her Gun" }
            elseif ($dd9.Count -gt 0) {
                $_calismaGun = ($dd9 | ForEach-Object {
                    $k = $_.ToString().Trim()
                    if ($_gunTR.ContainsKey($k)) { $_gunTR[$k] } else { $k }
                }) -join ", "
            }
        }
    }
} catch {}

$_saat = "N/A"
try {
    $dt9 = $sched.OptionsDaily.TimeLocal
    if ($dt9 -and $dt9 -is [datetime] -and $dt9.Year -gt 1) { $_saat = $dt9.ToString("HH:mm") }
} catch {}
if ($_saat -eq "N/A") {
    try {
        $dt9 = $sched.StartDateTimeLocal
        if ($dt9 -and $dt9 -is [datetime] -and $dt9.Year -gt 1970) { $_saat = $dt9.ToString("HH:mm") }
    } catch {}
}

$_objects  = @($job.GetObjectsInJob() | Sort-Object Name)
$_gfsParts = if ($_parts.Count -gt 0) { $_parts -join " | " } else { "GFS Yok" }

if     ($_synthEnabled   -and $_synthDay)   { $_fullBackupStr = $_synthDay }
elseif ($_actFullEnabled -and $_actFullDay) { $_fullBackupStr = $_actFullDay }
elseif ($_synthDay)                         { $_fullBackupStr = "$_synthDay [KAPALI]" }
elseif ($_actFullDay)                       { $_fullBackupStr = "$_actFullDay [KAPALI]" }
else                                        { $_fullBackupStr = "Yok" }

# ─── 10. REFLECTION — SYNTHETIC FULL GIZLI PROPERTY'LER ─────────────────────
Sec "10. REFLECTION — OptionSyntheticFull / OptionPeriodicalFull"

Write-Host "`n  --- OptionSyntheticFull .NET GetType + reflection ---" -ForegroundColor Yellow
try {
    $sf2 = $sched.OptionSyntheticFull
    if ($null -eq $sf2) { Write-Host "  <OptionSyntheticFull NULL>" }
    else {
        Write-Host "  Tip: $($sf2.GetType().FullName)"
        $sf2.GetType().GetProperties() | Sort-Object Name | ForEach-Object {
            try { Row $_.Name ($_.GetValue($sf2, $null)) } catch { Row $_.Name "<hata>" }
        }
    }
} catch { Write-Host "  <reflection hatasi: $_>" }

Write-Host "`n  --- OptionPeriodicalFull .NET reflection ---" -ForegroundColor Yellow
try {
    $pf2 = $sched.OptionPeriodicalFull
    if ($null -eq $pf2) { Write-Host "  <OptionPeriodicalFull NULL>" }
    else {
        Write-Host "  Tip: $($pf2.GetType().FullName)"
        $pf2.GetType().GetProperties() | Sort-Object Name | ForEach-Object {
            try { Row $_.Name ($_.GetValue($pf2, $null)) } catch { Row $_.Name "<hata>" }
        }
    }
} catch { Write-Host "  <reflection hatasi: $_>" }

Write-Host "`n  --- ScheduleOptions .NET reflection (tum public property) ---" -ForegroundColor Yellow
try {
    $sched.GetType().GetProperties() | Sort-Object Name | ForEach-Object {
        try { Row $_.Name ($_.GetValue($sched, $null)) } catch { Row $_.Name "<hata>" }
    }
} catch { Write-Host "  <ScheduleOptions reflection hatasi: $_>" }

Write-Host "`n  --- GenerationPolicy.CompactFullBackupDays (join) ---" -ForegroundColor Yellow
Try-Get "CompactFullBackupDays"  { $job.Options.GenerationPolicy.CompactFullBackupDays -join ", " }
Try-Get "EnableCompactFull"      { $job.Options.GenerationPolicy.EnableCompactFull }

# ─── 11. SYNTHETIC FULL GUN ARAŞTIRMASI ──────────────────────────────────────
Sec "11. SYNTHETIC FULL GUN — ALTERNATIF KAYNAKLAR"

Write-Host "`n  --- BackupStorageOptions .NET reflection (gizli property'ler) ---" -ForegroundColor Yellow
try {
    $bso2 = $job.BackupStorageOptions
    $bso2.GetType().GetProperties() | Sort-Object Name | ForEach-Object {
        try { Row $_.Name ($_.GetValue($bso2, $null)) } catch { Row $_.Name "<hata>" }
    }
} catch { Write-Host "  <BSO reflection hatasi: $_>" }

Write-Host "`n  --- job.Options.JobOptions .NET reflection ---" -ForegroundColor Yellow
try {
    $jo = $job.Options.JobOptions
    if ($null -eq $jo) { Write-Host "  <JobOptions NULL>" }
    else {
        $jo.GetType().GetProperties() | Sort-Object Name | ForEach-Object {
            try { Row $_.Name ($_.GetValue($jo, $null)) } catch { Row $_.Name "<hata>" }
        }
    }
} catch { Write-Host "  <JobOptions reflection hatasi: $_>" }

Write-Host "`n  --- Get-VBRJobScheduleOptions ---" -ForegroundColor Yellow
try {
    $vbrSched = Get-VBRJobScheduleOptions -Job $job
    if ($null -eq $vbrSched) { Write-Host "  <NULL dondu>" }
    else {
        $vbrSched.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value }
    }
} catch { Write-Host "  <cmdlet hatasi: $_>" }

Write-Host "`n  --- ScheduleOptions NonPublic fields (reflection) ---" -ForegroundColor Yellow
try {
    $flags = [System.Reflection.BindingFlags]'NonPublic,Instance'
    $sched.GetType().GetFields($flags) | Sort-Object Name | ForEach-Object {
        try { Row $_.Name ($_.GetValue($sched)) } catch { Row $_.Name "<hata>" }
    }
} catch { Write-Host "  <nonpublic reflection hatasi: $_>" }

Write-Host "`n  --- job.Options.Options (CDomContainer) PSObject.Properties ---" -ForegroundColor Yellow
try {
    $opts = $job.Options.Options
    if ($null -eq $opts) { Write-Host "  <Options NULL>" }
    else {
        $opts.PSObject.Properties | Sort-Object Name | ForEach-Object { Row $_.Name $_.Value }
    }
} catch { Write-Host "  <Options hatasi: $_>" }

# ─── 12. SYNTHETIC FULL GUN — KAPSAMLI ARAMA ────────────────────────────────
Sec "12. SYNTHETIC FULL GUN — KAPSAMLI ARAMA"

# ── A) GenerationPolicy gizli .NET property'leri ─────────────────────────────
Write-Host "`n  --- GenerationPolicy .NET GetProperties() (tum public) ---" -ForegroundColor Yellow
try {
    $gp12 = $job.Options.GenerationPolicy
    $gp12.GetType().GetProperties() | Sort-Object Name | ForEach-Object {
        try { Row $_.Name ($_.GetValue($gp12, $null)) } catch { Row $_.Name "<hata>" }
    }
} catch { Write-Host "  <hata: $_>" }

# ── B) GenerationPolicy NonPublic field'lar ───────────────────────────────────
Write-Host "`n  --- GenerationPolicy NonPublic fields ---" -ForegroundColor Yellow
try {
    $gp12 = $job.Options.GenerationPolicy
    $flags = [System.Reflection.BindingFlags]'NonPublic,Instance'
    $gp12.GetType().GetFields($flags) | Sort-Object Name | ForEach-Object {
        try { Row $_.Name ($_.GetValue($gp12)) } catch { Row $_.Name "<hata>" }
    }
} catch { Write-Host "  <hata: $_>" }

# ── C) XML'de gun adlari — Friday / Cuma vb. ─────────────────────────────────
Write-Host "`n  --- XML'de gun adlari araması ---" -ForegroundColor Yellow
try {
    $xmlDoc12 = $job.Options.Options.Document
    $allXml12 = $xmlDoc12.OuterXml
    Write-Host "  XML uzunlugu: $($allXml12.Length) karakter" -ForegroundColor Cyan

    $gunler12 = @("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
    foreach ($g12 in $gunler12) {
        $ms = [regex]::Matches($allXml12, ".{0,80}$g12.{0,80}", 'IgnoreCase')
        if ($ms.Count -gt 0) {
            Write-Host "  '$g12' BULUNDU ($($ms.Count) yer):" -ForegroundColor Green
            foreach ($m12 in $ms) { Write-Host "    ...$(($m12.Value).Trim())..." -ForegroundColor Green }
        }
    }
} catch { Write-Host "  <hata: $_>" }

# ── D) XML tam döküm ──────────────────────────────────────────────────────────
Write-Host "`n  --- XML tam dump ---" -ForegroundColor Yellow
try {
    $xmlDoc12b = $job.Options.Options.Document
    $xmlStr12  = $xmlDoc12b.OuterXml
    # Formatted (her attribute'u yeni satira al) — okunabilirlik icin
    $xw = [System.Xml.XmlWriterSettings]::new()
    $xw.Indent = $true; $xw.IndentChars = "  "
    $sb12 = [System.Text.StringBuilder]::new()
    $xwr  = [System.Xml.XmlWriter]::Create($sb12, $xw)
    $xmlDoc12b.Save($xwr); $xwr.Close()
    $formatted = $sb12.ToString()
    Write-Host $formatted.Substring(0, [Math]::Min($formatted.Length, 8000))
    if ($formatted.Length -gt 8000) { Write-Host "  ... ($(($formatted.Length - 8000)) karakter kesildi)" -ForegroundColor DarkYellow }
} catch { Write-Host "  <hata: $_>" }

# ─── ANA SCRIPT CIKTISI (her zaman en sonda) ─────────────────────────────────
Write-Host "`n$("=" * 60)" -ForegroundColor Green
Write-Host "  ANA SCRIPT CIKTISI (veeam_backup_list.ps1 formati)" -ForegroundColor Green
Write-Host ("=" * 60) -ForegroundColor Green
Write-Host ""
Write-Host "JOB: $($job.Name)" -ForegroundColor Yellow
Write-Host "  Durum              : $_durum"
Write-Host "  Son Calisma        : $_sonCalisma"
Write-Host "  Calisma Gunleri    : $_calismaGun"
Write-Host "  Baslama Saati      : $_saat"
Write-Host "  Retention Policy   : $_retPolicy"
Write-Host "  Full Backup        : $_fullBackupStr"
Write-Host "  GFS                : $_gfsParts"
Write-Host "  VM Sayisi          : $($_objects.Count)"
foreach ($_obj in $_objects) { Write-Host "    - $($_obj.Name)" -ForegroundColor Gray }
Write-Host ""
Write-Host "=== TEST TAMAMLANDI ===" -ForegroundColor Green
