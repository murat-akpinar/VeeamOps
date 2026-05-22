# Veeam - Backup alinan sanal makina isimlerini listeler

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

$jobs = Get-VBRJob | Where-Object { $_.JobType -eq "Backup" }

if ($jobs.Count -eq 0) {
    Write-Host "Hic backup job bulunamadi." -ForegroundColor Yellow
    exit 0
}

$vmler = $jobs | ForEach-Object { $_.GetObjectsInJob() } |
         Select-Object -ExpandProperty Name |
         Sort-Object -Unique

Write-Host "`n=== VEEAM - BACKUP ALINAN SANAL MAKINALAR ===" -ForegroundColor Cyan
Write-Host "Toplam: $($vmler.Count) VM`n" -ForegroundColor Green

$vmler | ForEach-Object { Write-Host $_ }

Write-Host ""
