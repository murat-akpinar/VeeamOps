# V E E A M  O P S

**Veeam Backup & Replication job analizi ve raporlama için PowerShell scriptleri.**

![Veeam](https://img.shields.io/badge/yap%C4%B1ld%C4%B1-Veeam%20B%26R-00B336?logo=veeam&logoColor=ffffff)
![PowerShell](https://img.shields.io/badge/ile%20cal%C4%B1s%C4%B1r-PowerShell-5391FE?logo=powershell&logoColor=ffffff)
![Windows](https://img.shields.io/badge/platform-Windows%20Server-0078D6?logo=windows&logoColor=ffffff)
![Excel](https://img.shields.io/badge/excel%20export-COM%20gerektirmez-217346?logo=microsoftexcel&logoColor=ffffff)

> English version: [README.md](README.md)

---

## Genel Bakış

Veeam Backup & Replication sunucusuna bağlanarak job yapılandırma detaylarını konsola ve Excel'e aktaran küçük bir PowerShell script koleksiyonu. Temel amaçlar:

- **Full Backup** veya **GFS** yapılandırması eksik olan job'ları hızlıca tespit etmek.
- Tüm job'ları ve korunan VM'leri Excel yüklü olmadan **.xlsx** olarak dışa aktarmak.
- Tek bir job'u derinlemesine incelemek için ham **debug görünümü** sunmak.

Tüm scriptler Veeam Backup Server üzerinde (ya da Veeam yönetim konsolu kurulu bir makinede) çalıştırılmalıdır.

---

## Gereksinimler

| Gereksinim | Notlar |
|---|---|
| Veeam Backup & Replication | v11 veya üzeri önerilir |
| PowerShell | 5.1+ |
| Çalıştırma konumu | Veeam Backup Server veya Veeam konsolu kurulu makine |
| Excel | **Gerekmiyor** — Excel export saf OOXML (ZIP) kullanır |

Scriptler Veeam modülünü şu sırayla yüklemeye çalışır:
1. `VeeamPSSnapIn` (eski sürümler)
2. `Veeam.Backup.PowerShell` modülü (v12+)

---

## Scriptler

### `veeam_vm_list.ps1` — Korunan VM Listesi

En az bir aktif Backup job'unda yer alan benzersiz VM isimlerini listeler.

```powershell
.\veeam_vm_list.ps1
```

**Çıktı:** konsola sıralanmış ve tekilleştirilmiş VM isim listesi.

---

### `veeam_job_status.ps1` — Job Durum Tablosu

Tüm Backup job'larını renkli tablo olarak gösterir. Full Backup ve/veya GFS rotasyonu tanımlanmamış job'ları vurgular.

```powershell
.\veeam_job_status.ps1
```

| Satır rengi | Anlam |
|---|---|
| Beyaz | Full Backup **ve** GFS ikisi de tanımlı |
| Sarı | Birisi eksik |
| Kırmızı | **Her ikisi de** tanımlanmamış |

Alttaki özet bölümü toplam sayıları ve sorunlu job isimlerini listeler; hızlıca kopyalanıp takip edilebilir.

**Kolonlar:** No · Job Adı · Durum · VM Sayısı · Full Backup · GFS

Full Backup kolon formatı: `Synth:Cuma` (synthetic full), `Aktif:Pazartesi` (active full), `Yok` (hiçbiri).  
GFS kolon formatı: `H:Cuma(4h)  A:Son(12a)  Y:Ara(1y)` — H=Haftalık, A=Aylık, Y=Yıllık.

---

### `veeam_backup_list.ps1` — Tam Rapor + Excel Export

Tüm Backup job'larının ve korunan VM'lerin detaylı bilgisini konsola basar; `C:\scripts\` klasörüne `.xlsx` dosyası yazar.

```powershell
.\veeam_backup_list.ps1
```

**Excel kolonları:**

| # | Kolon | Açıklama |
|---|---|---|
| 1 | No | Job sıra numarası |
| 2 | Job Adı / VM Adı | Job başlık satırı + girintili VM satırları |
| 3 | Durum | Aktif / Pasif |
| 4 | Son Çalışma | En son çalışma zaman damgası |
| 5 | VM Sayısı | Job'daki nesne sayısı |
| 6 | Çalışma Günleri | Job'un hangi günler çalıştığı |
| 7 | Başlama Saati | Zamanlanmış başlangıç saati |
| 8 | Retention Policy | Restore point sayısı veya gün sayısı |
| 9 | Full Backup | Synthetic veya Active full backup günü |
| 10 | GFS | GFS Haftalık / Aylık / Yıllık retention detayları |

Excel, saf OOXML (`.xlsx` dosyası içi XML barındıran bir ZIP arşivi) ile üretilir. **COM kullanılmaz, Excel kurulumuna gerek yoktur.**

Excel'deki satır renkleri:
- Koyu mavi başlık satırı
- Lacivert job satırları (Pasif job'lar kırmızı)
- Alternatif açık mavi VM satırları

---

### `veeam_test_job.ps1` — Tek Job Debug Görünümü

Belirli bir job'un erişilebilir tüm property'lerini konsola döker. Belgelenmemiş Veeam API alanlarını keşfetmek için kullanılır.

```powershell
# Önce dosyanın başındaki $TEST_JOB değişkenini düzenleyin:
$TEST_JOB = "Job-Adiniz-Buraya"
.\veeam_test_job.ps1
```

Kapsanan bölümler: temel bilgiler · tüm üst düzey property'ler · schedule seçenekleri · full backup planı · GFS politikası · restore point'ler · retention politikası · reflection (public olmayan .NET alanları) · ham XML dump.

---

## Notlar

- **Veeam iç yazım hatası:** Synthetic full için XML node adı `TransformFullToSyntethic`'tir ("Syntethic"de bir `n` eksik). Bu Veeam'ın kendi içindeki bir yazım hatasıdır; scriptler bunu doğru şekilde işler.
- **GFS veri kaynağı:** GFS ayarları önce `job.Options.GfsPolicy`'den (güncel API), bulunamazsa eski job formatları için `job.Options.GenerationPolicy`'den okunur.
- **Schedule saati:** `ScheduleOptions.StartDateTimeLocal`'den alınır; eski job sürümlerinde `OptionsDaily.TimeLocal`'e geri düşer.
