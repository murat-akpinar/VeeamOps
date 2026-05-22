# V E E A M  O P S

**PowerShell scripts for Veeam Backup & Replication job analysis and reporting.**

![Veeam](https://img.shields.io/badge/made%20for-Veeam%20B%26R-00B336?logo=veeam&logoColor=ffffff)
![PowerShell](https://img.shields.io/badge/powered%20by-PowerShell-5391FE?logo=powershell&logoColor=ffffff)
![Windows](https://img.shields.io/badge/runs%20on-Windows%20Server-0078D6?logo=windows&logoColor=ffffff)
![Excel](https://img.shields.io/badge/excel%20export-COM%20free-217346?logo=microsoftexcel&logoColor=ffffff)

> Turkish version: [README_TR.md](README_TR.md)

---

## Overview

A small collection of PowerShell scripts that connect to the local Veeam Backup & Replication server and surface job configuration details that are otherwise buried inside the Veeam console. The main goals are:

- Quickly identify jobs that are **missing Full Backup** or **GFS** configuration.
- Generate an **Excel report** of all jobs and their protected VMs without needing Excel installed.
- Provide a raw **debug view** for deep-diving into a single job's internal properties.

All scripts must be run on the Veeam Backup Server (or a machine with the Veeam console installed).

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Veeam Backup & Replication | v11 or later recommended |
| PowerShell | 5.1+ |
| Run location | Veeam Backup Server or a machine with the Veeam management console |
| Excel | **Not required** — Excel export uses pure OOXML (ZIP) |

The scripts attempt to load the Veeam snap-in in this order:
1. `VeeamPSSnapIn` (older versions)
2. `Veeam.Backup.PowerShell` module (v12+)

---

## Scripts

### `veeam_vm_list.ps1` — Protected VM List

Lists every unique VM name that appears in at least one active Backup job.

```powershell
.\veeam_vm_list.ps1
```

**Output:** sorted, deduplicated list of VM names printed to the console.

---

### `veeam_job_status.ps1` — Job Status Table

Displays all Backup jobs as a color-coded table. Highlights jobs where Full Backup and/or GFS rotation is not configured.

```powershell
.\veeam_job_status.ps1
```

| Row color | Meaning |
|---|---|
| White | Full Backup **and** GFS both configured |
| Yellow | One of them is missing |
| Red | **Neither** Full Backup nor GFS is configured |

A summary at the bottom lists totals and names the problematic jobs for quick copy-paste follow-up.

**Columns:** No · Job Name · Status · VM Count · Full Backup · GFS

Full Backup column format: `Synth:Friday` (synthetic full), `Aktif:Monday` (active full), `Yok` (none).  
GFS column format: `H:Friday(4w)  A:Last(12m)  Y:Dec(1y)` — H=Weekly, A=Monthly, Y=Yearly.

---

### `veeam_backup_list.ps1` — Full Report + Excel Export

Collects detailed information for every Backup job and all its protected VMs, prints it to the console, and writes a `.xlsx` file to `C:\scripts\`.

```powershell
.\veeam_backup_list.ps1
```

**Excel columns:**

| # | Column | Description |
|---|---|---|
| 1 | No | Job index |
| 2 | Job Name / VM Name | Job header row + indented VM rows |
| 3 | Status | Aktif / Pasif |
| 4 | Last Run | Most recent run timestamp |
| 5 | VM Count | Number of objects in the job |
| 6 | Schedule Days | Which days the job runs |
| 7 | Start Time | Scheduled start time |
| 8 | Retention Policy | Restore points or days |
| 9 | Full Backup | Synthetic or Active full backup day |
| 10 | GFS | GFS Weekly / Monthly / Yearly retention details |

Excel is generated using pure OOXML (the `.xlsx` file is a ZIP with XML inside). **No COM / no Excel installation required.**

Row colors in Excel:
- Dark blue header row
- Navy job rows (red for Disabled/Pasif jobs)
- Alternating light blue VM rows

---

### `veeam_test_job.ps1` — Single Job Debug View

Dumps every accessible property of a single job to the console. Useful for exploring undocumented Veeam API fields.

```powershell
# Edit the $TEST_JOB variable at the top of the file first:
$TEST_JOB = "Your-Job-Name-Here"
.\veeam_test_job.ps1
```

Sections covered: basic info · all top-level properties · schedule options · full backup schedule · GFS policy · restore points · retention policy · reflection (non-public .NET fields) · raw XML dump.

---

## Notes

- **Typo in Veeam internals:** The XML node name for synthetic full is `TransformFullToSyntethic` (one `n` in "Syntethic"). This is a typo inside Veeam itself; the scripts handle it correctly.
- **GFS data source:** GFS settings are read from `job.Options.GfsPolicy` (current API) with a fallback to `job.Options.GenerationPolicy` for older job formats.
- **Schedule time:** Retrieved from `ScheduleOptions.StartDateTimeLocal`; falls back to `OptionsDaily.TimeLocal` on older job versions.
