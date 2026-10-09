# SQLBackupMonitor (PowerShell Edition)

> **Enterprise SQL Server Backup SLA Monitoring & Automated PDF/Email Reporting Platform**  
> **100% Pure PowerShell &bull; Zero External Executables (`.exe`) &bull; Zero Third-Party DLLs &bull; DLP & Enterprise Policy Compliant**

---

## 📌 Overview

**`SQLBackupMonitor.ps1`** is a standalone, single-file PowerShell tool designed for Database Administrators (DBAs) to centrally monitor, evaluate, audit, and alert on Microsoft SQL Server backup schedules across dozens or hundreds of database instances.

Because it runs **100% natively in Windows PowerShell**, it requires **no `.exe` installations, no administrator privileges, no external DLLs, and no internet access**, making it ideal for restricted corporate and air-gapped environments.

---

## ✨ Key Features

- 🖥️ **Embedded Glassmorphism Web Dashboard**: Serves a full dark-mode responsive web interface via native .NET `HttpListener` (`http://localhost:5000` or network IP).
- 🗂️ **Server-First Accordion Hierarchy**: Dashboard groups databases by Server Instance with instant SLA breach indicators and one-click drill-down to database details.
- ⚡ **Multi-Instance Parallel SQL Polling**: Non-blocking telemetry collection querying `msdb.dbo.backupset` and `sys.databases`.
- 🛡️ **3-Tier Hierarchical SLA Policy Engine**: Global baseline SLA, Group-wise SLA (e.g., *Production*, *DR*, *Staging*), and Server-level custom SLA overrides.
- 🔄 **AlwaysOn Availability Groups (HADR)**: Discovers replica roles (Primary/Secondary), sync health, and backup preferences.
- 📄 **Native Executive PDF Reports**: Pure PowerShell landscape PDF generator with compliance scores, KPI summaries, and zero text overflow.
- 📧 **Automated SMTP SLA Alerting**: Sends HTML email notifications with attached PDF reports upon SLA breaches, plus daily scheduled executive summaries.
- 💾 **Historical Repository Database**: Optional centralized SQL Server repository for historical trend analytics and compliance auditing.
- ⚙️ **Persistent Inventory & Configuration**: Saves all servers, SLA thresholds, and SMTP credentials in a local [`config.json`](file:///d:/Backup_Monitoring/config.json).

---

## 🚀 Quick Start

### Option 1: One-Click Launcher
Double-click [`Run_SQLBackupMonitor.bat`](file:///d:/Backup_Monitoring/Run_SQLBackupMonitor.bat) in Windows Explorer.

### Option 2: Launch via PowerShell
Open PowerShell and run:
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\SQLBackupMonitor.ps1
```
*The web server will start on port `5000` and automatically open your default browser to `http://localhost:5000`.*

---

## 🕹️ Operational Modes

The script supports dedicated operational modes via the `-Mode` parameter:

| Mode | Command | Description |
|---|---|---|
| **Web** *(Default)* | `.\SQLBackupMonitor.ps1 -Mode Web` | Starts the embedded HTTP dashboard and REST API. |
| **Scan** | `.\SQLBackupMonitor.ps1 -Mode Scan` | Performs an instant telemetry scan and outputs a summary table to the console. |
| **PdfReport** | `.\SQLBackupMonitor.ps1 -Mode PdfReport` | Generates a landscape executive PDF audit report (`BackupReport.pdf`). |
| **Report** | `.\SQLBackupMonitor.ps1 -Mode Report` | Generates a standalone offline HTML audit report (`BackupReport.html`). |
| **Loop** | `.\SQLBackupMonitor.ps1 -Mode Loop -IntervalSeconds 60` | Runs a continuous background polling loop with automated SMTP SLA breach dispatch. |
| **EmailReport** | `.\SQLBackupMonitor.ps1 -Mode EmailReport` | Forces immediate generation and SMTP delivery of the Daily Executive PDF Report. |

---

## 📋 Command-Line Parameter Reference

```powershell
param(
    [string]$Mode = "Web",           # Web | Scan | Report | PdfReport | Loop | EmailReport
    [int]$Port = 5000,               # HTTP server port for Web Dashboard
    [string]$ConfigFile = "",        # Path to config.json (defaults to script directory)
    [switch]$NoBrowser,              # Prevents automatic browser launch on startup
    [string]$ReportPath = "",        # Custom output file path for PDF/HTML reports
    [int]$IntervalSeconds = 60       # Polling frequency in seconds for Loop mode
)
```

---

## ⚙️ Configuration File (`config.json`)

All server inventories, SLA policy thresholds, SMTP mail settings, and central repository database settings are stored in [`config.json`](file:///d:/Backup_Monitoring/config.json).

### Do I have to re-enter SMTP settings every time?
**No.** Once configured in [`config.json`](file:///d:/Backup_Monitoring/config.json) or via the **Policies & SMTP** tab in the Web UI, all settings persist permanently and are loaded automatically on every run.

### Moving to Another Server:
To migrate to another machine, copy both [`SQLBackupMonitor.ps1`](file:///d:/Backup_Monitoring/SQLBackupMonitor.ps1) and [`config.json`](file:///d:/Backup_Monitoring/config.json) into the target folder. All registered servers and alert settings will immediately be available.

### Sample `config.json`:
```json
{
  "GlobalPolicies": {
    "FullBackupWarningHours": 20,
    "FullBackupCriticalHours": 24,
    "DiffBackupWarningHours": 10,
    "DiffBackupCriticalHours": 12,
    "LogBackupWarningMinutes": 20,
    "LogBackupCriticalMinutes": 30,
    "AlertOnMissingBackups": true,
    "IgnoreSimpleRecoveryLogs": true,
    "AutoRefreshIntervalSec": 60
  },
  "GroupPolicies": [
    {
      "GroupName": "Production",
      "FullBackupWarningHours": 20,
      "FullBackupCriticalHours": 24,
      "DiffBackupWarningHours": 10,
      "DiffBackupCriticalHours": 12,
      "LogBackupWarningMinutes": 20,
      "LogBackupCriticalMinutes": 30,
      "IgnoreSimpleRecoveryLogs": true
    }
  ],
  "EmailSettings": {
    "IsEnabled": true,
    "SmtpServer": "smtp.office365.com",
    "SmtpPort": 587,
    "EnableSsl": true,
    "SenderEmail": "dba-alerts@yourcompany.com",
    "SenderDisplayName": "SQLBackupMonitor Alerts",
    "SmtpUsername": "dba-alerts@yourcompany.com",
    "SmtpPassword": "YourAppPasswordHere",
    "RecipientEmails": "dba-team@yourcompany.com",
    "AlertOnCritical": true,
    "AlertOnWarning": false,
    "AttachPdfReport": true,
    "DailyReportEnabled": true,
    "DailyReportTime": "08:00",
    "CooldownMinutes": 60
  },
  "Servers": [
    {
      "Id": "a1b2c3d4-e5f6-7890-abcd-1234567890ab",
      "Name": "Primary Production Cluster",
      "ServerAddress": "SQLPROD01.corp.local",
      "Port": 1433,
      "AuthType": "Windows",
      "GroupName": "Production",
      "IsEnabled": true
    }
  ]
}
```

---

## 🔒 SQL Server Permissions Required

`SQLBackupMonitor.ps1` requires only lightweight read permissions on target SQL Server instances:

```sql
-- Minimum required permissions:
GRANT VIEW SERVER STATE TO [Domain\YourServiceAccount];
GRANT VIEW ANY DATABASE TO [Domain\YourServiceAccount];
USE msdb;
GRANT SELECT ON msdb.dbo.backupset TO [Domain\YourServiceAccount];
GRANT SELECT ON msdb.dbo.backupmediafamily TO [Domain\YourServiceAccount];
```

---

## ⏰ Automating with Windows Task Scheduler

You can schedule automated runs without third-party services:

### 1. Continuous Background Service Loop:
```cmd
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "D:\Backup_Monitoring\SQLBackupMonitor.ps1" -Mode Loop -IntervalSeconds 60
```

### 2. Daily 08:00 AM Executive PDF Email Dispatch:
```cmd
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "D:\Backup_Monitoring\SQLBackupMonitor.ps1" -Mode EmailReport
```

---

## 📄 File Inventory

- [`SQLBackupMonitor.ps1`](file:///d:/Backup_Monitoring/SQLBackupMonitor.ps1) — Complete standalone PowerShell platform script.
- [`Run_SQLBackupMonitor.bat`](file:///d:/Backup_Monitoring/Run_SQLBackupMonitor.bat) — One-click batch launcher.
- [`config.json`](file:///d:/Backup_Monitoring/config.json) — Local configuration and server inventory file.
- [`POWERSHELL_USER_GUIDE.md`](file:///d:/Backup_Monitoring/POWERSHELL_USER_GUIDE.md) — Comprehensive technical documentation and administrator guide.
