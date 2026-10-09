# SQLBackupMonitor - PowerShell Edition User Guide

> **Enterprise SQL Server Backup Monitoring & SLA Platform (Single-File Standalone Edition)**  
> **100% Native PowerShell &bull; Zero External Executables (`.exe`) Required &bull; Enterprise DLP / Security Compliant**

---

## 📑 Table of Contents
1. [Overview](#1-overview)
2. [Key Advantages in Corporate Environments](#2-key-advantages-in-corporate-environments)
3. [Prerequisites & System Requirements](#3-prerequisites--system-requirements)
4. [Quick Start](#4-quick-start)
5. [Operational Modes](#5-operational-modes)
   - [Mode 1: Self-Hosted Web Dashboard (Default)](#mode-1-self-hosted-web-dashboard-default)
   - [Mode 2: Terminal Console Scan](#mode-2-terminal-console-scan)
   - [Mode 3: Offline HTML Audit Report](#mode-3-offline-html-audit-report)
   - [Mode 4: Standalone Native PDF Report](#mode-4-standalone-native-pdf-report)
   - [Mode 5: Continuous Monitoring Service Loop](#mode-5-continuous-monitoring-service-loop)
   - [Mode 6: Daily Executive PDF Report Email Delivery](#mode-6-daily-executive-pdf-report-email-delivery)
6. [Command-Line Parameters Reference](#6-command-line-parameters-reference)
7. [Hierarchical SLA Policy Engine (Group-Wise & Server-Wise SLA)](#7-hierarchical-sla-policy-engine-group-wise--server-wise-sla)
8. [Configuration & Inventory Management (`config.json`)](#8-configuration--inventory-management-configjson)
9. [Automated Email & Daily Scheduled PDF Reports]bel(#9-automated-email--daily-scheduled-pdf-reports)
10. [SQL Server Permissions Required](#10-sql-server-permissions-required)
11. [Automating with Windows Task Scheduler](#11-automating-with-windows-task-scheduler)
12. [Troubleshooting & FAQ](#12-troubleshooting--faq)

---

## 1. Overview

[`SQLBackupMonitor.ps1`](file:///d:/Backup_Monitoring/SQLBackupMonitor.ps1) is an all-in-one, enterprise-grade PowerShell tool designed to monitor, audit, and evaluate backup health across Microsoft SQL Server instances across your network without needing compiled `.exe` files, third-party DLLs, or administrative software installations.

It incorporates:
- An **embedded HTTP server & REST API** powered by native .NET `System.Net.HttpListener`.
- An **embedded dark-mode glassmorphism web dashboard** rendered entirely in the browser.
- A **high-performance ADO.NET SQL query engine** using native `System.Data.SqlClient`.
- A **granular SLA evaluation engine** with Recovery Model awareness (`SIMPLE` vs `FULL`).
- An **interactive offline HTML report generator**.

---

## 2. Key Advantages in Corporate Environments

| Challenge in Corporate / Office Networks | How `SQLBackupMonitor.ps1` Solves It |
|---|---|
| **DLP / USB / Email blocks on `.exe` files** | Runs purely as a text-based `.ps1` script; no `.exe` binaries to copy. |
| **Endpoint Privilege Restrictions** | Runs entirely in standard user space without requiring local administrator rights. |
| **Missing .NET Runtimes / SDKs** | Uses built-in Windows PowerShell 5.1 / 7+ and standard Windows ADO.NET. |
| **Air-gapped / Offline Networks** | Self-contained UI and script; functions 100% offline without internet access. |

---

## 3. Prerequisites & System Requirements

- **Operating System**: Windows 10, Windows 11, Windows Server 2012 R2 / 2016 / 2019 / 2022 / 2025.
- **PowerShell Version**: Windows PowerShell 5.1 (built into Windows) or PowerShell 7+.
- **Network / SQL Access**: TCP connectivity to SQL Server instances (default port `1433` or custom ports).

---

## 4. Quick Start

### Method A: One-Click Batch Launcher
Double-click [`Run_SQLBackupMonitor.bat`](file:///d:/Backup_Monitoring/Run_SQLBackupMonitor.bat) in Windows Explorer.

### Method B: Launch from PowerShell Console
Open PowerShell and run:
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\SQLBackupMonitor.ps1
```
*The script automatically initiates a telemetry scan, starts the local web server on port `5000`, and opens your default browser to `http://localhost:5000`.*

---

## 5. Operational Modes

`SQLBackupMonitor.ps1` supports 4 dedicated execution modes via the `-Mode` parameter:

```
                          ┌────────────────────────┐
                          │  SQLBackupMonitor.ps1  │
                          └───────────┬────────────┘
                                      │
        ┌───────────────┬─────────────┴──────────────┬───────────────┐
        ▼               ▼                            ▼               ▼
   [-Mode Web]    [-Mode Scan]                 [-Mode Report]  [-Mode Loop]
  Web Dashboard   Console Table               Stand-Alone HTML Continuous Loop
```

---

### Mode 1: Self-Hosted Web Dashboard (Default)

Starts the background HTTP server and serves the full interactive dashboard.

```powershell
# Default launch on port 5000:
.\SQLBackupMonitor.ps1 -Mode Web

# Custom port without automatically opening browser:
.\SQLBackupMonitor.ps1 -Mode Web -Port 8080 -NoBrowser
```

#### Dashboard Features:
1. **📊 Real-Time KPI Cards**: Total Instances, Online Databases, SLA Compliance %, Critical Alerts, Total Backup Volume.
2. **⏱️ Programmable Auto-Pull Telemetry**: Configure exact background polling & UI refresh interval (in seconds) directly from the Web UI with live badge indicator.
3. **🗄️ Database Backup Matrix**: Search and filter by Server, Database Name, Recovery Model, or SLA Status (`Healthy`, `Warning`, `Critical`) with 10-row pagination.
4. **🖥️ Instance Manager**: Add, edit, remove, and live-test SQL Server connections with latency indicators.
5. **⚙️ SLA Policy Configuration**: Fine-tune warning and critical hour/minute thresholds for Full, Differential, and Transaction Log backups.
6. **📥 Export**: One-click download of self-contained PDF and HTML audit reports.

---

### Mode 2: Terminal Console Scan

Executes a single scan across all configured SQL instances and prints a formatted, color-coded health summary directly in the command prompt.

```powershell
.\SQLBackupMonitor.ps1 -Mode Scan
```

#### Sample Output:
```text
Running single SQLBackupMonitor scan across configured instances...

--- SQL BACKUP HEALTH & SLA AUDIT ---

Status   ServerName       DatabaseName       RecoveryModel LastFullBackup      LastLogBackup       LastBackupSizeGB StatusReason
------   ----------       ------------       ------------- --------------      -------------       ---------------- ------------
Healthy  SQL-PROD-01      AppDatabase        FULL          2026-10-05 02:00:00 2026-10-05 10:30:00 45.120           All backups within SLA thresholds
Warning  SQL-PROD-01      SalesArchive       FULL          2026-10-04 06:00:00 2026-10-05 09:45:00 120.450          Full backup is 28.5h old (Warn: >20h)
Critical SQL-PROD-02      PaymentGateway     FULL          2026-10-05 01:00:00 Never               8.200            No Log Backup found (Recovery: FULL)
Healthy  SQL-PROD-02      Reporting_Staging  SIMPLE        2026-10-05 03:30:00 N/A (Simple)        15.800           All backups within SLA thresholds

Summary:
Total Databases: 4 | Healthy: 2 | Warning: 1 | Critical: 1 | Compliance: 50.0%
```

---

### Mode 3: Offline HTML Audit Report

Scans all instances and exports an offline, beautifully styled HTML document (`BackupReport.html`) with KPI statistics and database audit rows.

```powershell
# Generate report and automatically open it:
.\SQLBackupMonitor.ps1 -Mode Report

# Specify custom output path without opening browser:
.\SQLBackupMonitor.ps1 -Mode Report -ReportPath "C:\Reports\Daily_Backup_Audit.html" -NoBrowser
```

---

### Mode 4: Native PDF Audit Report Generation

Generates a standalone, pixel-perfect executive PDF report (`BackupReport.pdf`) directly using 100% native PowerShell without needing external converters or `.exe` dependencies:

```powershell
# Generate PDF and open automatically:
.\SQLBackupMonitor.ps1 -Mode PdfReport

# Custom destination:
.\SQLBackupMonitor.ps1 -Mode PdfReport -ReportPath "C:\AuditReports\Executive_Backup_SLA.pdf" -NoBrowser
```

---

### Mode 5: Continuous Monitoring Service Loop

Runs continuously in the terminal, polling SQL Server targets at a specified interval, checking daily scheduled email delivery, and writing real-time timestamped health summaries.

```powershell
# Poll every 60 seconds (default):
.\SQLBackupMonitor.ps1 -Mode Loop

# Poll every 5 minutes (300 seconds):
.\SQLBackupMonitor.ps1 -Mode Loop -IntervalSeconds 300
```

---

### Mode 6: Daily Executive PDF Report Email Delivery

Runs an on-demand or automated scan, compiles the executive PDF audit report, and sends an HTML executive summary email with the PDF attached directly to the configured recipient email list:

```powershell
# Send scheduled daily PDF audit report email immediately:
.\SQLBackupMonitor.ps1 -Mode EmailReport
```

---

## 6. Command-Line Parameters Reference

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-Mode` | `string` | `"Web"` | Operational mode: `Web`, `Scan`, `Report`, `PdfReport`, `Loop`, or `EmailReport`. |
| `-Port` | `int` | `5000` | TCP port for the web dashboard (`-Mode Web`). |
| `-ConfigFile` | `string` | `config.json` | Path to the JSON configuration and server inventory file. |
| `-NoBrowser` | `switch` | `$false` | Suppresses automatic launching of the web browser. |
| `-ReportPath` | `string` | `BackupReport.html` or `BackupReport.pdf` | Destination path for exported reports. |
| `-IntervalSeconds`| `int` | `60` | Polling frequency in seconds for `-Mode Loop`. |

---

## 7. Hierarchical SLA Policy Engine (Group-Wise & Server-Wise SLA)

`SQLBackupMonitor.ps1` supports a 3-tier hierarchical SLA evaluation engine allowing DBAs to define tailored Recovery Point Objective (RPO) rules across diverse server environments:

```
                    ┌─────────────────────────────────────────┐
                    │    1. Server-Wise Custom SLA Override   │
                    │   (Enabled directly on target instance) │
                    └────────────────────┬────────────────────┘
                                         │ (If disabled)
                                         ▼
                    ┌─────────────────────────────────────────┐
                    │    2. Group-Wise SLA Policy             │
                    │   (Matches server Group: e.g. Prod, DR) │
                    └────────────────────┬────────────────────┘
                                         │ (If unmatched)
                                         ▼
                    ┌─────────────────────────────────────────┐
                    │    3. Global Default SLA Policy         │
                    │   (Baseline fallback for all instances) │
                    └─────────────────────────────────────────┘
```

### SLA Resolution Logic:
1. **Server Override (`UseCustomSla = true`)**: If enabled on an instance, the system uses the instance's custom Warning and Critical hours/minutes, overriding any group or global policy.
2. **Group-Level Policy (`GroupPolicies`)**: If no instance override is present, the engine matches the server's `GroupName` (case-insensitive) against configured group policies (e.g. `Production`, `Non-Production`, `DR`, `Staging`, `Tier-1`).
3. **Global Default Policy (`GlobalPolicies`)**: Fallback policy applied to servers whose group has no specific policy defined.

---

## 8. Configuration & Inventory Management (`config.json`)

The script automatically generates and maintains a local `config.json` file. You can edit this file manually or manage it via the **Instances**, **SLA Policies**, **Email Alerts**, and **Repository DB** tabs in the web interface.

### Connection Encryption & Certificate Trust:
Both SQL Server Instances and the Repository Database support granular TLS encryption settings:
- **`Encryption`**:
  - `Optional`: Standard default connection without enforcing TLS encryption.
  - `Mandatory` (`Encrypt=True`): Requires TLS encryption for all client-server traffic.
  - `Strict` (`Encrypt=Strict`): Enforces modern TDS 8.0 / TLS 1.3 encryption.
- **`TrustServerCertificate`**: Checkbox / boolean (`$true` / `$false`) allowing self-signed or enterprise CA certificates without certificate chain errors.

### Sample `config.json` with Group-Wise & Server-Wise SLA:
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
    },
    {
      "GroupName": "Non-Production",
      "FullBackupWarningHours": 48,
      "FullBackupCriticalHours": 72,
      "DiffBackupWarningHours": 24,
      "DiffBackupCriticalHours": 48,
      "LogBackupWarningMinutes": 120,
      "LogBackupCriticalMinutes": 240,
      "IgnoreSimpleRecoveryLogs": true
    }
  ],
  "RepositorySettings": {
    "IsEnabled": true,
    "ServerAddress": "SQLREPO01",
    "Port": 1433,
    "DatabaseName": "SQLBackupMonitorDB",
    "AuthType": "Windows",
    "Username": "",
    "Password": "",
    "Encryption": "Optional",
    "TrustServerCertificate": true,
    "RetentionDays": 90,
    "AutoCreateSchema": true
  },
  "EmailSettings": {
    "IsEnabled": true,
    "SmtpServer": "smtp.office365.com",
    "SmtpPort": 587,
    "EnableSsl": true,
    "SenderEmail": "dba-alerts@yourcompany.com",
    "SenderDisplayName": "SQLBackupMonitor Alerts",
    "SmtpUsername": "dba-alerts@yourcompany.com",
    "SmtpPassword": "YourAppPasswordHere",
    "RecipientEmails": "dba-team@yourcompany.com, oncall-dba@yourcompany.com",
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
      "Username": "",
      "Password": "",
      "Encryption": "Optional",
      "TrustServerCertificate": true,
      "IsEnabled": true,
      "GroupName": "Production",
      "Environment": "Production",
      "UseCustomSla": false,
      "CustomPolicies": null
    },
    {
      "Id": "b2c3d4e5-f6a7-8901-bcde-2345678901bc",
      "Name": "Critical Payment Gateway SQL",
      "ServerAddress": "SQLPAY01.corp.local",
      "Port": 1433,
      "AuthType": "Windows",
      "Username": "",
      "Password": "",
      "Encryption": "Mandatory",
      "TrustServerCertificate": true,
      "IsEnabled": true,
      "GroupName": "Production",
      "Environment": "Production",
      "UseCustomSla": true,
      "CustomPolicies": {
        "FullBackupWarningHours": 4,
        "FullBackupCriticalHours": 6,
        "DiffBackupWarningHours": 2,
        "DiffBackupCriticalHours": 3,
        "LogBackupWarningMinutes": 10,
        "LogBackupCriticalMinutes": 15,
        "IgnoreSimpleRecoveryLogs": true
      }
    }
  ]
}
```

---

## 8. Automated Email & Daily Scheduled PDF Reports

`SQLBackupMonitor.ps1` includes an integrated, native SMTP alerting engine that dispatches rich, HTML-formatted email alerts and executive daily PDF reports.

### Key Capabilities:
- **Daily Scheduled PDF Report Delivery**: Automatically generate and dispatch a morning executive summary with attached PDF report at a configured time (e.g. `08:00`).
- **Zero Third-Party Dependencies**: Uses native Windows `System.Net.Mail.SmtpClient` supporting TLS/SSL encryption and authenticated SMTP (Office 365, Gmail, Exchange Server, internal relay).
- **Intelligent Alert Cooldown**: Prevents alert flooding. Default 60-minute cooldown per breached database ensures you are notified when a problem starts without getting spammed every polling cycle.
- **Configurable Trigger Levels**: Selectively trigger on **Critical Breaches** and/or **Warning Thresholds**.
- **Formatted Diagnostic Tables**: Emails include a dark-themed summary table listing Server, Database, Recovery Model, Last Full/Log Backup timestamps, and exact diagnostic reasons.
- **One-Click Delivery & Test**: In the Web Dashboard, click **Send Daily PDF Report Now** or **Send Test Alert Email** to verify your mail server configuration immediately.

### Sample Email Configuration in `config.json`:
```json
{
  "EmailSettings": {
    "IsEnabled": true,
    "SmtpServer": "smtp.office365.com",
    "SmtpPort": 587,
    "EnableSsl": true,
    "SenderEmail": "sql-alerts@yourcompany.com",
    "SenderDisplayName": "SQLBackupMonitor DBA Alerts",
    "SmtpUsername": "sql-alerts@yourcompany.com",
    "SmtpPassword": "YourAppPasswordHere",
    "RecipientEmails": "dba-team@yourcompany.com, oncall-dba@yourcompany.com",
    "AlertOnCritical": true,
    "AlertOnWarning": false,
    "CooldownMinutes": 60
  }
}
```

---

## 9. SQL Server Permissions Required

`SQLBackupMonitor` only requires **read-only metadata permissions**. It does **not** modify databases or execute backups.

### Minimum Required Permissions Script:
Run the following T-SQL script on your target SQL Server instances to create a dedicated read-only audit user:

```sql
-- Create Login
CREATE LOGIN [backup_monitor_svc] WITH PASSWORD = 'StrongPassword!2026', CHECK_POLICY = ON;
GO

-- Grant View Server State & Connect
GRANT CONNECT SQL TO [backup_monitor_svc];
GRANT VIEW SERVER STATE TO [backup_monitor_svc];
GO

-- Grant Read Access to msdb backup history
USE [msdb];
GO
CREATE USER [backup_monitor_svc] FOR LOGIN [backup_monitor_svc];
GRANT SELECT ON [dbo].[backupset] TO [backup_monitor_svc];
GRANT SELECT ON [dbo].[backupmediafamily] TO [backup_monitor_svc];
GO

-- Grant Read Access to master databases metadata
USE [master];
GO
CREATE USER [backup_monitor_svc] FOR LOGIN [backup_monitor_svc];
GRANT SELECT ON sys.databases TO [backup_monitor_svc];
GO
```

> **Note on Windows Authentication**: If running under your domain account with Windows Authentication (`"AuthType": "Windows"`), ensure your Active Directory account has `VIEW SERVER STATE` and `db_datareader` on `msdb`.

---

## 10. Automating with Windows Task Scheduler

You can schedule automated daily HTML audit reports or background checks without any user interaction:

1. Open **Task Scheduler** (`taskschd.msc`).
2. Click **Create Task** (e.g. `SQL Backup Daily Audit Report`).
3. Set **Trigger**: Daily at 07:00 AM.
4. Set **Action**:
   - **Program/script**: `powershell.exe`
   - **Add arguments**:
     ```text
     -NoProfile -ExecutionPolicy Bypass -File "D:\Backup_Monitoring\SQLBackupMonitor.ps1" -Mode Report -ReportPath "D:\Backup_Monitoring\Reports\DailyAudit.html" -NoBrowser
     ```
   - **Start in**: `D:\Backup_Monitoring`
5. Check **Run whether user is logged on or not**.

---

## 11. Troubleshooting & FAQ

### Q1: "Running scripts is disabled on this system" error
**Solution**: Run the script with `-ExecutionPolicy Bypass`:
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\SQLBackupMonitor.ps1
```

### Q2: Port 5000 is already in use
**Solution**: Specify an alternate port with the `-Port` parameter:
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\SQLBackupMonitor.ps1 -Port 5050
```

### Q3: Cannot connect to a Named Instance (e.g. `SERVER\INSTANCE`)
**Solution**: In the Instance Manager or `config.json`, enter `SERVER\INSTANCE` in the `ServerAddress` field and set `Port` to `0` or `1433`. Ensure the SQL Server Browser service is running on the target machine.

### Q4: Are SQL credentials encrypted?
**Solution**: In corporate environments using Windows Integrated Authentication (`"AuthType": "Windows"`), no passwords are stored at all. For SQL Authentication, passwords in `config.json` are read directly; ensure NTFS permissions on `config.json` restrict access to authorized administrators.

---

*Documentation maintained for SQLBackupMonitor v2.5.0 Standalone PowerShell Edition.*
