<#
================================================================================
 SQLBackupMonitor - Enterprise SQL Server Backup Monitoring Platform
 Single-File Standalone PowerShell Edition (.ps1)
 
 100% Native Windows PowerShell - Zero External Executables (.exe) Required
 Features:
  - Multi-Instance SQL Server Backup Telemetry & SLA Tracking
  - AlwaysOn Availability Groups (AG) Health & Replica Sync
  - Granular Thresholds, Maintenance Windows & Policies
  - Active Alert Logs & SLA Breach Management
  - Native PDF Report Generation & HTML Export
  - Email / SMTP SLA Alerts with Attached PDF Audit Reports
  - SQL Server Repository Database & Historical Trend Analytics
  - Activity & System Audit Logging
  - Embedded Dark-Mode Glassmorphism Web Dashboard & REST API
================================================================================
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet("Web", "Scan", "Report", "PdfReport", "Loop", "EmailReport")]
    [string]$Mode = "Web",

    [Parameter()]
    [int]$Port = 5000,

    [Parameter()]
    [string]$ConfigFile = "",

    [Parameter()]
    [switch]$NoBrowser,

    [Parameter()]
    [string]$ReportPath = "",

    [Parameter()]
    [int]$IntervalSeconds = 60
)

# Resolve default paths safely
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($ConfigFile)) { $ConfigFile = Join-Path $scriptDir "config.json" }
if ([string]::IsNullOrWhiteSpace($ReportPath)) { 
    $ReportPath = if ($Mode -eq "PdfReport") { Join-Path $scriptDir "BackupReport.pdf" } else { Join-Path $scriptDir "BackupReport.html" }
}

# Set TLS protocols for compatibility
try {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
} catch {}

# Global State
$Global:AppVersion = "2.6.0-ps"
$Global:StartTime = [DateTime]::UtcNow
$Global:LastScanTime = $null
$Global:IsScanning = $false
$Global:CachedResults = @()
$Global:CachedAlwaysOn = @()
$Global:CachedMetrics = @{}
$Global:ActiveAlerts = @()
$Global:AuditLogs = @()
$Global:ScanLock = [System.Object]::new()

$Global:LastAlertSentTimes = @{} # "ServerId:DbName:Status" -> [DateTime]

# Audit Logger Helper
function Add-AuditLog {
    param(
        [string]$Action,
        [string]$Details,
        [string]$User = "System"
    )
    $log = [PSCustomObject]@{
        Id        = [Guid]::NewGuid().ToString()
        Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        User      = $User
        Action    = $Action
        Details   = $Details
    }
    $Global:AuditLogs = @($log) + @($Global:AuditLogs | Select-Object -First 99)
}

# ==============================================================================
# CONFIGURATION MANAGEMENT
# ==============================================================================
function Get-DefaultConfig {
    return [PSCustomObject]@{
        GlobalPolicies = [PSCustomObject]@{
            FullBackupWarningHours    = 20
            FullBackupCriticalHours   = 24
            DiffBackupWarningHours    = 10
            DiffBackupCriticalHours   = 12
            LogBackupWarningMinutes   = 20
            LogBackupCriticalMinutes  = 30
            AlertOnMissingBackups     = $true
            IgnoreSimpleRecoveryLogs  = $true
            AutoRefreshIntervalSec    = 60
            ExcludeWeekends           = $false
            MaintenanceWindowStart    = ""
            MaintenanceWindowEnd      = ""
        }
        GroupPolicies = @(
            [PSCustomObject]@{
                GroupName                = "Production"
                FullBackupWarningHours   = 20
                FullBackupCriticalHours  = 24
                DiffBackupWarningHours   = 10
                DiffBackupCriticalHours  = 12
                LogBackupWarningMinutes  = 20
                LogBackupCriticalMinutes = 30
                IgnoreSimpleRecoveryLogs = $true
            },
            [PSCustomObject]@{
                GroupName                = "Non-Production"
                FullBackupWarningHours   = 48
                FullBackupCriticalHours  = 72
                DiffBackupWarningHours   = 24
                DiffBackupCriticalHours  = 48
                LogBackupWarningMinutes  = 120
                LogBackupCriticalMinutes = 240
                IgnoreSimpleRecoveryLogs = $true
            }
        )
        RepositorySettings = [PSCustomObject]@{
            IsEnabled               = $false
            ServerAddress           = "localhost"
            Port                    = 1433
            DatabaseName            = "SQLBackupMonitorDB"
            AuthType                = "Windows"   # Windows | SqlPassword
            Username                = ""
            Password                = ""
            Encryption              = "Optional"  # Optional | Mandatory | Strict
            TrustServerCertificate  = $true
            RetentionDays           = 90
            AutoCreateSchema        = $true
        }
        EmailSettings = [PSCustomObject]@{
            IsEnabled           = $false
            SmtpServer          = "smtp.office365.com"
            SmtpPort            = 587
            EnableSsl           = $true
            SenderEmail         = "dba-alerts@yourcompany.com"
            SenderDisplayName   = "SQLBackupMonitor Alerts"
            SmtpUsername        = ""
            SmtpPassword        = ""
            RecipientEmails     = "dba-team@yourcompany.com"
            AlertOnCritical     = $true
            AlertOnWarning      = $false
            AttachPdfReport     = $true
            DailyReportEnabled  = $false
            DailyReportTime     = "08:00"
            LastDailyReportDate = ""
            CooldownMinutes     = 60
        }
        Servers = @()
    }
}

function Load-AppConfig {
    if (Test-Path $ConfigFile) {
        try {
            $json = Get-Content -Path $ConfigFile -Raw -Encoding UTF8
            $cfg = $json | ConvertFrom-Json
            if ($null -eq $cfg.GlobalPolicies) { $cfg | Add-Member -MemberType NoteProperty -Name "GlobalPolicies" -Value (Get-DefaultConfig).GlobalPolicies -Force }
            if ($null -eq $cfg.GroupPolicies) { $cfg | Add-Member -MemberType NoteProperty -Name "GroupPolicies" -Value (Get-DefaultConfig).GroupPolicies -Force }
            if ($null -eq $cfg.RepositorySettings) { $cfg | Add-Member -MemberType NoteProperty -Name "RepositorySettings" -Value (Get-DefaultConfig).RepositorySettings -Force }
            if ($null -eq $cfg.EmailSettings) { $cfg | Add-Member -MemberType NoteProperty -Name "EmailSettings" -Value (Get-DefaultConfig).EmailSettings -Force }
            if ($null -eq $cfg.Servers) { 
                $cfg | Add-Member -MemberType NoteProperty -Name "Servers" -Value @() -Force 
            } else {
                # Deduplicate and ensure valid unique IDs for all servers
                $seen = @{}
                $cleanServers = [System.Collections.Generic.List[PSCustomObject]]::new()
                foreach ($s in @($cfg.Servers)) {
                    if ($null -eq $s) { continue }
                    if (-not $s.Id -or [string]::IsNullOrWhiteSpace($s.Id)) {
                        $s | Add-Member -MemberType NoteProperty -Name "Id" -Value ([Guid]::NewGuid().ToString()) -Force
                    }
                    $sKey = "$($s.ServerAddress):$($s.Port)".ToLower()
                    $nKey = if ($s.Name) { $s.Name.ToLower() } else { $sKey }
                    if (-not $seen.ContainsKey($sKey) -and -not $seen.ContainsKey($nKey)) {
                        $seen[$sKey] = $true
                        $seen[$nKey] = $true
                        $cleanServers.Add($s)
                    }
                }
                $cfg.Servers = $cleanServers.ToArray()
            }
            return $cfg
        } catch {
            Write-Warning "Failed to parse $ConfigFile. Using default configuration."
        }
    }
    $defaultCfg = Get-DefaultConfig
    Save-AppConfig -Config $defaultCfg
    return $defaultCfg
}

function Save-AppConfig {
    param([Parameter(Mandatory=$true)]$Config)
    try {
        $json = $Config | ConvertTo-Json -Depth 10
        $json | Set-Content -Path $ConfigFile -Encoding UTF8 -Force
        Add-AuditLog -Action "Configuration Saved" -Details "Updated application configuration in $ConfigFile"
    } catch {
        Write-Error "Failed to save configuration to $ConfigFile : $_"
    }
}

function Get-EffectiveSlaPolicy {
    param([Parameter(Mandatory=$true)]$ServerObj)

    # 1. Check Server Custom SLA Override
    if ($ServerObj.UseCustomSla -and $ServerObj.CustomPolicies) {
        $cp = $ServerObj.CustomPolicies
        return [PSCustomObject]@{
            Tier                     = "Server Override ($($ServerObj.Name))"
            FullBackupWarningHours   = if ($cp.FullBackupWarningHours) { [double]$cp.FullBackupWarningHours } else { 20.0 }
            FullBackupCriticalHours  = if ($cp.FullBackupCriticalHours) { [double]$cp.FullBackupCriticalHours } else { 24.0 }
            DiffBackupWarningHours   = if ($cp.DiffBackupWarningHours) { [double]$cp.DiffBackupWarningHours } else { 10.0 }
            DiffBackupCriticalHours  = if ($cp.DiffBackupCriticalHours) { [double]$cp.DiffBackupCriticalHours } else { 12.0 }
            LogBackupWarningMinutes  = if ($cp.LogBackupWarningMinutes) { [double]$cp.LogBackupWarningMinutes } else { 20.0 }
            LogBackupCriticalMinutes = if ($cp.LogBackupCriticalMinutes) { [double]$cp.LogBackupCriticalMinutes } else { 30.0 }
            IgnoreSimpleRecoveryLogs = if ($null -ne $cp.IgnoreSimpleRecoveryLogs) { [bool]$cp.IgnoreSimpleRecoveryLogs } else { $true }
        }
    }

    # 2. Check Group-Level SLA Policy
    if ($ServerObj.GroupName -and $Global:Config.GroupPolicies) {
        $grpTarget = $ServerObj.GroupName.Trim().ToLower()
        $matched = @($Global:Config.GroupPolicies | Where-Object { $_.GroupName -and $_.GroupName.Trim().ToLower() -eq $grpTarget })
        if ($matched.Count -gt 0) {
            $gp = $matched[0]
            return [PSCustomObject]@{
                Tier                     = "Group ($($gp.GroupName))"
                FullBackupWarningHours   = if ($gp.FullBackupWarningHours) { [double]$gp.FullBackupWarningHours } else { 20.0 }
                FullBackupCriticalHours  = if ($gp.FullBackupCriticalHours) { [double]$gp.FullBackupCriticalHours } else { 24.0 }
                DiffBackupWarningHours   = if ($gp.DiffBackupWarningHours) { [double]$gp.DiffBackupWarningHours } else { 10.0 }
                DiffBackupCriticalHours  = if ($gp.DiffBackupCriticalHours) { [double]$gp.DiffBackupCriticalHours } else { 12.0 }
                LogBackupWarningMinutes  = if ($gp.LogBackupWarningMinutes) { [double]$gp.LogBackupWarningMinutes } else { 20.0 }
                LogBackupCriticalMinutes = if ($gp.LogBackupCriticalMinutes) { [double]$gp.LogBackupCriticalMinutes } else { 30.0 }
                IgnoreSimpleRecoveryLogs = if ($null -ne $gp.IgnoreSimpleRecoveryLogs) { [bool]$gp.IgnoreSimpleRecoveryLogs } else { $true }
            }
        }
    }

    # 3. Fallback to Global SLA Policy
    $glob = $Global:Config.GlobalPolicies
    return [PSCustomObject]@{
        Tier                     = "Global Policy"
        FullBackupWarningHours   = if ($glob.FullBackupWarningHours) { [double]$glob.FullBackupWarningHours } else { 20.0 }
        FullBackupCriticalHours  = if ($glob.FullBackupCriticalHours) { [double]$glob.FullBackupCriticalHours } else { 24.0 }
        DiffBackupWarningHours   = if ($glob.DiffBackupWarningHours) { [double]$glob.DiffBackupWarningHours } else { 10.0 }
        DiffBackupCriticalHours  = if ($glob.DiffBackupCriticalHours) { [double]$glob.DiffBackupCriticalHours } else { 12.0 }
        LogBackupWarningMinutes  = if ($glob.LogBackupWarningMinutes) { [double]$glob.LogBackupWarningMinutes } else { 20.0 }
        LogBackupCriticalMinutes = if ($glob.LogBackupCriticalMinutes) { [double]$glob.LogBackupCriticalMinutes } else { 30.0 }
        IgnoreSimpleRecoveryLogs = if ($null -ne $glob.IgnoreSimpleRecoveryLogs) { [bool]$glob.IgnoreSimpleRecoveryLogs } else { $true }
    }
}

$Global:Config = Load-AppConfig
Add-AuditLog -Action "System Start" -Details "SQLBackupMonitor standalone platform started (Version $Global:AppVersion)"

# ==============================================================================
# SQL SERVER QUERYING ENGINE (ADO.NET)
# ==============================================================================
function Build-SqlConnectionString {
    param(
        [Parameter(Mandatory=$true)]$ServerObj,
        [string]$Database = "master"
    )

    $target = $ServerObj.ServerAddress
    if ($ServerObj.Port -and $ServerObj.Port -ne 1433 -and -not ($target -match ",")) {
        $target = "$target,$($ServerObj.Port)"
    }

    $builder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new()
    $builder["Data Source"] = $target
    $builder["Initial Catalog"] = $Database
    $builder["Connect Timeout"] = 8

    # Trust Server Certificate
    if ($null -ne $ServerObj.TrustServerCertificate) {
        $builder["TrustServerCertificate"] = [bool]$ServerObj.TrustServerCertificate
    } else {
        $builder["TrustServerCertificate"] = $true
    }

    # Encryption Mode: Optional | Mandatory | Strict
    $enc = if ($ServerObj.Encryption) { $ServerObj.Encryption } else { "Optional" }
    if ($enc -eq "Mandatory" -or $enc -eq "Strict" -or $enc -eq $true) {
        $builder["Encrypt"] = $true
    } else {
        $builder["Encrypt"] = $false
    }

    $builder["Application Name"] = "SQLBackupMonitor-PS"

    if ($ServerObj.AuthType -eq "SqlPassword" -or (-not [string]::IsNullOrWhiteSpace($ServerObj.Username))) {
        $builder["Integrated Security"] = $false
        $builder["User ID"] = $ServerObj.Username
        $builder["Password"] = $ServerObj.Password
    } else {
        $builder["Integrated Security"] = $true
    }

    return $builder.ConnectionString
}

function Test-SqlServerConnection {
    param([Parameter(Mandatory=$true)]$ServerObj)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $connStr = Build-SqlConnectionString -ServerObj $ServerObj

    $result = [PSCustomObject]@{
        Success        = $false
        ServerName     = $ServerObj.Name
        ProductVersion = "Unknown"
        Edition        = "Unknown"
        IsHadrEnabled  = $false
        LatencyMs      = 0
        ErrorMessage   = $null
    }

    try {
        $conn = [System.Data.SqlClient.SqlConnection]::new($connStr)
        $conn.Open()
        
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = @"
SELECT 
    CAST(SERVERPROPERTY('ServerName') AS NVARCHAR(128)) AS ServerName,
    CAST(SERVERPROPERTY('ProductVersion') AS NVARCHAR(128)) AS ProductVersion,
    CAST(SERVERPROPERTY('ProductLevel') AS NVARCHAR(128)) AS ProductLevel,
    CAST(SERVERPROPERTY('Edition') AS NVARCHAR(128)) AS Edition,
    CAST(ISNULL(SERVERPROPERTY('IsHadrEnabled'), 0) AS INT) AS IsHadrEnabled;
"@
        $cmd.CommandTimeout = 10
        $reader = $cmd.ExecuteReader()
        if ($reader.Read()) {
            $result.Success = $true
            $result.ProductVersion = "$($reader['ProductVersion']) ($($reader['ProductLevel']))"
            $result.Edition = "$($reader['Edition'])"
            $result.IsHadrEnabled = ([int]$reader['IsHadrEnabled'] -eq 1)
        }
        $reader.Close()
        $conn.Close()
    } catch {
        $result.Success = $false
        $result.ErrorMessage = $_.Exception.Message
    } finally {
        $sw.Stop()
        $result.LatencyMs = [Math]::Round($sw.Elapsed.TotalMilliseconds, 1)
    }

    return $result
}

$Global:ServerTelemetryScriptBlock = {
    param($ServerObj, $GlobalPolicies, $GroupPolicies, $RepoDbName)

    $dbResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    $agResults = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Calculate SLA policy for this server
    $policy = $null
    if ($ServerObj.UseCustomSla -and $ServerObj.CustomPolicies) {
        $cp = $ServerObj.CustomPolicies
        $policy = [PSCustomObject]@{
            Tier                     = "Server Override ($($ServerObj.Name))"
            FullBackupWarningHours   = if ($cp.FullBackupWarningHours) { [double]$cp.FullBackupWarningHours } else { 20.0 }
            FullBackupCriticalHours  = if ($cp.FullBackupCriticalHours) { [double]$cp.FullBackupCriticalHours } else { 24.0 }
            DiffBackupWarningHours   = if ($cp.DiffBackupWarningHours) { [double]$cp.DiffBackupWarningHours } else { 10.0 }
            DiffBackupCriticalHours  = if ($cp.DiffBackupCriticalHours) { [double]$cp.DiffBackupCriticalHours } else { 12.0 }
            LogBackupWarningMinutes  = if ($cp.LogBackupWarningMinutes) { [double]$cp.LogBackupWarningMinutes } else { 20.0 }
            LogBackupCriticalMinutes = if ($cp.LogBackupCriticalMinutes) { [double]$cp.LogBackupCriticalMinutes } else { 30.0 }
            IgnoreSimpleRecoveryLogs = if ($null -ne $cp.IgnoreSimpleRecoveryLogs) { [bool]$cp.IgnoreSimpleRecoveryLogs } else { $true }
        }
    } elseif ($ServerObj.GroupName -and $GroupPolicies) {
        $grpTarget = $ServerObj.GroupName.Trim().ToLower()
        $matched = @($GroupPolicies | Where-Object { $_.GroupName -and $_.GroupName.Trim().ToLower() -eq $grpTarget })
        if ($matched.Count -gt 0) {
            $gp = $matched[0]
            $policy = [PSCustomObject]@{
                Tier                     = "Group ($($gp.GroupName))"
                FullBackupWarningHours   = if ($gp.FullBackupWarningHours) { [double]$gp.FullBackupWarningHours } else { 20.0 }
                FullBackupCriticalHours  = if ($gp.FullBackupCriticalHours) { [double]$gp.FullBackupCriticalHours } else { 24.0 }
                DiffBackupWarningHours   = if ($gp.DiffBackupWarningHours) { [double]$gp.DiffBackupWarningHours } else { 10.0 }
                DiffBackupCriticalHours  = if ($gp.DiffBackupCriticalHours) { [double]$gp.DiffBackupCriticalHours } else { 12.0 }
                LogBackupWarningMinutes  = if ($gp.LogBackupWarningMinutes) { [double]$gp.LogBackupWarningMinutes } else { 20.0 }
                LogBackupCriticalMinutes = if ($gp.LogBackupCriticalMinutes) { [double]$gp.LogBackupCriticalMinutes } else { 30.0 }
                IgnoreSimpleRecoveryLogs = if ($null -ne $gp.IgnoreSimpleRecoveryLogs) { [bool]$gp.IgnoreSimpleRecoveryLogs } else { $true }
            }
        }
    }
    if ($null -eq $policy) {
        $glob = $GlobalPolicies
        $policy = [PSCustomObject]@{
            Tier                     = "Global Policy"
            FullBackupWarningHours   = if ($glob.FullBackupWarningHours) { [double]$glob.FullBackupWarningHours } else { 20.0 }
            FullBackupCriticalHours  = if ($glob.FullBackupCriticalHours) { [double]$glob.FullBackupCriticalHours } else { 24.0 }
            DiffBackupWarningHours   = if ($glob.DiffBackupWarningHours) { [double]$glob.DiffBackupWarningHours } else { 10.0 }
            DiffBackupCriticalHours  = if ($glob.DiffBackupCriticalHours) { [double]$glob.DiffBackupCriticalHours } else { 12.0 }
            LogBackupWarningMinutes  = if ($glob.LogBackupWarningMinutes) { [double]$glob.LogBackupWarningMinutes } else { 20.0 }
            LogBackupCriticalMinutes = if ($glob.LogBackupCriticalMinutes) { [double]$glob.LogBackupCriticalMinutes } else { 30.0 }
            IgnoreSimpleRecoveryLogs = if ($null -ne $glob.IgnoreSimpleRecoveryLogs) { [bool]$glob.IgnoreSimpleRecoveryLogs } else { $true }
        }
    }

    $target = $ServerObj.ServerAddress
    if ($ServerObj.Port -and $ServerObj.Port -ne 1433 -and -not ($target -match ",")) {
        $target = "$target,$($ServerObj.Port)"
    }

    $builder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new()
    $builder["Data Source"] = $target
    $builder["Initial Catalog"] = "master"
    $builder["Connect Timeout"] = 5
    $builder["TrustServerCertificate"] = if ($null -ne $ServerObj.TrustServerCertificate) { [bool]$ServerObj.TrustServerCertificate } else { $true }
    $enc = if ($ServerObj.Encryption) { $ServerObj.Encryption } else { "Optional" }
    $builder["Encrypt"] = ($enc -eq "Mandatory" -or $enc -eq "Strict" -or $enc -eq $true)
    $builder["Application Name"] = "SQLBackupMonitor-PS"
    if ($ServerObj.AuthType -eq "SqlPassword" -or (-not [string]::IsNullOrWhiteSpace($ServerObj.Username))) {
        $builder["Integrated Security"] = $false
        $builder["User ID"] = $ServerObj.Username
        $builder["Password"] = $ServerObj.Password
    } else {
        $builder["Integrated Security"] = $true
    }

    $connStr = $builder.ConnectionString

    $batchQuery = @"
SELECT 
    db.database_id AS DatabaseId,
    db.name AS DatabaseName,
    CAST(DATABASEPROPERTYEX(db.name, 'Recovery') AS NVARCHAR(50)) AS RecoveryModel,
    db.state_desc AS StateDesc,
    MAX(CASE WHEN bs.type = 'D' THEN bs.backup_finish_date END) AS LastFullBackup,
    MAX(CASE WHEN bs.type = 'I' THEN bs.backup_finish_date END) AS LastDifferentialBackup,
    MAX(CASE WHEN bs.type = 'L' THEN bs.backup_finish_date END) AS LastLogBackup,
    DATEDIFF(MINUTE, MAX(CASE WHEN bs.type = 'D' THEN bs.backup_finish_date END), GETDATE()) / 60.0 AS FullBackupAgeHours,
    DATEDIFF(MINUTE, MAX(CASE WHEN bs.type = 'I' THEN bs.backup_finish_date END), GETDATE()) / 60.0 AS DiffBackupAgeHours,
    DATEDIFF(MINUTE, MAX(CASE WHEN bs.type = 'L' THEN bs.backup_finish_date END), GETDATE()) * 1.0 AS LogBackupAgeMinutes,
    MAX(CASE WHEN bs.type = 'D' THEN CAST(bs.backup_size / 1073741824.0 AS DECIMAL(18,3)) END) AS LastBackupSizeGB
FROM sys.databases db
LEFT JOIN msdb.dbo.backupset bs ON db.name COLLATE DATABASE_DEFAULT = bs.database_name COLLATE DATABASE_DEFAULT
WHERE db.database_id > 4 
  AND db.name NOT IN ('master', 'tempdb', 'model', 'msdb', 'distribution', 'SSISDB')
  AND (@repoDbName = '' OR db.name <> @repoDbName)
  AND db.state_desc = 'ONLINE' 
  AND db.is_read_only = 0
  AND db.is_in_standby = 0
GROUP BY db.database_id, db.name, db.state_desc
ORDER BY db.name ASC;

IF SERVERPROPERTY('IsHadrEnabled') = 1
BEGIN
    SELECT 
        ag.name AS GroupName,
        ag.automated_backup_preference_desc AS BackupPreference,
        ar.replica_server_name AS ReplicaServerName,
        ars.role_desc AS Role,
        ars.synchronization_health_desc AS SyncHealth,
        ars.connected_state_desc AS ConnectedState,
        db.name AS DatabaseName,
        drs.synchronization_state_desc AS DbSyncState,
        ISNULL(drs.is_suspended, 0) AS IsSuspended
    FROM sys.availability_groups ag
    INNER JOIN sys.availability_replicas ar ON ag.group_id = ar.group_id
    INNER JOIN sys.dm_hadr_availability_replica_states ars ON ar.replica_id = ars.replica_id
    LEFT JOIN sys.dm_hadr_database_replica_states drs ON ar.replica_id = drs.replica_id
    LEFT JOIN sys.databases db ON drs.database_id = db.database_id
    WHERE db.database_id IS NULL OR (db.database_id > 4 AND db.name NOT IN ('master', 'tempdb', 'model', 'msdb', 'distribution', 'SSISDB'))
    ORDER BY ag.name, ar.replica_server_name, db.name;
END
"@

    try {
        $conn = [System.Data.SqlClient.SqlConnection]::new($connStr)
        $conn.Open()

        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $batchQuery
        $cmd.Parameters.AddWithValue("@repoDbName", $(if ($RepoDbName) { $RepoDbName } else { "" })) | Out-Null
        $cmd.CommandTimeout = 8
        $reader = $cmd.ExecuteReader()

        # 1. Process Database Result Set
        $critFullHours = $policy.FullBackupCriticalHours
        $warnFullHours = $policy.FullBackupWarningHours
        $critDiffHours = $policy.DiffBackupCriticalHours
        $warnDiffHours = $policy.DiffBackupWarningHours
        $critLogMin    = $policy.LogBackupCriticalMinutes
        $warnLogMin    = $policy.LogBackupWarningMinutes
        $ignoreSimple  = $policy.IgnoreSimpleRecoveryLogs
        $policyTier    = $policy.Tier

        while ($reader.Read()) {
            $dbName   = [string]$reader["DatabaseName"]
            $recModel = if ($reader["RecoveryModel"] -ne [DBNull]::Value) { [string]$reader["RecoveryModel"] } else { "SIMPLE" }
            $fullAge  = if ($reader["FullBackupAgeHours"] -ne [DBNull]::Value) { [double]$reader["FullBackupAgeHours"] } else { $null }
            $diffAge  = if ($reader["DiffBackupAgeHours"] -ne [DBNull]::Value) { [double]$reader["DiffBackupAgeHours"] } else { $null }
            $logAge   = if ($reader["LogBackupAgeMinutes"] -ne [DBNull]::Value) { [double]$reader["LogBackupAgeMinutes"] } else { $null }
            $lastFull = if ($reader["LastFullBackup"] -ne [DBNull]::Value) { ([DateTime]$reader["LastFullBackup"]).ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
            $lastDiff = if ($reader["LastDifferentialBackup"] -ne [DBNull]::Value) { ([DateTime]$reader["LastDifferentialBackup"]).ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
            $lastLog  = if ($reader["LastLogBackup"] -ne [DBNull]::Value) { ([DateTime]$reader["LastLogBackup"]).ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
            $sizeGB   = if ($reader["LastBackupSizeGB"] -ne [DBNull]::Value) { [Math]::Round([double]$reader["LastBackupSizeGB"], 2) } else { 0.0 }

            $status = "Healthy"
            $reasons = [System.Collections.Generic.List[string]]::new()

            if ($null -eq $fullAge -or $fullAge -lt 0) {
                $status = "Critical"
                $reasons.Add("No Full Backup found on record")
            }
            elseif ($fullAge -ge $critFullHours) {
                $status = "Critical"
                $daysOld = [Math]::Round($fullAge / 24.0, 1)
                $reasons.Add("Full Backup is $([Math]::Round($fullAge,1))h ($daysOld days) old (${policyTier} SLA: ${critFullHours}h)")
            }
            elseif ($fullAge -ge $warnFullHours) {
                if ($status -ne "Critical") { $status = "Warning" }
                $reasons.Add("Full Backup is $([Math]::Round($fullAge,1))h old (${policyTier} SLA: ${warnFullHours}h)")
            }

            if ($recModel -ne "SIMPLE" -or (-not $ignoreSimple)) {
                if ($null -eq $logAge -or $logAge -lt 0) {
                    $status = "Critical"
                    $reasons.Add("Missing Log Backup ($recModel recovery model)")
                }
                elseif ($logAge -ge $critLogMin) {
                    $status = "Critical"
                    $reasons.Add("Log Backup is $([Math]::Round($logAge,0))m old (${policyTier} SLA: ${critLogMin}m)")
                }
                elseif ($logAge -ge $warnLogMin) {
                    if ($status -eq "Healthy") { $status = "Warning" }
                    $reasons.Add("Log Backup is $([Math]::Round($logAge,0))m old (${policyTier} SLA: ${warnLogMin}m)")
                }
            }

            $finalReason = if ($reasons.Count -gt 0) { $reasons -join " | " } else { "Backup policy within SLA limits" }

            $dbResults.Add([PSCustomObject]@{
                ServerId               = $ServerObj.Id
                ServerName             = $ServerObj.Name
                DatabaseName           = $dbName
                RecoveryModel          = $recModel
                LastFullBackup         = $lastFull
                LastDifferentialBackup = $lastDiff
                LastLogBackup          = $lastLog
                FullBackupAgeHours     = if ($fullAge) { [Math]::Round($fullAge, 1) } else { -1 }
                LogBackupAgeMinutes    = if ($logAge) { [Math]::Round($logAge, 1) } else { -1 }
                LastBackupSizeGB       = $sizeGB
                Status                 = $status
                StatusReason           = $finalReason
                SlaTier                = $policyTier
                Timestamp              = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            })
        }

        # 2. Process AlwaysOn Result Set (if present)
        if ($reader.NextResult()) {
            while ($reader.Read()) {
                $agResults.Add([PSCustomObject]@{
                    ServerName        = $ServerObj.Name
                    GroupName         = [string]$reader["GroupName"]
                    BackupPreference  = [string]$reader["BackupPreference"]
                    ReplicaServerName = [string]$reader["ReplicaServerName"]
                    Role              = [string]$reader["Role"]
                    SyncHealth        = [string]$reader["SyncHealth"]
                    ConnectedState    = [string]$reader["ConnectedState"]
                    DatabaseName      = if ($reader["DatabaseName"] -ne [DBNull]::Value) { [string]$reader["DatabaseName"] } else { "All Replicas" }
                    DbSyncState       = if ($reader["DbSyncState"] -ne [DBNull]::Value) { [string]$reader["DbSyncState"] } else { "N/A" }
                    IsSuspended       = ([bool]$reader["IsSuspended"])
                })
            }
        }

        $reader.Close()
        $conn.Close()
    } catch {
        $dbResults.Add([PSCustomObject]@{
            ServerId               = $ServerObj.Id
            ServerName             = $ServerObj.Name
            DatabaseName           = "CONNECTION_FAILED"
            RecoveryModel          = "N/A"
            LastFullBackup         = "N/A"
            LastDifferentialBackup = "N/A"
            LastLogBackup          = "N/A"
            FullBackupAgeHours     = -1
            LogBackupAgeMinutes    = -1
            LastBackupSizeGB       = 0
            Status                 = "Critical"
            StatusReason           = "Connection Error: $($_.Exception.Message)"
            Timestamp              = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        })
    }

    return [PSCustomObject]@{
        Databases  = $dbResults
        AlwaysOn   = $agResults
        ServerId   = $ServerObj.Id
        ServerName = $ServerObj.Name
    }
}

function Get-SqlServerTelemetryCombined {
    param([Parameter(Mandatory=$true)]$ServerObj)
    $repoDb = if ($Global:Config.RepositorySettings -and $Global:Config.RepositorySettings.DatabaseName) { $Global:Config.RepositorySettings.DatabaseName.Trim() } else { "" }
    return & $Global:ServerTelemetryScriptBlock $ServerObj $Global:Config.GlobalPolicies $Global:Config.GroupPolicies $repoDb
}

function Get-SqlServerBackupData {
    param([Parameter(Mandatory=$true)]$ServerObj)
    $res = Get-SqlServerTelemetryCombined -ServerObj $ServerObj
    return $res.Databases
}

function Get-SqlServerAlwaysOnData {
    param([Parameter(Mandatory=$true)]$ServerObj)
    $res = Get-SqlServerTelemetryCombined -ServerObj $ServerObj
    return $res.AlwaysOn
}

# ==============================================================================
# REPOSITORY DATABASE & HISTORICAL TREND ENGINE
# ==============================================================================
function Build-RepoConnectionString {
    param(
        [Parameter(Mandatory=$true)]$RepoCfg,
        [string]$Database = ""
    )

    $dbName = if ($Database) { $Database } elseif ($RepoCfg.DatabaseName) { $RepoCfg.DatabaseName } else { "SQLBackupMonitorDB" }
    $target = $RepoCfg.ServerAddress
    if ($RepoCfg.Port -and $RepoCfg.Port -ne 1433 -and -not ($target -match ",")) {
        $target = "$target,$($RepoCfg.Port)"
    }

    $builder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new()
    $builder["Data Source"] = $target
    $builder["Initial Catalog"] = $dbName
    $builder["Connect Timeout"] = 8

    # Trust Server Certificate
    if ($null -ne $RepoCfg.TrustServerCertificate) {
        $builder["TrustServerCertificate"] = [bool]$RepoCfg.TrustServerCertificate
    } else {
        $builder["TrustServerCertificate"] = $true
    }

    # Encryption Mode: Optional | Mandatory | Strict
    $enc = if ($RepoCfg.Encryption) { $RepoCfg.Encryption } else { "Optional" }
    if ($enc -eq "Mandatory" -or $enc -eq "Strict" -or $enc -eq $true) {
        $builder["Encrypt"] = $true
    } else {
        $builder["Encrypt"] = $false
    }

    $builder["Application Name"] = "SQLBackupMonitor-Repo"

    if ($RepoCfg.AuthType -eq "SqlPassword" -or (-not [string]::IsNullOrWhiteSpace($RepoCfg.Username))) {
        $builder["Integrated Security"] = $false
        $builder["User ID"] = $RepoCfg.Username
        $builder["Password"] = $RepoCfg.Password
    } else {
        $builder["Integrated Security"] = $true
    }

    return $builder.ConnectionString
}

function Test-RepositoryConnection {
    param([Parameter(Mandatory=$true)]$RepoCfg)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $masterConnStr = Build-RepoConnectionString -RepoCfg $RepoCfg -Database "master"
    $dbName = if ($RepoCfg.DatabaseName) { $RepoCfg.DatabaseName } else { "SQLBackupMonitorDB" }

    $result = [PSCustomObject]@{
        Success        = $false
        ServerName     = $RepoCfg.ServerAddress
        DatabaseName   = $dbName
        DatabaseExists = $false
        SchemaReady    = $false
        ProductVersion = "Unknown"
        LatencyMs      = 0
        ErrorMessage   = $null
    }

    try {
        $conn = [System.Data.SqlClient.SqlConnection]::new($masterConnStr)
        $conn.Open()
        
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = @"
SELECT 
    CAST(SERVERPROPERTY('ServerName') AS NVARCHAR(128)) AS ServerName,
    CAST(SERVERPROPERTY('ProductVersion') AS NVARCHAR(128)) AS ProductVersion,
    CAST(SERVERPROPERTY('Edition') AS NVARCHAR(128)) AS Edition,
    (SELECT COUNT(*) FROM sys.databases WHERE name = @dbName) AS DbExists;
"@
        $cmd.Parameters.AddWithValue("@dbName", $dbName) | Out-Null
        $reader = $cmd.ExecuteReader()
        if ($reader.Read()) {
            $result.Success = $true
            $result.ProductVersion = "$($reader['ProductVersion']) ($($reader['ProductLevel']))"
            $result.DatabaseExists = ([int]$reader['DbExists'] -gt 0)
        }
        $reader.Close()
        $conn.Close()

        if ($result.DatabaseExists) {
            $dbConnStr = Build-RepoConnectionString -RepoCfg $RepoCfg -Database $dbName
            $dbConn = [System.Data.SqlClient.SqlConnection]::new($dbConnStr)
            $dbConn.Open()
            $checkCmd = $dbConn.CreateCommand()
            $checkCmd.CommandText = "SELECT COUNT(*) FROM sys.tables WHERE name IN ('SlaMetricsTrend', 'DatabaseBackupSnapshots')"
            $tblCount = [int]$checkCmd.ExecuteScalar()
            $result.SchemaReady = ($tblCount -ge 2)
            $dbConn.Close()
        }
    } catch {
        $result.Success = $false
        $result.ErrorMessage = $_.Exception.Message
    } finally {
        $sw.Stop()
        $result.LatencyMs = [Math]::Round($sw.Elapsed.TotalMilliseconds, 1)
    }

    return $result
}

function Initialize-RepositorySchema {
    param([Parameter(Mandatory=$true)]$RepoCfg)

    $dbName = if ($RepoCfg.DatabaseName) { $RepoCfg.DatabaseName } else { "SQLBackupMonitorDB" }
    
    try {
        $masterConnStr = Build-RepoConnectionString -RepoCfg $RepoCfg -Database "master"
        $connMaster = [System.Data.SqlClient.SqlConnection]::new($masterConnStr)
        $connMaster.Open()
        
        $createDbCmd = $connMaster.CreateCommand()
        $createDbCmd.CommandText = @"
IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = @dbName)
BEGIN
    DECLARE @sql NVARCHAR(MAX) = N'CREATE DATABASE [' + REPLACE(@dbName, ']', ']]') + N']';
    EXEC sp_executesql @sql;
END

-- Ensure Read Committed Snapshot Isolation (RCSI) is enabled to eliminate reader/writer blocking
IF EXISTS (SELECT 1 FROM sys.databases WHERE name = @dbName AND is_read_committed_snapshot_on = 0)
BEGIN
    DECLARE @rcsiSql NVARCHAR(MAX) = N'ALTER DATABASE [' + REPLACE(@dbName, ']', ']]') + N'] SET READ_COMMITTED_SNAPSHOT ON WITH ROLLBACK IMMEDIATE;';
    BEGIN TRY
        EXEC sp_executesql @rcsiSql;
    END TRY
    BEGIN CATCH
    END CATCH
END
"@
        $createDbCmd.Parameters.AddWithValue("@dbName", $dbName) | Out-Null
        $createDbCmd.ExecuteNonQuery() | Out-Null
        $connMaster.Close()
    } catch {
        return @{ Success = $false; ErrorMessage = "Failed creating database '$dbName': $($_.Exception.Message)" }
    }

    try {
        $dbConnStr = Build-RepoConnectionString -RepoCfg $RepoCfg -Database $dbName
        $connDb = [System.Data.SqlClient.SqlConnection]::new($dbConnStr)
        $connDb.Open()

        $schemaSql = @"
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'SlaMetricsTrend')
BEGIN
    CREATE TABLE SlaMetricsTrend (
        Id BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY CLUSTERED,
        SnapshotTime DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
        TotalServers INT NOT NULL,
        TotalDatabases INT NOT NULL,
        HealthyDatabases INT NOT NULL,
        WarningDatabases INT NOT NULL,
        CriticalDatabases INT NOT NULL,
        CompliancePct FLOAT NOT NULL,
        TotalBackupSizeGB FLOAT NOT NULL
    );
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_SlaMetricsTrend_SnapshotTime' AND object_id = OBJECT_ID('SlaMetricsTrend'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_SlaMetricsTrend_SnapshotTime 
    ON SlaMetricsTrend(SnapshotTime DESC)
    INCLUDE (CompliancePct, TotalDatabases, HealthyDatabases, WarningDatabases, CriticalDatabases, TotalBackupSizeGB);
END

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'DatabaseBackupSnapshots')
BEGIN
    CREATE TABLE DatabaseBackupSnapshots (
        Id BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY CLUSTERED,
        SnapshotTime DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
        ServerId NVARCHAR(100) NULL,
        ServerName NVARCHAR(255) NOT NULL,
        DatabaseName NVARCHAR(255) NOT NULL,
        RecoveryModel NVARCHAR(50) NULL,
        Status NVARCHAR(50) NOT NULL,
        LastFullBackup NVARCHAR(50) NULL,
        LastDifferentialBackup NVARCHAR(50) NULL,
        LastLogBackup NVARCHAR(50) NULL,
        FullBackupAgeHours FLOAT NULL,
        LogBackupAgeMinutes FLOAT NULL,
        LastBackupSizeGB FLOAT NULL,
        StatusReason NVARCHAR(MAX) NULL
    );
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_DbSnapshots_SnapshotTime' AND object_id = OBJECT_ID('DatabaseBackupSnapshots'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_DbSnapshots_SnapshotTime 
    ON DatabaseBackupSnapshots(SnapshotTime DESC);
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_DbSnapshots_Server_Db_Time' AND object_id = OBJECT_ID('DatabaseBackupSnapshots'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_DbSnapshots_Server_Db_Time 
    ON DatabaseBackupSnapshots(ServerName, DatabaseName, SnapshotTime DESC)
    INCLUDE (Status, RecoveryModel, LastFullBackup, LastDifferentialBackup, LastLogBackup, LastBackupSizeGB);
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_DbSnapshots_Status_Time' AND object_id = OBJECT_ID('DatabaseBackupSnapshots'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_DbSnapshots_Status_Time 
    ON DatabaseBackupSnapshots(Status, SnapshotTime DESC)
    INCLUDE (ServerName, DatabaseName, FullBackupAgeHours, StatusReason);
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_DbSnapshots_ServerId' AND object_id = OBJECT_ID('DatabaseBackupSnapshots'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_DbSnapshots_ServerId 
    ON DatabaseBackupSnapshots(ServerId, SnapshotTime DESC);
END
"@
        $cmdSchema = $connDb.CreateCommand()
        $cmdSchema.CommandText = $schemaSql
        $cmdSchema.ExecuteNonQuery() | Out-Null
        $connDb.Close()

        Add-AuditLog -Action "Repository Initialized" -Details "Schema and indexes verified in database $dbName"
        return @{ Success = $true; Message = "Repository database '$dbName' schema and covering indexes initialized successfully." }
    } catch {
        return @{ Success = $false; ErrorMessage = "Failed initializing schema in '$dbName': $($_.Exception.Message)" }
    }
}

function Clear-RepositoryStaleData {
    param(
        [switch]$PurgeAll
    )
    $repoCfg = $Global:Config.RepositorySettings
    if (-not $repoCfg.IsEnabled) {
        return @{ Success = $false; ErrorMessage = "Repository database is not enabled." }
    }

    $dbName = if ($repoCfg.DatabaseName) { $repoCfg.DatabaseName } else { "SQLBackupMonitorDB" }
    try {
        $connStr = Build-RepoConnectionString -RepoCfg $repoCfg -Database $dbName
        $conn = [System.Data.SqlClient.SqlConnection]::new($connStr)
        $conn.Open()

        $cmd = $conn.CreateCommand()
        if ($PurgeAll) {
            $cmd.CommandText = "TRUNCATE TABLE DatabaseBackupSnapshots; TRUNCATE TABLE SlaMetricsTrend;"
            $cmd.ExecuteNonQuery() | Out-Null
            $msg = "All historical snapshots and SLA metrics trends purged successfully."
        } else {
            $activeServerNames = @($Global:Config.Servers | ForEach-Object { "'" + $_.Name.Replace("'", "''") + "'" })
            $serverFilterSql = if ($activeServerNames.Count -gt 0) {
                "AND ServerName NOT IN (" + ($activeServerNames -join ",") + ")"
            } else { "" }

            $retentionDays = if ($repoCfg.RetentionDays) { [int]$repoCfg.RetentionDays } else { 90 }
            $cmd.CommandText = @"
DECLARE @batchSize INT = 5000;
WHILE (1 = 1)
BEGIN
    DELETE TOP (@batchSize) FROM DatabaseBackupSnapshots 
    WHERE SnapshotTime < DATEADD(DAY, -@days, SYSUTCDATETIME()) $serverFilterSql;
    IF @@ROWCOUNT < @batchSize BREAK;
END;
WHILE (1 = 1)
BEGIN
    DELETE TOP (@batchSize) FROM SlaMetricsTrend 
    WHERE SnapshotTime < DATEADD(DAY, -@days, SYSUTCDATETIME());
    IF @@ROWCOUNT < @batchSize BREAK;
END;
"@
            $cmd.Parameters.AddWithValue("@days", $retentionDays) | Out-Null
            $deleted = $cmd.ExecuteNonQuery()
            $msg = "Purged old historical data older than $retentionDays days ($deleted records cleaned)."
        }
        $conn.Close()
        Add-AuditLog -Action "Repository Cleaned" -Details $msg
        return @{ Success = $true; Message = $msg }
    } catch {
        return @{ Success = $false; ErrorMessage = $_.Exception.Message }
    }
}

function Save-RepositoryHistoricalData {
    param(
        [Parameter(Mandatory=$true)]$Metrics,
        [Parameter(Mandatory=$true)]$Databases
    )

    $repoCfg = $Global:Config.RepositorySettings
    if (-not $repoCfg.IsEnabled) { return }

    $dbName = if ($repoCfg.DatabaseName) { $repoCfg.DatabaseName } else { "SQLBackupMonitorDB" }

    try {
        $connStr = Build-RepoConnectionString -RepoCfg $repoCfg -Database $dbName
        $conn = [System.Data.SqlClient.SqlConnection]::new($connStr)
        $conn.Open()

        # Insert Metrics Trend
        $cmdTrend = $conn.CreateCommand()
        $cmdTrend.CommandText = @"
INSERT INTO SlaMetricsTrend (SnapshotTime, TotalServers, TotalDatabases, HealthyDatabases, WarningDatabases, CriticalDatabases, CompliancePct, TotalBackupSizeGB)
VALUES (SYSUTCDATETIME(), @servers, @dbs, @healthy, @warn, @crit, @compliance, @size);
"@
        $cmdTrend.Parameters.AddWithValue("@servers", [int]$Metrics.TotalServers) | Out-Null
        $cmdTrend.Parameters.AddWithValue("@dbs", [int]$Metrics.TotalDatabases) | Out-Null
        $cmdTrend.Parameters.AddWithValue("@healthy", [int]$Metrics.HealthyDatabases) | Out-Null
        $cmdTrend.Parameters.AddWithValue("@warn", [int]$Metrics.WarningDatabases) | Out-Null
        $cmdTrend.Parameters.AddWithValue("@crit", [int]$Metrics.CriticalDatabases) | Out-Null
        $cmdTrend.Parameters.AddWithValue("@compliance", [double]$Metrics.CompliancePct) | Out-Null
        $cmdTrend.Parameters.AddWithValue("@size", [double]$Metrics.TotalBackupSizeGB) | Out-Null
        $cmdTrend.ExecuteNonQuery() | Out-Null

        # Fast Bulk Insert using SqlBulkCopy for thousands of databases in < 100ms
        if ($Databases -and $Databases.Count -gt 0) {
            $dt = [System.Data.DataTable]::new()
            [void]$dt.Columns.Add("SnapshotTime", [DateTime])
            [void]$dt.Columns.Add("ServerId", [string])
            [void]$dt.Columns.Add("ServerName", [string])
            [void]$dt.Columns.Add("DatabaseName", [string])
            [void]$dt.Columns.Add("RecoveryModel", [string])
            [void]$dt.Columns.Add("Status", [string])
            [void]$dt.Columns.Add("LastFullBackup", [string])
            [void]$dt.Columns.Add("LastDifferentialBackup", [string])
            [void]$dt.Columns.Add("LastLogBackup", [string])
            [void]$dt.Columns.Add("FullBackupAgeHours", [double])
            [void]$dt.Columns.Add("LogBackupAgeMinutes", [double])
            [void]$dt.Columns.Add("LastBackupSizeGB", [double])
            [void]$dt.Columns.Add("StatusReason", [string])

            $nowUtc = [DateTime]::UtcNow
            foreach ($db in $Databases) {
                if ($null -eq $db) { continue }
                $row = $dt.NewRow()
                $row["SnapshotTime"]           = $nowUtc
                $row["ServerId"]               = if ($db.ServerId) { $db.ServerId } else { "" }
                $row["ServerName"]             = if ($db.ServerName) { $db.ServerName } else { "Unknown" }
                $row["DatabaseName"]           = if ($db.DatabaseName) { $db.DatabaseName } else { "Unknown" }
                $row["RecoveryModel"]          = if ($db.RecoveryModel) { $db.RecoveryModel } else { "SIMPLE" }
                $row["Status"]                 = if ($db.Status) { $db.Status } else { "Healthy" }
                $row["LastFullBackup"]         = if ($db.LastFullBackup) { $db.LastFullBackup } else { "Never" }
                $row["LastDifferentialBackup"] = if ($db.LastDifferentialBackup) { $db.LastDifferentialBackup } else { "Never" }
                $row["LastLogBackup"]          = if ($db.LastLogBackup) { $db.LastLogBackup } else { "Never" }
                $row["FullBackupAgeHours"]     = if ($db.FullBackupAgeHours) { [double]$db.FullBackupAgeHours } else { -1.0 }
                $row["LogBackupAgeMinutes"]    = if ($db.LogBackupAgeMinutes) { [double]$db.LogBackupAgeMinutes } else { -1.0 }
                $row["LastBackupSizeGB"]       = if ($db.LastBackupSizeGB) { [double]$db.LastBackupSizeGB } else { 0.0 }
                $row["StatusReason"]           = if ($db.StatusReason) { $db.StatusReason } else { "" }
                $dt.Rows.Add($row)
            }

            $bulkCopy = [System.Data.SqlClient.SqlBulkCopy]::new($conn)
            $bulkCopy.DestinationTableName = "DatabaseBackupSnapshots"
            $bulkCopy.BulkCopyTimeout = 30
            $bulkCopy.BatchSize = 1000

            $bulkCopy.ColumnMappings.Add("SnapshotTime", "SnapshotTime") | Out-Null
            $bulkCopy.ColumnMappings.Add("ServerId", "ServerId") | Out-Null
            $bulkCopy.ColumnMappings.Add("ServerName", "ServerName") | Out-Null
            $bulkCopy.ColumnMappings.Add("DatabaseName", "DatabaseName") | Out-Null
            $bulkCopy.ColumnMappings.Add("RecoveryModel", "RecoveryModel") | Out-Null
            $bulkCopy.ColumnMappings.Add("Status", "Status") | Out-Null
            $bulkCopy.ColumnMappings.Add("LastFullBackup", "LastFullBackup") | Out-Null
            $bulkCopy.ColumnMappings.Add("LastDifferentialBackup", "LastDifferentialBackup") | Out-Null
            $bulkCopy.ColumnMappings.Add("LastLogBackup", "LastLogBackup") | Out-Null
            $bulkCopy.ColumnMappings.Add("FullBackupAgeHours", "FullBackupAgeHours") | Out-Null
            $bulkCopy.ColumnMappings.Add("LogBackupAgeMinutes", "LogBackupAgeMinutes") | Out-Null
            $bulkCopy.ColumnMappings.Add("LastBackupSizeGB", "LastBackupSizeGB") | Out-Null
            $bulkCopy.ColumnMappings.Add("StatusReason", "StatusReason") | Out-Null

            $bulkCopy.WriteToServer($dt)
            $bulkCopy.Close()
        }

        # Retention Cleanup (Batched to prevent lock escalation)
        $retentionDays = if ($repoCfg.RetentionDays) { [int]$repoCfg.RetentionDays } else { 90 }
        $cmdPurge = $conn.CreateCommand()
        $cmdPurge.CommandText = @"
DECLARE @batchSize INT = 5000;
WHILE (1 = 1)
BEGIN
    DELETE TOP (@batchSize) FROM DatabaseBackupSnapshots 
    WHERE SnapshotTime < DATEADD(DAY, -@days, SYSUTCDATETIME());
    IF @@ROWCOUNT < @batchSize BREAK;
END;
WHILE (1 = 1)
BEGIN
    DELETE TOP (@batchSize) FROM SlaMetricsTrend 
    WHERE SnapshotTime < DATEADD(DAY, -@days, SYSUTCDATETIME());
    IF @@ROWCOUNT < @batchSize BREAK;
END;
"@
        $cmdPurge.Parameters.AddWithValue("@days", $retentionDays) | Out-Null
        $cmdPurge.ExecuteNonQuery() | Out-Null

        $conn.Close()
    } catch {
        Write-Warning "Failed to persist historical telemetry to repository database: $($_.Exception.Message)"
    }
}

function Get-RepositoryHistoricalTrends {
    param([int]$Days = 7)

    $repoCfg = $Global:Config.RepositorySettings
    if (-not $repoCfg.IsEnabled) { return @() }

    $dbName = if ($repoCfg.DatabaseName) { $repoCfg.DatabaseName } else { "SQLBackupMonitorDB" }
    $trends = @()

    try {
        $connStr = Build-RepoConnectionString -RepoCfg $repoCfg -Database $dbName
        $conn = [System.Data.SqlClient.SqlConnection]::new($connStr)
        $conn.Open()

        $cmd = $conn.CreateCommand()
        $cmd.CommandText = @"
SELECT TOP 100
    CONVERT(VARCHAR(19), SnapshotTime, 120) AS SnapshotTime,
    TotalServers,
    TotalDatabases,
    HealthyDatabases,
    WarningDatabases,
    CriticalDatabases,
    CompliancePct,
    TotalBackupSizeGB
FROM SlaMetricsTrend WITH (NOLOCK)
WHERE SnapshotTime >= DATEADD(DAY, -@days, SYSUTCDATETIME())
ORDER BY SnapshotTime ASC;
"@
        $cmd.Parameters.AddWithValue("@days", $Days) | Out-Null
        $reader = $cmd.ExecuteReader()
        while ($reader.Read()) {
            $trends += [PSCustomObject]@{
                SnapshotTime      = [string]$reader["SnapshotTime"]
                TotalServers      = [int]$reader["TotalServers"]
                TotalDatabases    = [int]$reader["TotalDatabases"]
                HealthyDatabases  = [int]$reader["HealthyDatabases"]
                WarningDatabases  = [int]$reader["WarningDatabases"]
                CriticalDatabases = [int]$reader["CriticalDatabases"]
                CompliancePct     = [double]$reader["CompliancePct"]
                TotalBackupSizeGB = [double]$reader["TotalBackupSizeGB"]
            }
        }
        $reader.Close()
        $conn.Close()
    } catch {
        Write-Warning "Failed to retrieve historical trends: $($_.Exception.Message)"
    }

    return $trends
}

function Get-RepositoryHistoricalSnapshots {
    param(
        [int]$Limit = 50,
        [string]$Status = "",
        [string]$Search = ""
    )

    $repoCfg = $Global:Config.RepositorySettings
    if (-not $repoCfg.IsEnabled) { return @() }

    $dbName = if ($repoCfg.DatabaseName) { $repoCfg.DatabaseName } else { "SQLBackupMonitorDB" }
    $snapshots = @()

    try {
        $connStr = Build-RepoConnectionString -RepoCfg $repoCfg -Database $dbName
        $conn = [System.Data.SqlClient.SqlConnection]::new($connStr)
        $conn.Open()

        $cmd = $conn.CreateCommand()
        $cmd.CommandText = @"
SELECT TOP (@limit)
    CONVERT(VARCHAR(19), SnapshotTime, 120) AS SnapshotTime,
    ServerName,
    DatabaseName,
    RecoveryModel,
    Status,
    LastFullBackup,
    LastDifferentialBackup,
    LastLogBackup,
    LastBackupSizeGB,
    StatusReason
FROM DatabaseBackupSnapshots WITH (NOLOCK)
WHERE (@status = '' OR Status = @status)
  AND (@search = '' OR ServerName LIKE '%' + @search + '%' OR DatabaseName LIKE '%' + @search + '%')
ORDER BY SnapshotTime DESC, Id DESC;
"@
        $cmd.Parameters.AddWithValue("@limit", $Limit) | Out-Null
        $cmd.Parameters.AddWithValue("@status", $(if ($Status) { $Status } else { "" })) | Out-Null
        $cmd.Parameters.AddWithValue("@search", $(if ($Search) { $Search } else { "" })) | Out-Null

        $reader = $cmd.ExecuteReader()
        while ($reader.Read()) {
            $snapshots += [PSCustomObject]@{
                SnapshotTime           = [string]$reader["SnapshotTime"]
                ServerName             = [string]$reader["ServerName"]
                DatabaseName           = [string]$reader["DatabaseName"]
                RecoveryModel          = [string]$reader["RecoveryModel"]
                Status                 = [string]$reader["Status"]
                LastFullBackup         = [string]$reader["LastFullBackup"]
                LastDifferentialBackup = [string]$reader["LastDifferentialBackup"]
                LastLogBackup          = [string]$reader["LastLogBackup"]
                LastBackupSizeGB       = [double]$reader["LastBackupSizeGB"]
                StatusReason           = [string]$reader["StatusReason"]
            }
        }
        $reader.Close()
        $conn.Close()
    } catch {
        Write-Warning "Failed to retrieve historical snapshots: $($_.Exception.Message)"
    }

    return $snapshots
}

# ==============================================================================
# NATIVE PDF AUDIT REPORT GENERATOR (100% PURE POWERSHELL - ZERO EXTERNAL EXEs)
# ==============================================================================
# NATIVE PDF AUDIT REPORT GENERATOR (100% PURE POWERSHELL - ZERO EXTERNAL EXEs)
# ==============================================================================
function Export-BackupPdfReport {
    param([string]$OutputPath = "$PSScriptRoot\BackupReport.pdf")

    $metrics = $Global:CachedMetrics
    $results = $Global:CachedResults
    $generatedTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

    # Clean string helper for PDF
    function Escape-PdfText([string]$text) {
        if ([string]::IsNullOrEmpty($text)) { return "" }
        return $text.Replace("\", "\\").Replace("(", "\(").Replace(")", "\)")
    }

    # Safe text truncator helper to prevent text overflow in PDF tables
    function Truncate-PdfText([string]$text, [int]$maxLen) {
        if ([string]::IsNullOrWhiteSpace($text)) { return "" }
        $t = $text.ToString().Trim()
        if ($t.Length -gt $maxLen) {
            if ($maxLen -gt 3) {
                return $t.Substring(0, $maxLen - 3) + "..."
            } else {
                return $t.Substring(0, $maxLen)
            }
        }
        return $t
    }

    # Dynamic row capacity per page (Landscape: 792 x 612 pt)
    # Page 1 contains top header + KPI summary box + table header => fits 21 data rows
    # Pages 2+ contain top header + table header => fits 25 data rows
    $totalRows = $results.Count
    $totalPages = 1
    if ($totalRows -gt 21) {
        $totalPages = 1 + [Math]::Max(1, [int][Math]::Ceiling(($totalRows - 21) / 25.0))
    }

    $pdfObjects = [System.Collections.Generic.List[string]]::new()
    $pageObjectIds = [System.Collections.Generic.List[int]]::new()

    # Base PDF Objects
    # Object 1: Catalog
    $pdfObjects.Add("1 0 obj`n<< /Type /Catalog /Pages 2 0 R >>`nendobj")
    # Object 2: Pages (Will be set after pages created)
    $pdfObjects.Add("") # Placeholder for Pages obj
    # Object 3: Helvetica
    $pdfObjects.Add("3 0 obj`n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>`nendobj")
    # Object 4: Helvetica-Bold
    $pdfObjects.Add("4 0 obj`n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>`nendobj")
    # Object 5: Courier
    $pdfObjects.Add("5 0 obj`n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>`nendobj")

    $currentObjId = 6
    $rowIdx = 0

    for ($p = 0; $p -lt $totalPages; $p++) {
        $pageObjId = $currentObjId++
        $contentObjId = $currentObjId++
        $pageObjectIds.Add($pageObjId)

        $pageCapacity = if ($p -eq 0) { 21 } else { 25 }
        $pageRows = @()
        if ($totalRows -gt 0 -and $rowIdx -lt $totalRows) {
            $takeCount = [Math]::Min($pageCapacity, $totalRows - $rowIdx)
            $endIdx = $rowIdx + $takeCount - 1
            $pageRows = $results[$rowIdx..$endIdx]
            $rowIdx += $takeCount
        }

        # Build Stream Content for this page (Landscape 792 x 612 pt)
        $sb = [System.Text.StringBuilder]::new()

        # Background Canvas: Dark navy #0b0f19
        [void]$sb.AppendLine("0.043 0.059 0.098 rg 0 0 792 612 re f")

        # Top Header Bar: Gradient look #1e293b
        [void]$sb.AppendLine("0.118 0.161 0.231 rg 0 546 792 66 re f")
        [void]$sb.AppendLine("0.231 0.510 0.965 RG 2 w 0 546 m 792 546 l S")

        # Header Title & Meta
        [void]$sb.AppendLine("BT /F2 15 Tf 0.95 0.96 0.98 rg 36 584 Td (SQLBackupMonitor - Enterprise Backup SLA Report) Tj ET")
        [void]$sb.AppendLine("BT /F1 9 Tf 0.58 0.64 0.72 rg 36 564 Td (Generated: $(Escape-PdfText $generatedTime)  |  SLA Compliance: $($metrics.CompliancePct)%  |  Total Monitored DBs: $($metrics.TotalDatabases)) Tj ET")
        [void]$sb.AppendLine("BT /F1 9 Tf 0.58 0.64 0.72 rg 680 584 Td (Page $($p+1) of $totalPages) Tj ET")

        $curY = 530

        # If First Page: Draw Summary KPI Cards
        if ($p -eq 0) {
            # KPI Container Box (Width 720, from X=36 to X=756)
            [void]$sb.AppendLine("0.071 0.094 0.149 rg 36 470 720 52 re f")
            [void]$sb.AppendLine("0.18 0.22 0.31 RG 1 w 36 470 720 52 re S")

            # KPI 1: Servers
            [void]$sb.AppendLine("BT /F1 8 Tf 0.58 0.64 0.72 rg 52 506 Td (SERVERS) Tj ET")
            [void]$sb.AppendLine("BT /F2 14 Tf 0.22 0.74 0.97 rg 52 485 Td ($($metrics.TotalServers)) Tj ET")

            # KPI 2: Databases
            [void]$sb.AppendLine("BT /F1 8 Tf 0.58 0.64 0.72 rg 165 506 Td (DATABASES) Tj ET")
            [void]$sb.AppendLine("BT /F2 14 Tf 0.95 0.96 0.98 rg 165 485 Td ($($metrics.TotalDatabases)) Tj ET")

            # KPI 3: Healthy
            [void]$sb.AppendLine("BT /F1 8 Tf 0.58 0.64 0.72 rg 285 506 Td (HEALTHY) Tj ET")
            [void]$sb.AppendLine("BT /F2 14 Tf 0.20 0.83 0.60 rg 285 485 Td ($($metrics.HealthyDatabases)) Tj ET")

            # KPI 4: Warning
            [void]$sb.AppendLine("BT /F1 8 Tf 0.58 0.64 0.72 rg 405 506 Td (WARNINGS) Tj ET")
            [void]$sb.AppendLine("BT /F2 14 Tf 0.96 0.62 0.14 rg 405 485 Td ($($metrics.WarningDatabases)) Tj ET")

            # KPI 5: Critical
            [void]$sb.AppendLine("BT /F1 8 Tf 0.58 0.64 0.72 rg 525 506 Td (CRITICAL) Tj ET")
            [void]$sb.AppendLine("BT /F2 14 Tf 0.94 0.27 0.27 rg 525 485 Td ($($metrics.CriticalDatabases)) Tj ET")

            # KPI 6: Volume
            [void]$sb.AppendLine("BT /F1 8 Tf 0.58 0.64 0.72 rg 640 506 Td (TOTAL VOLUME) Tj ET")
            [void]$sb.AppendLine("BT /F2 12 Tf 0.65 0.55 0.98 rg 640 485 Td ($($metrics.TotalBackupSizeGB) GB) Tj ET")

            $curY = 456
        }

        # Table Header Bar (Width 720, from X=36 to X=756)
        [void]$sb.AppendLine("0.118 0.161 0.231 rg 36 $($curY - 18) 720 18 re f")
        [void]$sb.AppendLine("0.20 0.26 0.36 RG 1 w 36 $($curY - 18) 720 18 re S")

        [void]$sb.AppendLine("BT /F2 8 Tf 0.70 0.76 0.85 rg")
        [void]$sb.AppendLine("44 $($curY - 12) Td (STATUS) Tj")
        [void]$sb.AppendLine("95 $($curY - 12) Td (SQL SERVER INSTANCE) Tj")
        [void]$sb.AppendLine("235 $($curY - 12) Td (DATABASE NAME) Tj")
        [void]$sb.AppendLine("375 $($curY - 12) Td (REC) Tj")
        [void]$sb.AppendLine("418 $($curY - 12) Td (LAST FULL BACKUP) Tj")
        [void]$sb.AppendLine("520 $($curY - 12) Td (SIZE) Tj")
        [void]$sb.AppendLine("570 $($curY - 12) Td (SLA DIAGNOSTIC REASON) Tj")
        [void]$sb.AppendLine("ET")

        $curY -= 20

        # Table Rows
        foreach ($row in $pageRows) {
            $rowH = 17
            # Alternating row background
            [void]$sb.AppendLine("0.063 0.086 0.137 rg 36 $($curY - $rowH) 720 $rowH re f")
            [void]$sb.AppendLine("0.12 0.16 0.24 RG 0.5 w 36 $($curY - $rowH) 720 $rowH re S")

            # Status Badge Pill
            $badgeColor = if ($row.Status -eq "Critical") { "0.94 0.27 0.27" } elseif ($row.Status -eq "Warning") { "0.96 0.62 0.14" } else { "0.06 0.72 0.50" }
            [void]$sb.AppendLine("$badgeColor rg 42 $($curY - $rowH + 3) 44 11 re f")
            $statusStr = Escape-PdfText (Truncate-PdfText $row.Status.ToUpper() 8)
            [void]$sb.AppendLine("BT /F2 6.5 Tf 1 1 1 rg 46 $($curY - $rowH + 5.5) Td ($statusStr) Tj ET")

            # Truncated text columns to completely prevent PDF text overflow
            $srvText = Escape-PdfText (Truncate-PdfText $row.ServerName 20)
            $dbText = Escape-PdfText (Truncate-PdfText $row.DatabaseName 22)
            $recText = Escape-PdfText (Truncate-PdfText $row.RecoveryModel 6)
            $fullText = Escape-PdfText (Truncate-PdfText $row.LastFullBackup 19)
            $sizeText = Escape-PdfText (Truncate-PdfText "$($row.LastBackupSizeGB) GB" 8)
            $reasonText = Escape-PdfText (Truncate-PdfText $row.StatusReason 36)

            [void]$sb.AppendLine("BT /F2 7.5 Tf 0.95 0.96 0.98 rg 95 $($curY - $rowH + 5.5) Td ($srvText) Tj ET")
            [void]$sb.AppendLine("BT /F1 7.5 Tf 0.22 0.74 0.97 rg 235 $($curY - $rowH + 5.5) Td ($dbText) Tj ET")
            [void]$sb.AppendLine("BT /F3 7 Tf 0.70 0.76 0.85 rg 375 $($curY - $rowH + 5.5) Td ($recText) Tj ET")
            [void]$sb.AppendLine("BT /F3 7 Tf 0.80 0.85 0.92 rg 418 $($curY - $rowH + 5.5) Td ($fullText) Tj ET")
            [void]$sb.AppendLine("BT /F3 7 Tf 0.70 0.76 0.85 rg 520 $($curY - $rowH + 5.5) Td ($sizeText) Tj ET")

            $reasonColor = if ($row.Status -eq "Critical") { "0.98 0.55 0.55" } elseif ($row.Status -eq "Warning") { "0.98 0.80 0.35" } else { "0.60 0.68 0.78" }
            [void]$sb.AppendLine("BT /F1 6.5 Tf $reasonColor rg 570 $($curY - $rowH + 5.5) Td ($reasonText) Tj ET")

            $curY -= $rowH
        }

        # Footer Bar
        [void]$sb.AppendLine("BT /F1 7.5 Tf 0.45 0.50 0.60 rg 36 22 Td (SQLBackupMonitor Standalone PowerShell Platform  |  Zero External Binaries Required  |  Official SLA Audit Snapshot) Tj ET")

        $streamBytes = [System.Text.Encoding]::ASCII.GetBytes($sb.ToString())
        $streamLen = $streamBytes.Length

        # Add Page Object & Content Object (MediaBox: Landscape 792 x 612 pt)
        $pdfObjects.Add("$pageObjId 0 obj`n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 792 612] /Contents $contentObjId 0 R /Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R >> >> >>`nendobj")
        $pdfObjects.Add("$contentObjId 0 obj`n<< /Length $streamLen >>`nstream`n$($sb.ToString())`nendstream`nendobj")
    }

    # Now define Pages Object (Object 2)
    $kidsStr = ($pageObjectIds | ForEach-Object { "$_ 0 R" }) -join " "
    $pdfObjects[1] = "2 0 obj`n<< /Type /Pages /Kids [ $kidsStr ] /Count $totalPages >>`nendobj"

    # Assemble complete PDF Binary Stream with precise XRef Table
    $ms = [System.IO.MemoryStream]::new()
    $writer = [System.IO.StreamWriter]::new($ms, [System.Text.Encoding]::ASCII)
    $writer.NewLine = "`n"

    $writer.WriteLine("%PDF-1.4")
    $writer.Flush()

    $offsets = [System.Collections.Generic.List[long]]::new()
    $offsets.Add(0) # Object 0

    for ($i = 0; $i -lt $pdfObjects.Count; $i++) {
        $writer.Flush()
        $offsets.Add($ms.Position)
        $writer.WriteLine($pdfObjects[$i])
        $writer.Flush()
    }

    $xrefPos = $ms.Position
    $writer.WriteLine("xref")
    $writer.WriteLine("0 $($pdfObjects.Count + 1)")
    $writer.WriteLine("0000000000 65535 f ")
    for ($i = 1; $i -le $pdfObjects.Count; $i++) {
        $off = $offsets[$i]
        $writer.WriteLine(("{0:D10} 00000 n " -f $off))
    }
    $writer.WriteLine("trailer")
    $writer.WriteLine("<< /Size $($pdfObjects.Count + 1) /Root 1 0 R >>")
    $writer.WriteLine("startxref")
    $writer.WriteLine("$xrefPos")
    $writer.WriteLine("%%EOF")
    $writer.Flush()

    # Write to target destination
    [System.IO.File]::WriteAllBytes($OutputPath, $ms.ToArray())
    $writer.Close()
    $ms.Close()

    Add-AuditLog -Action "PDF Generated" -Details "Generated native PDF audit report at $OutputPath"
    return $OutputPath
}

# ==============================================================================
# EMAIL NOTIFICATIONS & SCHEDULED REPORTS WITH PDF ATTACHMENT
# ==============================================================================
function Send-SlaEmailAlert {
    param(
        [Parameter(Mandatory=$true)]$BreachedDatabases,
        [switch]$IsTest
    )

    $emailCfg = $Global:Config.EmailSettings
    if (-not $emailCfg.IsEnabled -and -not $IsTest) { return }

    $now = [DateTime]::UtcNow
    $dbsToAlert = @()

    foreach ($db in $BreachedDatabases) {
        $key = "$($db.ServerId):$($db.DatabaseName):$($db.Status)"
        $cooldown = if ($emailCfg.CooldownMinutes) { [int]$emailCfg.CooldownMinutes } else { 60 }

        if ($Global:LastAlertSentTimes.ContainsKey($key)) {
            $lastSent = $Global:LastAlertSentTimes[$key]
            if (($now - $lastSent).TotalMinutes -lt $cooldown -and -not $IsTest) {
                continue
            }
        }

        $dbsToAlert += $db
        $Global:LastAlertSentTimes[$key] = $now
    }

    if ($dbsToAlert.Count -eq 0 -and -not $IsTest) {
        return @{ Success = $true; Skipped = $true; Message = "All breached items are within the alert cooldown window." }
    }

    try {
        $mail = [System.Net.Mail.MailMessage]::new()
        $from = if ($emailCfg.SenderDisplayName) {
            [System.Net.Mail.MailAddress]::new($emailCfg.SenderEmail, $emailCfg.SenderDisplayName)
        } else {
            [System.Net.Mail.MailAddress]::new($emailCfg.SenderEmail)
        }
        $mail.From = $from

        # Add recipients
        $recipients = $emailCfg.RecipientEmails -split "[,;]" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        foreach ($r in $recipients) {
            $mail.To.Add($r.Trim())
        }

        $critCount = @($dbsToAlert | Where-Object { $_.Status -eq "Critical" }).Count
        $warnCount = @($dbsToAlert | Where-Object { $_.Status -eq "Warning" }).Count

        $subjectPrefix = if ($IsTest) { "[TEST ALERT] " } elseif ($critCount -gt 0) { "[CRITICAL SLA BREACH] " } else { "[WARNING SLA ALERT] " }
        $mail.Subject = "$subjectPrefix SQLBackupMonitor: $critCount Critical, $warnCount Warning Database(s)"

        # Generate sleek HTML body
        $rowsHtml = ($dbsToAlert | ForEach-Object {
            $badgeBg = if ($_.Status -eq "Critical") { "#ef4444" } else { "#f59e0b" }
            @"
            <tr>
                <td style="padding:10px 14px; border-bottom:1px solid #334155;"><span style="background:${badgeBg}; color:#ffffff; padding:3px 8px; border-radius:4px; font-weight:bold; font-size:11px;">$($_.Status)</span></td>
                <td style="padding:10px 14px; border-bottom:1px solid #334155; font-weight:bold; color:#f8fafc;">$([System.Web.HttpUtility]::HtmlEncode($_.ServerName))</td>
                <td style="padding:10px 14px; border-bottom:1px solid #334155; color:#38bdf8;">$([System.Web.HttpUtility]::HtmlEncode($_.DatabaseName))</td>
                <td style="padding:10px 14px; border-bottom:1px solid #334155; font-family:monospace; color:#cbd5e1;">$($_.RecoveryModel)</td>
                <td style="padding:10px 14px; border-bottom:1px solid #334155; font-family:monospace; color:#cbd5e1;">$($_.LastFullBackup)</td>
                <td style="padding:10px 14px; border-bottom:1px solid #334155; font-family:monospace; color:#cbd5e1;">$($_.LastLogBackup)</td>
                <td style="padding:10px 14px; border-bottom:1px solid #334155; color:#fca5a5; font-size:12px;">$([System.Web.HttpUtility]::HtmlEncode($_.StatusReason))</td>
            </tr>
"@
        }) -join "`n"

        $bodyHtml = @"
<!DOCTYPE html>
<html>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background-color: #0b0f19; color: #f3f4f6; margin:0; padding: 24px;">
    <div style="max-width: 900px; margin: 0 auto; background-color: #121826; border: 1px solid #1e293b; border-radius: 12px; overflow: hidden; box-shadow: 0 10px 25px rgba(0,0,0,0.5);">
        <div style="background: linear-gradient(135deg, #1e293b, #0f172a); padding: 20px 24px; border-bottom: 2px solid #ef4444;">
            <h2 style="margin: 0; color: #f8fafc; font-size: 20px; display: flex; align-items: center; gap: 8px;">
                <span style="color:#ef4444; font-size:22px;">[!]</span> SQL Server Backup SLA Alert
            </h2>
            <p style="margin: 6px 0 0 0; color: #94a3b8; font-size: 13px;">Generated automatically by SQLBackupMonitor PowerShell Platform at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        </div>
        <div style="padding: 24px;">
            <p style="margin-top:0; font-size: 14px; color: #e2e8f0;">The following database backups have breached configured Recovery Point Objectives (RPO) / SLA policy thresholds:</p>
            <table style="width: 100%; border-collapse: collapse; background-color: #0f172a; border-radius: 8px; font-size: 13px; text-align: left;">
                <thead>
                    <tr style="background-color: #1e293b; color: #94a3b8; text-transform: uppercase; font-size: 11px;">
                        <th style="padding: 10px 14px;">Status</th>
                        <th style="padding: 10px 14px;">Server</th>
                        <th style="padding: 10px 14px;">Database</th>
                        <th style="padding: 10px 14px;">Recovery</th>
                        <th style="padding: 10px 14px;">Last Full</th>
                        <th style="padding: 10px 14px;">Last Log</th>
                        <th style="padding: 10px 14px;">Violation Details</th>
                    </tr>
                </thead>
                <tbody>
                    $rowsHtml
                </tbody>
            </table>
            <div style="margin-top: 24px; padding: 14px; background-color: #1e293b; border-radius: 6px; font-size: 12px; color: #94a3b8;">
                <strong>Action Required:</strong> Check SQL Agent backup jobs or execute immediate ad-hoc backups on target SQL instances.
            </div>
        </div>
    </div>
</body>
</html>
"@
        $mail.Body = $bodyHtml
        $mail.IsBodyHtml = $true

        # Attach Native PDF Report if configured
        $tempPdf = $null
        if ($emailCfg.AttachPdfReport -ne $false) {
            $tempPdf = [System.IO.Path]::GetTempFileName() + ".pdf"
            Export-BackupPdfReport -OutputPath $tempPdf | Out-Null
            if (Test-Path $tempPdf) {
                $att = [System.Net.Mail.Attachment]::new($tempPdf, "application/pdf")
                $att.Name = "SQLBackup_AuditReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').pdf"
                $mail.Attachments.Add($att)
            }
        }

        $client = [System.Net.Mail.SmtpClient]::new($emailCfg.SmtpServer, [int]$emailCfg.SmtpPort)
        $client.EnableSsl = [bool]$emailCfg.EnableSsl
        $client.Timeout = 15000

        if (-not [string]::IsNullOrWhiteSpace($emailCfg.SmtpUsername)) {
            $client.Credentials = [System.Net.NetworkCredential]::new($emailCfg.SmtpUsername, $emailCfg.SmtpPassword)
        }

        $client.Send($mail)
        $mail.Dispose()
        $client.Dispose()

        if ($tempPdf -and (Test-Path $tempPdf)) {
            Remove-Item $tempPdf -Force -ErrorAction SilentlyContinue
        }

        Write-Host "[OK] SLA breach email alert with PDF attachment sent to $($emailCfg.RecipientEmails)" -ForegroundColor Green
        Add-AuditLog -Action "Email Sent" -Details "SLA breach notification (PDF attached) sent to $($emailCfg.RecipientEmails)"
        return @{ Success = $true; Message = "Email alert sent successfully with attached PDF report." }
    } catch {
        Write-Warning "Failed to send SLA email alert: $_"
        return @{ Success = $false; ErrorMessage = $_.Exception.Message }
    }
}

function Send-DailyPdfEmailReport {
    param(
        [string]$RecipientOverride = ""
    )

    $emailCfg = $Global:Config.EmailSettings
    $recipients = if ($RecipientOverride) { $RecipientOverride } else { $emailCfg.RecipientEmails }

    if ([string]::IsNullOrWhiteSpace($recipients) -or [string]::IsNullOrWhiteSpace($emailCfg.SmtpServer)) {
        return @{ Success = $false; ErrorMessage = "SMTP Server or Recipient Email(s) not configured." }
    }

    try {
        Invoke-FullScan

        $mail = [System.Net.Mail.MailMessage]::new()
        $sender = if ($emailCfg.SenderEmail) { $emailCfg.SenderEmail } else { "sqlbackupmonitor@corp.local" }
        $displayName = if ($emailCfg.SenderDisplayName) { $emailCfg.SenderDisplayName } else { "SQLBackupMonitor Daily Digest" }
        $mail.From = [System.Net.Mail.MailAddress]::new($sender, $displayName)

        foreach ($r in ($recipients -split "[,;]")) {
            $cleaned = $r.Trim()
            if ($cleaned) { $mail.To.Add($cleaned) }
        }

        $m = $Global:CachedMetrics
        $mail.Subject = "[DAILY EXECUTIVE REPORT] SQLBackupMonitor: $($m.CompliancePct)% SLA Compliance ($($m.TotalDatabases) Databases)"

        $rowsHtml = ($Global:CachedResults | ForEach-Object {
            $badgeBg = if ($_.Status -eq "Critical") { "#ef4444" } elseif ($_.Status -eq "Warning") { "#f59e0b" } else { "#10b981" }
            @"
            <tr>
                <td style="padding:8px 12px; border-bottom:1px solid #334155;"><span style="background:${badgeBg}; color:#ffffff; padding:2px 7px; border-radius:4px; font-weight:bold; font-size:11px;">$($_.Status)</span></td>
                <td style="padding:8px 12px; border-bottom:1px solid #334155; font-weight:bold; color:#f8fafc;">$([System.Web.HttpUtility]::HtmlEncode($_.ServerName))</td>
                <td style="padding:8px 12px; border-bottom:1px solid #334155; color:#38bdf8;">$([System.Web.HttpUtility]::HtmlEncode($_.DatabaseName))</td>
                <td style="padding:8px 12px; border-bottom:1px solid #334155; font-family:monospace; color:#cbd5e1;">$($_.RecoveryModel)</td>
                <td style="padding:8px 12px; border-bottom:1px solid #334155; font-family:monospace; color:#cbd5e1;">$($_.LastFullBackup)</td>
                <td style="padding:8px 12px; border-bottom:1px solid #334155; font-family:monospace; color:#cbd5e1;">$($_.LastLogBackup)</td>
                <td style="padding:8px 12px; border-bottom:1px solid #334155; color:#94a3b8; font-size:12px;">$([System.Web.HttpUtility]::HtmlEncode($_.StatusReason))</td>
            </tr>
"@
        }) -join "`n"

        $bodyHtml = @"
<!DOCTYPE html>
<html>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background-color: #0b0f19; color: #f3f4f6; margin:0; padding: 24px;">
    <div style="max-width: 900px; margin: 0 auto; background-color: #121826; border: 1px solid #1e293b; border-radius: 12px; overflow: hidden;">
        <div style="background: linear-gradient(135deg, #1e293b, #0f172a); padding: 20px 24px; border-bottom: 2px solid #3b82f6;">
            <h2 style="margin: 0; color: #f8fafc; font-size: 20px;">Daily SQL Server Backup Executive Audit</h2>
            <p style="margin: 6px 0 0 0; color: #94a3b8; font-size: 13px;">Generated at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') &bull; Executive PDF Report Attached</p>
        </div>
        <div style="padding: 24px;">
            <div style="display:flex; gap:16px; margin-bottom:20px;">
                <div style="flex:1; background:#1e293b; padding:12px 16px; border-radius:8px; text-align:center;">
                    <div style="color:#94a3b8; font-size:11px; text-transform:uppercase;">Compliance</div>
                    <div style="color:#38bdf8; font-size:22px; font-weight:bold;">$($m.CompliancePct)%</div>
                </div>
                <div style="flex:1; background:#1e293b; padding:12px 16px; border-radius:8px; text-align:center;">
                    <div style="color:#94a3b8; font-size:11px; text-transform:uppercase;">Healthy DBs</div>
                    <div style="color:#10b981; font-size:22px; font-weight:bold;">$($m.HealthyDatabases) / $($m.TotalDatabases)</div>
                </div>
                <div style="flex:1; background:#1e293b; padding:12px 16px; border-radius:8px; text-align:center;">
                    <div style="color:#94a3b8; font-size:11px; text-transform:uppercase;">Critical Breaches</div>
                    <div style="color:#ef4444; font-size:22px; font-weight:bold;">$($m.CriticalDatabases)</div>
                </div>
                <div style="flex:1; background:#1e293b; padding:12px 16px; border-radius:8px; text-align:center;">
                    <div style="color:#94a3b8; font-size:11px; text-transform:uppercase;">Backup Pool</div>
                    <div style="color:#a78bfa; font-size:22px; font-weight:bold;">$($m.TotalBackupSizeGB) GB</div>
                </div>
            </div>

            <h3 style="color:#e2e8f0; font-size:14px; margin-bottom:10px;">Database Backup Status Overview</h3>
            <table style="width: 100%; border-collapse: collapse; background-color: #0f172a; border-radius: 8px; font-size: 12px; text-align: left;">
                <thead>
                    <tr style="background-color: #1e293b; color: #94a3b8; text-transform: uppercase; font-size: 11px;">
                        <th style="padding: 8px 12px;">Status</th>
                        <th style="padding: 8px 12px;">Server</th>
                        <th style="padding: 8px 12px;">Database</th>
                        <th style="padding: 8px 12px;">Recovery</th>
                        <th style="padding: 8px 12px;">Last Full</th>
                        <th style="padding: 8px 12px;">Last Log</th>
                        <th style="padding: 8px 12px;">Diagnostics</th>
                    </tr>
                </thead>
                <tbody>
                    $rowsHtml
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>
"@
        $mail.Body = $bodyHtml
        $mail.IsBodyHtml = $true

        # Attach PDF Report
        $tempPdf = [System.IO.Path]::GetTempFileName() + ".pdf"
        Export-BackupPdfReport -OutputPath $tempPdf | Out-Null
        if (Test-Path $tempPdf) {
            $att = [System.Net.Mail.Attachment]::new($tempPdf, "application/pdf")
            $att.Name = "Executive_SQLBackup_Report_$(Get-Date -Format 'yyyyMMdd').pdf"
            $mail.Attachments.Add($att)
        }

        $client = [System.Net.Mail.SmtpClient]::new($emailCfg.SmtpServer, [int]$emailCfg.SmtpPort)
        $client.EnableSsl = [bool]$emailCfg.EnableSsl
        $client.Timeout = 20000

        if (-not [string]::IsNullOrWhiteSpace($emailCfg.SmtpUsername)) {
            $client.Credentials = [System.Net.NetworkCredential]::new($emailCfg.SmtpUsername, $emailCfg.SmtpPassword)
        }

        $client.Send($mail)
        $mail.Dispose()
        $client.Dispose()

        if ($tempPdf -and (Test-Path $tempPdf)) {
            Remove-Item $tempPdf -Force -ErrorAction SilentlyContinue
        }

        $Global:Config.EmailSettings.LastDailyReportDate = (Get-Date -Format "yyyy-MM-dd")
        Save-AppConfig -Config $Global:Config

        Write-Host "[OK] Daily executive PDF report email sent to $recipients" -ForegroundColor Green
        Add-AuditLog -Action "Daily Report Sent" -Details "Executive daily PDF report emailed to $recipients"
        return @{ Success = $true; Message = "Daily executive PDF report emailed successfully to $recipients." }
    } catch {
        Write-Warning "Failed to send daily PDF email report: $_"
        return @{ Success = $false; ErrorMessage = $_.Exception.Message }
    }
}

function Invoke-FullScan {
    [System.Threading.Monitor]::Enter($Global:ScanLock)
    try {
        $Global:IsScanning = $true
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        # Deduplicate active registered servers
        $seenServers = @{}
        $uniqueServers = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($s in @($Global:Config.Servers)) {
            if ($null -eq $s -or $s.IsEnabled -eq $false) { continue }
            $sKey = "$($s.ServerAddress):$($s.Port)".ToLower()
            $nKey = if ($s.Name) { $s.Name.ToLower() } else { $sKey }
            if (-not $seenServers.ContainsKey($sKey) -and -not $seenServers.ContainsKey($nKey)) {
                $seenServers[$sKey] = $true
                $seenServers[$nKey] = $true
                $uniqueServers.Add($s)
            }
        }
        $servers = $uniqueServers.ToArray()

        if ($servers.Count -eq 0) {
            $Global:CachedResults = @()
            $Global:CachedAlwaysOn = @()
            $Global:ActiveAlerts = @()
            $Global:LastScanTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $Global:CachedMetrics = [PSCustomObject]@{
                TotalServers         = 0
                TotalDatabases       = 0
                HealthyDatabases     = 0
                WarningDatabases     = 0
                CriticalDatabases    = 0
                CompliancePct        = 100
                TotalBackupSizeGB    = 0
                LastScanTime         = $Global:LastScanTime
            }
            return
        }

        # Optimal Concurrency pool (up to 35 parallel worker threads)
        $concurrency = [Math]::Min(35, [Math]::Max(4, [int]($servers.Count / 2)))
        if ($servers.Count -gt 50) { $concurrency = 30 }
        if ($servers.Count -gt 150) { $concurrency = 40 }

        # Setup Parallel Runspace Pool
        $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $concurrency, $sessionState, $Host)
        $pool.Open()

        $allResults = [System.Collections.Generic.List[PSCustomObject]]::new()
        $allAg = [System.Collections.Generic.List[PSCustomObject]]::new()
        $tasks = [System.Collections.Generic.List[PSCustomObject]]::new()

        $repoDbName = ""
        if ($Global:Config.RepositorySettings -and $Global:Config.RepositorySettings.DatabaseName) {
            $repoDbName = $Global:Config.RepositorySettings.DatabaseName.Trim()
        }

        $globalPolicies = $Global:Config.GlobalPolicies
        $groupPolicies = $Global:Config.GroupPolicies

        foreach ($s in $servers) {
            $ps = [System.Management.Automation.PowerShell]::Create()
            $ps.RunspacePool = $pool
            [void]$ps.AddScript($Global:ServerTelemetryScriptBlock)
            [void]$ps.AddArgument($s)
            [void]$ps.AddArgument($globalPolicies)
            [void]$ps.AddArgument($groupPolicies)
            [void]$ps.AddArgument($repoDbName)

            $asyncResult = $ps.BeginInvoke()
            $tasks.Add([PSCustomObject]@{
                PowerShell  = $ps
                AsyncResult = $asyncResult
                ServerName  = $s.Name
            })
        }

        # Collect async parallel results
        foreach ($t in $tasks) {
            try {
                $out = $t.PowerShell.EndInvoke($t.AsyncResult)
                if ($out -and $out.Count -gt 0) {
                    foreach ($item in $out) {
                        if ($item.Databases) {
                            foreach ($db in $item.Databases) { $allResults.Add($db) }
                        }
                        if ($item.AlwaysOn) {
                            foreach ($ag in $item.AlwaysOn) { $allAg.Add($ag) }
                        }
                    }
                }
            } catch {
                Write-Warning "Error processing async scan for $($t.ServerName): $_"
            } finally {
                $t.PowerShell.Dispose()
            }
        }

        $pool.Close()
        $pool.Dispose()

        # Deduplicate database telemetry results by ServerName + DatabaseName
        $seenDb = @{}
        $dedupedResults = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($r in $allResults) {
            if ($null -eq $r) { continue }
            $dbKey = "$($r.ServerName)|$($r.DatabaseName)".ToLower()
            if (-not $seenDb.ContainsKey($dbKey)) {
                $seenDb[$dbKey] = $true
                $dedupedResults.Add($r)
            }
        }
        $finalResults = $dedupedResults.ToArray()

        $Global:CachedResults = $finalResults
        $Global:CachedAlwaysOn = $allAg.ToArray()
        $Global:LastScanTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

        # Compute Metrics
        $totalDb = $finalResults.Count
        $healthyDb = @($finalResults | Where-Object { $_.Status -eq "Healthy" }).Count
        $warnDb = @($finalResults | Where-Object { $_.Status -eq "Warning" }).Count
        $critDb = @($finalResults | Where-Object { $_.Status -eq "Critical" }).Count
        $compliance = if ($totalDb -gt 0) { [Math]::Round(($healthyDb / [double]$totalDb) * 100.0, 1) } else { 100 }
        $totalSizeGB = [Math]::Round((($finalResults | Measure-Object -Property LastBackupSizeGB -Sum).Sum), 2)

        $Global:CachedMetrics = [PSCustomObject]@{
            TotalServers         = $servers.Count
            TotalDatabases       = $totalDb
            HealthyDatabases     = $healthyDb
            WarningDatabases     = $warnDb
            CriticalDatabases    = $critDb
            CompliancePct        = $compliance
            TotalBackupSizeGB    = $totalSizeGB
            LastScanTime         = $Global:LastScanTime
        }

        # Track Active Alerts
        $active = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($db in $finalResults) {
            if ($db.Status -ne "Healthy") {
                $active.Add([PSCustomObject]@{
                    Id           = [Guid]::NewGuid().ToString()
                    Timestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    ServerName   = $db.ServerName
                    DatabaseName = $db.DatabaseName
                    Severity     = $db.Status
                    Message      = $db.StatusReason
                    IsResolved   = $false
                })
            }
        }
        $Global:ActiveAlerts = $active.ToArray()

        $sw.Stop()
        $durationMs = [Math]::Round($sw.Elapsed.TotalMilliseconds, 0)
        Add-AuditLog -Action "Telemetry Scan Complete" -Details "Scanned $($servers.Count) servers ($totalDb databases) across $concurrency parallel runspaces in ${durationMs}ms"

        # Persist to Repository Database for Historical Trends
        if ($Global:Config.RepositorySettings -and $Global:Config.RepositorySettings.IsEnabled) {
            Save-RepositoryHistoricalData -Metrics $Global:CachedMetrics -Databases $finalResults
        }

        # Check for Email SLA Breaches
        $breached = @($finalResults | Where-Object { 
            ($_.Status -eq "Critical" -and $Global:Config.EmailSettings.AlertOnCritical) -or 
            ($_.Status -eq "Warning" -and $Global:Config.EmailSettings.AlertOnWarning) 
        })

        if ($breached.Count -gt 0) {
            Send-SlaEmailAlert -BreachedDatabases $breached
        }
    } finally {
        $Global:IsScanning = $false
        [System.Threading.Monitor]::Exit($Global:ScanLock)
    }
}

# ==============================================================================
# EMBEDDED DASHBOARD WEB CLIENT (HTML/CSS/JS)
# ==============================================================================
function Get-EmbeddedHtmlDashboard {
    return @'
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SQLBackupMonitor - Enterprise Dashboard (PowerShell Edition)</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-primary: #0b0f19;
            --bg-card: rgba(18, 24, 38, 0.75);
            --bg-card-hover: rgba(28, 38, 58, 0.85);
            --border-color: rgba(255, 255, 255, 0.08);
            --border-glow: rgba(59, 130, 246, 0.3);
            --accent-blue: #3b82f6;
            --accent-cyan: #06b6d4;
            --accent-purple: #8b5cf6;
            --success: #10b981;
            --warning: #f59e0b;
            --danger: #ef4444;
            --text-main: #f3f4f6;
            --text-muted: #9ca3af;
            --font-main: 'Outfit', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            --font-mono: 'JetBrains Mono', 'Cascadia Code', Consolas, monospace;
        }

        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            font-family: var(--font-main);
            background-color: var(--bg-primary);
            background-image: 
                radial-gradient(at 0% 0%, rgba(59, 130, 246, 0.12) 0px, transparent 50%),
                radial-gradient(at 100% 100%, rgba(139, 92, 246, 0.08) 0px, transparent 50%),
                radial-gradient(at 50% 50%, rgba(6, 182, 212, 0.05) 0px, transparent 50%);
            background-attachment: fixed;
            color: var(--text-main);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
        }

        header {
            background: rgba(11, 15, 25, 0.85);
            backdrop-filter: blur(16px);
            border-bottom: 1px solid var(--border-color);
            position: sticky;
            top: 0;
            z-index: 100;
            padding: 0.75rem 1.5rem;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1rem;
        }

        .brand-box {
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }

        .brand-logo {
            width: 38px;
            height: 38px;
            background: linear-gradient(135deg, var(--accent-blue), var(--accent-purple));
            border-radius: 9px;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 4px 14px rgba(59, 130, 246, 0.4);
            color: #ffffff;
        }

        .brand-logo svg { width: 22px; height: 22px; stroke: #ffffff; }

        .brand-text h1 {
            font-size: 1.15rem;
            font-weight: 700;
            letter-spacing: -0.02em;
            background: linear-gradient(to right, #fff, #93c5fd);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .brand-badge {
            font-size: 0.65rem;
            font-weight: 600;
            background: rgba(59, 130, 246, 0.2);
            color: #60a5fa;
            border: 1px solid rgba(59, 130, 246, 0.4);
            padding: 2px 7px;
            border-radius: 9999px;
            text-transform: uppercase;
        }

        .nav-tabs {
            display: flex;
            flex-wrap: wrap;
            gap: 0.3rem;
            background: rgba(255, 255, 255, 0.03);
            padding: 3px;
            border-radius: 10px;
            border: 1px solid var(--border-color);
        }

        .tab-btn {
            background: transparent;
            border: none;
            color: var(--text-muted);
            padding: 6px 12px;
            border-radius: 7px;
            cursor: pointer;
            font-family: inherit;
            font-size: 0.82rem;
            font-weight: 500;
            display: flex;
            align-items: center;
            gap: 6px;
            transition: all 0.2s ease;
        }

        .tab-btn svg { width: 14px; height: 14px; stroke: currentColor; }

        .tab-btn:hover { color: var(--text-main); background: rgba(255, 255, 255, 0.05); }

        .tab-btn.active {
            color: #fff;
            background: linear-gradient(135deg, rgba(59, 130, 246, 0.4), rgba(139, 92, 246, 0.4));
            border: 1px solid rgba(255, 255, 255, 0.15);
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.3);
        }

        .header-actions {
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .btn-primary {
            background: linear-gradient(135deg, var(--accent-blue), #2563eb);
            color: #fff;
            border: none;
            padding: 7px 14px;
            border-radius: 8px;
            font-family: inherit;
            font-weight: 600;
            font-size: 0.82rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 6px;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
            transition: all 0.2s;
        }

        .btn-primary svg { width: 14px; height: 14px; stroke: currentColor; }

        .btn-secondary {
            background: rgba(255, 255, 255, 0.06);
            color: var(--text-main);
            border: 1px solid var(--border-color);
            padding: 7px 12px;
            border-radius: 8px;
            font-family: inherit;
            font-weight: 500;
            font-size: 0.82rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 6px;
            transition: all 0.2s;
        }

        .btn-secondary svg { width: 14px; height: 14px; stroke: currentColor; }

        main {
            max-width: 1560px;
            margin: 0 auto;
            padding: 1.5rem;
            width: 100%;
            flex: 1;
        }

        .tab-content { display: none; animation: fadeIn 0.25s ease; }
        .tab-content.active { display: block; }

        @keyframes fadeIn { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: translateY(0); } }

        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
            gap: 1rem;
            margin-bottom: 1.5rem;
        }

        .kpi-card {
            background: var(--bg-card);
            backdrop-filter: blur(12px);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 1.15rem 1.25rem;
            position: relative;
            overflow: hidden;
            transition: all 0.2s;
        }

        .kpi-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.3);
        }

        .kpi-card::before {
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0; height: 3px;
            background: var(--card-accent, var(--accent-blue));
        }

        .kpi-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.4rem; }
        .kpi-title { font-size: 0.75rem; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em; }
        .kpi-icon { width: 30px; height: 30px; border-radius: 7px; background: rgba(255, 255, 255, 0.04); display: flex; align-items: center; justify-content: center; }
        .kpi-icon svg { width: 18px; height: 18px; }
        .kpi-value { font-size: 1.85rem; font-weight: 700; font-family: var(--font-mono); }
        .kpi-sub { font-size: 0.75rem; color: var(--text-muted); margin-top: 0.25rem; }

        .glass-panel {
            background: var(--bg-card);
            backdrop-filter: blur(12px);
            border: 1px solid var(--border-color);
            border-radius: 14px;
            padding: 1.25rem;
            margin-bottom: 1.5rem;
        }

        .panel-header {
            display: flex;
            flex-wrap: wrap;
            justify-content: space-between;
            align-items: center;
            gap: 0.75rem;
            margin-bottom: 1rem;
        }

        .panel-title { font-size: 1.05rem; font-weight: 600; display: flex; align-items: center; gap: 8px; }
        .panel-title svg { width: 18px; height: 18px; stroke: var(--accent-blue); }

        .search-filter-bar { display: flex; flex-wrap: wrap; gap: 0.6rem; align-items: center; }

        .input-search-box { position: relative; display: flex; align-items: center; }
        .input-search-box svg { position: absolute; left: 10px; width: 13px; height: 13px; stroke: var(--text-muted); }
        .input-search {
            background: rgba(0, 0, 0, 0.3);
            border: 1px solid var(--border-color);
            color: #fff;
            padding: 7px 12px 7px 30px;
            border-radius: 7px;
            font-size: 0.82rem;
            width: 240px;
            outline: none;
        }

        .select-filter {
            background: rgba(18, 24, 38, 0.9);
            border: 1px solid var(--border-color);
            color: #fff;
            padding: 7px 10px;
            border-radius: 7px;
            font-size: 0.82rem;
            outline: none;
            cursor: pointer;
        }

        .table-responsive { width: 100%; overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; font-size: 0.85rem; text-align: left; }
        th {
            background: rgba(255, 255, 255, 0.02);
            color: var(--text-muted);
            font-weight: 600;
            padding: 10px 12px;
            border-bottom: 1px solid var(--border-color);
            text-transform: uppercase;
            font-size: 0.72rem;
            white-space: nowrap;
        }
        td { padding: 10px 12px; border-bottom: 1px solid rgba(255, 255, 255, 0.04); color: var(--text-main); }
        tr:hover td { background: rgba(255, 255, 255, 0.02); }

        .badge-status { display: inline-flex; align-items: center; gap: 5px; padding: 2px 8px; border-radius: 9999px; font-size: 0.72rem; font-weight: 600; text-transform: uppercase; }
        .badge-healthy { background: rgba(16, 185, 129, 0.15); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.3); }
        .badge-warning { background: rgba(245, 158, 11, 0.15); color: #fbbf24; border: 1px solid rgba(245, 158, 11, 0.3); }
        .badge-critical { background: rgba(239, 68, 68, 0.15); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3); }
        .badge-recovery { background: rgba(255, 255, 255, 0.06); color: #d1d5db; padding: 2px 6px; border-radius: 4px; font-size: 0.7rem; font-family: var(--font-mono); }
        .mono { font-family: var(--font-mono); font-size: 0.8rem; }

        .chart-box { background: rgba(0, 0, 0, 0.25); border: 1px solid var(--border-color); border-radius: 10px; padding: 1rem; margin-bottom: 1.25rem; }
        .chart-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.75rem; }
        .chart-svg-container { width: 100%; height: 190px; }

        .modal-overlay { position: fixed; inset: 0; background: rgba(0, 0, 0, 0.75); backdrop-filter: blur(8px); z-index: 200; display: none; align-items: center; justify-content: center; padding: 1rem; }
        .modal-overlay.active { display: flex; }
        .modal-box { background: #111827; border: 1px solid var(--border-color); border-radius: 14px; width: 100%; max-width: 560px; box-shadow: 0 20px 40px rgba(0, 0, 0, 0.6); overflow: hidden; }
        .modal-header { padding: 1.15rem 1.25rem; border-bottom: 1px solid var(--border-color); display: flex; justify-content: space-between; align-items: center; }
        .modal-body { padding: 1.25rem; }
        .modal-footer { padding: 0.85rem 1.25rem; background: rgba(0, 0, 0, 0.2); border-top: 1px solid var(--border-color); display: flex; justify-content: flex-end; gap: 0.6rem; }

        .form-group { margin-bottom: 1rem; }
        .form-group label { display: block; font-size: 0.78rem; font-weight: 500; color: var(--text-muted); margin-bottom: 0.35rem; }
        .form-control { width: 100%; background: rgba(0, 0, 0, 0.3); border: 1px solid var(--border-color); color: #fff; padding: 8px 11px; border-radius: 7px; font-size: 0.85rem; outline: none; }
        .form-control:focus { border-color: var(--accent-blue); }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 0.85rem; }

        /* Server-First SLA Hierarchy & Accordion */
        .server-card {
            margin-bottom: 0.85rem;
            border-radius: 10px;
            background: rgba(15, 23, 42, 0.55);
            border: 1px solid var(--border-color);
            overflow: hidden;
            transition: all 0.2s ease;
        }
        .server-card:hover {
            border-color: rgba(56, 189, 248, 0.4);
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3);
        }
        .server-card.status-critical {
            border-left: 4px solid #ef4444;
        }
        .server-card.status-warning {
            border-left: 4px solid #f59e0b;
        }
        .server-card.status-healthy {
            border-left: 4px solid #10b981;
        }
        .server-card-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0.85rem 1.15rem;
            cursor: pointer;
            user-select: none;
            background: rgba(255, 255, 255, 0.02);
            gap: 1rem;
            flex-wrap: wrap;
        }
        .server-card-header:hover {
            background: rgba(255, 255, 255, 0.05);
        }
        .server-title-group {
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }
        .server-chevron-icon {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 22px;
            height: 22px;
            border-radius: 4px;
            background: rgba(255, 255, 255, 0.06);
            color: var(--text-muted);
            transition: transform 0.25s cubic-bezier(0.4, 0, 0.2, 1), background 0.2s ease, color 0.2s ease;
        }
        .server-chevron-icon svg {
            width: 14px;
            height: 14px;
        }
        .server-card.expanded .server-chevron-icon {
            transform: rotate(90deg);
            background: var(--accent-blue);
            color: #fff;
        }
        .server-name-text {
            font-size: 0.95rem;
            font-weight: 600;
            color: #f8fafc;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .server-stats-group {
            display: flex;
            align-items: center;
            gap: 0.6rem;
            flex-wrap: wrap;
            margin-left: auto;
        }
        .server-card-body {
            display: none;
            border-top: 1px solid rgba(255, 255, 255, 0.06);
            background: rgba(0, 0, 0, 0.22);
            padding: 0;
            animation: fadeInAccordion 0.2s ease forwards;
        }
        .server-card.expanded .server-card-body {
            display: block !important;
        }
        @keyframes fadeInAccordion {
            from { opacity: 0; transform: translateY(-4px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .pulse-dot { width: 7px; height: 7px; border-radius: 50%; display: inline-block; }
        .pulse-dot.green { background: #10b981; box-shadow: 0 0 6px #10b981; }
        .pulse-dot.amber { background: #f59e0b; box-shadow: 0 0 6px #f59e0b; }
        .pulse-dot.red { background: #ef4444; box-shadow: 0 0 6px #ef4444; }

        .pagination-container {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 0.85rem 0.25rem 0.25rem 0.25rem;
            font-size: 0.82rem;
            color: var(--text-muted);
            border-top: 1px solid rgba(255, 255, 255, 0.05);
            margin-top: 0.6rem;
            flex-wrap: wrap;
            gap: 0.5rem;
        }
        .pagination-info { font-size: 0.82rem; color: var(--text-muted); }
        .pagination-controls { display: flex; gap: 0.35rem; align-items: center; }
        .btn-page {
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-color);
            color: var(--text-main);
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 0.78rem;
            cursor: pointer;
            transition: all 0.2s ease;
            display: inline-flex;
            align-items: center;
            justify-content: center;
        }
        .btn-page:hover:not(:disabled) {
            background: var(--accent-blue);
            color: #fff;
            border-color: var(--accent-blue);
        }
        .btn-page:disabled {
            opacity: 0.35;
            cursor: not-allowed;
        }
        .page-indicator {
            padding: 4px 8px;
            font-size: 0.78rem;
            font-weight: 500;
            color: var(--text-main);
            background: rgba(0, 0, 0, 0.25);
            border-radius: 5px;
            border: 1px solid var(--border-color);
        }

        footer { border-top: 1px solid var(--border-color); padding: 1rem 1.5rem; text-align: center; font-size: 0.78rem; color: var(--text-muted); background: rgba(11, 15, 25, 0.6); }
    </style>
</head>
<body>

    <!-- Header Navigation -->
    <header>
        <div class="brand-box">
            <div class="brand-logo" title="SQLBackupMonitor">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"></path>
                    <polyline points="9 12 11 14 15 10"></polyline>
                </svg>
            </div>
            <div class="brand-text">
                <h1>SQLBackupMonitor</h1>
            </div>
            <span class="brand-badge">PowerShell Core v2.6</span>
        </div>

        <div class="nav-tabs">
            <button class="tab-btn active" onclick="switchTab('dashboard')">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="9"></rect><rect x="14" y="3" width="7" height="5"></rect><rect x="14" y="12" width="7" height="9"></rect><rect x="3" y="16" width="7" height="5"></rect></svg>
                Dashboard
            </button>
            <button class="tab-btn" onclick="switchTab('servers')">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="2" width="20" height="8" rx="2"></rect><rect x="2" y="14" width="20" height="8" rx="2"></rect><line x1="6" y1="6" x2="6.01" y2="6"></line><line x1="6" y1="18" x2="6.01" y2="18"></line></svg>
                Instances
            </button>
            <button class="tab-btn" onclick="switchTab('alwayson')">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg>
                AlwaysOn AGs
            </button>
            <button class="tab-btn" onclick="switchTab('alerts')">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>
                Alerts
            </button>
            <button class="tab-btn" onclick="switchTab('reports')">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line></svg>
                Reports &amp; PDF
            </button>
            <button class="tab-btn" onclick="switchTab('trends')">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"></polyline><polyline points="17 6 23 6 23 12"></polyline></svg>
                Historical Trends
            </button>
            <button class="tab-btn" onclick="switchTab('policies')">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path></svg>
                Policies &amp; SMTP
            </button>
            <button class="tab-btn" onclick="switchTab('repository')">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><ellipse cx="12" cy="5" rx="9" ry="3"></ellipse><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"></path><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"></path></svg>
                Repository DB
            </button>
            <button class="tab-btn" onclick="switchTab('audit')">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg>
                Audit Log
            </button>
        </div>

        <div class="header-actions">
            <span id="autoPullBadge" title="Configured background polling & UI refresh interval" style="font-size:0.75rem; color:var(--text-muted); background:rgba(255,255,255,0.05); border:1px solid var(--border-color); padding:6px 12px; border-radius:6px; display:inline-flex; align-items:center; gap:6px;">
                <span style="width:7px; height:7px; border-radius:50%; background:var(--accent-green); display:inline-block; box-shadow: 0 0 8px rgba(16,185,129,0.6);"></span>
                Auto-Pull: <strong id="lblAutoPullSec" style="color:var(--text-bright);">60s</strong>
            </span>
            <button class="btn-secondary" onclick="exportReportPdf()" title="Download Clean PDF Report">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                PDF Report
            </button>
            <button class="btn-primary" id="btnRefresh" onclick="triggerScan()">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg>
                Poll Now
            </button>
        </div>
    </header>

    <!-- Main Workspace -->
    <main>
        <!-- TAB 1: DASHBOARD -->
        <div id="tab-dashboard" class="tab-content active">
            <div class="kpi-grid">
                <div class="kpi-card" style="--card-accent: var(--accent-blue); cursor:pointer;" onclick="filterByKpi('ALL')" title="Click to view all monitored instances">
                    <div class="kpi-header"><span class="kpi-title">Monitored Instances</span><div class="kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="var(--accent-blue)" stroke-width="2"><rect x="2" y="2" width="20" height="8" rx="2"></rect><rect x="2" y="14" width="20" height="8" rx="2"></rect></svg></div></div>
                    <div class="kpi-value" id="kpiServers">0</div>
                    <div class="kpi-sub">Target SQL Servers</div>
                </div>
                <div class="kpi-card" style="--card-accent: var(--accent-cyan); cursor:pointer;" onclick="filterByKpi('ALL')" title="Click to view all databases">
                    <div class="kpi-header"><span class="kpi-title">Total Databases</span><div class="kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="var(--accent-cyan)" stroke-width="2"><ellipse cx="12" cy="5" rx="9" ry="3"></ellipse><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"></path><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"></path></svg></div></div>
                    <div class="kpi-value" id="kpiDatabases">0</div>
                    <div class="kpi-sub">Online User DBs</div>
                </div>
                <div class="kpi-card" style="--card-accent: var(--success); cursor:pointer;" onclick="filterByKpi('Healthy')" title="Click to filter to healthy instances">
                    <div class="kpi-header"><span class="kpi-title">SLA Compliance</span><div class="kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="var(--success)" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg></div></div>
                    <div class="kpi-value" id="kpiCompliance">100%</div>
                    <div class="kpi-sub" id="kpiHealthyCount">0 within policy</div>
                </div>
                <div class="kpi-card" style="--card-accent: var(--danger); cursor:pointer;" onclick="filterByKpi('Critical')" title="Click to filter directly to critical SLA breaches">
                    <div class="kpi-header"><span class="kpi-title">Critical Breaches</span><div class="kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="var(--danger)" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path></svg></div></div>
                    <div class="kpi-value" id="kpiCritical" style="color: var(--danger)">0</div>
                    <div class="kpi-sub" id="kpiWarningCount">0 warnings</div>
                </div>
                <div class="kpi-card" style="--card-accent: var(--accent-purple)">
                    <div class="kpi-header"><span class="kpi-title">Total Backup Size</span><div class="kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="var(--accent-purple)" stroke-width="2"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"></path></svg></div></div>
                    <div class="kpi-value" id="kpiSize">0 GB</div>
                    <div class="kpi-sub">Last full backup pool</div>
                </div>
            </div>

            <!-- Server-Grouped Backup Health & SLA Hierarchy -->
            <div class="glass-panel">
                <div class="panel-header">
                    <div class="panel-title">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="3" width="20" height="14" rx="2"></rect><line x1="8" y1="21" x2="16" y2="21"></line></svg>
                        <span>Server Backup Health &amp; SLA Tracking</span>
                        <span id="scanStatusBadge" style="font-size:0.75rem; color:var(--text-muted); font-weight:normal;"></span>
                    </div>
                    <div class="search-filter-bar">
                        <div class="input-search-box">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
                            <input type="text" id="searchInput" class="input-search" placeholder="Search server or database..." oninput="renderDatabaseTable(true)">
                        </div>
                        <select id="statusFilter" class="select-filter" onchange="renderDatabaseTable(true)">
                            <option value="ALL">All Statuses</option>
                            <option value="Critical">Critical Breaches Only</option>
                            <option value="Warning">Warning Alerts Only</option>
                            <option value="Healthy">Healthy Instances Only</option>
                        </select>
                        <select id="serverFilter" class="select-filter" onchange="renderDatabaseTable(true)">
                            <option value="ALL">All Servers</option>
                        </select>
                        <button class="btn-secondary" id="btnToggleAll" style="font-size:0.78rem; padding:6px 12px; display:inline-flex; align-items:center; gap:6px;" onclick="toggleAllServerAccordions()">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px; height:14px;"><polyline points="6 9 12 15 18 9"></polyline></svg>
                            <span id="lblToggleAll">Expand All</span>
                        </button>
                    </div>
                </div>

                <!-- Interactive Server-First Accordion Container -->
                <div id="serverAccordionContainer" style="padding-top:0.35rem;">
                    <div style="text-align:center; padding: 2.5rem; color: var(--text-muted)">Loading backup telemetry...</div>
                </div>

                <div class="pagination-container" id="dbPagination" style="display:none;">
                    <div class="pagination-info" id="dbPaginationInfo">Showing 0 of 0 servers</div>
                    <div class="pagination-controls">
                        <button class="btn-page" id="dbFirstBtn" onclick="changeDbPage('first')">&laquo; First</button>
                        <button class="btn-page" id="dbPrevBtn" onclick="changeDbPage('prev')">&lsaquo; Prev</button>
                        <span class="page-indicator" id="dbPageIndicator">Page 1 of 1</span>
                        <button class="btn-page" id="dbNextBtn" onclick="changeDbPage('next')">Next &rsaquo;</button>
                        <button class="btn-page" id="dbLastBtn" onclick="changeDbPage('last')">Last &raquo;</button>
                    </div>
                </div>
            </div>
        </div>

        <!-- TAB 2: INSTANCES -->
        <div id="tab-servers" class="tab-content">
            <div class="glass-panel">
                <div class="panel-header">
                    <div class="panel-title">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="2" width="20" height="8" rx="2"></rect><rect x="2" y="14" width="20" height="8" rx="2"></rect></svg>
                        <span>Configured SQL Server Instances</span>
                    </div>
                    <button class="btn-primary" onclick="openServerModal()">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>
                        Add SQL Instance
                    </button>
                </div>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th>Display Name</th>
                                <th>Server Address</th>
                                <th>Authentication</th>
                                <th>Group / Env</th>
                                <th>Status</th>
                                <th>Actions</th>
                            </tr>
                        </thead>
                        <tbody id="serversTableBody"></tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- TAB 3: ALWAYSON AGs -->
        <div id="tab-alwayson" class="tab-content">
            <div class="glass-panel">
                <div class="panel-header">
                    <div class="panel-title">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg>
                        <span>AlwaysOn High Availability &amp; Disaster Recovery (HADR) Groups</span>
                    </div>
                </div>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th>SQL Instance</th>
                                <th>Availability Group</th>
                                <th>Replica Node</th>
                                <th>Role</th>
                                <th>Backup Preference</th>
                                <th>Sync Health</th>
                                <th>Database Name</th>
                                <th>DB Sync State</th>
                            </tr>
                        </thead>
                        <tbody id="alwaysOnTableBody">
                            <tr><td colspan="8" style="text-align:center; padding:2rem; color:var(--text-muted)">No AlwaysOn AG clusters discovered or HADR is not enabled on monitored instances.</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- TAB 4: ALERTS & INCIDENTS -->
        <div id="tab-alerts" class="tab-content">
            <div class="glass-panel">
                <div class="panel-header">
                    <div class="panel-title">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path></svg>
                        <span>Active SLA Alerts &amp; Breach Incidents</span>
                    </div>
                </div>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th>Detected At</th>
                                <th>Severity</th>
                                <th>Server</th>
                                <th>Database</th>
                                <th>Violation Details</th>
                                <th>Action</th>
                            </tr>
                        </thead>
                        <tbody id="alertsTableBody">
                            <tr><td colspan="6" style="text-align:center; padding:2rem; color:var(--text-muted)">All database backups are currently within SLA policy limits.</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- TAB 5: REPORTS & PDF -->
        <div id="tab-reports" class="tab-content">
            <div class="glass-panel" style="max-width: 900px;">
                <div class="panel-header">
                    <div class="panel-title">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>
                        <span>Executive Backup Compliance Reports &amp; PDF Export</span>
                    </div>
                </div>

                <div style="background:rgba(59, 130, 246, 0.08); border:1px solid rgba(59, 130, 246, 0.25); border-radius:10px; padding:1.25rem; margin-bottom:1.5rem;">
                    <h4 style="color:#38bdf8; margin-bottom:0.5rem; font-size:0.95rem;">Executive High-Resolution PDF Audit Report</h4>
                    <p style="font-size:0.85rem; color:var(--text-muted); margin-bottom:1rem;">
                        Generates a formal, pixel-perfect executive PDF report with full compliance scoring, SLA breach diagnostics, individual server tables, and volume summaries.
                    </p>
                    <div style="display:flex; flex-wrap:wrap; gap:0.75rem;">
                        <button class="btn-primary" onclick="exportReportPdf()">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                            Download PDF Report
                        </button>
                        <button class="btn-secondary" onclick="exportReportHtml()">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path></svg>
                            View Standalone HTML Report
                        </button>
                        <button class="btn-secondary" onclick="emailPdfReportNow()" style="color:#34d399; border-color:rgba(16, 185, 129, 0.4);">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#34d399" stroke-width="2"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"></path><polyline points="22,6 12,13 2,6"></polyline></svg>
                            Email PDF Report to DBA Team
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <!-- TAB 6: HISTORICAL TRENDS -->
        <div id="tab-trends" class="tab-content">
            <div class="glass-panel">
                <div class="panel-header">
                    <div class="panel-title">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"></polyline><polyline points="17 6 23 6 23 12"></polyline></svg>
                        <span>SLA Compliance &amp; Volume Trends (Repository Backed)</span>
                    </div>
                    <div style="display:flex; gap:0.5rem; align-items:center;">
                        <span id="repoStatusIndicator" style="font-size:0.8rem; color:var(--text-muted)">Checking Repository...</span>
                        <select id="trendRangeSelect" class="select-filter" onchange="fetchTrends()">
                            <option value="1">Last 24 Hours</option>
                            <option value="7" selected>Last 7 Days</option>
                            <option value="30">Last 30 Days</option>
                            <option value="90">Last 90 Days</option>
                        </select>
                    </div>
                </div>

                <div class="chart-box">
                    <div class="chart-header">
                        <span style="font-size:0.85rem; font-weight:600; color:#38bdf8;">SLA Compliance Trend (%) Over Time</span>
                        <span id="trendLatestCompliance" style="font-size:0.8rem; color:var(--text-muted);"></span>
                    </div>
                    <div class="chart-svg-container">
                        <svg id="complianceSvg" width="100%" height="190" style="overflow:visible;"></svg>
                    </div>
                </div>

                <div class="chart-box">
                    <div class="chart-header">
                        <span style="font-size:0.85rem; font-weight:600; color:#a78bfa;">Total Backup Pool Volume (GB) Over Time</span>
                        <span id="trendLatestSize" style="font-size:0.8rem; color:var(--text-muted);"></span>
                    </div>
                    <div class="chart-svg-container">
                        <svg id="sizeSvg" width="100%" height="190" style="overflow:visible;"></svg>
                    </div>
                </div>

                <div class="panel-header" style="margin-top:1.5rem;">
                    <div class="panel-title"><span>Historical SLA Snapshot Log</span></div>
                </div>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th>Snapshot Time</th>
                                <th>Status</th>
                                <th>Server</th>
                                <th>Database</th>
                                <th>Recovery</th>
                                <th>Last Full</th>
                                <th>Size</th>
                                <th>Diagnostic Reason</th>
                            </tr>
                        </thead>
                        <tbody id="historyTableBody">
                            <tr><td colspan="8" style="text-align:center; padding:2rem; color:var(--text-muted)">Configure and enable the Repository Database to store historical trends.</td></tr>
                        </tbody>
                    </table>
                </div>

                <div class="pagination-container" id="histPagination" style="display:none;">
                    <div class="pagination-info" id="histPaginationInfo">Showing 0 of 0 snapshots</div>
                    <div class="pagination-controls">
                        <button class="btn-page" id="histFirstBtn" onclick="changeHistPage('first')">&laquo; First</button>
                        <button class="btn-page" id="histPrevBtn" onclick="changeHistPage('prev')">&lsaquo; Prev</button>
                        <span class="page-indicator" id="histPageIndicator">Page 1 of 1</span>
                        <button class="btn-page" id="histNextBtn" onclick="changeHistPage('next')">Next &rsaquo;</button>
                        <button class="btn-page" id="histLastBtn" onclick="changeHistPage('last')">Last &raquo;</button>
                    </div>
                </div>
            </div>
        </div>

        <!-- TAB 7: POLICIES & SMTP -->
        <div id="tab-policies" class="tab-content">
            <!-- Global Fallback Policy -->
            <div class="glass-panel" style="max-width: 840px;">
                <div class="panel-header">
                    <div class="panel-title">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path></svg>
                        <span>Global SLA Policy Thresholds (Default Fallback)</span>
                    </div>
                </div>
                <form id="policyForm" onsubmit="savePolicies(event)">
                    <div class="form-row">
                        <div class="form-group">
                            <label>Full Backup Warning Threshold (Hours)</label>
                            <input type="number" class="form-control" id="cfgFullWarn" required min="1">
                        </div>
                        <div class="form-group">
                            <label>Full Backup Critical Threshold (Hours)</label>
                            <input type="number" class="form-control" id="cfgFullCrit" required min="1">
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label>Differential Backup Warning (Hours)</label>
                            <input type="number" class="form-control" id="cfgDiffWarn" required min="1">
                        </div>
                        <div class="form-group">
                            <label>Differential Backup Critical (Hours)</label>
                            <input type="number" class="form-control" id="cfgDiffCrit" required min="1">
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label>Log Backup Warning (Minutes)</label>
                            <input type="number" class="form-control" id="cfgLogWarn" required min="1">
                        </div>
                        <div class="form-group">
                            <label>Log Backup Critical (Minutes)</label>
                            <input type="number" class="form-control" id="cfgLogCrit" required min="1">
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group" style="flex:1;">
                            <label>Automatic Telemetry Pulling &amp; Refresh Interval (Seconds)</label>
                            <input type="number" class="form-control" id="cfgAutoRefresh" required min="5" max="86400" placeholder="60">
                            <small style="color:var(--text-muted); font-size:0.75rem; display:block; margin-top:4px;">Frequency in seconds for the background engine to query SQL Server instances and refresh Web UI telemetry (e.g. 30, 60, 300).</small>
                        </div>
                    </div>
                    <div style="margin-top: 1rem;">
                        <button type="submit" class="btn-primary">Save Global Policies &amp; Interval</button>
                    </div>
                </form>
            </div>

            <!-- Group-wise SLA Policies Card -->
            <div class="glass-panel" style="max-width: 840px; margin-top: 1.5rem;">
                <div class="panel-header">
                    <div class="panel-title">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
                        <span>Group-Wise SLA Policies</span>
                    </div>
                    <button class="btn-primary" style="font-size:0.78rem; padding:5px 12px;" onclick="openGroupPolicyModal()">+ Add Group Policy</button>
                </div>
                <p style="font-size:0.82rem; color:var(--text-muted); margin-bottom:1rem;">
                    Define distinct SLA threshold rules for specific server groups (e.g. <strong>Production</strong>, <strong>Non-Production</strong>, <strong>DR</strong>, <strong>Staging</strong>, <strong>Tier-1</strong>). Any server assigned to a group automatically inherits that group's backup SLA rules unless overridden.
                </p>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th>Group Name</th>
                                <th>Full Backup (Warn / Crit)</th>
                                <th>Diff Backup (Warn / Crit)</th>
                                <th>Log Backup (Warn / Crit)</th>
                                <th>Actions</th>
                            </tr>
                        </thead>
                        <tbody id="groupPoliciesTableBody">
                            <tr><td colspan="5" style="text-align:center; padding:1.5rem; color:var(--text-muted)">Loading Group SLA Policies...</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Email / SMTP Alerting Settings -->
            <div class="glass-panel" style="max-width: 840px; margin-top: 1.5rem;">
                <div class="panel-header">
                    <div class="panel-title">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"></path><polyline points="22,6 12,13 2,6"></polyline></svg>
                        <span>Email &amp; SMTP Notification Settings</span>
                    </div>
                </div>
                <form id="emailForm" onsubmit="saveEmailSettings(event)">
                    <div class="form-group" style="margin-bottom: 1.25rem;">
                        <label style="display:flex; align-items:center; gap: 8px; cursor: pointer; color:#fff; font-size:0.92rem;">
                            <input type="checkbox" id="emailEnabled" style="width:18px; height:18px;">
                            <strong>Enable Automated Email Alerts on SLA Breaches</strong>
                        </label>
                    </div>
                    <div class="form-row">
                        <div class="form-group"><label>SMTP Server / Host</label><input type="text" class="form-control" id="smtpServer" placeholder="smtp.office365.com, smtp.gmail.com"></div>
                        <div class="form-group"><label>SMTP Port (e.g. 587, 25, 465)</label><input type="number" class="form-control" id="smtpPort" value="587"></div>
                    </div>
                    <div class="form-row">
                        <div class="form-group"><label>Sender Email Address</label><input type="email" class="form-control" id="senderEmail" placeholder="dba-alerts@yourcompany.com"></div>
                        <div class="form-group"><label>Sender Display Name</label><input type="text" class="form-control" id="senderDisplayName" placeholder="SQLBackupMonitor DBA Alerts"></div>
                    </div>
                    <div class="form-row">
                        <div class="form-group"><label>SMTP Username (Optional)</label><input type="text" class="form-control" id="smtpUsername" placeholder="Optional or Office365 Account"></div>
                        <div class="form-group"><label>SMTP Password (Optional)</label><input type="password" class="form-control" id="smtpPassword" placeholder="********"></div>
                    </div>
                    <div class="form-group"><label>Recipient Email(s) (Comma-separated)</label><input type="text" class="form-control" id="recipientEmails" placeholder="dba1@corp.com, oncall@corp.com"></div>
                    
                    <div class="form-row" style="margin-top: 0.75rem;">
                        <div class="form-group">
                            <label>Alert Trigger &amp; Attachment Options</label>
                            <div style="display:flex; flex-wrap:wrap; gap:1.25rem; margin-top:0.35rem;">
                                <label style="display:flex; align-items:center; gap:6px; cursor:pointer;"><input type="checkbox" id="alertOnCrit" checked> [!] Critical</label>
                                <label style="display:flex; align-items:center; gap:6px; cursor:pointer;"><input type="checkbox" id="alertOnWarn"> [i] Warning</label>
                                <label style="display:flex; align-items:center; gap:6px; cursor:pointer;"><input type="checkbox" id="smtpSsl" checked> SSL / TLS</label>
                                <label style="display:flex; align-items:center; gap:6px; cursor:pointer;"><input type="checkbox" id="emailAttachPdf" checked> [PDF] Attach PDF Report</label>
                            </div>
                        </div>
                        <div class="form-group"><label>Alert Cooldown (Minutes)</label><input type="number" class="form-control" id="cooldownMin" value="60" min="5"></div>
                    </div>

                    <!-- Daily Scheduled Executive PDF Report Email -->
                    <div class="form-group" style="background: rgba(16, 185, 129, 0.08); border: 1px solid rgba(16, 185, 129, 0.25); border-radius: 8px; padding: 1rem; margin-top: 1.25rem; margin-bottom: 1rem;">
                        <label style="display:flex; align-items:center; gap: 10px; cursor: pointer; color:#fff; font-size:0.92rem;">
                            <input type="checkbox" id="dailyReportEnabled" style="width:18px; height:18px;">
                            <strong>Enable Daily Scheduled Executive PDF Report Email</strong>
                        </label>
                        <div class="form-row" style="margin-top:0.75rem; margin-bottom:0;">
                            <div class="form-group" style="margin-bottom:0;">
                                <label style="font-size:0.8rem; color:#94a3b8;">Daily Send Time (24h Format HH:mm, e.g. 08:00)</label>
                                <input type="time" class="form-control" id="dailyReportTime" value="08:00" style="max-width:200px;">
                            </div>
                            <div class="form-group" style="display:flex; align-items:flex-end; margin-bottom:0;">
                                <button type="button" class="btn-secondary" onclick="emailPdfReportNow()" style="color:#34d399;">Send Daily PDF Report Now</button>
                            </div>
                        </div>
                    </div>

                    <div style="margin-top: 1.25rem; display: flex; gap: 0.75rem;">
                        <button type="button" class="btn-secondary" onclick="sendTestEmail()">Send Test Alert Email</button>
                        <button type="submit" class="btn-primary">Save Email Settings</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- TAB 8: REPOSITORY SETTINGS -->
        <div id="tab-repository" class="tab-content">
            <div class="glass-panel" style="max-width: 840px;">
                <div class="panel-header">
                    <div class="panel-title">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><ellipse cx="12" cy="5" rx="9" ry="3"></ellipse><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"></path><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"></path></svg>
                        <span>SQL Server Repository Database Configuration</span>
                    </div>
                </div>
                <form id="repoForm" onsubmit="saveRepoSettings(event)">
                    <div class="form-group" style="background: rgba(59, 130, 246, 0.08); border: 1px solid rgba(59, 130, 246, 0.25); border-radius: 8px; padding: 1rem; margin-bottom:1.25rem;">
                        <label style="display:flex; align-items:center; gap: 10px; cursor: pointer; color:#fff; font-size:0.95rem;">
                            <input type="checkbox" id="repoEnabled" style="width:18px; height:18px;">
                            <strong>Enable SQL Server Repository Database for Historical Trend Storage</strong>
                        </label>
                    </div>
                    <div class="form-row">
                        <div class="form-group"><label>Repository SQL Server Address / Host</label><input type="text" class="form-control" id="repoServer" placeholder="localhost, SQLREPO01" required></div>
                        <div class="form-group"><label>Port (Default: 1433)</label><input type="number" class="form-control" id="repoPort" value="1433"></div>
                    </div>
                    <div class="form-row">
                        <div class="form-group"><label>Repository Database Name</label><input type="text" class="form-control" id="repoDatabase" value="SQLBackupMonitorDB" required></div>
                        <div class="form-group"><label>Data Retention Period (Days)</label><input type="number" class="form-control" id="repoRetention" value="90" min="1" max="3650"></div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label>Connection Encryption</label>
                            <select id="repoEncryption" class="form-control">
                                <option value="Optional">Optional (Default / No Encryption)</option>
                                <option value="Mandatory">Mandatory (Encrypt=True)</option>
                                <option value="Strict">Strict (TDS 8.0 / TLS 1.3)</option>
                            </select>
                        </div>
                        <div class="form-group" style="display:flex; align-items:flex-end; padding-bottom:0.5rem;">
                            <label style="display:flex; align-items:center; gap:8px; cursor:pointer; color:#fff; font-size:0.88rem;">
                                <input type="checkbox" id="repoTrustCert" checked style="width:18px; height:18px;">
                                <strong>Trust Server Certificate</strong>
                            </label>
                        </div>
                    </div>
                    <div class="form-group">
                        <label>Authentication Mode</label>
                        <select id="repoAuthType" class="form-control" onchange="toggleRepoAuthFields()">
                            <option value="Windows">Windows Authentication (Current Credentials)</option>
                            <option value="SqlPassword">SQL Server Authentication (Username &amp; Password)</option>
                        </select>
                    </div>
                    <div id="repoSqlAuthBox" style="display:none;">
                        <div class="form-row">
                            <div class="form-group"><label>SQL Username</label><input type="text" id="repoUsername" class="form-control" placeholder="sa or repo_user"></div>
                            <div class="form-group"><label>SQL Password</label><input type="password" id="repoPassword" class="form-control" placeholder="********"></div>
                        </div>
                    </div>
                    <div style="margin-top: 1.25rem; display: flex; flex-wrap: wrap; gap: 0.75rem;">
                        <button type="button" class="btn-secondary" onclick="testRepoConnection()">Test Connection</button>
                        <button type="button" class="btn-secondary" onclick="initializeRepoSchema()" style="color:#38bdf8;">Initialize Schema</button>
                        <button type="button" class="btn-secondary" onclick="cleanupRepoData()" style="color:#f87171;">Purge Old Data</button>
                        <button type="submit" class="btn-primary" style="margin-left:auto;">Save Repository Settings</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- TAB 9: AUDIT LOG -->
        <div id="tab-audit" class="tab-content">
            <div class="glass-panel">
                <div class="panel-header">
                    <div class="panel-title">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg>
                        <span>System Activity &amp; Audit Trail</span>
                    </div>
                </div>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th>Timestamp</th>
                                <th>User</th>
                                <th>Action</th>
                                <th>Operation Details</th>
                            </tr>
                        </thead>
                        <tbody id="auditTableBody"></tbody>
                    </table>
                </div>
            </div>
        </div>
    </main>

    <!-- Modal: Add / Edit Server -->
    <div class="modal-overlay" id="serverModal">
        <div class="modal-box">
            <div class="modal-header">
                <h3 id="modalTitle" style="font-size:1.05rem;">Add SQL Server Instance</h3>
                <button onclick="closeServerModal()" style="background:none; border:none; color:var(--text-muted); font-size:1.2rem; cursor:pointer;" title="Close">&times;</button>
            </div>
            <form id="serverForm" onsubmit="saveServer(event)">
                <div class="modal-body">
                    <input type="hidden" id="srvId">
                    <div class="form-group"><label>Friendly Name</label><input type="text" id="srvName" class="form-control" placeholder="e.g. Production Cluster 01" required></div>
                    <div class="form-row">
                        <div class="form-group"><label>SQL Address / Hostname</label><input type="text" id="srvAddress" class="form-control" placeholder="e.g. localhost, SQLPROD01" required></div>
                        <div class="form-group"><label>Port (Default: 1433)</label><input type="number" id="srvPort" class="form-control" value="1433"></div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label>Connection Encryption</label>
                            <select id="srvEncryption" class="form-control">
                                <option value="Optional">Optional (Default / No Encryption)</option>
                                <option value="Mandatory">Mandatory (Encrypt=True)</option>
                                <option value="Strict">Strict (TDS 8.0 / TLS 1.3)</option>
                            </select>
                        </div>
                        <div class="form-group" style="display:flex; align-items:flex-end; padding-bottom:0.5rem;">
                            <label style="display:flex; align-items:center; gap:8px; cursor:pointer; color:#fff; font-size:0.88rem;">
                                <input type="checkbox" id="srvTrustCert" checked style="width:18px; height:18px;">
                                <strong>Trust Server Certificate</strong>
                            </label>
                        </div>
                    </div>
                    <div class="form-group">
                        <label>Authentication Mode</label>
                        <select id="srvAuthType" class="form-control" onchange="toggleAuthFields()">
                            <option value="Windows">Windows Authentication (Current Credentials)</option>
                            <option value="SqlPassword">SQL Server Authentication (Username &amp; Password)</option>
                        </select>
                    </div>
                    <div id="sqlAuthBox" style="display:none;">
                        <div class="form-row">
                            <div class="form-group"><label>SQL Username</label><input type="text" id="srvUsername" class="form-control" placeholder="sa or backup_svc"></div>
                            <div class="form-group"><label>SQL Password</label><input type="password" id="srvPassword" class="form-control" placeholder="********"></div>
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group"><label>Group / Category</label><input type="text" id="srvGroup" class="form-control" value="Production" oninput="updateServerSlaNotice()"></div>
                        <div class="form-group"><label>Environment</label><input type="text" id="srvEnv" class="form-control" value="Production"></div>
                    </div>

                    <!-- Server-Wise Custom SLA Override Panel -->
                    <div style="background: rgba(59, 130, 246, 0.06); border: 1px solid rgba(59, 130, 246, 0.2); border-radius: 8px; padding: 0.9rem; margin-top: 1rem;">
                        <label style="display:flex; align-items:center; gap: 8px; cursor: pointer; color:#fff; font-size:0.88rem; font-weight:600;">
                            <input type="checkbox" id="srvUseCustomSla" onchange="toggleServerCustomSla()" style="width:16px; height:16px;">
                            Override SLA: Custom SLA Thresholds for this Instance
                        </label>
                        <div id="srvCustomSlaNotice" style="font-size:0.75rem; color:#94a3b8; margin-top:4px;">
                            Currently inheriting SLA rules from Group policy or Global default.
                        </div>
                        <div id="srvCustomSlaFields" style="display:none; margin-top:0.8rem; border-top:1px solid rgba(255,255,255,0.08); padding-top:0.8rem;">
                            <div class="form-row">
                                <div class="form-group"><label style="font-size:0.75rem;">Full Warning (Hours)</label><input type="number" id="srvFullWarn" class="form-control" value="20" min="1"></div>
                                <div class="form-group"><label style="font-size:0.75rem;">Full Critical (Hours)</label><input type="number" id="srvFullCrit" class="form-control" value="24" min="1"></div>
                            </div>
                            <div class="form-row">
                                <div class="form-group"><label style="font-size:0.75rem;">Diff Warning (Hours)</label><input type="number" id="srvDiffWarn" class="form-control" value="10" min="1"></div>
                                <div class="form-group"><label style="font-size:0.75rem;">Diff Critical (Hours)</label><input type="number" id="srvDiffCrit" class="form-control" value="12" min="1"></div>
                            </div>
                            <div class="form-row">
                                <div class="form-group"><label style="font-size:0.75rem;">Log Warning (Minutes)</label><input type="number" id="srvLogWarn" class="form-control" value="20" min="1"></div>
                                <div class="form-group"><label style="font-size:0.75rem;">Log Critical (Minutes)</label><input type="number" id="srvLogCrit" class="form-control" value="30" min="1"></div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn-secondary" onclick="testModalConnection()">Test Connection</button>
                    <button type="submit" class="btn-primary">Save Server</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Modal: Add / Edit Group SLA Policy -->
    <div class="modal-overlay" id="groupPolicyModal">
        <div class="modal-box" style="max-width: 520px;">
            <div class="modal-header">
                <h3 id="grpModalTitle" style="font-size:1.05rem;">Add Group SLA Policy</h3>
                <button onclick="closeGroupPolicyModal()" style="background:none; border:none; color:var(--text-muted); font-size:1.2rem; cursor:pointer;" title="Close">&times;</button>
            </div>
            <form id="groupPolicyForm" onsubmit="saveGroupPolicy(event)">
                <div class="modal-body">
                    <input type="hidden" id="grpOriginalName">
                    <div class="form-group">
                        <label>Group Name (e.g. Production, Non-Production, DR, Staging)</label>
                        <input type="text" id="grpName" class="form-control" placeholder="Production" required>
                    </div>
                    <div class="form-row">
                        <div class="form-group"><label>Full Warning (Hours)</label><input type="number" id="grpFullWarn" class="form-control" value="20" min="1" required></div>
                        <div class="form-group"><label>Full Critical (Hours)</label><input type="number" id="grpFullCrit" class="form-control" value="24" min="1" required></div>
                    </div>
                    <div class="form-row">
                        <div class="form-group"><label>Diff Warning (Hours)</label><input type="number" id="grpDiffWarn" class="form-control" value="10" min="1" required></div>
                        <div class="form-group"><label>Diff Critical (Hours)</label><input type="number" id="grpDiffCrit" class="form-control" value="12" min="1" required></div>
                    </div>
                    <div class="form-row">
                        <div class="form-group"><label>Log Warning (Minutes)</label><input type="number" id="grpLogWarn" class="form-control" value="20" min="1" required></div>
                        <div class="form-group"><label>Log Critical (Minutes)</label><input type="number" id="grpLogCrit" class="form-control" value="30" min="1" required></div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn-secondary" onclick="closeGroupPolicyModal()">Cancel</button>
                    <button type="submit" class="btn-primary">Save Group Policy</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Footer -->
    <footer>
        SQLBackupMonitor (Standalone PowerShell Platform) &bull; Zero External Executables Required &bull; Auto-Refresh Active
    </footer>

    <!-- Frontend Script -->
    <script>
        let allDatabases = [];
        let allServers = [];
        let allAlwaysOn = [];
        let allAlerts = [];
        let allAudit = [];
        let globalConfig = null;
        let trendData = [];

        // Pagination state
        let dbCurrentPage = 1;
        const dbPageSize = 10;
        let filteredDatabases = [];

        let allHistoryLogs = [];
        let histCurrentPage = 1;
        const histPageSize = 10;

        function switchTab(tabId) {
            document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
            document.getElementById('tab-' + tabId).classList.add('active');
            if (event && event.currentTarget) {
                event.currentTarget.classList.add('active');
            }
            if (tabId === 'trends') { fetchTrends(); fetchHistoryLogs(); }
            if (tabId === 'alwayson') { fetchAlwaysOn(); }
            if (tabId === 'alerts') { fetchAlerts(); }
            if (tabId === 'audit') { fetchAuditLogs(); }
        }

        async function fetchMetrics() {
            try {
                const res = await fetch('/api/dashboard/metrics');
                if (!res.ok) return;
                const m = await res.json();
                document.getElementById('kpiServers').innerText = m.TotalServers !== undefined ? m.TotalServers : 0;
                document.getElementById('kpiDatabases').innerText = m.TotalDatabases !== undefined ? m.TotalDatabases : 0;
                const compVal = (m.CompliancePct !== undefined && m.CompliancePct !== null) ? m.CompliancePct : (m.TotalDatabases === 0 ? 100 : 0);
                document.getElementById('kpiCompliance').innerText = compVal + '%';
                const compKpi = document.getElementById('kpiCompliance');
                if (compVal >= 90) { compKpi.style.color = 'var(--success)'; }
                else if (compVal >= 70) { compKpi.style.color = 'var(--warning)'; }
                else { compKpi.style.color = 'var(--danger)'; }
                document.getElementById('kpiHealthyCount').innerText = `${m.HealthyDatabases || 0} within SLA policy`;
                document.getElementById('kpiCritical').innerText = m.CriticalDatabases || 0;
                document.getElementById('kpiWarningCount').innerText = `${m.WarningDatabases || 0} warning alerts`;
                document.getElementById('kpiSize').innerText = (m.TotalBackupSizeGB || 0) + ' GB';
                if (m.LastScanTime) {
                    document.getElementById('scanStatusBadge').innerText = `Last Polled: ${m.LastScanTime}`;
                }
            } catch (err) { console.error("Failed to load metrics:", err); }
        }

        async function fetchDatabases() {
            try {
                const res = await fetch('/api/dashboard/databases');
                if (!res.ok) return;
                allDatabases = await res.json();
                populateServerFilterDropdown();
                renderDatabaseTable(false);
            } catch (err) { console.error("Failed to load databases:", err); }
        }

        async function fetchAlwaysOn() {
            try {
                const res = await fetch('/api/alwayson');
                if (!res.ok) return;
                allAlwaysOn = await res.json();
                renderAlwaysOnTable();
            } catch (err) {}
        }

        async function fetchAlerts() {
            try {
                const res = await fetch('/api/alerts');
                if (!res.ok) return;
                allAlerts = await res.json();
                renderAlertsTable();
            } catch (err) {}
        }

        async function fetchAuditLogs() {
            try {
                const res = await fetch('/api/audit');
                if (!res.ok) return;
                allAudit = await res.json();
                renderAuditTable();
            } catch (err) {}
        }

        async function fetchServers() {
            try {
                const res = await fetch('/api/servers');
                if (!res.ok) return;
                allServers = await res.json();
                renderServersTable();
            } catch (err) {}
        }

        async function fetchConfig() {
            try {
                const res = await fetch('/api/config');
                if (!res.ok) return;
                globalConfig = await res.json();
                
                if (globalConfig.GlobalPolicies) {
                    const p = globalConfig.GlobalPolicies;
                    document.getElementById('cfgFullWarn').value = p.FullBackupWarningHours;
                    document.getElementById('cfgFullCrit').value = p.FullBackupCriticalHours;
                    document.getElementById('cfgDiffWarn').value = p.DiffBackupWarningHours;
                    document.getElementById('cfgDiffCrit').value = p.DiffBackupCriticalHours;
                    document.getElementById('cfgLogWarn').value = p.LogBackupWarningMinutes;
                    document.getElementById('cfgLogCrit').value = p.LogBackupCriticalMinutes;
                    const pullSec = p.AutoRefreshIntervalSec || 60;
                    document.getElementById('cfgAutoRefresh').value = pullSec;
                    if (typeof updateAutoRefreshInterval === 'function') updateAutoRefreshInterval(pullSec);
                }

                renderGroupPoliciesTable();

                if (globalConfig.RepositorySettings) {
                    const r = globalConfig.RepositorySettings;
                    document.getElementById('repoEnabled').checked = !!r.IsEnabled;
                    document.getElementById('repoServer').value = r.ServerAddress || 'localhost';
                    document.getElementById('repoPort').value = r.Port || 1433;
                    document.getElementById('repoDatabase').value = r.DatabaseName || 'SQLBackupMonitorDB';
                    document.getElementById('repoRetention').value = r.RetentionDays || 90;
                    document.getElementById('repoEncryption').value = r.Encryption || 'Optional';
                    document.getElementById('repoTrustCert').checked = (r.TrustServerCertificate !== false);
                    document.getElementById('repoAuthType').value = r.AuthType || 'Windows';
                    document.getElementById('repoUsername').value = r.Username || '';
                    document.getElementById('repoPassword').value = r.Password || '';
                    toggleRepoAuthFields();

                    const indicator = document.getElementById('repoStatusIndicator');
                    if (indicator) {
                        indicator.innerHTML = r.IsEnabled 
                            ? '<span style="color:#34d399">&#9679; Repository Connected (' + escapeHtml(r.DatabaseName) + ')</span>'
                            : '<span style="color:#fbbf24">&#9675; Repository Disabled (Enable in Repository DB tab)</span>';
                    }
                }

                if (globalConfig.EmailSettings) {
                    const e = globalConfig.EmailSettings;
                    document.getElementById('emailEnabled').checked = !!e.IsEnabled;
                    document.getElementById('smtpServer').value = e.SmtpServer || '';
                    document.getElementById('smtpPort').value = e.SmtpPort || 587;
                    document.getElementById('senderEmail').value = e.SenderEmail || '';
                    document.getElementById('senderDisplayName').value = e.SenderDisplayName || '';
                    document.getElementById('smtpUsername').value = e.SmtpUsername || '';
                    document.getElementById('smtpPassword').value = e.SmtpPassword || '';
                    document.getElementById('recipientEmails').value = e.RecipientEmails || '';
                    document.getElementById('alertOnCrit').checked = (e.AlertOnCritical !== false);
                    document.getElementById('alertOnWarn').checked = !!e.AlertOnWarning;
                    document.getElementById('smtpSsl').checked = (e.EnableSsl !== false);
                    document.getElementById('emailAttachPdf').checked = (e.AttachPdfReport !== false);
                    document.getElementById('dailyReportEnabled').checked = !!e.DailyReportEnabled;
                    document.getElementById('dailyReportTime').value = e.DailyReportTime || '08:00';
                    document.getElementById('cooldownMin').value = e.CooldownMinutes || 60;
                }
            } catch (err) {}
        }

        async function fetchTrends() {
            const days = document.getElementById('trendRangeSelect').value || 7;
            try {
                const res = await fetch(`/api/repository/trends?days=${days}`);
                if (!res.ok) return;
                trendData = await res.json();
                renderTrendCharts(trendData);
            } catch (err) {}
        }

        function changeHistPage(action) {
            const totalPages = Math.max(1, Math.ceil(allHistoryLogs.length / histPageSize));
            if (action === 'first') histCurrentPage = 1;
            else if (action === 'prev') histCurrentPage = Math.max(1, histCurrentPage - 1);
            else if (action === 'next') histCurrentPage = Math.min(totalPages, histCurrentPage + 1);
            else if (action === 'last') histCurrentPage = totalPages;
            renderHistoryTable(false);
        }

        function renderHistoryTable(resetPage = true) {
            if (resetPage) histCurrentPage = 1;
            const tbody = document.getElementById('historyTableBody');
            const totalRecords = allHistoryLogs.length;
            const totalPages = Math.max(1, Math.ceil(totalRecords / histPageSize));
            if (histCurrentPage > totalPages) histCurrentPage = totalPages;

            const startIdx = (histCurrentPage - 1) * histPageSize;
            const endIdx = Math.min(startIdx + histPageSize, totalRecords);
            const pagedLogs = allHistoryLogs.slice(startIdx, endIdx);

            const pagEl = document.getElementById('histPagination');
            if (pagEl) {
                pagEl.style.display = totalRecords > 0 ? 'flex' : 'none';
                document.getElementById('histPaginationInfo').innerText = totalRecords > 0 
                    ? `Showing ${startIdx + 1}–${endIdx} of ${totalRecords} snapshots` 
                    : `Showing 0 of 0 snapshots`;
                document.getElementById('histPageIndicator').innerText = `Page ${histCurrentPage} of ${totalPages}`;
                document.getElementById('histFirstBtn').disabled = (histCurrentPage <= 1);
                document.getElementById('histPrevBtn').disabled = (histCurrentPage <= 1);
                document.getElementById('histNextBtn').disabled = (histCurrentPage >= totalPages);
                document.getElementById('histLastBtn').disabled = (histCurrentPage >= totalPages);
            }

            if (!allHistoryLogs || allHistoryLogs.length === 0) {
                tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; padding:2rem; color:var(--text-muted)">No historical snapshots recorded yet. Ensure Repository DB is enabled.</td></tr>`;
                return;
            }

            tbody.innerHTML = pagedLogs.map(l => {
                let badgeCls = 'badge-healthy';
                let pulse = 'green';
                if (l.Status === 'Warning') { badgeCls = 'badge-warning'; pulse = 'amber'; }
                if (l.Status === 'Critical') { badgeCls = 'badge-critical'; pulse = 'red'; }
                return `
                    <tr>
                        <td class="mono" style="font-size:0.8rem; color:#94a3b8;">${escapeHtml(l.SnapshotTime)}</td>
                        <td><span class="badge-status ${badgeCls}"><span class="pulse-dot ${pulse}"></span>${l.Status}</span></td>
                        <td style="font-weight:600">${escapeHtml(l.ServerName)}</td>
                        <td style="font-weight:500">${escapeHtml(l.DatabaseName)}</td>
                        <td><span class="badge-recovery">${escapeHtml(l.RecoveryModel)}</span></td>
                        <td class="mono">${escapeHtml(l.LastFullBackup)}</td>
                        <td class="mono">${l.LastBackupSizeGB > 0 ? l.LastBackupSizeGB + ' GB' : '-'}</td>
                        <td style="color:${l.Status === 'Critical' ? '#f87171' : (l.Status === 'Warning' ? '#fbbf24' : '#9ca3af')}; font-size:0.8rem;">
                            ${escapeHtml(l.StatusReason)}
                        </td>
                    </tr>
                `;
            }).join('');
        }

        async function fetchHistoryLogs() {
            try {
                const res = await fetch('/api/repository/history?limit=300');
                if (!res.ok) return;
                allHistoryLogs = await res.json();
                renderHistoryTable(true);
            } catch (err) {}
        }

        function renderTrendCharts(data) {
            const compSvg = document.getElementById('complianceSvg');
            const sizeSvg = document.getElementById('sizeSvg');

            if (!data || data.length === 0) {
                compSvg.innerHTML = `<text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" fill="#64748b" font-size="13">No historical telemetry points available. Enable Repository DB to begin recording.</text>`;
                sizeSvg.innerHTML = `<text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" fill="#64748b" font-size="13">No volume telemetry recorded.</text>`;
                return;
            }

            const latest = data[data.length - 1];
            document.getElementById('trendLatestCompliance').innerText = `Current: ${latest.CompliancePct}% (${data.length} snapshots recorded)`;
            document.getElementById('trendLatestSize').innerText = `Current: ${latest.TotalBackupSizeGB} GB`;

            drawSvgLineChart(compSvg, data, 'CompliancePct', '#38bdf8', '%', 0, 100);
            const maxSize = Math.max(...data.map(d => d.TotalBackupSizeGB), 10);
            drawSvgLineChart(sizeSvg, data, 'TotalBackupSizeGB', '#a78bfa', ' GB', 0, maxSize * 1.15);
        }

        function drawSvgLineChart(svgEl, data, field, strokeColor, unit, minVal, maxVal) {
            const w = svgEl.clientWidth || 800;
            const h = 170;
            const padL = 45;
            const padR = 25;
            const padT = 15;
            const padB = 25;

            const innerW = w - padL - padR;
            const innerH = h - padT - padB;

            let pts = [];
            data.forEach((d, idx) => {
                const x = padL + (idx / Math.max(data.length - 1, 1)) * innerW;
                const val = d[field];
                const normalizedY = (val - minVal) / Math.max(maxVal - minVal, 1);
                const y = padT + innerH - (normalizedY * innerH);
                pts.push({ x, y, val, time: d.SnapshotTime });
            });

            const pointsStr = pts.map(p => `${p.x},${p.y}`).join(' ');

            let html = `
                <line x1="${padL}" y1="${padT}" x2="${w - padR}" y2="${padT}" stroke="rgba(255,255,255,0.06)" stroke-dasharray="4"/>
                <line x1="${padL}" y1="${padT + innerH/2}" x2="${w - padR}" y2="${padT + innerH/2}" stroke="rgba(255,255,255,0.06)" stroke-dasharray="4"/>
                <line x1="${padL}" y1="${padT + innerH}" x2="${w - padR}" y2="${padT + innerH}" stroke="rgba(255,255,255,0.1)"/>
                
                <text x="${padL - 8}" y="${padT + 4}" fill="#64748b" font-size="10" text-anchor="end">${Math.round(maxVal)}${unit}</text>
                <text x="${padL - 8}" y="${padT + innerH/2 + 4}" fill="#64748b" font-size="10" text-anchor="end">${Math.round((maxVal+minVal)/2)}${unit}</text>
                <text x="${padL - 8}" y="${padT + innerH + 4}" fill="#64748b" font-size="10" text-anchor="end">${Math.round(minVal)}${unit}</text>

                <polygon points="${padL},${padT + innerH} ${pointsStr} ${pts[pts.length-1].x},${padT + innerH}" fill="${strokeColor}" fill-opacity="0.12"/>
                <polyline points="${pointsStr}" fill="none" stroke="${strokeColor}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>
            `;

            pts.forEach((p, idx) => {
                if (idx % Math.ceil(pts.length / 8) === 0 || idx === pts.length - 1) {
                    const shortTime = p.time.length > 10 ? p.time.substring(5, 16) : p.time;
                    html += `
                        <circle cx="${p.x}" cy="${p.y}" r="3.5" fill="${strokeColor}" stroke="#0b0f19" stroke-width="2">
                            <title>${p.time}: ${p.val}${unit}</title>
                        </circle>
                        <text x="${p.x}" y="${h - 4}" fill="#64748b" font-size="9" text-anchor="middle">${shortTime}</text>
                    `;
                }
            });

            svgEl.innerHTML = html;
        }

        function populateServerFilterDropdown() {
            const select = document.getElementById('serverFilter');
            const currentVal = select.value;
            const uniqueServers = [...new Set(allDatabases.map(d => d.ServerName))];
            select.innerHTML = '<option value="ALL">All Servers</option>';
            uniqueServers.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s;
                opt.innerText = s;
                select.appendChild(opt);
            });
            select.value = currentVal;
        }

        let expandedServerNames = new Set();
        let cachedServerGroupsList = [];

        function filterByKpi(status) {
            const select = document.getElementById('statusFilter');
            if (select) {
                select.value = status;
                renderDatabaseTable(true);
            }
        }

        function toggleServerCard(headerEl) {
            const card = headerEl.closest('.server-card');
            if (!card) return;
            const isNowExpanded = card.classList.toggle('expanded');
            const srvName = card.getAttribute('data-server-name');
            if (srvName) {
                if (isNowExpanded) {
                    expandedServerNames.add(srvName);
                } else {
                    expandedServerNames.delete(srvName);
                }
            }
            const actionLabel = card.querySelector('.server-action-label');
            if (actionLabel) {
                actionLabel.innerText = isNowExpanded ? 'Collapse ▲' : 'View Databases ▼';
            }
            const allCards = document.querySelectorAll('.server-card');
            const btn = document.getElementById('lblToggleAll');
            if (btn && allCards.length > 0) {
                const anyCollapsed = Array.from(allCards).some(c => !c.classList.contains('expanded'));
                btn.innerText = anyCollapsed ? 'Expand All' : 'Collapse All';
            }
        }

        function toggleAllServerAccordions() {
            const cards = document.querySelectorAll('.server-card');
            const btn = document.getElementById('lblToggleAll');
            if (!cards || cards.length === 0) return;

            const anyCollapsed = Array.from(cards).some(c => !c.classList.contains('expanded'));

            cards.forEach(card => {
                const srvName = card.getAttribute('data-server-name');
                const actionLabel = card.querySelector('.server-action-label');
                if (anyCollapsed) {
                    card.classList.add('expanded');
                    if (srvName) expandedServerNames.add(srvName);
                    if (actionLabel) actionLabel.innerText = 'Collapse ▲';
                } else {
                    card.classList.remove('expanded');
                    if (srvName) expandedServerNames.delete(srvName);
                    if (actionLabel) actionLabel.innerText = 'View Databases ▼';
                }
            });

            if (btn) btn.innerText = anyCollapsed ? 'Collapse All' : 'Expand All';
        }

        function changeDbPage(action) {
            const totalPages = Math.max(1, Math.ceil(cachedServerGroupsList.length / dbPageSize));
            if (action === 'first') dbCurrentPage = 1;
            else if (action === 'prev') dbCurrentPage = Math.max(1, dbCurrentPage - 1);
            else if (action === 'next') dbCurrentPage = Math.min(totalPages, dbCurrentPage + 1);
            else if (action === 'last') dbCurrentPage = totalPages;
            renderDatabaseTable(false);
        }

        function renderDatabaseTable(resetPage = true) {
            if (resetPage) dbCurrentPage = 1;
            const container = document.getElementById('serverAccordionContainer');
            if (!container) return;

            const search = (document.getElementById('searchInput').value || '').toLowerCase();
            const statusFilter = document.getElementById('statusFilter').value;
            const serverFilter = document.getElementById('serverFilter').value;

            filteredDatabases = allDatabases.filter(d => {
                const matchSearch = (d.DatabaseName && d.DatabaseName.toLowerCase().includes(search)) || 
                                     (d.ServerName && d.ServerName.toLowerCase().includes(search));
                const matchStatus = (statusFilter === 'ALL') || (d.Status === statusFilter);
                const matchServer = (serverFilter === 'ALL') || (d.ServerName === serverFilter);
                return matchSearch && matchStatus && matchServer;
            });

            // If user is searching, auto-expand servers that contain the search hits
            if (search.length > 0) {
                filteredDatabases.forEach(d => expandedServerNames.add(d.ServerName));
            }

            // Group filtered databases by Server
            const serverMap = {};
            filteredDatabases.forEach(d => {
                const sName = d.ServerName || 'Unknown Instance';
                if (!serverMap[sName]) serverMap[sName] = [];
                serverMap[sName].push(d);
            });

            cachedServerGroupsList = Object.keys(serverMap).map(sName => {
                const dbs = serverMap[sName];
                const critCount = dbs.filter(d => d.Status === 'Critical').length;
                const warnCount = dbs.filter(d => d.Status === 'Warning').length;
                const healthyCount = dbs.filter(d => d.Status === 'Healthy').length;
                const totalSizeGB = dbs.reduce((sum, d) => sum + (parseFloat(d.LastBackupSizeGB) || 0), 0).toFixed(2);
                const serverStatus = critCount > 0 ? 'Critical' : (warnCount > 0 ? 'Warning' : 'Healthy');
                const srvObj = allServers.find(s => s.Name === sName || s.ServerAddress === sName);
                const groupName = srvObj ? (srvObj.GroupName || 'Default') : 'Default';

                return {
                    name: sName,
                    status: serverStatus,
                    group: groupName,
                    critCount,
                    warnCount,
                    healthyCount,
                    totalDbs: dbs.length,
                    totalSizeGB,
                    databases: dbs
                };
            });

            // Sort: Critical SLA breach servers first, then Warning, then Healthy
            cachedServerGroupsList.sort((a, b) => {
                const score = s => s.status === 'Critical' ? 3 : (s.status === 'Warning' ? 2 : 1);
                return score(b) - score(a) || a.name.localeCompare(b.name);
            });

            const totalServers = cachedServerGroupsList.length;
            const totalPages = Math.max(1, Math.ceil(totalServers / dbPageSize));
            if (dbCurrentPage > totalPages) dbCurrentPage = totalPages;

            const startIdx = (dbCurrentPage - 1) * dbPageSize;
            const endIdx = Math.min(startIdx + dbPageSize, totalServers);
            const pagedServers = cachedServerGroupsList.slice(startIdx, endIdx);

            const pagEl = document.getElementById('dbPagination');
            if (pagEl) {
                pagEl.style.display = totalServers > dbPageSize ? 'flex' : 'none';
                document.getElementById('dbPaginationInfo').innerText = totalServers > 0 
                    ? `Showing ${startIdx + 1}–${endIdx} of ${totalServers} servers (${filteredDatabases.length} databases)` 
                    : `Showing 0 of 0 servers`;
                document.getElementById('dbPageIndicator').innerText = `Page ${dbCurrentPage} of ${totalPages}`;
                document.getElementById('dbFirstBtn').disabled = (dbCurrentPage <= 1);
                document.getElementById('dbPrevBtn').disabled = (dbCurrentPage <= 1);
                document.getElementById('dbNextBtn').disabled = (dbCurrentPage >= totalPages);
                document.getElementById('dbLastBtn').disabled = (dbCurrentPage >= totalPages);
            }

            const btnToggleAll = document.getElementById('lblToggleAll');
            if (btnToggleAll && totalServers > 0) {
                btnToggleAll.innerText = expandedServerNames.size >= totalServers ? 'Collapse All' : 'Expand All';
            }

            if (pagedServers.length === 0) {
                if (allDatabases.length === 0) {
                    if (allServers.length === 0) {
                        container.innerHTML = `<div style="text-align:center; padding: 2.5rem; color: #94a3b8; background:rgba(0,0,0,0.2); border-radius:10px; border:1px solid var(--border-color);">
                            <div style="font-size:1.05rem; font-weight:600; color:#f1f5f9; margin-bottom:6px;">No SQL Server Instances Registered</div>
                            <div style="font-size:0.85rem; margin-bottom:14px; color:var(--text-muted);">Add your monitored database instances in the <strong>SQL Instances</strong> tab to start tracking backup SLA health.</div>
                            <button class="btn-primary" style="display:inline-flex; font-size:0.82rem; margin:0 auto;" onclick="switchTab('servers'); openServerModal();">Add First SQL Instance</button>
                        </div>`;
                    } else {
                        container.innerHTML = `<div style="text-align:center; padding: 2rem; color: var(--text-muted); background:rgba(0,0,0,0.2); border-radius:10px;">No databases discovered on configured instances. Click "Poll Now" to query telemetry.</div>`;
                    }
                } else {
                    container.innerHTML = `<div style="text-align:center; padding: 2rem; color: var(--text-muted); background:rgba(0,0,0,0.2); border-radius:10px;">No servers or databases matching active filters.</div>`;
                }
                return;
            }

            container.innerHTML = pagedServers.map(srv => {
                const isExpanded = expandedServerNames.has(srv.name);
                const cardStatusCls = srv.status === 'Critical' ? 'status-critical' : (srv.status === 'Warning' ? 'status-warning' : 'status-healthy');
                const badgeCls = srv.status === 'Critical' ? 'badge-critical' : (srv.status === 'Warning' ? 'badge-warning' : 'badge-healthy');
                const pulse = srv.status === 'Critical' ? 'red' : (srv.status === 'Warning' ? 'amber' : 'green');
                const statusLabel = srv.status === 'Critical' ? 'Critical SLA Breach' : (srv.status === 'Warning' ? 'Warning Alert' : 'Healthy SLA');

                const dbRowsHtml = srv.databases.map(d => {
                    let dbBadgeCls = 'badge-healthy';
                    let dbPulse = 'green';
                    if (d.Status === 'Warning') { dbBadgeCls = 'badge-warning'; dbPulse = 'amber'; }
                    if (d.Status === 'Critical') { dbBadgeCls = 'badge-critical'; dbPulse = 'red'; }

                    return `
                        <tr>
                            <td style="width:110px;"><span class="badge-status ${dbBadgeCls}"><span class="pulse-dot ${dbPulse}"></span>${d.Status}</span></td>
                            <td style="font-weight:600; color:#38bdf8;">${escapeHtml(d.DatabaseName)}</td>
                            <td><span class="badge-recovery">${escapeHtml(d.RecoveryModel)}</span></td>
                            <td class="mono">${escapeHtml(d.LastFullBackup || '-')}</td>
                            <td class="mono">${escapeHtml(d.LastDifferentialBackup || '-')}</td>
                            <td class="mono">${escapeHtml(d.LastLogBackup || '-')}</td>
                            <td class="mono">${d.LastBackupSizeGB > 0 ? d.LastBackupSizeGB + ' GB' : '-'}</td>
                            <td style="color: ${d.Status === 'Critical' ? '#f87171' : (d.Status === 'Warning' ? '#fbbf24' : '#9ca3af')}; font-size:0.8rem;">
                                ${escapeHtml(d.StatusReason || '')}
                            </td>
                        </tr>
                    `;
                }).join('');

                return `
                    <div class="server-card ${cardStatusCls} ${isExpanded ? 'expanded' : ''}" data-server-name="${escapeHtml(srv.name)}">
                        <div class="server-card-header" onclick="toggleServerCard(this)">
                            <div class="server-title-group">
                                <span class="server-chevron-icon">
                                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="9 18 15 12 9 6"></polyline></svg>
                                </span>
                                <div>
                                    <span class="server-name-text">${escapeHtml(srv.name)}</span>
                                </div>
                                <span class="badge-recovery" style="margin-left:4px;">${escapeHtml(srv.group)}</span>
                            </div>

                            <div class="server-stats-group">
                                <span class="badge-status ${badgeCls}"><span class="pulse-dot ${pulse}"></span>${statusLabel}</span>
                                ${srv.critCount > 0 ? `<span class="badge-status badge-critical" style="font-size:0.7rem; padding:2px 7px;">${srv.critCount} Critical</span>` : ''}
                                ${srv.warnCount > 0 ? `<span class="badge-status badge-warning" style="font-size:0.7rem; padding:2px 7px;">${srv.warnCount} Warning</span>` : ''}
                                ${srv.healthyCount > 0 ? `<span class="badge-status badge-healthy" style="font-size:0.7rem; padding:2px 7px;">${srv.healthyCount} Healthy</span>` : ''}
                                <span style="font-size:0.76rem; color:var(--text-muted); font-family:var(--font-mono); margin-left:4px;">${srv.totalDbs} DBs &bull; ${srv.totalSizeGB} GB</span>
                                <span class="server-action-label" style="font-size:0.75rem; color:var(--accent-blue); display:inline-flex; align-items:center; gap:4px; margin-left:6px; font-weight:500;">
                                    ${isExpanded ? 'Collapse ▲' : 'View Databases ▼'}
                                </span>
                            </div>
                        </div>

                        <div class="server-card-body">
                            <div class="table-responsive" style="margin:0; border:none; border-radius:0;">
                                <table style="width:100%; border-collapse:collapse;">
                                    <thead>
                                        <tr>
                                            <th style="padding-left:1.25rem;">Status</th>
                                            <th>Database Name</th>
                                            <th>Recovery</th>
                                            <th>Last Full Backup</th>
                                            <th>Last Diff Backup</th>
                                            <th>Last Log Backup</th>
                                            <th>Backup Size</th>
                                            <th>SLA Diagnostic Reason</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        ${dbRowsHtml}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                `;
            }).join('');
        }

        function renderAlwaysOnTable() {
            const tbody = document.getElementById('alwaysOnTableBody');
            if (!allAlwaysOn || allAlwaysOn.length === 0) {
                tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; padding:2rem; color:var(--text-muted)">No AlwaysOn AG clusters discovered or HADR not enabled.</td></tr>`;
                return;
            }
            tbody.innerHTML = allAlwaysOn.map(a => `
                <tr>
                    <td style="font-weight:600">${escapeHtml(a.ServerName)}</td>
                    <td style="color:#38bdf8; font-weight:500;">${escapeHtml(a.GroupName)}</td>
                    <td class="mono">${escapeHtml(a.ReplicaServerName)}</td>
                    <td><span class="badge-recovery">${escapeHtml(a.Role)}</span></td>
                    <td>${escapeHtml(a.BackupPreference)}</td>
                    <td><span class="badge-status ${a.SyncHealth === 'HEALTHY' ? 'badge-healthy' : 'badge-warning'}">${escapeHtml(a.SyncHealth)}</span></td>
                    <td style="font-weight:500;">${escapeHtml(a.DatabaseName)}</td>
                    <td><span class="mono">${escapeHtml(a.DbSyncState)}</span></td>
                </tr>
            `).join('');
        }

        function renderAlertsTable() {
            const tbody = document.getElementById('alertsTableBody');
            if (!allAlerts || allAlerts.length === 0) {
                tbody.innerHTML = `<tr><td colspan="6" style="text-align:center; padding:2rem; color:var(--text-muted)">All database backups are currently within SLA policy limits.</td></tr>`;
                return;
            }
            tbody.innerHTML = allAlerts.map(a => `
                <tr>
                    <td class="mono" style="font-size:0.8rem; color:#94a3b8;">${escapeHtml(a.Timestamp)}</td>
                    <td><span class="badge-status ${a.Severity === 'Critical' ? 'badge-critical' : 'badge-warning'}">${a.Severity}</span></td>
                    <td style="font-weight:600">${escapeHtml(a.ServerName)}</td>
                    <td style="font-weight:500">${escapeHtml(a.DatabaseName)}</td>
                    <td style="color:#fca5a5; font-size:0.82rem;">${escapeHtml(a.Message)}</td>
                    <td><button class="btn-secondary" style="padding:3px 8px; font-size:0.72rem;" onclick="dismissAlert('${a.Id}')">Dismiss</button></td>
                </tr>
            `).join('');
        }

        function renderAuditTable() {
            const tbody = document.getElementById('auditTableBody');
            if (!allAudit || allAudit.length === 0) {
                tbody.innerHTML = `<tr><td colspan="4" style="text-align:center; padding:2rem; color:var(--text-muted)">No system activity recorded yet.</td></tr>`;
                return;
            }
            tbody.innerHTML = allAudit.map(u => `
                <tr>
                    <td class="mono" style="font-size:0.8rem; color:#94a3b8;">${escapeHtml(u.Timestamp)}</td>
                    <td style="color:#60a5fa; font-weight:600;">${escapeHtml(u.User)}</td>
                    <td style="font-weight:500;">${escapeHtml(u.Action)}</td>
                    <td style="color:var(--text-muted); font-size:0.82rem;">${escapeHtml(u.Details)}</td>
                </tr>
            `).join('');
        }

        function renderGroupPoliciesTable() {
            const tbody = document.getElementById('groupPoliciesTableBody');
            if (!tbody) return;
            const groups = (globalConfig && globalConfig.GroupPolicies) ? globalConfig.GroupPolicies : [];
            if (!groups || groups.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding:1.5rem; color:var(--text-muted)">No Group SLA Policies defined. Servers inherit the default Global Policy.</td></tr>`;
                return;
            }

            tbody.innerHTML = groups.map(g => `
                <tr>
                    <td style="font-weight:600; color:#38bdf8;">
                        <span style="display:inline-flex; align-items:center; gap:6px;">
                            <span style="display:inline-block; width:8px; height:8px; border-radius:50%; background:#38bdf8;"></span>
                            ${escapeHtml(g.GroupName)}
                        </span>
                    </td>
                    <td class="mono">${g.FullBackupWarningHours}h <span style="color:#64748b">/</span> <strong style="color:#f87171">${g.FullBackupCriticalHours}h</strong></td>
                    <td class="mono">${g.DiffBackupWarningHours}h <span style="color:#64748b">/</span> <strong style="color:#f87171">${g.DiffBackupCriticalHours}h</strong></td>
                    <td class="mono">${g.LogBackupWarningMinutes}m <span style="color:#64748b">/</span> <strong style="color:#f87171">${g.LogBackupCriticalMinutes}m</strong></td>
                    <td>
                        <button class="btn-secondary" style="padding:3px 8px; font-size:0.75rem; margin-right:4px;" onclick='openGroupPolicyModal(${JSON.stringify(g)})'>Edit</button>
                        <button class="btn-secondary" style="padding:3px 8px; font-size:0.75rem; color:#f87171;" onclick='deleteGroupPolicy(${JSON.stringify(g.GroupName)})'>Delete</button>
                    </td>
                </tr>
            `).join('');
        }

        function openGroupPolicyModal(groupPolicy = null) {
            document.getElementById('groupPolicyModal').classList.add('active');
            if (groupPolicy) {
                document.getElementById('grpModalTitle').innerText = 'Edit Group SLA Policy: ' + groupPolicy.GroupName;
                document.getElementById('grpOriginalName').value = groupPolicy.GroupName || '';
                document.getElementById('grpName').value = groupPolicy.GroupName || '';
                document.getElementById('grpFullWarn').value = groupPolicy.FullBackupWarningHours || 20;
                document.getElementById('grpFullCrit').value = groupPolicy.FullBackupCriticalHours || 24;
                document.getElementById('grpDiffWarn').value = groupPolicy.DiffBackupWarningHours || 10;
                document.getElementById('grpDiffCrit').value = groupPolicy.DiffBackupCriticalHours || 12;
                document.getElementById('grpLogWarn').value = groupPolicy.LogBackupWarningMinutes || 20;
                document.getElementById('grpLogCrit').value = groupPolicy.LogBackupCriticalMinutes || 30;
            } else {
                document.getElementById('grpModalTitle').innerText = 'Add Group SLA Policy';
                document.getElementById('grpOriginalName').value = '';
                document.getElementById('grpName').value = '';
                document.getElementById('grpFullWarn').value = 20;
                document.getElementById('grpFullCrit').value = 24;
                document.getElementById('grpDiffWarn').value = 10;
                document.getElementById('grpDiffCrit').value = 12;
                document.getElementById('grpLogWarn').value = 20;
                document.getElementById('grpLogCrit').value = 30;
            }
        }

        function closeGroupPolicyModal() {
            document.getElementById('groupPolicyModal').classList.remove('active');
        }

        async function saveGroupPolicy(e) {
            e.preventDefault();
            const originalName = document.getElementById('grpOriginalName').value.trim();
            const groupName = document.getElementById('grpName').value.trim();
            if (!groupName) {
                alert("Please enter a valid Group Name.");
                return;
            }

            const policy = {
                OriginalName: originalName,
                GroupName: groupName,
                FullBackupWarningHours: parseInt(document.getElementById('grpFullWarn').value) || 20,
                FullBackupCriticalHours: parseInt(document.getElementById('grpFullCrit').value) || 24,
                DiffBackupWarningHours: parseInt(document.getElementById('grpDiffWarn').value) || 10,
                DiffBackupCriticalHours: parseInt(document.getElementById('grpDiffCrit').value) || 12,
                LogBackupWarningMinutes: parseInt(document.getElementById('grpLogWarn').value) || 20,
                LogBackupCriticalMinutes: parseInt(document.getElementById('grpLogCrit').value) || 30,
                IgnoreSimpleRecoveryLogs: true
            };

            try {
                const res = await fetch('/api/policies/groups', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(policy)
                });

                if (res.ok) {
                    closeGroupPolicyModal();
                    await fetchConfig();
                    triggerScan();
                } else {
                    const err = await res.text();
                    alert("Failed to save group policy: " + err);
                }
            } catch (err) {
                alert("Error saving group policy: " + err.message);
            }
        }

        async function deleteGroupPolicy(groupName) {
            if (!groupName) return;
            if (!confirm(`Are you sure you want to delete the SLA policy for group "${groupName}"? Servers in this group will fall back to the default Global SLA policy.`)) return;

            try {
                let res = await fetch(`/api/policies/groups/${encodeURIComponent(groupName)}`, { method: 'DELETE' });
                if (!res.ok) {
                    res = await fetch('/api/policies/groups/delete', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ GroupName: groupName })
                    });
                }
                if (res.ok) {
                    await fetchConfig();
                    triggerScan();
                } else {
                    alert("Failed to delete group policy.");
                }
            } catch (err) {
                alert("Error deleting group policy: " + err.message);
            }
        }

        function toggleServerCustomSla() {
            const isCustom = document.getElementById('srvUseCustomSla').checked;
            document.getElementById('srvCustomSlaFields').style.display = isCustom ? 'block' : 'none';
            updateServerSlaNotice();
        }

        function updateServerSlaNotice() {
            const isCustom = document.getElementById('srvUseCustomSla').checked;
            const noticeEl = document.getElementById('srvCustomSlaNotice');
            if (!noticeEl) return;
            if (isCustom) {
                noticeEl.innerHTML = '<span style="color:#38bdf8; font-weight:600;">Custom SLA Active:</span> This server will use the instance-specific thresholds below, overriding all group and global policies.';
                return;
            }

            const grpName = (document.getElementById('srvGroup').value || '').trim().toLowerCase();
            const matchedGrp = (globalConfig && globalConfig.GroupPolicies) ? globalConfig.GroupPolicies.find(g => g.GroupName && g.GroupName.trim().toLowerCase() === grpName) : null;

            if (matchedGrp) {
                noticeEl.innerHTML = `Inheriting from Group policy <strong style="color:#38bdf8;">"${escapeHtml(matchedGrp.GroupName)}"</strong>: Full &gt; ${matchedGrp.FullBackupCriticalHours}h, Diff &gt; ${matchedGrp.DiffBackupCriticalHours}h, Log &gt; ${matchedGrp.LogBackupCriticalMinutes}m.`;
            } else {
                const glob = (globalConfig && globalConfig.GlobalPolicies) ? globalConfig.GlobalPolicies : { FullBackupCriticalHours: 24, DiffBackupCriticalHours: 12, LogBackupCriticalMinutes: 30 };
                noticeEl.innerHTML = `Inheriting <strong style="color:#a78bfa;">Global Default SLA</strong>: Full &gt; ${glob.FullBackupCriticalHours}h, Diff &gt; ${glob.DiffBackupCriticalHours}h, Log &gt; ${glob.LogBackupCriticalMinutes}m.`;
            }
        }

        function renderServersTable() {
            const tbody = document.getElementById('serversTableBody');
            if (!allServers || allServers.length === 0) {
                tbody.innerHTML = `<tr><td colspan="6" style="text-align:center; padding:2rem; color:var(--text-muted)">No SQL servers configured yet. Click "Add SQL Instance" above.</td></tr>`;
                return;
            }

            tbody.innerHTML = allServers.map(s => {
                const srvId = s.Id || s.Name;
                const slaBadge = s.UseCustomSla 
                    ? `<span class="badge-status badge-warning" style="font-size:0.68rem; padding:1px 5px; margin-left:4px;" title="Instance Custom SLA Override">Custom SLA</span>`
                    : ``;
                return `
                <tr>
                    <td style="font-weight:600">${escapeHtml(s.Name)}</td>
                    <td class="mono">${escapeHtml(s.ServerAddress)}:${s.Port || 1433}</td>
                    <td>${s.AuthType === 'SqlPassword' ? 'SQL Auth (' + escapeHtml(s.Username) + ')' : 'Windows Integrated'}</td>
                    <td><span class="badge-recovery">${escapeHtml(s.GroupName || 'Default')} / ${escapeHtml(s.Environment || 'Prod')}</span>${slaBadge}</td>
                    <td><span class="badge-status ${s.IsEnabled ? 'badge-healthy' : 'badge-warning'}">${s.IsEnabled ? 'Active' : 'Disabled'}</span></td>
                    <td>
                        <button class="btn-secondary" style="padding:3px 8px; font-size:0.75rem; margin-right:4px; display:inline-flex;" onclick='testSingleServer(${JSON.stringify(srvId)})'>Test</button>
                        <button class="btn-secondary" style="padding:3px 8px; font-size:0.75rem; margin-right:4px; display:inline-flex;" onclick='editServer(${JSON.stringify(srvId)})'>Edit</button>
                        <button class="btn-secondary" style="padding:3px 8px; font-size:0.75rem; color:#f87171; display:inline-flex;" onclick='deleteServer(${JSON.stringify(srvId)})'>Delete</button>
                    </td>
                </tr>
                `;
            }).join('');
        }

        async function triggerScan() {
            const btn = document.getElementById('btnRefresh');
            btn.disabled = true;
            btn.innerText = "Scanning...";
            try {
                await fetch('/api/scan', { method: 'POST' });
                await refreshAll();
            } finally {
                btn.disabled = false;
                btn.innerText = "Poll Now";
            }
        }

        async function refreshAll() {
            await Promise.all([fetchMetrics(), fetchDatabases(), fetchServers(), fetchConfig()]);
            if (document.getElementById('tab-trends').classList.contains('active')) { fetchTrends(); fetchHistoryLogs(); }
            if (document.getElementById('tab-alwayson').classList.contains('active')) { fetchAlwaysOn(); }
            if (document.getElementById('tab-alerts').classList.contains('active')) { fetchAlerts(); }
            if (document.getElementById('tab-audit').classList.contains('active')) { fetchAuditLogs(); }
        }

        function exportReportHtml() { window.open('/api/export/html', '_blank'); }
        function exportReportPdf() { window.open('/api/export/pdf', '_blank'); }

        async function emailPdfReportNow() {
            if (!confirm("Generate executive PDF report and email to configured recipients now?")) return;
            try {
                const res = await fetch('/api/email/report', { method: 'POST' });
                const r = await res.json();
                if (r.Success) {
                    alert("PDF Report generated and emailed successfully!");
                } else {
                    alert("Failed to send report: " + r.ErrorMessage);
                }
            } catch (err) { alert("Error: " + err.message); }
        }

        function dismissAlert(alertId) {
            allAlerts = allAlerts.filter(a => a.Id !== alertId);
            renderAlertsTable();
        }

        function toggleAuthFields() {
            const authType = document.getElementById('srvAuthType').value;
            document.getElementById('sqlAuthBox').style.display = (authType === 'SqlPassword') ? 'block' : 'none';
        }

        function toggleRepoAuthFields() {
            const authType = document.getElementById('repoAuthType').value;
            document.getElementById('repoSqlAuthBox').style.display = (authType === 'SqlPassword') ? 'block' : 'none';
        }

        function openServerModal(server = null) {
            document.getElementById('serverModal').classList.add('active');
            if (server) {
                document.getElementById('modalTitle').innerText = 'Edit SQL Server';
                document.getElementById('srvId').value = server.Id || '';
                document.getElementById('srvName').value = server.Name || '';
                document.getElementById('srvAddress').value = server.ServerAddress || 'localhost';
                document.getElementById('srvPort').value = server.Port || 1433;
                document.getElementById('srvEncryption').value = server.Encryption || 'Optional';
                document.getElementById('srvTrustCert').checked = (server.TrustServerCertificate !== false);
                document.getElementById('srvAuthType').value = server.AuthType || 'Windows';
                document.getElementById('srvUsername').value = server.Username || '';
                document.getElementById('srvPassword').value = server.Password || '';
                document.getElementById('srvGroup').value = server.GroupName || 'Production';
                document.getElementById('srvEnv').value = server.Environment || 'Production';

                const isCustom = !!server.UseCustomSla;
                document.getElementById('srvUseCustomSla').checked = isCustom;
                if (server.CustomPolicies) {
                    const cp = server.CustomPolicies;
                    document.getElementById('srvFullWarn').value = cp.FullBackupWarningHours || 20;
                    document.getElementById('srvFullCrit').value = cp.FullBackupCriticalHours || 24;
                    document.getElementById('srvDiffWarn').value = cp.DiffBackupWarningHours || 10;
                    document.getElementById('srvDiffCrit').value = cp.DiffBackupCriticalHours || 12;
                    document.getElementById('srvLogWarn').value = cp.LogBackupWarningMinutes || 20;
                    document.getElementById('srvLogCrit').value = cp.LogBackupCriticalMinutes || 30;
                } else {
                    document.getElementById('srvFullWarn').value = 20;
                    document.getElementById('srvFullCrit').value = 24;
                    document.getElementById('srvDiffWarn').value = 10;
                    document.getElementById('srvDiffCrit').value = 12;
                    document.getElementById('srvLogWarn').value = 20;
                    document.getElementById('srvLogCrit').value = 30;
                }
            } else {
                document.getElementById('modalTitle').innerText = 'Add SQL Server';
                document.getElementById('srvId').value = '';
                document.getElementById('srvName').value = '';
                document.getElementById('srvAddress').value = 'localhost';
                document.getElementById('srvPort').value = 1433;
                document.getElementById('srvEncryption').value = 'Optional';
                document.getElementById('srvTrustCert').checked = true;
                document.getElementById('srvAuthType').value = 'Windows';
                document.getElementById('srvUsername').value = '';
                document.getElementById('srvPassword').value = '';
                document.getElementById('srvGroup').value = 'Production';
                document.getElementById('srvEnv').value = 'Production';
                document.getElementById('srvUseCustomSla').checked = false;
                document.getElementById('srvFullWarn').value = 20;
                document.getElementById('srvFullCrit').value = 24;
                document.getElementById('srvDiffWarn').value = 10;
                document.getElementById('srvDiffCrit').value = 12;
                document.getElementById('srvLogWarn').value = 20;
                document.getElementById('srvLogCrit').value = 30;
            }
            toggleAuthFields();
            toggleServerCustomSla();
            updateServerSlaNotice();
        }

        function closeServerModal() { document.getElementById('serverModal').classList.remove('active'); }

        function editServer(id) {
            const srv = allServers.find(s => s.Id === id || s.Name === id);
            if (srv) openServerModal(srv);
        }

        async function saveServer(e) {
            e.preventDefault();
            const id = document.getElementById('srvId').value;
            const useCustomSla = document.getElementById('srvUseCustomSla').checked;
            const customPolicies = useCustomSla ? {
                FullBackupWarningHours: parseInt(document.getElementById('srvFullWarn').value) || 20,
                FullBackupCriticalHours: parseInt(document.getElementById('srvFullCrit').value) || 24,
                DiffBackupWarningHours: parseInt(document.getElementById('srvDiffWarn').value) || 10,
                DiffBackupCriticalHours: parseInt(document.getElementById('srvDiffCrit').value) || 12,
                LogBackupWarningMinutes: parseInt(document.getElementById('srvLogWarn').value) || 20,
                LogBackupCriticalMinutes: parseInt(document.getElementById('srvLogCrit').value) || 30,
                IgnoreSimpleRecoveryLogs: true
            } : null;

            const payload = {
                Id: id || null,
                Name: document.getElementById('srvName').value.trim(),
                ServerAddress: document.getElementById('srvAddress').value.trim(),
                Port: parseInt(document.getElementById('srvPort').value) || 1433,
                Encryption: document.getElementById('srvEncryption').value,
                TrustServerCertificate: document.getElementById('srvTrustCert').checked,
                AuthType: document.getElementById('srvAuthType').value,
                Username: document.getElementById('srvUsername').value.trim(),
                Password: document.getElementById('srvPassword').value,
                GroupName: document.getElementById('srvGroup').value.trim() || 'Production',
                Environment: document.getElementById('srvEnv').value.trim() || 'Production',
                UseCustomSla: useCustomSla,
                CustomPolicies: customPolicies,
                IsEnabled: true
            };

            if (!payload.Name || !payload.ServerAddress) {
                alert("Please provide a valid Instance Name and Server Address.");
                return;
            }

            const method = id ? 'PUT' : 'POST';
            const url = id ? `/api/servers/${encodeURIComponent(id)}` : '/api/servers';

            try {
                const res = await fetch(url, {
                    method: method,
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                if (res.ok) {
                    closeServerModal();
                    await refreshAll();
                } else {
                    const err = await res.text();
                    alert("Failed to save server instance: " + err);
                }
            } catch (err) {
                alert("Error saving server: " + err.message);
            }
        }

        async function deleteServer(id) {
            if (!id) {
                alert("Invalid server identifier.");
                return;
            }
            if (!confirm("Are you sure you want to remove this SQL instance?")) return;
            try {
                let res = await fetch(`/api/servers/${encodeURIComponent(id)}`, { method: 'DELETE' });
                if (!res.ok) {
                    res = await fetch('/api/servers/delete', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ Id: id })
                    });
                }
                if (res.ok) {
                    await refreshAll();
                } else {
                    alert("Failed to remove server.");
                }
            } catch (err) {
                alert("Error removing server: " + err.message);
            }
        }

        async function testModalConnection() {
            const payload = {
                Name: document.getElementById('srvName').value || 'TestServer',
                ServerAddress: document.getElementById('srvAddress').value,
                Port: parseInt(document.getElementById('srvPort').value) || 1433,
                Encryption: document.getElementById('srvEncryption').value,
                TrustServerCertificate: document.getElementById('srvTrustCert').checked,
                AuthType: document.getElementById('srvAuthType').value,
                Username: document.getElementById('srvUsername').value,
                Password: document.getElementById('srvPassword').value
            };

            try {
                const res = await fetch('/api/servers/test', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const r = await res.json();
                if (r.Success) {
                    alert(`[SUCCESS] Connection Successful!\nVersion: ${r.ProductVersion}\nEdition: ${r.Edition}\nAlwaysOn Enabled: ${r.IsHadrEnabled ? 'YES' : 'NO'}\nLatency: ${r.LatencyMs}ms`);
                } else {
                    alert(`[FAILED] Connection Failed:\n${r.ErrorMessage}`);
                }
            } catch (err) { alert("Connection test error: " + err.message); }
        }

        async function testSingleServer(id) {
            const srv = allServers.find(s => s.Id === id);
            if (!srv) return;
            try {
                const res = await fetch('/api/servers/test', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(srv)
                });
                const r = await res.json();
                if (r.Success) {
                    alert(`[SUCCESS] [${srv.Name}] Connection Successful!\nVersion: ${r.ProductVersion}\nLatency: ${r.LatencyMs}ms`);
                } else {
                    alert(`[FAILED] [${srv.Name}] Connection Failed:\n${r.ErrorMessage}`);
                }
            } catch (err) { alert("Test failed: " + err.message); }
        }

        async function savePolicies(e) {
            e.preventDefault();
            const pullSec = parseInt(document.getElementById('cfgAutoRefresh').value) || 60;
            const policies = {
                FullBackupWarningHours: parseInt(document.getElementById('cfgFullWarn').value),
                FullBackupCriticalHours: parseInt(document.getElementById('cfgFullCrit').value),
                DiffBackupWarningHours: parseInt(document.getElementById('cfgDiffWarn').value),
                DiffBackupCriticalHours: parseInt(document.getElementById('cfgDiffCrit').value),
                LogBackupWarningMinutes: parseInt(document.getElementById('cfgLogWarn').value),
                LogBackupCriticalMinutes: parseInt(document.getElementById('cfgLogCrit').value),
                AlertOnMissingBackups: true,
                IgnoreSimpleRecoveryLogs: true,
                AutoRefreshIntervalSec: pullSec
            };

            const res = await fetch('/api/config', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ GlobalPolicies: policies })
            });

            if (res.ok) {
                if (typeof updateAutoRefreshInterval === 'function') updateAutoRefreshInterval(pullSec);
                alert(`SLA Policies & Telemetry Pull Interval (${pullSec}s) saved successfully!`);
                triggerScan();
            } else {
                alert("Failed to save SLA policies.");
            }
        }

        async function saveRepoSettings(e) {
            e.preventDefault();
            const repoCfg = {
                IsEnabled: document.getElementById('repoEnabled').checked,
                ServerAddress: document.getElementById('repoServer').value.trim(),
                Port: parseInt(document.getElementById('repoPort').value) || 1433,
                DatabaseName: document.getElementById('repoDatabase').value.trim() || 'SQLBackupMonitorDB',
                RetentionDays: parseInt(document.getElementById('repoRetention').value) || 90,
                Encryption: document.getElementById('repoEncryption').value,
                TrustServerCertificate: document.getElementById('repoTrustCert').checked,
                AuthType: document.getElementById('repoAuthType').value,
                Username: document.getElementById('repoUsername').value.trim(),
                Password: document.getElementById('repoPassword').value
            };

            const res = await fetch('/api/repository', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(repoCfg)
            });

            if (res.ok) {
                alert("SQL Server Repository settings saved successfully!");
                await fetchConfig();
            } else { alert("Failed to save repository settings."); }
        }

        async function testRepoConnection() {
            const repoCfg = {
                ServerAddress: document.getElementById('repoServer').value.trim(),
                Port: parseInt(document.getElementById('repoPort').value) || 1433,
                DatabaseName: document.getElementById('repoDatabase').value.trim() || 'SQLBackupMonitorDB',
                Encryption: document.getElementById('repoEncryption').value,
                TrustServerCertificate: document.getElementById('repoTrustCert').checked,
                AuthType: document.getElementById('repoAuthType').value,
                Username: document.getElementById('repoUsername').value.trim(),
                Password: document.getElementById('repoPassword').value
            };

            try {
                const res = await fetch('/api/repository/test', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(repoCfg)
                });
                const r = await res.json();
                if (r.Success) {
                    alert(`[SUCCESS] Connected to Repository SQL Server!\nVersion: ${r.ProductVersion}\nDatabase [${r.DatabaseName}] Exists: ${r.DatabaseExists ? 'YES' : 'NO'}\nSchema Ready: ${r.SchemaReady ? 'YES' : 'NO'}\nLatency: ${r.LatencyMs}ms`);
                } else {
                    alert(`[FAILED] Repository Connection Failed:\n${r.ErrorMessage}`);
                }
            } catch (err) { alert("Repository test error: " + err.message); }
        }

        async function initializeRepoSchema() {
            const repoCfg = {
                ServerAddress: document.getElementById('repoServer').value.trim(),
                Port: parseInt(document.getElementById('repoPort').value) || 1433,
                DatabaseName: document.getElementById('repoDatabase').value.trim() || 'SQLBackupMonitorDB',
                Encryption: document.getElementById('repoEncryption').value,
                TrustServerCertificate: document.getElementById('repoTrustCert').checked,
                AuthType: document.getElementById('repoAuthType').value,
                Username: document.getElementById('repoUsername').value.trim(),
                Password: document.getElementById('repoPassword').value
            };

            try {
                const res = await fetch('/api/repository/init', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(repoCfg)
                });
                const r = await res.json();
                if (r.Success) {
                    alert(`[SUCCESS] ${r.Message}`);
                } else {
                    alert(`[FAILED] Failed to initialize schema:\n${r.ErrorMessage}`);
                }
            } catch (err) { alert("Schema initialization error: " + err.message); }
        }

        async function cleanupRepoData() {
            if (!confirm("Are you sure you want to purge old historical snapshot data and remove snapshots from inactive/deleted servers?")) return;
            try {
                const res = await fetch('/api/repository/cleanup', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ PurgeAll: false })
                });
                const r = await res.json();
                if (r.Success) {
                    alert("[SUCCESS] " + r.Message);
                    fetchTrends();
                    fetchHistoryLogs();
                } else {
                    alert("[FAILED] " + r.ErrorMessage);
                }
            } catch (err) { alert("Cleanup error: " + err.message); }
        }

        async function saveEmailSettings(e) {
            e.preventDefault();
            const emailSettings = {
                IsEnabled: document.getElementById('emailEnabled').checked,
                SmtpServer: document.getElementById('smtpServer').value.trim(),
                SmtpPort: parseInt(document.getElementById('smtpPort').value) || 587,
                SenderEmail: document.getElementById('senderEmail').value.trim(),
                SenderDisplayName: document.getElementById('senderDisplayName').value.trim(),
                SmtpUsername: document.getElementById('smtpUsername').value.trim(),
                SmtpPassword: document.getElementById('smtpPassword').value,
                RecipientEmails: document.getElementById('recipientEmails').value.trim(),
                AlertOnCritical: document.getElementById('alertOnCrit').checked,
                AlertOnWarning: document.getElementById('alertOnWarn').checked,
                EnableSsl: document.getElementById('smtpSsl').checked,
                AttachPdfReport: document.getElementById('emailAttachPdf').checked,
                DailyReportEnabled: document.getElementById('dailyReportEnabled').checked,
                DailyReportTime: document.getElementById('dailyReportTime').value || '08:00',
                CooldownMinutes: parseInt(document.getElementById('cooldownMin').value) || 60
            };

            const res = await fetch('/api/config', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ EmailSettings: emailSettings })
            });

            if (res.ok) { alert("Email & SMTP settings saved successfully!"); }
            else { alert("Failed to save email settings."); }
        }

        async function sendTestEmail() {
            const emailSettings = {
                IsEnabled: true,
                SmtpServer: document.getElementById('smtpServer').value.trim(),
                SmtpPort: parseInt(document.getElementById('smtpPort').value) || 587,
                SenderEmail: document.getElementById('senderEmail').value.trim(),
                SenderDisplayName: document.getElementById('senderDisplayName').value.trim(),
                SmtpUsername: document.getElementById('smtpUsername').value.trim(),
                SmtpPassword: document.getElementById('smtpPassword').value,
                RecipientEmails: document.getElementById('recipientEmails').value.trim(),
                AlertOnCritical: true,
                AlertOnWarning: true,
                EnableSsl: document.getElementById('smtpSsl').checked,
                AttachPdfReport: document.getElementById('emailAttachPdf').checked,
                CooldownMinutes: 0
            };

            if (!emailSettings.SmtpServer || !emailSettings.RecipientEmails) {
                alert("Please specify SMTP Server and Recipient Email(s) before testing.");
                return;
            }

            const btn = event.currentTarget;
            btn.disabled = true;
            btn.innerText = "Sending Test Email with PDF...";

            try {
                const res = await fetch('/api/email/test', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(emailSettings)
                });
                const r = await res.json();
                if (r.Success) {
                    alert("Test alert email (with PDF attachment) sent successfully!\nCheck inbox at: " + emailSettings.RecipientEmails);
                } else {
                    alert("Failed to send test email:\n" + r.ErrorMessage);
                }
            } catch (err) { alert("Error: " + err.message); }
            finally {
                btn.disabled = false;
                btn.innerText = "Send Test Alert Email";
            }
        }

        function escapeHtml(text) {
            if (!text) return '';
            return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
        }

        // Dynamic Telemetry Polling & Auto-Refresh Timers
        let metricsTimer = null;
        let dbTimer = null;

        function updateAutoRefreshInterval(sec) {
            const interval = Math.max(5, parseInt(sec) || 60);
            const lbl = document.getElementById('lblAutoPullSec');
            if (lbl) lbl.innerText = interval + 's';

            if (metricsTimer) clearInterval(metricsTimer);
            if (dbTimer) clearInterval(dbTimer);

            // Metrics update smoothly (every 5-10s or interval), databases poll according to configured interval
            metricsTimer = setInterval(fetchMetrics, Math.max(3000, Math.min(interval * 1000, 10000)));
            dbTimer = setInterval(fetchDatabases, interval * 1000);
        }

        // Initial Load
        refreshAll();
        updateAutoRefreshInterval(60);
    </script>
</body>
</html>
'@
}

# ==============================================================================
# STANDALONE HTML REPORT GENERATOR
# ==============================================================================
function Export-BackupHtmlReport {
    param([string]$OutputPath = "$PSScriptRoot\BackupReport.html")

    $metrics = $Global:CachedMetrics
    $results = $Global:CachedResults
    $generatedTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

    $rowsHtml = ($results | ForEach-Object {
        $statusClass = if ($_.Status -eq "Critical") { "crit" } elseif ($_.Status -eq "Warning") { "warn" } else { "ok" }
        @"
        <tr>
            <td><span class="badge $statusClass">$($_.Status)</span></td>
            <td><strong>$([System.Web.HttpUtility]::HtmlEncode($_.ServerName))</strong></td>
            <td>$([System.Web.HttpUtility]::HtmlEncode($_.DatabaseName))</td>
            <td><span class="code">$($_.RecoveryModel)</span></td>
            <td class="code">$($_.LastFullBackup)</td>
            <td class="code">$($_.LastDifferentialBackup)</td>
            <td class="code">$($_.LastLogBackup)</td>
            <td class="code">$($_.LastBackupSizeGB) GB</td>
            <td class="reason">$([System.Web.HttpUtility]::HtmlEncode($_.StatusReason))</td>
        </tr>
"@
    }) -join "`n"

    $reportHtml = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>SQL Server Backup Audit & Compliance Report</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background: #0f172a; color: #f8fafc; margin: 0; padding: 2rem; }
        .container { max-width: 1400px; margin: 0 auto; }
        .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #334155; padding-bottom: 1rem; margin-bottom: 2rem; }
        .header h1 { margin: 0; font-size: 1.6rem; color: #38bdf8; display: flex; align-items: center; gap: 10px; }
        .header .meta { color: #94a3b8; font-size: 0.9rem; }
        .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 1rem; margin-bottom: 2rem; }
        .card { background: #1e293b; border: 1px solid #334155; border-radius: 8px; padding: 1rem; }
        .card-title { font-size: 0.75rem; text-transform: uppercase; color: #94a3b8; letter-spacing: 0.05em; }
        .card-value { font-size: 1.8rem; font-weight: bold; margin-top: 0.25rem; font-family: monospace; }
        table { width: 100%; border-collapse: collapse; background: #1e293b; border-radius: 8px; overflow: hidden; font-size: 0.85rem; }
        th { background: #0f172a; color: #94a3b8; text-align: left; padding: 10px 12px; font-weight: 600; text-transform: uppercase; font-size: 0.75rem; }
        td { padding: 10px 12px; border-bottom: 1px solid #334155; }
        tr:hover td { background: #243247; }
        .badge { display: inline-block; padding: 2px 8px; border-radius: 9999px; font-size: 0.7rem; font-weight: 600; text-transform: uppercase; }
        .badge.ok { background: rgba(16, 185, 129, 0.2); color: #34d399; border: 1px solid #10b981; }
        .badge.warn { background: rgba(245, 158, 11, 0.2); color: #fbbf24; border: 1px solid #f59e0b; }
        .badge.crit { background: rgba(239, 68, 68, 0.2); color: #f87171; border: 1px solid #ef4444; }
        .code { font-family: monospace; }
        .reason { color: #cbd5e1; font-size: 0.8rem; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div>
                <h1>
                    <svg style="width:28px;height:28px;" viewBox="0 0 24 24" fill="none" stroke="#38bdf8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"></path>
                    </svg>
                    SQL Server Backup Audit Report
                </h1>
                <div class="meta">Automated Compliance Snapshot &bull; Generated by SQLBackupMonitor PS</div>
            </div>
            <div class="meta" style="text-align:right;">
                <div>Timestamp: <strong>$generatedTime</strong></div>
                <div>Status: <strong>$($metrics.CompliancePct)% SLA Compliance</strong></div>
            </div>
        </div>

        <div class="cards">
            <div class="card"><div class="card-title">Servers</div><div class="card-value">$($metrics.TotalServers)</div></div>
            <div class="card"><div class="card-title">Databases</div><div class="card-value">$($metrics.TotalDatabases)</div></div>
            <div class="card"><div class="card-title">Healthy</div><div class="card-value" style="color:#34d399">$($metrics.HealthyDatabases)</div></div>
            <div class="card"><div class="card-title">Warnings</div><div class="card-value" style="color:#fbbf24">$($metrics.WarningDatabases)</div></div>
            <div class="card"><div class="card-title">Critical SLA</div><div class="card-value" style="color:#f87171">$($metrics.CriticalDatabases)</div></div>
            <div class="card"><div class="card-title">Total Volume</div><div class="card-value">$($metrics.TotalBackupSizeGB) GB</div></div>
        </div>

        <table>
            <thead>
                <tr>
                    <th>Status</th>
                    <th>SQL Server</th>
                    <th>Database Name</th>
                    <th>Recovery</th>
                    <th>Last Full Backup</th>
                    <th>Last Diff Backup</th>
                    <th>Last Log Backup</th>
                    <th>Size (GB)</th>
                    <th>Diagnostic SLA Notes</th>
                </tr>
            </thead>
            <tbody>
                $rowsHtml
            </tbody>
        </table>
    </div>
</body>
</html>
"@

    $reportHtml | Set-Content -Path $OutputPath -Encoding UTF8 -Force
    return $OutputPath
}

# ==============================================================================
# HTTP WEB SERVER DISPATCHER (System.Net.HttpListener)
# ==============================================================================
function Start-BackupMonitorWebServer {
    param([int]$HttpPort = 5000)

    $listener = [System.Net.HttpListener]::new()
    $prefix = "http://localhost:$HttpPort/"
    $listener.Prefixes.Add($prefix)

    try {
        $listener.Start()
    } catch {
        Write-Error "Could not bind HTTP listener to $prefix : $_"
        return
    }

    Write-Host "======================================================================" -ForegroundColor Cyan
    Write-Host " [x] SQLBackupMonitor - Standalone PowerShell Web Platform Started" -ForegroundColor Green
    Write-Host "======================================================================" -ForegroundColor Cyan
    Write-Host " [x] Local Web Dashboard:   http://localhost:$HttpPort" -ForegroundColor White
    Write-Host " [x] Configuration File:    $ConfigFile" -ForegroundColor Gray
    Write-Host " [x] Standalone Mode:       Zero .EXE Dependencies" -ForegroundColor Gray
    Write-Host " [x] PDF Report Engine:     Native Pure-PowerShell Compiler" -ForegroundColor Gray
    Write-Host " [x] Press CTRL+C in this console anytime to stop the server." -ForegroundColor Yellow
    Write-Host "======================================================================" -ForegroundColor Cyan

    # Initial Scan
    Write-Host "Performing initial SQL Server telemetry scan..." -ForegroundColor DarkGray
    Invoke-FullScan

    if (-not $NoBrowser) {
        try { Start-Process "http://localhost:$HttpPort" } catch {}
    }

    # Background Auto-Refresh & Dispatch Loop
    $lastAutoScan = [DateTime]::UtcNow
    $asyncContext = $listener.BeginGetContext($null, $null)

    while ($listener.IsListening) {
        try {
            if ($asyncContext.AsyncWaitHandle.WaitOne(1500)) {
                $context = $listener.EndGetContext($asyncContext)
                $asyncContext = $listener.BeginGetContext($null, $null)

                $request = $context.Request
                $response = $context.Response

                $urlPath = $request.Url.AbsolutePath
                $httpMethod = $request.HttpMethod

                # CORS Headers
                $response.AddHeader("Access-Control-Allow-Origin", "*")
                $response.AddHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
                $response.AddHeader("Access-Control-Allow-Headers", "Content-Type")

                if ($httpMethod -eq "OPTIONS") {
                    $response.StatusCode = 200
                    $response.Close()
                    continue
                }

                # ROUTING
                if ($urlPath -eq "/" -or $urlPath -eq "/index.html") {
                    $html = Get-EmbeddedHtmlDashboard
                    $buf = [System.Text.Encoding]::UTF8.GetBytes($html)
                    $response.ContentType = "text/html; charset=utf-8"
                    $response.ContentLength64 = $buf.Length
                    $response.OutputStream.Write($buf, 0, $buf.Length)
                }
                elseif ($urlPath -eq "/api/dashboard/metrics" -and $httpMethod -eq "GET") {
                    $json = $Global:CachedMetrics | ConvertTo-Json -Depth 5
                    $buf = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $response.ContentType = "application/json"
                    $response.ContentLength64 = $buf.Length
                    $response.OutputStream.Write($buf, 0, $buf.Length)
                }
                elseif ($urlPath -eq "/api/dashboard/databases" -and $httpMethod -eq "GET") {
                    $json = $Global:CachedResults | ConvertTo-Json -Depth 5
                    $buf = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $response.ContentType = "application/json"
                    $response.ContentLength64 = $buf.Length
                    $response.OutputStream.Write($buf, 0, $buf.Length)
                }
                elseif ($urlPath -eq "/api/alwayson" -and $httpMethod -eq "GET") {
                    $json = $Global:CachedAlwaysOn | ConvertTo-Json -Depth 5
                    $buf = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $response.ContentType = "application/json"
                    $response.ContentLength64 = $buf.Length
                    $response.OutputStream.Write($buf, 0, $buf.Length)
                }
                elseif ($urlPath -eq "/api/alerts" -and $httpMethod -eq "GET") {
                    $json = $Global:ActiveAlerts | ConvertTo-Json -Depth 5
                    $buf = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $response.ContentType = "application/json"
                    $response.ContentLength64 = $buf.Length
                    $response.OutputStream.Write($buf, 0, $buf.Length)
                }
                elseif ($urlPath -eq "/api/audit" -and $httpMethod -eq "GET") {
                    $json = $Global:AuditLogs | ConvertTo-Json -Depth 5
                    $buf = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $response.ContentType = "application/json"
                    $response.ContentLength64 = $buf.Length
                    $response.OutputStream.Write($buf, 0, $buf.Length)
                }
                elseif ($urlPath -eq "/api/servers" -and $httpMethod -eq "GET") {
                    $json = $Global:Config.Servers | ConvertTo-Json -Depth 5
                    $buf = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $response.ContentType = "application/json"
                    $response.ContentLength64 = $buf.Length
                    $response.OutputStream.Write($buf, 0, $buf.Length)
                }
                elseif ($urlPath -eq "/api/servers" -and $httpMethod -eq "POST") {
                    $bodyReader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
                    $bodyStr = $bodyReader.ReadToEnd()
                    $newSrv = $bodyStr | ConvertFrom-Json
                    if (-not $newSrv.Id -or [string]::IsNullOrWhiteSpace($newSrv.Id)) { 
                        $newSrv | Add-Member -MemberType NoteProperty -Name "Id" -Value ([Guid]::NewGuid().ToString()) -Force 
                    }
                    
                    # Deduplicate and update existing instance if matching
                    $srvList = [System.Collections.Generic.List[PSCustomObject]]::new()
                    $matched = $false
                    foreach ($s in @($Global:Config.Servers)) {
                        if ($null -eq $s) { continue }
                        $isMatch = ($s.Id -and $s.Id -eq $newSrv.Id) -or 
                                   ($s.Name -and $newSrv.Name -and $s.Name.ToLower() -eq $newSrv.Name.ToLower()) -or
                                   ($s.ServerAddress -and $newSrv.ServerAddress -and $s.ServerAddress.ToLower() -eq $newSrv.ServerAddress.ToLower() -and $s.Port -eq $newSrv.Port)
                        if ($isMatch) {
                            $newSrv.Id = $s.Id
                            $srvList.Add($newSrv)
                            $matched = $true
                        } else {
                            $srvList.Add($s)
                        }
                    }
                    if (-not $matched) {
                        $srvList.Add($newSrv)
                    }
                    $Global:Config.Servers = $srvList.ToArray()
                    Save-AppConfig -Config $Global:Config
                    Add-AuditLog -Action "Server Registered" -Details "Registered SQL Server instance '$($newSrv.Name)'"
                    
                    # Immediate scan
                    Invoke-FullScan

                    $respStr = @{ Success = $true; Id = $newSrv.Id } | ConvertTo-Json
                    $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                    $response.ContentType = "application/json"
                    $response.OutputStream.Write($buf, 0, $buf.Length)
                }
                elseif ($urlPath -match "^/api/servers/(.+)$" -and $httpMethod -eq "PUT") {
                    $id = [System.Uri]::UnescapeDataString($matches[1]).Trim('/')
                    $bodyReader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
                    $bodyStr = $bodyReader.ReadToEnd()
                    $updatedSrv = $bodyStr | ConvertFrom-Json

                    $srvList = [System.Collections.Generic.List[PSCustomObject]]::new()
                    foreach ($s in @($Global:Config.Servers)) {
                        if ($null -eq $s) { continue }
                        if ($s.Id -eq $id -or $s.Name -eq $id) {
                            $updatedSrv.Id = $s.Id
                            $srvList.Add($updatedSrv)
                        } else {
                            $srvList.Add($s)
                        }
                    }
                    $Global:Config.Servers = $srvList.ToArray()
                    Save-AppConfig -Config $Global:Config
                    Add-AuditLog -Action "Server Updated" -Details "Updated SQL Server instance '$($updatedSrv.Name)'"
                    
                    Invoke-FullScan

                    $respStr = @{ Success = $true } | ConvertTo-Json
                    $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                    $response.ContentType = "application/json"
                    $response.OutputStream.Write($buf, 0, $buf.Length)
                }
                elseif (($urlPath -match "^/api/servers/(.+)$" -and $httpMethod -eq "DELETE") -or ($urlPath -eq "/api/servers/delete" -and $httpMethod -eq "POST")) {
                    $id = ""
                    if ($httpMethod -eq "DELETE") {
                        $id = [System.Uri]::UnescapeDataString($matches[1]).Trim('/')
                    } else {
                        $bodyReader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
                        $bodyStr = $bodyReader.ReadToEnd()
                        $delReq = if ($bodyStr) { $bodyStr | ConvertFrom-Json } else { @{} }
                        $id = if ($delReq.Id) { [string]$delReq.Id } elseif ($delReq.Name) { [string]$delReq.Name } else { "" }
                    }

                    $removedNames = [System.Collections.Generic.List[string]]::new()
                    $srvList = [System.Collections.Generic.List[PSCustomObject]]::new()
                    foreach ($s in @($Global:Config.Servers)) {
                        if ($null -eq $s) { continue }
                        $match = ($s.Id -and $s.Id -eq $id) -or ($s.Name -and $s.Name -eq $id) -or ("$($s.ServerAddress):$($s.Port)" -eq $id)
                        if ($match) {
                            $removedNames.Add($s.Name)
                        } else {
                            $srvList.Add($s)
                        }
                    }
                    $Global:Config.Servers = $srvList.ToArray()
                    Save-AppConfig -Config $Global:Config

                    # Immediately clean cached telemetry of removed instances
                    $delArray = $removedNames.ToArray()
                    $Global:CachedResults = @($Global:CachedResults | Where-Object { 
                        $_.ServerId -ne $id -and (-not ($delArray -contains $_.ServerName)) 
                    })
                    $Global:CachedAlwaysOn = @($Global:CachedAlwaysOn | Where-Object { 
                        -not ($delArray -contains $_.ServerName) 
                    })
                    $Global:ActiveAlerts = @($Global:ActiveAlerts | Where-Object { 
                        -not ($delArray -contains $_.ServerName) 
                    })

                    Add-AuditLog -Action "Server Removed" -Details "Removed instance identifier: $id"
                    
                    # Refresh telemetry and metrics
                    Invoke-FullScan

                    $respStr = @{ Success = $true; Removed = $removedNames } | ConvertTo-Json
                    $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                    $response.ContentType = "application/json"
                    $response.OutputStream.Write($buf, 0, $buf.Length)
                }
            elseif ($urlPath -eq "/api/servers/test" -and $httpMethod -eq "POST") {
                $bodyReader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
                $bodyStr = $bodyReader.ReadToEnd()
                $testTarget = $bodyStr | ConvertFrom-Json

                $connTest = Test-SqlServerConnection -ServerObj $testTarget
                $respStr = $connTest | ConvertTo-Json
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/repository/test" -and $httpMethod -eq "POST") {
                $bodyReader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
                $bodyStr = $bodyReader.ReadToEnd()
                $testRepo = $bodyStr | ConvertFrom-Json

                $repoTest = Test-RepositoryConnection -RepoCfg $testRepo
                $respStr = $repoTest | ConvertTo-Json
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/repository/init" -and $httpMethod -eq "POST") {
                $bodyReader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
                $bodyStr = $bodyReader.ReadToEnd()
                $repoTarget = $bodyStr | ConvertFrom-Json

                $initRes = Initialize-RepositorySchema -RepoCfg $repoTarget
                $respStr = $initRes | ConvertTo-Json
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/repository" -and $httpMethod -eq "POST") {
                $bodyReader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
                $bodyStr = $bodyReader.ReadToEnd()
                $postedRepo = $bodyStr | ConvertFrom-Json
                $Global:Config.RepositorySettings = $postedRepo
                Save-AppConfig -Config $Global:Config

                if ($postedRepo.IsEnabled) {
                    Initialize-RepositorySchema -RepoCfg $postedRepo | Out-Null
                }

                $respStr = @{ Success = $true } | ConvertTo-Json
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/repository/cleanup" -and $httpMethod -eq "POST") {
                $bodyReader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
                $bodyStr = $bodyReader.ReadToEnd()
                $cleanOpt = if ($bodyStr) { $bodyStr | ConvertFrom-Json } else { @{} }
                $purgeAll = [bool]$cleanOpt.PurgeAll
                $res = Clear-RepositoryStaleData -PurgeAll:$purgeAll
                $respStr = $res | ConvertTo-Json
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/repository/trends" -and $httpMethod -eq "GET") {
                $days = 7
                if ($request.QueryString["days"]) { [int]::TryParse($request.QueryString["days"], [ref]$days) | Out-Null }
                $trends = Get-RepositoryHistoricalTrends -Days $days
                $respStr = $trends | ConvertTo-Json -Depth 5
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/repository/history" -and $httpMethod -eq "GET") {
                $limit = 50
                if ($request.QueryString["limit"]) { [int]::TryParse($request.QueryString["limit"], [ref]$limit) | Out-Null }
                $status = if ($request.QueryString["status"]) { $request.QueryString["status"] } else { "" }
                $search = if ($request.QueryString["search"]) { $request.QueryString["search"] } else { "" }
                
                $snaps = Get-RepositoryHistoricalSnapshots -Limit $limit -Status $status -Search $search
                $respStr = $snaps | ConvertTo-Json -Depth 5
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/scan" -and $httpMethod -eq "POST") {
                Invoke-FullScan
                $respStr = @{ Success = $true; LastScan = $Global:LastScanTime } | ConvertTo-Json
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/policies/groups" -and $httpMethod -eq "GET") {
                $respStr = $Global:Config.GroupPolicies | ConvertTo-Json -Depth 5
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/policies/groups" -and $httpMethod -eq "POST") {
                $bodyReader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
                $bodyStr = $bodyReader.ReadToEnd()
                $postedGrp = $bodyStr | ConvertFrom-Json
                $origName = if ($postedGrp.OriginalName) { $postedGrp.OriginalName.Trim() } else { $postedGrp.GroupName.Trim() }
                $targetName = $postedGrp.GroupName.Trim()

                $grpList = [System.Collections.Generic.List[PSCustomObject]]::new()
                $updated = $false
                foreach ($g in @($Global:Config.GroupPolicies)) {
                    if ($null -eq $g) { continue }
                    if (($g.GroupName -and $g.GroupName.ToLower() -eq $origName.ToLower()) -or ($g.GroupName -and $g.GroupName.ToLower() -eq $targetName.ToLower())) {
                        $grpList.Add($postedGrp)
                        $updated = $true
                    } else {
                        $grpList.Add($g)
                    }
                }
                if (-not $updated) {
                    $grpList.Add($postedGrp)
                }
                $Global:Config.GroupPolicies = $grpList.ToArray()
                Save-AppConfig -Config $Global:Config
                Add-AuditLog -Action "Group SLA Policy Saved" -Details "Saved SLA policy for group '$targetName'"
                
                Invoke-FullScan

                $respStr = @{ Success = $true } | ConvertTo-Json
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif (($urlPath -match "^/api/policies/groups/(.+)$" -and $httpMethod -eq "DELETE") -or ($urlPath -eq "/api/policies/groups/delete" -and $httpMethod -eq "POST")) {
                $grpName = ""
                if ($httpMethod -eq "DELETE") {
                    $grpName = [System.Uri]::UnescapeDataString($matches[1]).Trim('/')
                } else {
                    $bodyReader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
                    $bodyStr = $bodyReader.ReadToEnd()
                    $delReq = if ($bodyStr) { $bodyStr | ConvertFrom-Json } else { @{} }
                    $grpName = if ($delReq.GroupName) { [string]$delReq.GroupName } else { "" }
                }

                $grpList = [System.Collections.Generic.List[PSCustomObject]]::new()
                foreach ($g in @($Global:Config.GroupPolicies)) {
                    if ($null -eq $g) { continue }
                    if ($g.GroupName -and $g.GroupName.ToLower() -eq $grpName.ToLower()) {
                        continue
                    }
                    $grpList.Add($g)
                }
                $Global:Config.GroupPolicies = $grpList.ToArray()
                Save-AppConfig -Config $Global:Config
                Add-AuditLog -Action "Group SLA Policy Deleted" -Details "Deleted SLA policy for group '$grpName'"

                Invoke-FullScan

                $respStr = @{ Success = $true } | ConvertTo-Json
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/config" -and $httpMethod -eq "GET") {
                $respStr = $Global:Config | ConvertTo-Json -Depth 5
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/config" -and $httpMethod -eq "POST") {
                $bodyReader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
                $bodyStr = $bodyReader.ReadToEnd()
                $postedCfg = $bodyStr | ConvertFrom-Json
                if ($postedCfg.GlobalPolicies) { $Global:Config.GlobalPolicies = $postedCfg.GlobalPolicies }
                if ($postedCfg.GroupPolicies) { $Global:Config.GroupPolicies = $postedCfg.GroupPolicies }
                if ($postedCfg.RepositorySettings) { $Global:Config.RepositorySettings = $postedCfg.RepositorySettings }
                if ($postedCfg.EmailSettings) { $Global:Config.EmailSettings = $postedCfg.EmailSettings }
                Save-AppConfig -Config $Global:Config
                $respStr = @{ Success = $true } | ConvertTo-Json
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/email/test" -and $httpMethod -eq "POST") {
                $bodyReader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
                $bodyStr = $bodyReader.ReadToEnd()
                if (-not [string]::IsNullOrWhiteSpace($bodyStr)) {
                    $postedEmail = $bodyStr | ConvertFrom-Json
                    $Global:Config.EmailSettings = $postedEmail
                    Save-AppConfig -Config $Global:Config
                }

                $dummyBreach = @(
                    [PSCustomObject]@{
                        ServerId       = "TEST-01"
                        ServerName     = "SQL-TEST-CLUSTER"
                        DatabaseName   = "PaymentGateway_DB"
                        RecoveryModel  = "FULL"
                        Status         = "Critical"
                        LastFullBackup = (Get-Date).AddHours(-28).ToString("yyyy-MM-dd HH:mm:ss")
                        LastLogBackup  = "Never"
                        StatusReason   = "TEST ALERT: Full backup is 28.0h old (Threshold: 24h); No Log Backup"
                    }
                )

                $res = Send-SlaEmailAlert -BreachedDatabases $dummyBreach -IsTest
                $respStr = $res | ConvertTo-Json
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/email/report" -and $httpMethod -eq "POST") {
                $res = Send-DailyPdfEmailReport
                $respStr = $res | ConvertTo-Json
                $buf = [System.Text.Encoding]::UTF8.GetBytes($respStr)
                $response.ContentType = "application/json"
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            elseif ($urlPath -eq "/api/export/pdf" -and $httpMethod -eq "GET") {
                $tempPdf = [System.IO.Path]::GetTempFileName() + ".pdf"
                Export-BackupPdfReport -OutputPath $tempPdf | Out-Null
                $pdfBytes = [System.IO.File]::ReadAllBytes($tempPdf)
                Remove-Item $tempPdf -Force -ErrorAction SilentlyContinue

                $response.ContentType = "application/pdf"
                $response.AddHeader("Content-Disposition", "attachment; filename=SQLBackupReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').pdf")
                $response.ContentLength64 = $pdfBytes.Length
                $response.OutputStream.Write($pdfBytes, 0, $pdfBytes.Length)
            }
            elseif ($urlPath -eq "/api/export/html" -and $httpMethod -eq "GET") {
                $tempReport = Export-BackupHtmlReport
                $html = Get-Content $tempReport -Raw -Encoding UTF8
                $buf = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html; charset=utf-8"
                $response.AddHeader("Content-Disposition", "attachment; filename=SQLBackupReport.html")
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }
            else {
                $response.StatusCode = 404
                $buf = [System.Text.Encoding]::UTF8.GetBytes("Not Found")
                $response.OutputStream.Write($buf, 0, $buf.Length)
            }

                $response.Close()
            }
        } catch {
            Write-Warning "HTTP Request error: $_"
        }

        # Periodic Background Auto-Scan & Daily Email Schedule Check
        try {
            $now = [DateTime]::UtcNow
            $scanInterval = if ($Global:Config.GlobalPolicies.AutoRefreshIntervalSec) { [Math]::Max(5, [int]$Global:Config.GlobalPolicies.AutoRefreshIntervalSec) } else { 60 }
            if (($now - $lastAutoScan).TotalSeconds -ge $scanInterval) {
                $lastAutoScan = $now
                if (-not $Global:IsScanning -and $Global:Config.Servers.Count -gt 0) {
                    Invoke-FullScan
                }

                # Check Daily Scheduled Email
                $today = Get-Date -Format "yyyy-MM-dd"
                $currTime = Get-Date -Format "HH:mm"
                if ($Global:Config.EmailSettings.DailyReportEnabled -and 
                    $Global:Config.EmailSettings.LastDailyReportDate -ne $today -and 
                    $currTime -ge $Global:Config.EmailSettings.DailyReportTime) {
                    Send-DailyPdfEmailReport | Out-Null
                }
            }
        } catch {}
    }
}

# ==============================================================================
# MAIN ENTRYPOINT LOGIC
# ==============================================================================
switch ($Mode) {
    "Web" {
        Start-BackupMonitorWebServer -HttpPort $Port
    }
    "Scan" {
        Write-Host "Running single SQLBackupMonitor scan across configured instances..." -ForegroundColor Cyan
        Invoke-FullScan
        
        Write-Host "`n--- SQL BACKUP HEALTH & SLA AUDIT ---" -ForegroundColor Green
        if ($Global:CachedResults.Count -gt 0) {
            $Global:CachedResults | Format-Table -Property Status, ServerName, DatabaseName, RecoveryModel, LastFullBackup, LastLogBackup, LastBackupSizeGB, StatusReason -AutoSize
        } else {
            Write-Host "No registered servers or active databases found. Register instances using Web mode or config.json." -ForegroundColor DarkGray
        }
        
        if ($Global:CachedAlwaysOn.Count -gt 0) {
            Write-Host "`n--- ALWAYS-ON AVAILABILITY GROUPS (HADR) ---" -ForegroundColor Cyan
            $Global:CachedAlwaysOn | Format-Table -Property ServerName, GroupName, ReplicaServerName, Role, SyncHealth, DatabaseName, DbSyncState -AutoSize
        }

        Write-Host "`nSummary:" -ForegroundColor Yellow
        Write-Host "Total Databases: $($Global:CachedMetrics.TotalDatabases) | Healthy: $($Global:CachedMetrics.HealthyDatabases) | Warning: $($Global:CachedMetrics.WarningDatabases) | Critical: $($Global:CachedMetrics.CriticalDatabases) | Compliance: $($Global:CachedMetrics.CompliancePct)%"
    }
    "Report" {
        Write-Host "Generating standalone HTML audit report..." -ForegroundColor Cyan
        Invoke-FullScan
        $out = Export-BackupHtmlReport -OutputPath $ReportPath
        Write-Host "HTML report generated: $out" -ForegroundColor Green
        if (-not $NoBrowser) { Start-Process $out }
    }
    "PdfReport" {
        Write-Host "Generating standalone PDF audit report (Pure PowerShell)..." -ForegroundColor Cyan
        Invoke-FullScan
        $out = Export-BackupPdfReport -OutputPath $ReportPath
        Write-Host "PDF report generated: $out" -ForegroundColor Green
        if (-not $NoBrowser) { Start-Process $out }
    }
    "EmailReport" {
        Write-Host "Executing on-demand / scheduled Daily Executive PDF Email Report..." -ForegroundColor Cyan
        Invoke-FullScan
        $res = Send-DailyPdfEmailReport
        if ($res.Success) {
            Write-Host "[SUCCESS] $($res.Message)" -ForegroundColor Green
        } else {
            Write-Host "[ERROR] $($res.ErrorMessage)" -ForegroundColor Red
        }
    }
    "Loop" {
        $loopInterval = if ($Global:Config.GlobalPolicies.AutoRefreshIntervalSec) { [Math]::Max(5, [int]$Global:Config.GlobalPolicies.AutoRefreshIntervalSec) } else { $IntervalSeconds }
        Write-Host "Starting continuous background backup monitoring loop (Interval: ${loopInterval}s)..." -ForegroundColor Cyan
        while ($true) {
            Invoke-FullScan
            $m = $Global:CachedMetrics
            $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Write-Host "[$ts] Scan complete. DBs: $($m.TotalDatabases) | OK: $($m.HealthyDatabases) | Warn: $($m.WarningDatabases) | Crit: $($m.CriticalDatabases) | Compliance: $($m.CompliancePct)%" -ForegroundColor (if ($m.CriticalDatabases -gt 0) { "Red" } else { "Green" })
            
            # Check Daily Email Schedule
            $today = Get-Date -Format "yyyy-MM-dd"
            $currTime = Get-Date -Format "HH:mm"
            if ($Global:Config.EmailSettings.DailyReportEnabled -and 
                $Global:Config.EmailSettings.LastDailyReportDate -ne $today -and 
                $currTime -ge $Global:Config.EmailSettings.DailyReportTime) {
                Write-Host "[$ts] [SCHEDULED] Triggering Daily Executive PDF Email Report..." -ForegroundColor Cyan
                Send-DailyPdfEmailReport | Out-Null
            }

            Start-Sleep -Seconds $loopInterval
        }
    }
}
