@echo off
title SQLBackupMonitor (PowerShell Standalone Edition)
cd /d "%~dp0"
echo ==============================================================================
echo  SQLBackupMonitor - Enterprise SQL Server Backup Monitoring Platform
echo  PowerShell Standalone Edition (Zero .EXE Dependencies)
echo ==============================================================================
echo  Starting local web server & dashboard on port 5000...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0SQLBackupMonitor.ps1" -Mode Web -Port 5000
pause
