@echo off
chcp 65001 >nul
cd /d "%~dp0"
where wt.exe >nul 2>&1
if not errorlevel 1 if not defined WT_SESSION (
    start "CheckSentry" wt.exe -w new -d "%CD%" powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Start-ComplianceCheck.ps1"
    exit /b
)
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Start-ComplianceCheck.ps1"
pause
