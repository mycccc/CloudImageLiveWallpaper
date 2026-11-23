@echo off
setlocal enabledelayedexpansion
title FY4B Cloud Image - Setup

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Administrator privileges required.
    echo Right-click this file and select "Run as administrator".
    pause
    exit /b 1
)

set "DIR=%~dp0"
if "%DIR:~-1%"=="\" set "DIR=%DIR:~0,-1%"

set "PS_SCRIPT=%DIR%\3download_fy4b.ps1"
set "VBS_FILE=%DIR%\2run_hidden.vbs"
set "TASK_NAME=FY4B_CloudImage_Download"

if not exist "%PS_SCRIPT%" (
    echo [ERROR] PowerShell script not found: %PS_SCRIPT%
    echo Make sure 3download_fy4b.ps1 is in the same folder as this cmd.
    pause
    exit /b 1
)

echo ============================================
echo FY4B Cloud Image - Setup
echo ============================================
echo Folder: %DIR%
echo.
echo   [1] Enable  - generate vbs, register task, run first download
echo   [2] Disable - disable the scheduled task (keep task definition)
echo   [3] Exit
echo.

set "CHOICE="
set /p "CHOICE=Choose (1/2/3, default 1): "
if not defined CHOICE set "CHOICE=1"

if "%CHOICE%"=="2" goto DISABLE
if "%CHOICE%"=="3" exit /b 0
if "%CHOICE%"=="1" goto ENABLE

echo Invalid choice.
pause
exit /b 1

:ENABLE
echo.
set "HOURS=8"
set /p "INPUT_HOURS=Download hours back (default 8): "
if defined INPUT_HOURS set "HOURS=%INPUT_HOURS%"
echo.
echo Using HoursBack = %HOURS%
echo.

echo [1/3] Generating 2run_hidden.vbs ...
> "%VBS_FILE%" echo Set shell = CreateObject("WScript.Shell")
>> "%VBS_FILE%" echo shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""%PS_SCRIPT%"" -HoursBack %HOURS% -SaveDir ""%DIR%""", 0, False
if not exist "%VBS_FILE%" (
    echo [ERROR] Failed to create vbs launcher.
    pause
    exit /b 1
)
echo   Created: %VBS_FILE%

echo.
echo [2/3] Registering / enabling scheduled task ...
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -HoursBack %HOURS% -InstallTask
if %errorlevel% neq 0 (
    echo [ERROR] Failed to register scheduled task.
    pause
    exit /b 1
)

:: -InstallTask 覆盖注册后通常已恢复启用，这里再显式 enable 一次，
:: 确保任务在禁用→启用的切换场景下状态确定。
schtasks /change /tn "%TASK_NAME%" /enable >nul 2>&1

echo.
echo [3/3] Done.
echo ============================================
echo   Task name:   %TASK_NAME%
echo   State:       ENABLED
echo   Interval:    Every 15 minutes
echo   Hours back:  %HOURS%
echo   Folder:      %DIR%
echo ============================================
pause
exit /b 0

:DISABLE
echo.
schtasks /query /tn "%TASK_NAME%" >nul 2>&1
if %errorlevel% neq 0 (
    echo Task "%TASK_NAME%" does not exist. Nothing to disable.
    pause
    exit /b 0
)

echo Disabling scheduled task "%TASK_NAME%" ...
schtasks /change /tn "%TASK_NAME%" /disable
if %errorlevel% neq 0 (
    echo [ERROR] Failed to disable task.
    pause
    exit /b 1
)

echo.
echo Done. Task "%TASK_NAME%" is now DISABLED.
echo Task definition is kept; choose [1] Enable to restore it.
echo.
pause
exit /b 0