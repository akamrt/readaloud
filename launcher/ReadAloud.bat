@echo off
REM ReadAloud Launcher - Double-click this file to run ReadAloud
REM This script will automatically install Python if needed, download the latest code, and run the app.

echo ========================================
echo ReadAloud - Text-to-Speech Tool
echo ========================================
echo.

REM Check if PowerShell is available
where powershell >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: PowerShell is required but not found.
    echo Please install PowerShell or use Windows 10/11.
    pause
    exit /b 1
)

REM Run the PowerShell setup script
powershell -ExecutionPolicy Bypass -File "%~dp0setup.ps1"

if %errorlevel% neq 0 (
    echo.
    echo Setup failed. Please check the error messages above.
    pause
)

exit /b %errorlevel%