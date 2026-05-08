@echo off
chcp 65001 > nul

:: Check administrator privileges and re-launch with elevation if needed
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo Starting Windows environment collection script...
echo Output path: %~dp0
echo.

:: Check PS1 file exists
if not exist "%~dp0Collect-WindowsEnv.ps1" (
    echo [ERROR] Collect-WindowsEnv.ps1 not found.
    echo         Path: %~dp0
    pause
    exit /b 1
)

powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0Collect-WindowsEnv.ps1" -GenerateHtml
set PS_EXIT=%errorlevel%

echo.
if %PS_EXIT% equ 0 (
    echo [DONE] Completed successfully.
) else (
    echo [WARN] Exit code: %PS_EXIT%
)
echo.
pause
