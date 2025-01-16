REM filepath: MagickTessTK.bat
@echo off
color 0b
cls
title MagickTessTK OCR - Automated OCR Processing Tool

REM Get script directory with proper handling
set "scriptPath=%~dp0"
if "%scriptPath:~-1%"=="\" set "scriptPath=%scriptPath:~0,-1%"

REM Unblock files first
pwsh.exe -WindowStyle Hidden -Command "Get-ChildItem -Path '%scriptPath%' -File | Where-Object { $_.Name -match '(.ps1|.bat|.txt|.ini)$' } | Unblock-File"
pwsh.exe -WindowStyle Hidden -Command "if (Test-Path '%scriptPath%\setup') { Get-ChildItem -Path '%scriptPath%\setup' -Recurse -File | Unblock-File }"

REM Check for PowerShell Core first
where pwsh >nul 2>nul
if %errorlevel% equ 0 (
    pwsh.exe -NoProfile -ExecutionPolicy RemoteSigned -File "%scriptPath%\data\validatedir.ps1"
    pwsh.exe -NoProfile -ExecutionPolicy RemoteSigned -File "%scriptPath%\data\magicktesstk-gui.ps1"
) else (
    start start_process.bat
)

if %errorlevel% neq 0 (
    echo Error launching GUI application
    echo Please ensure PowerShell is installed and try again
    pause
    exit /b 1
)
