@echo off
color 0b
cls
title MagickTessTK OCR - Automated OCR Processing Tool
:START
set "scriptPath=%~dp0"
if "%scriptPath:~-1%"=="\" set "scriptPath=%scriptPath:~0,-1%"

where pwsh >nul 2>nul
if %errorlevel% neq 0 (
    powershell.exe -Command "Get-ChildItem -Path '%scriptPath%' -File | Where-Object { $_.Name -in @('magicktesstk-st.ps1', 'magicktesstk-mt.ps1', 'magicktesstk-gui.ps1', 'validatedir.ps1', 'settings.ini', 'start_process.bat', 'MagickTessTK.bat', 'ReadMe.txt') } | Unblock-File"
    powershell.exe -Command "Get-ChildItem -Path '%scriptPath%'\setup' -Recurse -File | Unblock-File"
    powershell.exe -ExecutionPolicy RemoteSigned -File "magicktesstk-gui.ps1"
) else (
    pwsh -NoProfile -Command "exit" >nul 2>nul
    if %errorlevel% neq 0 (
        powershell.exe -Command "Get-ChildItem -Path '%scriptPath%' -File | Where-Object { $_.Name -in @('magicktesstk-st.ps1', 'magicktesstk-mt.ps1', 'magicktesstk-gui.ps1', 'validatedir.ps1', 'settings.ini', 'start_process.bat', 'MagickTessTK.bat', 'ReadMe.txt') } | Unblock-File"
        powershell.exe -Command "Get-ChildItem -Path '%scriptPath%'\setup' -Recurse -File | Unblock-File"
        powershell.exe -ExecutionPolicy RemoteSigned -File "magicktesstk-gui.ps1"
    ) else (
        pwsh.exe -Command "Get-ChildItem -Path '%scriptPath%' -File | Where-Object { $_.Name -in @('magicktesstk-st.ps1', 'magicktesstk-mt.ps1', 'magicktesstk-gui.ps1', 'validatedir.ps1', 'settings.ini', 'start_process.bat', 'MagickTessTK.bat', 'ReadMe.txt') } | Unblock-File"
        pwsh.exe -Command "Get-ChildItem -Path '%scriptPath%'\setup' -Recurse -File | Unblock-File"
        pwsh.exe -ExecutionPolicy RemoteSigned -File "magicktesstk-gui.ps1"
    )
)

GOTO START
pause