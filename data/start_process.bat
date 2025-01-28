@echo off
color 0b
cls
REM Author: Copyright 2025 | Kyle F. Capistrano
title MagickTessTK OCR - Automated OCR Processing Tool
set "scriptPath=%~dp0"
if "%scriptPath:~-1%"=="\" set "scriptPath=%scriptPath:~0,-1%"
cd /d "%scriptPath%"
cls
where pwsh >nul 2>nul
if %errorlevel% neq 0 (
    powershell.exe -ExecutionPolicy RemoteSigned -File "%scriptPath%\validatedir.ps1"
) else (
    pwsh -NoProfile -Command "exit" >nul 2>nul
    if %errorlevel% neq 0 (
        powershell.exe -ExecutionPolicy RemoteSigned -File "%scriptPath%\validatedir.ps1"
    ) else (
        pwsh.exe -ExecutionPolicy RemoteSigned -File "%scriptPath%\validatedir.ps1"
    )
)

cls
echo.
echo MagickTessTK OCR - Automated OCR Processing Tool
echo.
echo The script will only process files as batch and will treat folders as the document where its
echo pages will be the tif files, read the ReadMe.txt file for proper documentation
echo ---------------------------------------------------------------------------------------------
echo Repository:
echo https://github.com/tesseract-ocr/tesseract
echo ---------------------------------------------------------------------------------------------
echo Command Line Installer:
echo ^> just run "setup/setup.bat"
echo ---------------------------------------------------------------------------------------------
echo Note:
echo ^> press "ctrl + c" in powershell to cancel
echo ---------------------------------------------------------------------------------------------
echo Folder List Generated:
echo ^> input - [BATCH OCR ONLY] Where your folders with tif files that will be processed for OCR
echo ^> archive - Where your folders in input folder will be moved after OCR
echo ^> output - Where your processed OCR files in pdf format
echo ^> logs - Where the logs are stored for the entire process
echo ---------------------------------------------------------------------------------------------
echo.
echo The script will process your input folder for OCR
cls

where pwsh >nul 2>nul
if %errorlevel% neq 0 (
    powershell.exe -ExecutionPolicy RemoteSigned -File "%scriptPath%\magicktesstk-st.ps1"
) else (
    pwsh -NoProfile -Command "exit" >nul 2>nul
    if %errorlevel% neq 0 (
        powershell.exe -ExecutionPolicy RemoteSigned -File "%scriptPath%\magicktesstk-st.ps1"
    ) else (
        pwsh.exe -ExecutionPolicy RemoteSigned -File "%scriptPath%\magicktesstk-mt.ps1"
    )
)
