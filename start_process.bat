@echo off
color 0b
cls
title MagickTessTK OCR - Automated OCR Processing Tool
:START
cls
:: -- DEPRECATED --
:: if not exist "input" mkdir input
:: if not exist "archive" mkdir archive
:: if not exist "output" mkdir output
:: if not exist "logs" mkdir logs
cls
echo.
echo NAPS2 on Steroids
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
pause
cls

set "scriptPath=%~dp0"
if "%scriptPath:~-1%"=="\" set "scriptPath=%scriptPath:~0,-1%"

where pwsh >nul 2>nul
if %errorlevel% neq 0 (
    powershell.exe -Command "Get-ChildItem -Path '%scriptPath%' -File | Where-Object { $_.Name -in @('magicktesstk-st.ps1', 'magicktesstk-mt.ps1', 'validatedir.ps1', 'settings.ini', 'start_process.bat', 'ReadMe.txt') } | Unblock-File"
    powershell.exe -Command "Get-ChildItem -Path '%scriptPath%'\setup' -Recurse -File | Unblock-File"
    powershell.exe -ExecutionPolicy RemoteSigned -File "magicktesstk-st.ps1"
) else (
    pwsh -NoProfile -Command "exit" >nul 2>nul
    if %errorlevel% neq 0 (
        powershell.exe -Command "Get-ChildItem -Path '%scriptPath%' -File | Where-Object { $_.Name -in @('magicktesstk-st.ps1', 'magicktesstkk-mt.ps1', 'validatedir.ps1', 'settings.ini', 'start_process.bat', 'ReadMe.txt') } | Unblock-File"
        powershell.exe -Command "Get-ChildItem -Path '%scriptPath%\setup' -Recurse -File | Unblock-File"
        powershell.exe -ExecutionPolicy RemoteSigned -File "magicktesstk-st.ps1"
    ) else (
        pwsh.exe -Command "Get-ChildItem -Path '%scriptPath%' -File | Where-Object { $_.Name -in @('magicktesstk-st.ps1', 'magicktesstk-mt.ps1', 'validatedir.ps1', 'settings.ini', 'start_process.bat', 'ReadMe.txt') } | Unblock-File"
        pwsh.exe -Command "Get-ChildItem -Path '%scriptPath%'\setup' -Recurse -File | Unblock-File"
        pwsh.exe -ExecutionPolicy RemoteSigned -File "magicktesstk-mt.ps1"
    )
)

GOTO START
pause