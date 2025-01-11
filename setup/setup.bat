@echo off
color 0b
title MagickTess - Tesseract-OCR + ImageMagick + PowerShell 7 Setup

:: Check for Administrator Privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo This script requires Administrator privileges.
    echo Please run as Administrator.
    echo.
    echo Attempting to restart the script with elevated privileges...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Change the current working directory to the location of the batch file
cd /d "%~dp0"

cls

echo Installing Tesseract-OCR
start /wait "" "data\tesseract-ocr-w64-setup-5.5.0.20241111.exe"
echo Tesseract-OCR installed.
echo Patching Tesseract-OCR with Trained Data LTSM Models...
robocopy "data\tessdata_best" "C:\Program Files\Tesseract-OCR\tessdata" /S /E /IS /R:0 /W:0
echo Tesseract-OCR Trained Data LTSM Models Patch Installed.

echo Installing ImageMagick ...
start /wait "" "data\ImageMagick-7.1.1-41-Q16-HDRI-x64-dll.exe"
echo ImageMagick installed.

echo Installing PowerShell 7 ...
start /wait "" "data\PowerShell-7.4.6-win-x64.msi"
echo PowerShell 7 installed.

echo Installing PDFtk Server 2.02 ...
start /wait "" "data\pdftk_server-2.02-win-setup.exe"
echo PDFtk Server 2.02 installed.

echo Installing Ghostscript 10.04.0 for Windows ...
start /wait "" "data\gs10040w64.exe"
echo PDFtk Server 2.02 installed.

echo Installation complete.
pause
