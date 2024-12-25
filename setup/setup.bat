@echo off
color 0b
title Testra - Tesseract-OCR + ImageMagick Setup

:: Check for Administrator Privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo This script requires Administrator privileges.
    echo Please run as Administrator.
    echo.
    echo Attempting to restart the script with elevated privileges...
    powershell -Command "Start-Process cmd -ArgumentList '/c %~s0' -Verb RunAs"
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

echo Installation complete.
pause
