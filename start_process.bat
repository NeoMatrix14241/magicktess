@echo off
color 0b
cls
title Testra OCR - Automated OCR Processing Tool
:START
cls
if not exist "input" mkdir input
if not exist "archive" mkdir archive
if not exist "output" mkdir output
if not exist "logs" mkdir logs
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
echo ^> setup/tesseract-ocr-w64-setup-5.5.0.20241111.exe
echo ---------------------------------------------------------------------------------------------
echo Note:
echo ^> press "ctrl + c" in powershell to cancel
echo ---------------------------------------------------------------------------------------------
echo Folder List:
echo ^> input - [BATCH OCR ONLY] Where your folders with tif files that will be processed for OCR
echo ^> archive - Where your folders in input folder will be moved after OCR
echo ^> output - Where your processed OCR files in pdf format
echo ^> logs - Where the logs are stored for the entire process
echo ---------------------------------------------------------------------------------------------
echo.
echo The script will process your input folder for OCR
pause
cls
powershell.exe -ExecutionPolicy RemoteSigned -File "testra.ps1" "input"
GOTO START
pause