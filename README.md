AUTOMATED OCR POWERSHELL SCRIPT MADE FOR DOE PROJECT AS ALTERNATIVE TO NAPS2 OCR
★ FOR BATCH FOLDER OCR ONLY ★

The "start_process.bat" will generate the needed folders as follows:
Folder List:
> input - [BATCH OCR ONLY] Where your folders with tif files that will be processed for OCR (folder with tif file would the name of the pdf)
> archive - Where your folders in input folder will be moved after OCR
> output - Where your processed OCR files in pdf format
> logs - Where the logs are stored for the entire process


---------------------------------------------------------------------
HOW TO SETUP:
1.) Go to Setup folder then install "ImageMagick-7.1.1-41-Q16-HDRI-x64-dll.exe" and "tesseract-ocr-w64-setup-5.5.0.20241111"
2.) Copy the ".traineddata" files in "setup/tessdata_best" then go to Tesseract-OCR/tessdata directory then paste and overwrite if needed
(Default: C:\Program Files\Tesseract-OCR\tessdata)
3.) Start "start_process.bat" to setup folders then proceed to usage below.
---------------------------------------------------------------------
How TO USE:
> put the folders with tif files inside the "input" folder
> just run "start_process.bat" then wait
> processed pdf files with OCR will be generated to "output" folder

## DISREGARD EXPERIMENTAL ONLY 
## > press "shift + right click" then click "Open powershell window here"
## > type in powershell: ".\testra.ps1 input"
---------------------------------------------------------------------
PDF Filename Convention:

Proper Folder Structure
input
   ├── folder1
   │    ├── subfolder1 ★
   │    │    ├── image1.tif
   │    │    └── image2.tif
   │    ├── subfolder2 ★
   │    │    ├── image1.tif
   │    │    └── image2.tif
   └── folder2 ★
        ├── image1.tif
        └── image2.tif
THE ★ would be the name of the pdf files since it is where the tif files were found and will be treated as batch

Do **NOT**  do this folder structure:
input
   ├── folder1
   │    ├── image1.tif ★
   │    ├── image2.tif ★
   │    ├── subfolder1 <!>
   │    │    ├── image1.tif
   │    │    └── image2.tif
   │    ├── subfolder2
   │    │    ├── image1.tif
   │    │    └── image2.tif
   └── folder2
        ├── image1.tif
        └── image2.tif
The ★ would cause conflict since there is a subfolder (<!>) in the directory, what happens it it will stop the operation after the folder1 is done ocr, disregarding every sub folders with it.


### PDF Filename Convention:

- The PDF filename corresponds directly to the folder name in which `.tif` files are found.

    - If `.tif` files are found in **folder1**, the generated PDF is named **folder1.pdf**.
    - If `.tif` files are found in **folder1/subfolder1**, the generated PDF is named **subfolder1.pdf**.
    - Similarly, if `.tif` files are found in **folder1/subfolder2**, the generated PDF is named **subfolder2.pdf**.
    - If `.tif` files are found in **folder2**, the generated PDF is named **folder2.pdf**.

### Example Output:

- **folder1.pdf** for `.tif` files directly inside **folder1**.
- **subfolder1.pdf** for `.tif` files inside **folder1/subfolder1**.
- **subfolder2.pdf** for `.tif` files inside **folder1/subfolder2**.
- **folder2.pdf** for `.tif` files inside **folder2**.
---------------------------------------------------------------------
Repository:
https://github.com/tesseract-ocr/tesseract
---------------------------------------------------------------------
Command Line Installer:
> setup/tesseract-ocr-w64-setup-5.5.0.20241111.exe
> setup/ImageMagick-7.1.1-41-Q16-HDRI-x64-dll
---------------------------------------------------------------------
Note:
> press "ctrl + c" in powershell to cancel
---------------------------------------------------------------------
