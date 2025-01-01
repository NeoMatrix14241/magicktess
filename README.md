# Automated OCR PowerShell Script

**⚡ Alternative to NAPS2 OCR - Designed for Batch OCR Process**

### 🚀 Batch Folder OCR Only

This PowerShell script automates OCR processing for batches of different image files, generating searchable PDFs.

---
### Dependencies for MagickTess

The PowerShell script utilizes the following software and models for Tesseract-OCR:

1. **[ImageMagick](https://imagemagick.org)**  
   - Version: ImageMagick-7.1.1-43-Q16-HDRI-x64-dll.exe  

2. **[Tesseract-OCR](https://github.com/UB-Mannheim/tesseract/wiki)**  
   - Version: tesseract-ocr-w64-setup-5.5.0.20241111.exe  

3. **[Tessdata](https://github.com/tesseract-ocr/tessdata/tree/main) for Tesseract-OCR**  
   - Models: `eng.traineddata`, `enm.traineddata`, `fil.traineddata`
   
4. **[PowerShell 7](https://github.com/PowerShell/PowerShell)
   - Version: PowerShell-7.4.6-win-x64.msi

---
### 📂 Folder Structure
When executed, `start_process.bat` will create the following folders:

- **input** – [Batch OCR Only] Place folders containing `.tif` files here. Each folder name becomes the resulting PDF name.
- **archive** – Processed folders from `input` are moved here after OCR.
- **output** – OCR-processed PDF files are saved here.
- **logs** – All process logs are stored here.

---
## 🛠️ Setup & Installation Instructions

1. Download and extract MagickTess zip file contents
   - [MagickTess v1.0.0.0 Release](https://github.com/NeoMatrix14241/magicktess/releases/download/MagickTess-v1.0.0.0/MagickTess-v1.0.0.0.zip)

2. Navigate to the `setup` folder and run "setup.bat".

3. Copy `.traineddata` files from `setup/tessdata_best` to Tesseract's tessdata directory:
   ```
   Default Location: C:\Program Files\Tesseract-OCR\tessdata
   ```
4. Run `start_process.bat` to set up the necessary folders.

---
## ⚙️ Folder Structure & PDF Naming

❌❌❌ **Avoid This Structure:**
```
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
```
- **Issue:** Files at the root of `folder1` (★) will interrupt processing of subfolders (<!>).
- **Solution:** Ensure `.tif` files are inside subfolders.

✔️✔️✔️ **Proper Folder Structure:**
```
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
```
- **PDF Name:** The folder marked with (★) becomes the PDF name.
- **Example:** `subfolder1` generates `subfolder1.pdf`.

---
## ▶️ Usage Instructions

1. Place folders containing image files into the `input` directory.
   ```
   Supported Image Extensions/Types:
   .bmp   .jpeg   .gif   .png
   .dib   .jpe    .tif   .heic
   .jpg   .jiff   .tiff
   ```
3. Run `start_process.bat` and wait for the process to complete.
4. OCR-processed PDF files will be saved in the `output` directory.

---
## 📄 PDF Naming Convention

- **folder1/subfolder1** → `subfolder1.pdf`
- **folder1/subfolder2** → `subfolder2.pdf`
- **folder2** → `folder2.pdf`

---
## 🧪 Experimental (Optional)
- Right-click in the `input` folder while holding **Shift** and select **Open PowerShell window here**.
- Run the following command:
   ```
   .\magicktess.ps1 input
   ```
