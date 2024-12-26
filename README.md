# Automated OCR PowerShell Script

**⚡ Alternative to NAPS2 OCR - Designed for Batch OCR Process**

### 🚀 Batch Folder OCR Only

This PowerShell script automates OCR processing for batches of different image files, generating searchable PDFs.

### Dependencies for MagickTes

The PowerShell script utilizes the following software and models for Tesseract-OCR:

1. **ImageMagick**  
   - Version: ImageMagick-7.1.1-43-Q16-HDRI-x64-dll.exe  
   - [Download ImageMagick](https://imagemagick.org/script/download.php)

2. **Tesseract-OCR**  
   - Version: tesseract-ocr-w64-setup-5.5.0.20241111.exe  
   - [Download Tesseract-OCR](https://github.com/UB-Mannheim/tesseract/wiki)

3. **Tessdata for Tesseract-OCR**  
   - Models: `eng.traineddata`, `enm.traineddata`, `fil.traineddata`  
   - [Download Tessdata models](https://github.com/tesseract-ocr/tessdata/tree/main)

### 📂 Folder Structure
When executed, `start_process.bat` will create the following folders:

- **input** – [Batch OCR Only] Place folders containing `.tif` files here. Each folder name becomes the resulting PDF name.
- **archive** – Processed folders from `input` are moved here after OCR.
- **output** – OCR-processed PDF files are saved here.
- **logs** – All process logs are stored here.

---
## 🛠️ Setup & Installation Instructions

1. Download and extract MagickTes zip file contents
   - [MagickTes v1.0.0.0 Release](https://github.com/NeoMatrix14241/magicktes/releases/tag/MagickTes-v1.0.0.0)

2. Navigate to the `setup` folder and run "setup.bat".

3. Copy `.traineddata` files from `setup/tessdata_best` to Tesseract's tessdata directory:
   ```
   Default Location: C:\Program Files\Tesseract-OCR\tessdata
   ```
4. Run `start_process.bat` to set up the necessary folders, then proceed to usage.

---
## ▶️ Usage Instructions

1. Place folders containing `.tif` files into the `input` directory.
2. Run `start_process.bat` and wait for the process to complete.
3. OCR-processed PDF files will be saved in the `output` directory.

---
## ⚙️ Folder Structure & PDF Naming

**Proper Folder Structure:**
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

**Avoid This Structure:**
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

---
## 📄 PDF Naming Convention

- **folder1/subfolder1** → `subfolder1.pdf`
- **folder1/subfolder2** → `subfolder2.pdf`
- **folder2** → `folder2.pdf`

---
## 🔗 Repository
[Tesseract-OCR Repository](https://github.com/tesseract-ocr/tesseract)

---
## ⚡ Command Line Installer
- `setup/tesseract-ocr-w64-setup-5.5.0.20241111.exe`
- `setup/ImageMagick-7.1.1-41-Q16-HDRI-x64-dll.exe`

---
## ❗ Notes
- Press `CTRL + C` in PowerShell to cancel the operation.

---
## 🧪 Experimental (Optional)
- Right-click in the `input` folder while holding **Shift** and select **Open PowerShell window here**.
- Run the following command:
   ```
   .\magicktes.ps1 input
   ```
