$global:cancelRequested = $false

$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    if ($global:cancelRequested) {
        Write-Log "Interrupt detected. Ignoring the exit request.", "Warning"
        $global:cancelRequested = $false
        $Host.UI.WriteLine("Process interrupted. Will continue execution.")
    } else {
        Write-Log "Process Stopped. --> Press 'Ctrl+C' again to exit OR type 'N' then press 'Enter' <--", "Error"
        $global:cancelRequested = $true
    }
}

$scriptDirectory = $PSScriptRoot
$logsFolder = Join-Path $scriptDirectory "logs"
if (-not (Test-Path -Path $logsFolder)) {
    New-Item -ItemType Directory -Path $logsFolder
    Write-Host "[32mCreated logs folder: $logsFolder[0m"
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = Join-Path $logsFolder "magicktesstk-log_$timestamp.log"

function Write-Log {
    param (
        [string]$message,
        [string]$level = "Info"
    )

    $logMessage = "$timestamp - [$level] $message"
    Add-Content -Path $logFilePath -Value $logMessage

    switch ($level) {
        "Info"    { Write-Host "[36m$logMessage[0m" }
        "Warning" { Write-Host "[33m$logMessage[0m" }
        "Error"   { Write-Host "[31m$logMessage[0m" }
        default   { Write-Host $logMessage }
    }
}

Write-Log "Script started." "Info"

if ($args.Length -eq 0) {
    Write-Log "Usage: .\magicktess.ps1 input" "Warning"
    exit
}

$rootFolder = Join-Path $scriptDirectory $args[0]
$outputFolder = Join-Path $scriptDirectory "output"
if (-not (Test-Path -Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder
    Write-Log "Created folder for processed PDFs: $outputFolder" "Info"
}

$archiveFolder = Join-Path $scriptDirectory "archive"
if (-not (Test-Path -Path $archiveFolder)) {
    New-Item -ItemType Directory -Path $archiveFolder
    Write-Log "Created archive folder: $archiveFolder" "Info"
}
# --------------------------------------------------------------------
# IMAGEMAGICK AND TESSERACT OCR PATH / INSTALLATION FOLDER
# --------------------------------------------------------------------
$imageMagickPath = "C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe"
$tesseractPath = "C:\Program Files\Tesseract-OCR\tesseract.exe"
# --------------------------------------------------------------------

if (-not (Test-Path $imageMagickPath)) {
    Write-Log "ImageMagick not found at: $imageMagickPath" "Error"
    exit
} else {
    Write-Log "ImageMagick found at: $imageMagickPath" "Info"
}

if (-not (Test-Path $tesseractPath)) {
    Write-Log "Tesseract not found at: $tesseractPath" "Error"
    exit
} else {
    Write-Log "Tesseract found at: $tesseractPath" "Info"
}

# Define the accepted image file extensions (case insensitive)
$imageExtensions = @(".bmp", ".dib", ".jpg", ".jpeg", ".jpe", ".jiff", ".gif", ".tif", ".tiff", ".png", ".heic")

# Get all subfolders in the root folder, excluding output folders
$subfolders = Get-ChildItem -Path $rootFolder -Recurse -Directory | Where-Object { $_.FullName -notlike "*\output*" }

# Capture the start time for performance tracking
$startTime = Get-Date

# Process each subfolder
foreach ($subfolder in $subfolders) {
    Write-Log "Processing subfolder: $($subfolder.FullName)"
    
    try {
        # Get all image files with the specified extensions (case insensitive)
        $imageFiles = Get-ChildItem -Path $subfolder.FullName -File | Where-Object { $imageExtensions -contains $_.Extension.ToLower() }
    } catch {
        # If there is an error accessing the subfolder, log a warning and skip it
        Write-Log "Warning: Cannot access subfolder $($subfolder.FullName). Skipping this subfolder."
        continue  # Skip this folder and continue with the next one
    }

    if ($imageFiles.Count -gt 0) {
        # If image files are found, process them
        Write-Log "Found $($imageFiles.Count) image files in subfolder $($subfolder.Name). Preprocessing and creating PDF..."

        $baseFileName = [System.IO.Path]::GetFileName($subfolder.FullName)
        
        $parentFolder = Split-Path $subfolder.FullName -Parent

        $relativeSubfolderPath = $parentFolder.Substring($rootFolder.Length).TrimStart('\')
        $outputSubfolder = Join-Path $outputFolder $relativeSubfolderPath
        
        # Create the output subfolder if it doesn't exist
        if (-not (Test-Path -Path $outputSubfolder)) {
            New-Item -ItemType Directory -Path $outputSubfolder -Force
            Write-Log "Created parent subfolder for output: $outputSubfolder"
        }

        # Define the output PDF file path
        $outputPdf = Join-Path $outputSubfolder "$baseFileName"

        # Create a temporary file to store the list of image files
        $tempImageList = [System.IO.Path]::GetTempFileName()

        try {
            # Preprocess each image file using ImageMagick deskew (add _preprocessed suffix)
            $preprocessedImageFiles = @()
            foreach ($imageFile in $imageFiles) {
                $preprocessedImageFile = [System.IO.Path]::Combine($imageFile.DirectoryName, "$($imageFile.BaseName)_preprocessed.tif")
                
                # --------------------------------------------------------------------
                # IMAGEMAGICK COMMAND AND PARAMETERS FOR IMAGE PREPROCESSING
                # --------------------------------------------------------------------
                & $imageMagickPath $imageFile.FullName +repage -deskew 40% $preprocessedImageFile
                # --------------------------------------------------------------------

                Write-Log "Deskewed image created: $preprocessedImageFile"

                $preprocessedImageFiles += $preprocessedImageFile
            }

            # Save the preprocessed image file paths to the temporary file
            $preprocessedImageFiles | Set-Content -Path $tempImageList

            Write-Log "Running Tesseract OCR on preprocessed image files in subfolder $($subfolder.Name)..."

            # --------------------------------------------------------------------
            # TESSERACT-OCR COMMAND AND PARAMETERS FOR DOCUMENT OCR
            # --------------------------------------------------------------------
            & $tesseractPath @($tempImageList) $outputPdf -l eng+enm+fil --oem 1 --psm 6 --loglevel ALL pdf
            # --------------------------------------------------------------------

            Write-Log "PDF created successfully: $outputPdf"

        } catch {
            Write-Log "Error during preprocessing or Tesseract execution on subfolder $($subfolder.Name): $_"
        } finally {
            # Remove the temporary file after processing
            Remove-Item -Path $tempImageList -Force
            Write-Log "Temporary image list file removed: $tempImageList"

            # Clean up preprocessed image files (remove _preprocessed suffix)
            foreach ($preprocessedImageFile in $preprocessedImageFiles) {
                Remove-Item -Path $preprocessedImageFile -Force
                Write-Log "Removed preprocessed image file: $preprocessedImageFile"
            }
        }
    } else {
        # If no image files are found in the subfolder, log a message and skip
        Write-Log "No image files found in subfolder $($subfolder.Name). Skipping..."
    }
}

# Move all input items to the archive folder after processing
Write-Log "Moving all contents from input folder to archive folder..."

try {
    $inputItems = Get-ChildItem -Path $rootFolder -Recurse
    foreach ($item in $inputItems) {
        # Define the destination path for each item
        $destinationPath = Join-Path $archiveFolder $item.FullName.Substring($rootFolder.Length).TrimStart('\')

        # Create the destination directory if it doesn't exist
        $destinationDir = [System.IO.Path]::GetDirectoryName($destinationPath)
        if (-not (Test-Path -Path $destinationDir)) {
            New-Item -ItemType Directory -Path $destinationDir -Force
            Write-Log "Created destination subfolder: $destinationDir"
        }

        # Move the item to the archive folder
        if (Test-Path -Path $item.FullName) {
            Write-Log "Moving item: $($item.FullName) to $destinationPath"
            Move-Item -Path $item.FullName -Destination $destinationPath -Force
        } else {
            Write-Log "Warning: File or folder '$($item.FullName)' already moved or no longer exists. Continuing process..."
            Write-Host "Warning: File or folder '$($item.FullName)' already moved or no longer exists. Continuing process..."
        }
    }
} catch {
    Write-Log "Error while moving items to archive: $_"
}

# Clean up any empty folders in the input directory
Write-Log "Cleaning up empty folders in the input directory..."

function Remove-EmptyFolders {
    param (
        [string]$path
    )
    
    # Get all subfolders and sort them in descending order to remove child folders first
    $subfolders = Get-ChildItem -Path $path -Recurse -Directory -Force | Sort-Object FullName -Descending

    foreach ($folder in $subfolders) {
        # If the folder is empty, remove it
        if (-not (Get-ChildItem -Path $folder.FullName -Force)) {
            Write-Log "Removing empty folder: $($folder.FullName)"
            Remove-Item -Path $folder.FullName -Force
        }
    }
}

# Call the function to remove empty folders
Remove-EmptyFolders -path $rootFolder

# Clean up _preprocessed suffix in the input and archive folders
Write-Log "Cleaning up _preprocessed suffix in the input and archive folders..."

function Remove-PreprocessedSuffix {
    param (
        [string]$path
    )

    # Get all image files with _preprocessed suffix
    $imageFiles = Get-ChildItem -Path $path -Recurse -Filter "*_preprocessed.tif"
    
    foreach ($file in $imageFiles) {
        # Remove the _preprocessed suffix and rename the file
        $newFileName = $file.Name.Replace("_preprocessed", "")
        $newFilePath = Join-Path $file.DirectoryName $newFileName
        
        Rename-Item -Path $file.FullName -NewName $newFilePath
        Write-Log "Renamed file: $($file.FullName) to $newFilePath"
    }
}

# Clean up _preprocessed files in input and archive
Remove-PreprocessedSuffix -path $rootFolder
Remove-PreprocessedSuffix -path $archiveFolder

# Calculate and log the elapsed time
$endTime = Get-Date
$elapsedTime = $endTime - $startTime
Write-Log "Processing complete. Total time: $($elapsedTime.TotalSeconds) seconds"

Write-Log "Script finished."

# Prompt the user to press Enter to exit
Read-Host -Prompt "Press enter to exit"
