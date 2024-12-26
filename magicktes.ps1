# Set a global variable to track if a cancel request is made
$global:cancelRequested = $false

# Register an event for when PowerShell exits
$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    if ($global:cancelRequested) {
        Write-Log "Interrupt detected. Ignoring the exit request."
        $global:cancelRequested = $false
        $Host.UI.WriteLine("Process interrupted. Will continue execution.")
    } else {
        Write-Log "Process Stopped. --> Press 'Ctrl+C' again to exit OR type 'N' then press 'Enter' <--"
        $global:cancelRequested = $true
    }
}

# Get the directory where the script is located
$scriptDirectory = $PSScriptRoot

# Define a folder for storing logs
$logsFolder = Join-Path $scriptDirectory "logs"
if (-not (Test-Path -Path $logsFolder)) {
    New-Item -ItemType Directory -Path $logsFolder
    Write-Host "Created logs folder: $logsFolder"
}

# Create a timestamp for log files
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = Join-Path $logsFolder "magicktes-log_$timestamp.log"

# Define a function to write messages to the log file
function Write-Log {
    param (
        [string]$message
    )
    $logMessage = "$timestamp - $message"
    Add-Content -Path $logFilePath -Value $logMessage  # Write to log file
    Write-Host $logMessage  # Display to console
}

Write-Log "Script started."

# Check if an input argument is provided; if not, show usage and exit
if ($args.Length -eq 0) {
    Write-Log "Usage: .\magicktes.ps1 input"
    Write-Host "Usage: .\magicktes.ps1 input"
    exit
}

# Define the root folder based on the provided argument
$rootFolder = Join-Path $scriptDirectory $args[0]

# Create an output folder for processed files if it doesn't exist
$outputFolder = Join-Path $scriptDirectory "output"
if (-not (Test-Path -Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder
    Write-Log "Created folder for processed PDFs: $outputFolder"
}

# Create an archive folder if it doesn't exist (to store original files after processing)
$archiveFolder = Join-Path $scriptDirectory "archive"
if (-not (Test-Path -Path $archiveFolder)) {
    New-Item -ItemType Directory -Path $archiveFolder
    Write-Log "Created archive folder: $archiveFolder"
}

# --------------------------------------------------------------------
# WHERE THE IMAGEMAGICK (magick) AND TESSERACT-OCR COMMAND LINE ARE INSTALLED
# --------------------------------------------------------------------
$imageMagickPath = "C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe"
$tesseractPath = "C:\Program Files\Tesseract-OCR\tesseract.exe"
# --------------------------------------------------------------------

# Check if ImageMagick and Tesseract are installed
if (-not (Test-Path $imageMagickPath)) {
    Write-Log "ImageMagick not found at: $imageMagickPath"
    Write-Host "ImageMagick not found at: $imageMagickPath"
    exit
} else {
    Write-Log "ImageMagick found at: $imageMagickPath"
}

if (-not (Test-Path $tesseractPath)) {
    Write-Log "Tesseract not found at: $tesseractPath"
    Write-Host "Tesseract not found at: $tesseractPath"
    exit
} else {
    Write-Log "Tesseract found at: $tesseractPath"
}

# Get all subfolders in the root folder, excluding output folders
$subfolders = Get-ChildItem -Path $rootFolder -Recurse -Directory | Where-Object { $_.FullName -notlike "*\output*" }

# Capture the start time for performance tracking
$startTime = Get-Date

# Process each subfolder
foreach ($subfolder in $subfolders) {
    Write-Log "Processing subfolder: $($subfolder.FullName)"
    
    try {
        # Try to get all TIF files from the current subfolder
        $tifFiles = Get-ChildItem -Path $subfolder.FullName -Filter "*.tif" -ErrorAction Stop
    } catch {
        # If there is an error accessing the subfolder, log a warning and skip it
        Write-Log "Warning: Cannot access subfolder $($subfolder.FullName). Skipping this subfolder."
        continue  # Skip this folder and continue with the next one
    }

    if ($tifFiles.Count -gt 0) {
        # If TIF files are found, process them
        Write-Log "Found $($tifFiles.Count) TIF files in subfolder $($subfolder.Name). Preprocessing and creating PDF..."

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

        # Create a temporary file to store the list of TIF images
        $tempImageList = [System.IO.Path]::GetTempFileName()

        try {
            # Preprocess each TIF file using ImageMagick deskew (add _preprocessed suffix)
            $preprocessedTifFiles = @()
            foreach ($tifFile in $tifFiles) {
                $preprocessedTifFile = [System.IO.Path]::Combine($tifFile.DirectoryName, "$($tifFile.BaseName)_preprocessed.tif")
                
                # --------------------------------------------------------------------
                # IMAGEMAGICK COMMAND AND PARAMETERS FOR IMAGE PREPROCESSING
                # --------------------------------------------------------------------
                & $imageMagickPath $tifFile.FullName +repage -deskew 40% $preprocessedTifFile
                # --------------------------------------------------------------------

                Write-Log "Deskewed image created: $preprocessedTifFile"

                $preprocessedTifFiles += $preprocessedTifFile
            }

            # Save the preprocessed TIF file paths to the temporary file
            $preprocessedTifFiles | Set-Content -Path $tempImageList

            Write-Log "Running Tesseract OCR on preprocessed TIF files in subfolder $($subfolder.Name)..."

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

            # Clean up preprocessed TIF files (remove _preprocessed suffix)
            foreach ($preprocessedTifFile in $preprocessedTifFiles) {
                Remove-Item -Path $preprocessedTifFile -Force
                Write-Log "Removed preprocessed TIF file: $preprocessedTifFile"
            }
        }
    } else {
        # If no TIF files are found in the subfolder, log a message and skip
        Write-Log "No TIF files found in subfolder $($subfolder.Name). Skipping..."
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

    # Get all TIF files with _preprocessed suffix
    $tifFiles = Get-ChildItem -Path $path -Recurse -Filter "*_preprocessed.tif"
    
    foreach ($file in $tifFiles) {
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
