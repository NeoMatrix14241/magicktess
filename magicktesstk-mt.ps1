# Script metadata and initialization
$currentUTC = [System.DateTime]::UtcNow
$currentUser = $env:USERNAME
Write-Host "Current Date and Time (UTC): $($currentUTC.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "Current User's Login: $currentUser"

$global:cancelRequested = $false

# Register an event that triggers when PowerShell is exiting
$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    if ($global:cancelRequested) {
        Write-Log "Interrupt detected. Ignoring the exit request.", "Warning"
        $global:cancelRequested = $false
        $Host.UI.WriteLine("Process interrupted. Will continue execution.")
    }
    else {
        Write-Log "Process Stopped. --> Press 'Ctrl+C' again to exit OR type 'N' then press 'Enter' <--", "Error"
        $global:cancelRequested = $true
    }
}

# Define the script directory and logs folder
$scriptDirectory = $PSScriptRoot
$logsFolder = Join-Path $scriptDirectory "logs"

# Create logs folder if it doesn't exist
if (-not (Test-Path -Path $logsFolder)) {
    New-Item -ItemType Directory -Path $logsFolder
    Write-Host "Created logs folder: $logsFolder"
}

# Generate timestamp and log file path
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = Join-Path $logsFolder "magicktesstk-log_$timestamp.log"

# Function to write messages to the log file and console
function Write-Log {
    param (
        [string]$message,
        [string]$level = "Info",
        [string]$tool = "None"  # New parameter to specify the tool
    )

    $logMessage = "$(Get-Date -Format 'yyyyMMdd_HHmmss') - [$level] $message"
    Add-Content -Path $logFilePath -Value $logMessage

    # Define color based on tool and level
    $color = switch ($tool) {
        "ImageMagick" { "Blue" }      # Blue for ImageMagick
        "Tesseract" { "Green" }       # Green for Tesseract
        "PDFtk" { "DarkYellow" }      # Orange (DarkYellow) for PDFtk
        default {
            switch ($level) {
                "Info" { "White" }     # Changed from Cyan to White
                "Warning" { "Yellow" }
                "Error" { "Red" }
                default { "White" }
            }
        }
    }
    
    Write-Host $logMessage -ForegroundColor $color
}

Write-Log "Script started." "Info"
Write-Log "Current UTC Time: $($currentUTC.ToString('yyyy-MM-dd HH:mm:ss'))" "Info"
Write-Log "Current User: $currentUser" "Info"

# Check if there are any arguments passed to the script
if ($args.Length -eq 0) {
    Write-Log "Usage: .\magicktess.ps1 input" "Warning"
    exit
}

# Define root, output, and archive folders
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

# Paths to ImageMagick, Tesseract OCR, and PDFtk
$imageMagickPath = "C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe"
$tesseractPath = "C:\Program Files\Tesseract-OCR\tesseract.exe"
$pdfTkPath = "C:\Program Files (x86)\PDFtk Server\bin\pdftk.exe"

# Check if ImageMagick and Tesseract OCR are installed
if (-not (Test-Path $imageMagickPath)) {
    Write-Log "ImageMagick not found at: $imageMagickPath" "Error"
    exit
}
else {
    Write-Log "ImageMagick found at: $imageMagickPath" "Info"
}

if (-not (Test-Path $tesseractPath)) {
    Write-Log "Tesseract not found at: $tesseractPath" "Error"
    exit
}
else {
    Write-Log "Tesseract found at: $tesseractPath" "Info"
}

# Define accepted image file extensions
$imageExtensions = @(".bmp", ".dib", ".jpg", ".jpeg", ".jpe", ".jiff", ".gif", ".tif", ".tiff", ".png", ".heic")

# ----------------------------------------------------------------------------------
# Get the number of logical processors (threads) and set a conservative limit
# ----------------------------------------------------------------------------------
$cpuThreads = (Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors
$maxThreads = [Math]::Max(1, [Math]::Floor($cpuThreads * 1.00)) # 0.75 or 75% is the default
Write-Log "Detected $cpuThreads CPU threads. Using $maxThreads threads for processing." "Info"
# ----------------------------------------------------------------------------------

# Function to verify image file integrity
function Test-ImageFileIntegrity {
    param (
        [string]$imagePath
    )
    
    try {
        # Use ImageMagick to identify the image
        $process = Start-Process -FilePath $imageMagickPath -ArgumentList "identify", "`"$imagePath`"" -Wait -NoNewWindow -PassThru
        return $process.ExitCode -eq 0
    }
    catch {
        return $false
    }
}

# Track corrupted files globally using thread-safe collection
$global:corruptedFiles = [System.Collections.Concurrent.ConcurrentBag[string]]::new()

# Add thread-safe counters
$global:processedFiles = [System.Collections.Concurrent.ConcurrentDictionary[string,byte]]::new()
$global:successfulFiles = [System.Collections.Concurrent.ConcurrentDictionary[string,byte]]::new()
$global:failedFiles = [System.Collections.Concurrent.ConcurrentDictionary[string,byte]]::new()

# Get all subfolders in the root folder excluding output folders
$subfolders = Get-ChildItem -Path $rootFolder -Recurse -Directory | Where-Object { $_.FullName -notlike "*\output*" }

Write-Log "Starting parallel file integrity verification phase..." "Info"

# Verify all files in parallel
$subfolders | ForEach-Object -ThrottleLimit $maxThreads -Parallel {
    # Import required variables from parent scope
    $rootFolder = $using:rootFolder
    $imageExtensions = $using:imageExtensions
    $imageMagickPath = $using:imageMagickPath
    $logFilePath = $using:logFilePath
    $corruptedFiles = $using:corruptedFiles

    # Create local function copies
    function Write-Log {
        param (
            [string]$message,
            [string]$level = "Info",
            [string]$tool = "None"
        )
        $logMessage = "$(Get-Date -Format 'yyyyMMdd_HHmmss') - [$level] $message"
        $mutex = [System.Threading.Mutex]::new($false, "LogFileMutex")
        try {
            $mutex.WaitOne() | Out-Null
            Add-Content -Path $logFilePath -Value $logMessage
        }
        finally {
            $mutex.ReleaseMutex()
        }
        
        $color = switch ($tool) {
            "ImageMagick" { "Blue" }
            "Tesseract" { "Green" }
            "PDFtk" { "DarkYellow" }
            default {
                switch ($level) {
                    "Info" { "White" }
                    "Warning" { "Yellow" }
                    "Error" { "Red" }
                    default { "White" }
                }
            }
        }
        Write-Host $logMessage -ForegroundColor $color
    }

    function Test-ImageFileIntegrity {
        param ([string]$imagePath)
        try {
            $process = Start-Process -FilePath $imageMagickPath -ArgumentList "identify", "`"$imagePath`"" -Wait -NoNewWindow -PassThru
            return $process.ExitCode -eq 0
        }
        catch {
            return $false
        }
    }
    
    $imageFiles = Get-ChildItem -Path $_.FullName -File | 
        Where-Object { $imageExtensions -contains $_.Extension.ToLower() }
    
    foreach ($imageFile in $imageFiles) {
        Write-Log "Verifying file integrity: $($imageFile.Name)" "Info"
        if (-not (Test-ImageFileIntegrity -imagePath $imageFile.FullName)) {
            Write-Log "Corrupted file detected: $($imageFile.FullName)" "Error"
            $corruptedFiles.Add($imageFile.FullName)
        }
    }
}

if ($global:corruptedFiles.Count -gt 0) {
    Write-Log "Found $($global:corruptedFiles.Count) corrupted files:" "Error"
    $global:corruptedFiles | ForEach-Object { Write-Log $_ "Error" }
    Write-Log "Will proceed with processing non-corrupted files only." "Warning"
}

# Capture the start time for performance tracking
$startTime = Get-Date

# Process each subfolder in parallel
$subfolders | ForEach-Object -ThrottleLimit $maxThreads -Parallel {
    # Get access to the using variables
    $rootFolder = $using:rootFolder
    $outputFolder = $using:outputFolder
    $imageMagickPath = $using:imageMagickPath
    $tesseractPath = $using:tesseractPath
    $pdfTkPath = $using:pdfTkPath
    $imageExtensions = $using:imageExtensions
    $logFilePath = $using:logFilePath
    $subfolder = $_  # Current item from the pipeline
    $processedFiles = $using:processedFiles
    $successfulFiles = $using:successfulFiles
    $failedFiles = $using:failedFiles

    # Function to write messages to the log file and console
    function Write-Log {
        param (
            [string]$message,
            [string]$level = "Info",
            [string]$tool = "None"
        )

        $logMessage = "$(Get-Date -Format 'yyyyMMdd_HHmmss') - [$level] $message"
        $mutex = [System.Threading.Mutex]::new($false, "LogFileMutex")
        try {
            $mutex.WaitOne() | Out-Null
            Add-Content -Path $logFilePath -Value $logMessage
        }
        finally {
            $mutex.ReleaseMutex()
        }

        # Define color based on tool and level
        $color = switch ($tool) {
            "ImageMagick" { "Blue" }
            "Tesseract" { "Green" }
            "PDFtk" { "DarkYellow" }
            default {
                switch ($level) {
                    "Info" { "White" }
                    "Warning" { "Yellow" }
                    "Error" { "Red" }
                    default { "White" }
                }
            }
        }
        
        Write-Host $logMessage -ForegroundColor $color
    }

    Write-Log "Processing subfolder: $($subfolder.FullName)"

    try {
        # Get all image files with the specified extensions
        $imageFiles = Get-ChildItem -Path $subfolder.FullName -File | 
        Where-Object { $imageExtensions -contains $_.Extension.ToLower() } |
        Sort-Object Name

        if ($imageFiles.Count -gt 0) {
            Write-Log "Found $($imageFiles.Count) image files in subfolder $($subfolder.Name)"

            # Create temp folder for processing with counter for duplicates
            $baseTempFolder = Join-Path $outputFolder "temp_$($subfolder.Name)"
            $tempFolder = $baseTempFolder
            $counter = 1
            while (Test-Path -Path $tempFolder) {
                $tempFolder = "${baseTempFolder}_${counter}"
                $counter++
            }
            New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
            Write-Log "Created temporary processing folder: $tempFolder"

            # Calculate the correct output path maintaining folder structure
            $relativePath = $subfolder.FullName.Substring($rootFolder.Length).TrimStart('\')
            $parentPath = Split-Path -Path $relativePath -Parent
            
            # If this is a root-level folder, put PDF directly in output
            # Otherwise, create the parent folder structure
            $outputPath = if ([string]::IsNullOrEmpty($parentPath)) {
                $outputFolder
            }
            else {
                $outputSubfolder = Join-Path $outputFolder $parentPath
                if (-not (Test-Path -Path $outputSubfolder)) {
                    New-Item -ItemType Directory -Path $outputSubfolder -Force | Out-Null
                    Write-Log "Created output subfolder: $outputSubfolder"
                }
                $outputSubfolder
            }

            # Set final PDF path
            $finalOutputPdf = Join-Path $outputPath "$($subfolder.Name).pdf"
            Write-Log "Will create PDF at: $finalOutputPdf"
            
            # Stage 1: ImageMagick Processing with strict verification
            Write-Log "Starting ImageMagick processing..."
            $preprocessedFiles = @()
            
            foreach ($imageFile in $imageFiles) {
                $processedFiles.TryAdd($imageFile.FullName, 0)
                $preprocessedImageFile = Join-Path $tempFolder "$($imageFile.BaseName)_preprocessed.tif"
                Write-Log "Processing image: $($imageFile.Name)"
                # ----------------------------------------------------------------------------------
                # Process current image
                # ----------------------------------------------------------------------------------
                & $imageMagickPath -quiet $imageFile.FullName -deskew 40% $preprocessedImageFile
                if ($LASTEXITCODE -eq 0) {
                    Write-Log "Successfully processed: $($imageFile.Name)" "Info" "ImageMagick"
                    $successfulFiles.TryAdd($imageFile.FullName, 0)
                    if (Select-String -Path $logFilePath -Pattern "negative image positions unsupported" -Quiet) {
                        Write-Log "Warning: TIFF negative image positions reported but file was successfully preprocessed" "Warning" "ImageMagick"
                    }
                } else {
                    $failedFiles.TryAdd($imageFile.FullName, 0)
                }
                # ----------------------------------------------------------------------------------
                
                # Wait and verify this specific file is processed
                $maxAttempts = 30
                $attempt = 0
                $fileProcessed = $false
                
                while (-not $fileProcessed -and $attempt -lt $maxAttempts) {
                    $attempt++
                    if (Test-Path $preprocessedImageFile) {
                        $fileProcessed = $true
                        $preprocessedFiles += $preprocessedImageFile
                        Write-Log "Verified completion of: $($imageFile.Name)"
                    } else {
                        Write-Log "Waiting for ImageMagick to process $($imageFile.Name)... Attempt $attempt of $maxAttempts" "Info"
                        Start-Sleep -Seconds 1
                    }
                }
                
                if (-not $fileProcessed) {
                    throw "ImageMagick processing timed out for file: $($imageFile.Name)"
                }
                
                # Add small delay between files
                Start-Sleep -Milliseconds 500
            }

            Write-Log "All ImageMagick processing complete. Starting Tesseract OCR..."
            
            # Stage 2: Tesseract Processing with strict verification
            $sortedPreprocessedFiles = $preprocessedFiles | Sort-Object
            $pdfFiles = @()
            $total = $sortedPreprocessedFiles.Count
            $current = 0

            foreach ($preprocessedImageFile in $sortedPreprocessedFiles) {
                $current++
                $progress = [math]::Round(($current / $total) * 100, 2)
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($preprocessedImageFile)
                $tempPdfPath = Join-Path (Split-Path $preprocessedImageFile) $baseName
                $expectedPdfFile = "$tempPdfPath.pdf"

                Write-Log "[TESSERACT][$progress%] Running Tesseract on: $preprocessedImageFile" "Info" "Tesseract"
                # ----------------------------------------------------------------------------------
                # :: language with latest tessdata ::
                # eng = english
                # enm = english old
                # fil = filipino
                # combine with eng+enm+fil in parameters if needed
                # ----------------------------------------------------------------------------------
                & $tesseractPath $preprocessedImageFile $tempPdfPath -l eng --oem 1 --psm 3 pdf
                # ----------------------------------------------------------------------------------
                
                # Wait and verify this specific PDF is generated
                $maxAttempts = 30
                $attempt = 0
                $pdfGenerated = $false
                
                while (-not $pdfGenerated -and $attempt -lt $maxAttempts) {
                    $attempt++
                    if (Test-Path $expectedPdfFile) {
                        $pdfGenerated = $true
                        $pdfFiles += $expectedPdfFile
                        Write-Log "Verified PDF generation for: $preprocessedImageFile"
                        Remove-Item -Path $preprocessedImageFile -Force
                    } else {
                        Write-Log "Waiting for Tesseract to generate PDF... Attempt $attempt of $maxAttempts" "Info"
                        Start-Sleep -Seconds 1
                    }
                }
                
                if (-not $pdfGenerated) {
                    throw "Tesseract processing timed out for file: $preprocessedImageFile"
                }
                
                # Add small delay between files
                Start-Sleep -Milliseconds 500
            }

            # Stage 3: PDFtk Processing
            Write-Log "Starting PDF merge..."
            $sortedPdfFiles = $pdfFiles | Where-Object { Test-Path $_ } | Sort-Object

            # Add verification of all PDFs before merging
            $missingFiles = $pdfFiles | Where-Object { -not (Test-Path $_) }
            if ($missingFiles.Count -gt 0) {
                Write-Log "Missing PDF files detected:" "Error"
                $missingFiles | ForEach-Object { Write-Log "Missing: $_" "Error" }
                throw "Cannot proceed with merge - missing PDF files"
            }

            if ($sortedPdfFiles.Count -eq $imageFiles.Count) {
                Write-Log "Verified all $($imageFiles.Count) images were converted to PDFs"
            } else {
                throw "PDF count mismatch. Expected: $($imageFiles.Count), Found: $($sortedPdfFiles.Count)"
            }

            if ($sortedPdfFiles.Count -gt 0) {
                try {
                    # Create output directory if needed
                    $finalOutputDir = Split-Path -Parent $finalOutputPdf
                    if (-not (Test-Path $finalOutputDir)) {
                        New-Item -ItemType Directory -Path $finalOutputDir -Force | Out-Null
                    }

                    Write-Log "[PDFTK] Executing merge..." "Info" "PDFtk"
                    # ----------------------------------------------------------------------------------
                    # Prepare PDFtk arguments
                    # ----------------------------------------------------------------------------------
                    $pdfTkArgs = @($sortedPdfFiles | ForEach-Object { "`"$_`"" })
                    $pdfTkArgs += "cat", "output", "`"$finalOutputPdf`""
                    Write-Host "`e[33m[PDFTK] Executing merge...`e[0m"
                    $process = Start-Process -FilePath $pdfTkPath -ArgumentList $pdfTkArgs -Wait -NoNewWindow -PassThru
                    # ----------------------------------------------------------------------------------

                    if ($process.ExitCode -eq 0 -and (Test-Path $finalOutputPdf)) {
                        Write-Log "Successfully created combined PDF: $finalOutputPdf"
                        
                        # Clean up individual PDF files
                        $sortedPdfFiles | ForEach-Object {
                            Remove-Item -Path $_ -Force
                            Write-Log "Cleaned up temporary PDF: $_"
                        }
                    } else {
                        throw "PDFtk failed with exit code: $($process.ExitCode)"
                    }
                }
                catch {
                    Write-Log "Error during PDF combination: $_" "Error"
                    throw
                }
            } else {
                throw "No PDF files found to merge"
            }
        }
        else {
            Write-Log "No image files found in subfolder $($subfolder.FullName). Skipping..."
        }
    }
    catch {
        Write-Log "Error processing subfolder $($subfolder.Name): $_" "Error"
        Write-Log $_.ScriptStackTrace "Error"
        
        # Cleanup temp folder in case of error
        if (Test-Path $tempFolder) {
            Remove-Item -Path $tempFolder -Recurse -Force
            Write-Log "Cleaned up temp folder after error: $tempFolder"
        }
    }
}

# Final cleanup of any remaining temp folders
Write-Log "Performing final cleanup..."
Get-ChildItem -Path $outputFolder -Include "temp_*" -Directory -Recurse | ForEach-Object {
    # Ensure all preprocessed files are removed from temp folders
    Get-ChildItem -Path $_.FullName -Filter "*_preprocessed.tif" | ForEach-Object {
        Remove-Item $_.FullName -Force
        Write-Log "Removed preprocessed file: $($_.FullName)" "Info"
    }
    Remove-Item $_.FullName -Recurse -Force
    Write-Log "Cleaned up remaining temp folder: $($_.FullName)"
}

Write-Log "Moving non-corrupted contents from input folder to archive folder..."

try {
    # Only move original files (exclude any preprocessed files that might have leaked)
    $inputItems = Get-ChildItem -Path $rootFolder -Recurse | 
                 Where-Object { 
                     $_.Name -notlike "*_preprocessed*" -and 
                     $_.FullName -notin $global:corruptedFiles 
                 }
    
    foreach ($item in $inputItems) {
        $destinationPath = Join-Path $archiveFolder $item.FullName.Substring($rootFolder.Length).TrimStart('\')
        $destinationDir = [System.IO.Path]::GetDirectoryName($destinationPath)
        if (-not (Test-Path -Path $destinationDir)) {
            New-Item -ItemType Directory -Path $destinationDir -Force
            Write-Log "Created destination subfolder: $destinationDir"
        }

        if (Test-Path -Path $item.FullName) {
            Write-Log "Moving item: $($item.FullName) to $destinationPath"
            Move-Item -Path $item.FullName -Destination $destinationPath -Force
        }
        else {
            Write-Log "Warning: File or folder '$($item.FullName)' already moved or no longer exists. Continuing process..."
        }
    }
}
catch {
    Write-Log "Error while moving items to archive: $_"
}

if ($global:corruptedFiles.Count -gt 0) {
    Write-Log "The following corrupted files remain in the input folder:" "Warning"
    $global:corruptedFiles | ForEach-Object { Write-Log $_ "Warning" }
}

Write-Log "Cleaning up empty folders in the input directory..."

function Remove-EmptyFolders {
    param (
        [string]$path
    )
    
    $subfolders = Get-ChildItem -Path $path -Recurse -Directory -Force | Sort-Object FullName -Descending

    foreach ($folder in $subfolders) {
        if (-not (Get-ChildItem -Path $folder.FullName -Force)) {
            Write-Log "Removing empty folder: $($folder.FullName)"
            Remove-Item -Path $folder.FullName -Force
        }
    }
}

Remove-EmptyFolders -path $rootFolder

$endTime = Get-Date
$elapsedTime = $endTime - $startTime

# Statistics summary at the end
Write-Log "========================================" "Info"
Write-Log "Final Processing Statistics:" "Info"
Write-Log "----------------------------------------" "Info"
Write-Log "Total Processed Files: $($global:processedFiles.Count)" "Info"
Write-Log "Successfully Processed: $($global:successfulFiles.Count)" "Info"
Write-Log "Failed/Corrupted Files: $($global:failedFiles.Count)" "Info"
Write-Log "Total Processing Time: $($elapsedTime.TotalSeconds) seconds" "Info"
Write-Log "========================================" "Info"

Write-Log "Script finished."

Read-Host -Prompt "Press enter to exit"