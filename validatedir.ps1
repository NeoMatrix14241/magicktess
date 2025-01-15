# Function to read INI file
function Get-IniContent {
    param (
        [string]$filePath
    )
    $ini = @{}
    $section = ""
    foreach ($line in Get-Content $filePath) {
        $line = $line.Trim()
        if ($line -match "^\[(.+)\]$") {
            $section = $matches[1]
            $ini[$section] = @{}
        } elseif ($line -match "^(.*)=(.*)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $ini[$section][$key] = $value
        }
    }
    return $ini
}

# Path to the INI file
$iniFilePath = Join-Path $PSScriptRoot "settings.ini"

# Read the INI file
$iniContent = Get-IniContent -filePath $iniFilePath

# Get folder paths and settings
$inputFolder = $iniContent["Folders"]["InputFolder"]
$outputFolder = $iniContent["Folders"]["OutputFolder"]
$archiveFolder = $iniContent["Folders"]["ArchiveFolder"]
$createMissingFolders = $iniContent["Folders"]["CreateMissingFolders"]

# Function to create folder if it doesn't exist
function New-FolderIfMissing {
    param (
        [string]$folderPath
    )
    if (-not (Test-Path -Path $folderPath)) {
        New-Item -ItemType Directory -Path $folderPath | Out-Null
        Write-Output "Created folder: $folderPath"
    } else {
        Write-Output "Folder already exists: $folderPath"
    }
}

# Check and create folders if necessary
if ($createMissingFolders -eq "ON") {
    New-FolderIfMissing -folderPath $inputFolder
    New-FolderIfMissing -folderPath $outputFolder
    New-FolderIfMissing -folderPath $archiveFolder
} else {
    Write-Output "CreateMissingFolders is set to OFF. No folders were created."
}