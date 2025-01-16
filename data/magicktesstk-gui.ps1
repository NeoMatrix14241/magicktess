# Add PInvoke definitions first, before any other code
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class User32 {
        [DllImport("user32.dll")]
        public static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
        
        [DllImport("user32.dll")]
        public static extern bool EnableMenuItem(IntPtr hMenu, uint uIDEnableItem, uint uEnable);
    }
"@ -ErrorAction Stop

# Regular imports
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Simplified cleanup handler
$pidFile = Join-Path $env:TEMP "magicktesstk.pid"
$cleanupScript = {
    if (Test-Path $pidFile) {
        try {
            $procId = Get-Content $pidFile
            $proc = Get-Process -Id $procId -ErrorAction SilentlyContinue
            if ($proc) {
                $proc | Stop-Process -Force
            }
        } catch {}
        Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:TEMP\output.txt" -Force -ErrorAction SilentlyContinue
    }
}

# Register cleanup for different exit scenarios
Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action $cleanupScript
trap { & $cleanupScript }

# Load current settings
$iniPath = Join-Path $PSScriptRoot "settings.ini"
$ini = @{}
$section = "NoSection"
if (Test-Path $iniPath) {
    Get-Content $iniPath | ForEach-Object {
        if ($_ -match '^\[(.+)\]$') {
            $section = $matches[1]
            $ini[$section] = @{}
        }
        elseif ($_ -match '^([^;].+?)=(.+)$') {
            $name, $value = $matches[1..2]
            $ini[$section][$name] = $value
        }
    }
}

# Get GUI settings with defaults
$guiSettings = $ini["GUI"] ?? @{
    WindowWidth = "900"
    WindowHeight = "500"
    BackgroundColor = "#001B1B"
    TextColor = "#00FF00"
    TitleFontSize = "24"
    WindowStartupLocation = "CenterScreen"
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MagickTessTK OCR - Automated OCR Processing Tool" 
        Height="$($guiSettings.WindowHeight)" Width="$($guiSettings.WindowWidth)"
        Background="$($guiSettings.BackgroundColor)" 
        WindowStartupLocation="$($guiSettings.WindowStartupLocation)"
        ResizeMode="CanMinimize">
    <Grid Name="MainGrid">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="400"/>
        </Grid.ColumnDefinitions>

        <!-- Left side - Controls -->
        <Grid Grid.Column="0" Margin="10">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <TextBlock Grid.Row="0" Text="NAPS2 on Steroids" 
                      Foreground="$($guiSettings.TextColor)" 
                      FontSize="$($guiSettings.TitleFontSize)" 
                      Margin="0,0,0,10"/>

            <GroupBox Grid.Row="2" Header="Folder Settings" Foreground="$($guiSettings.TextColor)" Margin="0,10">
                <StackPanel>
                    <DockPanel Margin="0,5">
                        <TextBlock Text="Input Folder: " Foreground="$($guiSettings.TextColor)" Width="100"/>
                        <Button Content="Browse" Width="70" DockPanel.Dock="Right" Name="btnInputBrowse"/>
                        <TextBox Name="txtInputFolder" Margin="5,0"/>
                    </DockPanel>
                    
                    <DockPanel Margin="0,5">
                        <TextBlock Text="Output Folder: " Foreground="$($guiSettings.TextColor)" Width="100"/>
                        <Button Content="Browse" Width="70" DockPanel.Dock="Right" Name="btnOutputBrowse"/>
                        <TextBox Name="txtOutputFolder" Margin="5,0"/>
                    </DockPanel>
                    
                    <DockPanel Margin="0,5">
                        <TextBlock Text="Archive Folder: " Foreground="$($guiSettings.TextColor)" Width="100"/>
                        <Button Content="Browse" Width="70" DockPanel.Dock="Right" Name="btnArchiveBrowse"/>
                        <TextBox Name="txtArchiveFolder" Margin="5,0"/>
                    </DockPanel>
                </StackPanel>
            </GroupBox>

            <GroupBox Grid.Row="3" Header="ImageMagick Settings" Foreground="$($guiSettings.TextColor)" Margin="0,10">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Compression: " Foreground="$($guiSettings.TextColor)" Width="120" Margin="0,5"/>
                    <ComboBox Grid.Row="0" Grid.Column="1" Grid.ColumnSpan="2" Name="cmbCompression" Margin="5,5">
                        <ComboBoxItem>LZW</ComboBoxItem>
                        <ComboBoxItem>ZIP</ComboBoxItem>
                        <ComboBoxItem>RLE</ComboBoxItem>
                        <ComboBoxItem>NONE</ComboBoxItem>
                        <ComboBoxItem>JPEG</ComboBoxItem>
                        <ComboBoxItem>WEBP</ComboBoxItem>
                    </ComboBox>

                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Quality (Default: 75%): " Foreground="$($guiSettings.TextColor)" Width="120" Margin="0,5"/>
                    <Grid Grid.Row="1" Grid.Column="1">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <Slider Name="sldQuality" Grid.Row="0" Minimum="0" Maximum="100" Margin="5,5" 
                               TickPlacement="BottomRight" 
                               TickFrequency="25"
                               IsMoveToPointEnabled="True"/>
                        
                    </Grid>
                    <TextBlock Grid.Row="1" Grid.Column="2" Name="txtQualityValue" Foreground="$($guiSettings.TextColor)" Margin="5,5" MinWidth="30"/>

                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Deskew (Default: 45%): " Foreground="$($guiSettings.TextColor)" Width="120" Margin="0,5"/>
                    <Grid Grid.Row="2" Grid.Column="1">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <Slider Name="sldDeskew" Grid.Row="0" Minimum="0" Maximum="100" Margin="5,5"
                               TickPlacement="BottomRight"
                               TickFrequency="20"
                               IsMoveToPointEnabled="True"/>
                    </Grid>
                    <TextBlock Grid.Row="2" Grid.Column="2" Name="txtDeskewValue" Foreground="$($guiSettings.TextColor)" Margin="5,5" MinWidth="30"/>

                    <TextBlock Grid.Row="3" Grid.Column="0" Text="Colorspace: " Foreground="$($guiSettings.TextColor)" Width="120" Margin="0,5"/>
                    <ComboBox Grid.Row="3" Grid.Column="1" Grid.ColumnSpan="2" Name="cmbColorspace" Margin="5,5"/>
                </Grid>
            </GroupBox>

            <GroupBox Grid.Row="4" Header="Tesseract Settings" Foreground="$($guiSettings.TextColor)" Margin="0,10">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Language: " Foreground="$($guiSettings.TextColor)" Width="100" Margin="0,5"/>
                    <ComboBox Grid.Row="0" Grid.Column="1" Name="cmbLanguage" Margin="5,5"/>

                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Engine Mode: " Foreground="$($guiSettings.TextColor)" Width="100" Margin="0,5"/>
                    <ComboBox Grid.Row="1" Grid.Column="1" Name="cmbEngineMode" Margin="5,5">
                        <ComboBoxItem>0 - Legacy engine only</ComboBoxItem>
                        <ComboBoxItem>1 - Neural nets LSTM engine only</ComboBoxItem>
                        <ComboBoxItem>2 - Legacy + LSTM engines</ComboBoxItem>
                        <ComboBoxItem>3 - Default</ComboBoxItem>
                    </ComboBox>

                    <TextBlock Grid.Row="2" Grid.Column="0" Text="PSM: " Foreground="$($guiSettings.TextColor)" Width="100" Margin="0,5"/>
                    <ComboBox Grid.Row="2" Grid.Column="1" Name="cmbPSM" Margin="5,5"/>
                </Grid>
            </GroupBox>

            <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10">
                <Button Name="btnOpenLogs" Content="Open Logs" Width="100" Margin="5,0"/>
                <Button Name="btnSave" Content="Save Settings" Width="100" Margin="5,0"/>
                <Button Name="btnStart" Content="Start OCR Process" Width="120" Margin="5,0"/>
                <Button Name="btnCancel" Content="Cancel" Width="80" IsEnabled="False"/>
            </StackPanel>
        </Grid>

        <!-- Splitter -->
        <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Center" VerticalAlignment="Stretch"/>

        <!-- Right side - Process output -->
        <GroupBox Grid.Column="2" Header="Process Output" Foreground="$($guiSettings.TextColor)" Margin="10">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <TextBox Name="txtOutput" 
                         IsReadOnly="True" 
                         Background="Black" 
                         Foreground="$($guiSettings.TextColor)"
                         FontFamily="Consolas"
                         TextWrapping="Wrap"
                         AcceptsReturn="True"
                         VerticalAlignment="Stretch"/>
            </ScrollViewer>
        </GroupBox>
    </Grid>
</Window>
"@

# Add error handling for XAML loading
try {
    Write-Host "Loading XAML..."
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    if ($null -eq $window) {
        throw "Failed to load XAML: Window is null"
    }
    Write-Host "XAML loaded successfully"

    # Disable close button in title bar
    Write-Host "Configuring window properties..."
    $windowHandle = (New-Object System.Windows.Interop.WindowInteropHelper($window)).Handle
    $closeMenuItem = 0xF060
    $windowStyle = [User32]::GetSystemMenu($windowHandle, $false)
    [User32]::EnableMenuItem($windowStyle, $closeMenuItem, 0x1)
    Write-Host "Window configured"

} catch {
    Write-Error "Error initializing GUI: $_"
    Write-Error $_.ScriptStackTrace
    Start-Sleep -Seconds 5
    exit 1
}

# Add this after XAML loading and before other event handlers
$window.Add_Closing({
    param($sender, $e)
    
    # Prevent window from closing with X button
    $e.Cancel = $true
})

# Get controls
$btnInputBrowse = $window.FindName("btnInputBrowse")
$btnOutputBrowse = $window.FindName("btnOutputBrowse")
$btnArchiveBrowse = $window.FindName("btnArchiveBrowse")
$txtInputFolder = $window.FindName("txtInputFolder")
$txtOutputFolder = $window.FindName("txtOutputFolder")
$txtArchiveFolder = $window.FindName("txtArchiveFolder")
$btnSave = $window.FindName("btnSave")
$btnStart = $window.FindName("btnStart")
$txtOutput = $window.FindName("txtOutput")

# Get additional controls
$cmbCompression = $window.FindName("cmbCompression")
$sldQuality = $window.FindName("sldQuality")
$sldDeskew = $window.FindName("sldDeskew")
$cmbColorspace = $window.FindName("cmbColorspace")
$cmbLanguage = $window.FindName("cmbLanguage")
$cmbEngineMode = $window.FindName("cmbEngineMode")
$cmbPSM = $window.FindName("cmbPSM")
$txtQualityValue = $window.FindName("txtQualityValue")
$txtDeskewValue = $window.FindName("txtDeskewValue")
$btnOpenLogs = $window.FindName("btnOpenLogs")

# Initialize ComboBox data - do this only once
$compressionTypes = @("LZW", "ZIP", "RLE", "NONE", "JPEG", "WEBP")
$colorspaces = @("Auto", "RGB", "CMYK", "GRAY", "sRGB", "scRGB", "Lab", "XYZ", "HSL", "HSB", "YCbCr", "CMY")
$psmOptions = 0..13 | ForEach-Object { "PSM $_" }

# Set ItemsSource for ComboBoxes
$cmbCompression.Items.Clear()
$compressionTypes | ForEach-Object { $cmbCompression.Items.Add($_) }

$cmbColorspace.Items.Clear()
$colorspaces | ForEach-Object { $cmbColorspace.Items.Add($_) }

$cmbPSM.Items.Clear()
$psmOptions | ForEach-Object { $cmbPSM.Items.Add($_) }

# Define language combinations with display names
$languageCombos = @(
    @{ Display = "English"; Value = "eng" }
    @{ Display = "English + English (Old)"; Value = "eng+enm" }
    @{ Display = "English + Filipino"; Value = "eng+fil" }
    @{ Display = "English + Filipino + English (Old)"; Value = "eng+fil+enm" }
    @{ Display = "Filipino"; Value = "fil" }
    @{ Display = "Filipino + English"; Value = "fil+eng" }
    @{ Display = "Filipino + English (Old)"; Value = "fil+enm" }
    @{ Display = "English (Old)"; Value = "enm" }
    @{ Display = "English (Old) + English"; Value = "enm+eng" }
    @{ Display = "English (Old) + Filipino"; Value = "enm+fil" }
)

# Populate language ComboBox
$cmbLanguage.Items.Clear()
foreach ($combo in $languageCombos) {
    $item = New-Object System.Windows.Controls.ComboBoxItem
    $item.Content = $combo.Display
    $item.Tag = $combo.Value
    $cmbLanguage.Items.Add($item)
}

# Load current settings
$iniPath = Join-Path $PSScriptRoot "settings.ini"
if (Test-Path $iniPath) {
    $ini = @{}
    $section = "NoSection"
    Get-Content $iniPath | ForEach-Object {
        if ($_ -match '^\[(.+)\]$') {
            $section = $matches[1]
            $ini[$section] = @{}
        }
        elseif ($_ -match '^([^;].+?)=(.+)$') {
            $name, $value = $matches[1..2]
            $ini[$section][$name] = $value
        }
    }

    # Load folder settings
    if ($ini.ContainsKey("Folders")) {
        $txtInputFolder.Text = $ini.Folders.InputFolder
        $txtOutputFolder.Text = $ini.Folders.OutputFolder
        $txtArchiveFolder.Text = $ini.Folders.ArchiveFolder
    }

    # Load ImageMagick settings
    if ($ini.ContainsKey("ImageMagick")) {
        try {
            # Use SelectedItem for ItemsSource-bound ComboBoxes
            $cmbCompression.SelectedItem = $ini.ImageMagick.CompressionType
            $sldQuality.Value = [double]$ini.ImageMagick.Quality
            $sldDeskew.Value = [double]($ini.ImageMagick.DeskewThreshold -replace '%')
            $cmbColorspace.SelectedItem = $ini.ImageMagick.Colorspace
        } catch {
            Write-Warning "Error loading ImageMagick settings: $_"
        }
    }

    # Load TesseractOCR settings
    if ($ini.ContainsKey("TesseractOCR")) {
        try {
            $savedLanguage = $ini.TesseractOCR.Language
            $cmbLanguage.SelectedItem = $cmbLanguage.Items | Where-Object { $_.Tag -eq $savedLanguage } | Select-Object -First 1
            $cmbEngineMode.SelectedIndex = [int]$ini.TesseractOCR.OCREngineMode
            $cmbPSM.SelectedIndex = [int]$ini.TesseractOCR.PageSegmentationMode
        } catch {
            Write-Warning "Error loading TesseractOCR settings: $_"
        }
    }
}

# Browse button handlers
$browse = {
    param($textBox)
    $folder = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folder.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBox.Text = $folder.SelectedPath
    }
}

$btnInputBrowse.Add_Click({ & $browse $txtInputFolder })
$btnOutputBrowse.Add_Click({ & $browse $txtOutputFolder })
$btnArchiveBrowse.Add_Click({ & $browse $txtArchiveFolder })

# Add global variables for process tracking
$script:currentProcess = $null
$script:outputTimer = $null

# Cleanup function
function Stop-CurrentProcess {
    if ($script:currentProcess -and !$script:currentProcess.HasExited) {
        Write-Host "Stopping current process and its children..."
        # Kill all child processes recursively
        function Kill-ProcessTree {
            param($ProcessId)
            Get-CimInstance Win32_Process | Where-Object { $_.ParentProcessId -eq $ProcessId } | ForEach-Object {
                Kill-ProcessTree $_.ProcessId
                Write-Host "Stopping child process: $($_.ProcessId)"
                Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
            }
        }

        # Kill the process tree starting from the main process
        Kill-ProcessTree $script:currentProcess.Id
        $script:currentProcess | Stop-Process -Force
        $script:currentProcess = $null
        Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
    }

    # Find and kill only magicktesstk-mt.ps1 related processes
    Get-Process | Where-Object { 
        $_.Name -like "*pwsh*" -and $_.Id -ne $PID
    } | ForEach-Object {
        try {
            $cmdLine = (Get-CimInstance Win32_Process -Filter "ProcessId = $($_.Id)").CommandLine
            if ($cmdLine -like "*magicktesstk-mt.ps1*") {
                Write-Host "Stopping OCR process: $($_.Id)"
                Stop-Process -Id $_.Id -Force
            }
        } catch {}
    }

    if ($script:outputTimer) {
        $script:outputTimer.Stop()
        $script:outputTimer = $null
    }
    if ($script:outputJob) {
        Stop-Job -Job $script:outputJob
        Remove-Job -Job $script:outputJob -Force
        $script:outputJob = $null
    }
    Remove-Item "$env:TEMP\output.txt" -ErrorAction SilentlyContinue
    $btnStart.IsEnabled = $true
    $btnCancel.IsEnabled = $false
}

# Replace the window.Add_Closing handler with this version
$window.Add_Closing({
    param($sender, $e)
    
    # Show confirmation dialog
    $result = [System.Windows.MessageBox]::Show(
        "Are you sure you want to close the application?",
        "Confirm Close",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        try {
            # Create wait message
            $waitMessage = New-Object System.Windows.Controls.TextBlock
            $waitMessage.Text = "Cleaning up processes..."
            $waitMessage.Foreground = "#00FF00"
            $waitMessage.HorizontalAlignment = "Center"
            $waitMessage.VerticalAlignment = "Center"
            $waitMessage.FontSize = 14
            
            # Add message to window
            $mainGrid = $window.FindName("MainGrid")
            if ($mainGrid) {
                $mainGrid.Children.Clear()
                $mainGrid.Children.Add($waitMessage)
            }
            
            # Stop all processes first
            Stop-CurrentProcess

            # Clean up event subscribers
            Get-EventSubscriber -SourceIdentifier PowerShell.Exiting -ErrorAction SilentlyContinue | 
                ForEach-Object {
                    Remove-Event -SourceIdentifier $_.SourceIdentifier -ErrorAction SilentlyContinue
                    Unregister-Event -SubscriptionId $_.SubscriptionId -ErrorAction SilentlyContinue
                }

            # Clean up temp files
            Remove-Item "$env:TEMP\output.txt" -Force -ErrorAction SilentlyContinue
            Remove-Item $pidFile -Force -ErrorAction SilentlyContinue

            # Allow window to close normally
            $e.Cancel = $false
        }
        catch {
            Write-Warning "Error during cleanup: $_"
            # Show error and let user decide
            $errorResult = [System.Windows.MessageBox]::Show(
                "Error during cleanup. Force close application?",
                "Error",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Error
            )
            
            if ($errorResult -eq [System.Windows.MessageBoxResult]::Yes) {
                Stop-Process -Id $PID -Force
            } else {
                $e.Cancel = $true
            }
        }
    }
    else {
        $e.Cancel = $true
    }
})

# Function to save settings
function Save-Settings {
    $content = @"
[Folders]
; -----------------------------------------------------------------------------
; Folder Path Settings
; -----------------------------------------------------------------------------

; Base working directories for the script
; Use absolute paths for best results
InputFolder=$($txtInputFolder.Text)
OutputFolder=$($txtOutputFolder.Text)
ArchiveFolder=$($txtArchiveFolder.Text)

; Create folders if they don't exist (ON/OFF)
CreateMissingFolders=ON

[ImageMagick]
; -----------------------------------------------------------------------------
; ImageMagick Processing Settings
; -----------------------------------------------------------------------------

; Compression type to use for image processing
; Valid options: 
;   - Lossless: LZW, ZIP, RLE, NONE
;   - Lossy: JPEG, WEBP
CompressionType=$($cmbCompression.SelectedItem)

; Quality setting for lossy compression (0-100)
; Only applies when using lossy compression types (JPEG/WEBP)
; Higher values = better quality but larger file size
; Recommended range: 70-100
; Ignored when using lossless compression
Quality=$($sldQuality.Value)

; Enable/Disable additional image processing parameters
; Valid options: ON, OFF
; IF ENABLED, DESKEW AND COLORSPACE WILL BE MARKED AND USED AS ACTIVE PARAMETERS
AdditionalParameters=ON

; Deskew threshold percentage
; Valid range: 0% to 100%
; Lower values are less aggressive in deskew correction
; Disabled: 0%
; Recommended: 40%
; Higher values are more aggressive in deskew correction
DeskewThreshold=$($sldDeskew.Value)%


; Color space conversion
; Valid options:
;   Auto          - Automatically detect and maintain original colorspace
;   RGB           - Red, Green, Blue
;   CMYK          - Cyan, Magenta, Yellow, Black
;   GRAY          - Grayscale
;   sRGB          - Standard RGB
;   scRGB         - Scene-referred RGB
;   Lab           - CIELAB
;   XYZ           - CIE XYZ
;   HSL           - Hue, Saturation, Lightness
;   HSB           - Hue, Saturation, Brightness
;   YCbCr         - Luminance, Chrominance
;   CMY           - Cyan, Magenta, Yellow
;   HCL           - Hue, Chroma, Luminance
;   HCLp          - Hue, Chroma, Luminance (polar)
;   HSI           - Hue, Saturation, Intensity
;   HSV           - Hue, Saturation, Value
;   HWB           - Hue, Whiteness, Blackness
;   Jzazbz        - Jzazbz color space
;   LCHab         - Luminance, Chroma, Hue (CIELAB)
;   LCHuv         - Luminance, Chroma, Hue (CIELUV)
;   LMS           - Long, Medium, Short (cone response)
;   Log           - Logarithmic
;   Luv           - CIELUV
;   OHTA          - Ohta color space
;   OkLab         - OkLab color space
;   OkLCH         - OkLCH color space
;   Rec601YCbCr   - Rec. 601 YCbCr
;   Rec709YCbCr   - Rec. 709 YCbCr
;   xyY           - CIE xyY
;   YCC           - Luminance, Chrominance (YCC)
;   YDbDr         - YDbDr color space
;   YIQ           - YIQ color space
;   YPbPr         - YPbPr color space
;   YUV           - YUV color space
;   Undefined     - Undefined color space
Colorspace=$($cmbColorspace.SelectedItem)

[TesseractOCR]
; -----------------------------------------------------------------------------
; Tesseract OCR Settings
; -----------------------------------------------------------------------------

; OCR Language
; Install additional language packs to use other options
; Multiple languages can be specified with '+' (e.g., eng+fil+enm)
; Valid options depend on installed language packs
; Common options:
;   eng     - English
;   fil     - Filipino
;   enm     - English Old
Language=$($cmbLanguage.SelectedItem.Tag)

; OCR Engine Mode (OEM)
; Valid options:
;   0 - Legacy engine only
;   1 - Neural nets LSTM engine only
;   2 - Legacy + LSTM engines
;   3 - Default, based on what is available
OCREngineMode=$($cmbEngineMode.SelectedIndex)

; Page Segmentation Mode (PSM)
; Valid options:
;   0  - Orientation and script detection (OSD) only
;   1  - Automatic page segmentation with OSD
;   2  - Automatic page segmentation, but no OSD, or OCR
;   3  - Fully automatic page segmentation, but no OSD (Default)
;   4  - Assume a single column of text of variable sizes
;   5  - Assume a single uniform block of vertically aligned text
;   6  - Assume a single uniform block of text
;   7  - Treat the image as a single text line
;   8  - Treat the image as a single word
;   9  - Treat the image as a single word in a circle
;   10 - Treat the image as a single character
;   11 - Sparse text. Find as much text as possible in no particular order
;   12 - Sparse text with OSD
;   13 - Raw line. Treat the image as a single text line
PageSegmentationMode=$($cmbPSM.SelectedIndex)

; Output file format
; Valid options: pdf, txt, hocr, tsv
; Note: Script is designed for PDF output, changing this may cause errors and is only there for debugging.
OutputType=pdf

[GUI]
; -----------------------------------------------------------------------------
; GUI Settings
; -----------------------------------------------------------------------------

WindowWidth=$($window.Width)
WindowHeight=$($window.Height)
BackgroundColor=$($guiSettings.BackgroundColor)
TextColor=$($guiSettings.TextColor)
TitleFontSize=$($guiSettings.TitleFontSize)
WindowStartupLocation=$($guiSettings.WindowStartupLocation)
"@
    $content | Set-Content $iniPath
}

# Button click handlers
$btnCancel = $window.FindName("btnCancel")

$btnSave.Add_Click({
    Save-Settings
    [System.Windows.MessageBox]::Show("Settings saved successfully!", "Success")
})

# Get the ScrollViewer control
$txtScrollViewer = ($txtOutput.Parent -as [System.Windows.Controls.ScrollViewer])

$btnStart.Add_Click({
    Save-Settings
    $startProcess = Join-Path $PSScriptRoot "start_process.bat"
    
    # Clear previous output
    $txtOutput.Text = ""
    
    # Start process with properly quoted path
    $script:currentProcess = Start-Process cmd -ArgumentList @(
        "/c",
        "`"$startProcess`""  # Add quotes around the path
    ) -RedirectStandardOutput "$env:TEMP\output.txt" -NoNewWindow -PassThru
    
    # Save process ID for cleanup
    $script:currentProcess.Id | Set-Content $pidFile
    
    # Track last position
    $script:lastOffset = 0
    
    # Update button states
    $btnStart.IsEnabled = $false
    $btnCancel.IsEnabled = $true
    
    # Create a timer to update UI
    $script:outputTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:outputTimer.Interval = [TimeSpan]::FromMilliseconds(100)
    $script:outputTimer.Add_Tick({
        try {
            if (Test-Path "$env:TEMP\output.txt") {
                $fs = [System.IO.File]::Open("$env:TEMP\output.txt", 'Open', 'Read', 'ReadWrite')
                $reader = New-Object System.IO.StreamReader($fs)
                
                if ($script:lastOffset -gt 0) {
                    $reader.BaseStream.Position = $script:lastOffset
                }
                
                while (!$reader.EndOfStream) {
                    $line = $reader.ReadLine()
                    if ($line) {
                        $txtOutput.Dispatcher.Invoke([Action]{
                            $txtOutput.AppendText("$line`r`n")
                            $txtScrollViewer.ScrollToBottom()
                        })
                    }
                }
                
                $script:lastOffset = $reader.BaseStream.Position
                
                $reader.Close()
                $fs.Close()
            }
            
            # Check if process has ended
            if ($script:currentProcess.HasExited) {
                Start-Sleep -Milliseconds 100  # Give time for final output
                Stop-CurrentProcess
            }
            
        } catch {
            Write-Warning "Error reading output: $_"
        }
    })
    $script:outputTimer.Start()
})

$btnCancel.Add_Click({
    try {
        # Create a single-threaded dispatcher if needed
        if (-not [System.Windows.Threading.Dispatcher]::CurrentDispatcher) {
            $dispatcherThread = [System.Windows.Threading.Dispatcher]::CurrentDispatcher
            $dispatcherThread.ShutdownStarted.Add({
                [System.Windows.Threading.Dispatcher]::CurrentDispatcher.InvokeShutdown()
            })
        }
        
        # Create and show the wait window synchronously
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{
            $waitWindow = New-Object System.Windows.Window
            $waitWindow.Title = "Please Wait"
            $waitWindow.Width = 300
            $waitWindow.Height = 100
            $waitWindow.WindowStyle = "None"
            $waitWindow.ResizeMode = "NoResize"
            $waitWindow.WindowStartupLocation = "CenterScreen"
            $waitWindow.Background = "#001B1B"
            $waitWindow.Topmost = $true

            $grid = New-Object System.Windows.Controls.Grid
            $waitWindow.Content = $grid

            $text = New-Object System.Windows.Controls.TextBlock
            $text.Text = "Please wait while the script is cancelling process..."
            $text.Foreground = "#00FF00"
            $text.HorizontalAlignment = "Center"
            $text.VerticalAlignment = "Center"
            $text.FontSize = 14
            $grid.Children.Add($text)

            # Show window
            $waitWindow.Show()
            
            # Process cancellation
            Stop-CurrentProcess
            $txtOutput.AppendText("`r`nProcess cancelled by user.")
            $waitWindow.Close()
        })

    } catch {
        Write-Warning "Error in cancel operation: $_"
        # Fallback to direct cancellation
        Stop-CurrentProcess
        $txtOutput.AppendText("`r`nProcess cancelled by user.")
    }
})

# Modified slider value update handlers
$sldQuality.Add_ValueChanged({
    $txtQualityValue.Text = [Math]::Round($sldQuality.Value)
})

$sldDeskew.Add_ValueChanged({
    $txtDeskewValue.Text = "$([Math]::Round($sldDeskew.Value))%"
})

$btnOpenLogs.Add_Click({
    $logsPath = Join-Path $PSScriptRoot "logs"
    if (!(Test-Path $logsPath)) {
        New-Item -ItemType Directory -Path $logsPath -Force | Out-Null
    }
    Start-Process "explorer.exe" -ArgumentList $logsPath
})

# Show window
try {
    Write-Host "Showing window..."
    $window.Topmost = $true
    $window.Activate()
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(100)
    $timer.Add_Tick({
        $window.Topmost = $false
        $timer.Stop()
    })
    $timer.Start()
    Write-Host "Starting GUI..."
    $window.ShowDialog() | Out-Null
} 
catch {
    Write-Error "Error showing window: $_"
    Write-Error $_.ScriptStackTrace
    Start-Sleep -Seconds 2
    Stop-Process -Id $PID -Force
}
