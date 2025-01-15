Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

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
        WindowStartupLocation="$($guiSettings.WindowStartupLocation)">
    <Grid Margin="10">
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
            <Button Name="btnSave" Content="Save Settings" Width="100" Margin="5,0"/>
            <Button Name="btnStart" Content="Start OCR Process" Width="120"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Add error handling for XAML loading
try {
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    if ($null -eq $window) {
        throw "Failed to load XAML"
    }
} catch {
    Write-Error "Failed to load XAML: $_"
    exit 1
}

# Get controls
$btnInputBrowse = $window.FindName("btnInputBrowse")
$btnOutputBrowse = $window.FindName("btnOutputBrowse")
$btnArchiveBrowse = $window.FindName("btnArchiveBrowse")
$txtInputFolder = $window.FindName("txtInputFolder")
$txtOutputFolder = $window.FindName("txtOutputFolder")
$txtArchiveFolder = $window.FindName("txtArchiveFolder")
$btnSave = $window.FindName("btnSave")
$btnStart = $window.FindName("btnStart")

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

# Populate colorspace options
$colorspaces = @("Auto", "RGB", "CMYK", "GRAY", "sRGB", "scRGB", "Lab", "XYZ", "HSL", "HSB", "YCbCr", "CMY")
$cmbColorspace.ItemsSource = $colorspaces

# Populate PSM options
$psmOptions = 0..13 | ForEach-Object { "PSM $_" }
$cmbPSM.ItemsSource = $psmOptions

# Create language mapping for display names
$languageMap = @{
    "eng" = "English"
    "enm" = "English (Old)"
    "fil" = "Filipino"
}

# Define combinations with display names
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

CompressionType=$($cmbCompression.SelectedItem)
Quality=$($sldQuality.Value)
AdditionalParameters=ON
DeskewThreshold=$($sldDeskew.Value)%
Colorspace=$($cmbColorspace.SelectedItem)

[TesseractOCR]
; -----------------------------------------------------------------------------
; Tesseract OCR Settings
; -----------------------------------------------------------------------------

Language=$($cmbLanguage.SelectedItem.Tag)
OCREngineMode=$($cmbEngineMode.SelectedIndex)
PageSegmentationMode=$($cmbPSM.SelectedIndex)
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
$btnSave.Add_Click({
    Save-Settings
    [System.Windows.MessageBox]::Show("Settings saved successfully!", "Success")
})

$btnStart.Add_Click({
    Save-Settings
    $startProcess = Join-Path $PSScriptRoot "start_process.bat"
    Start-Process -FilePath $startProcess
})

# Modified slider value update handlers
$sldQuality.Add_ValueChanged({
    $txtQualityValue.Text = [Math]::Round($sldQuality.Value)
})

$sldDeskew.Add_ValueChanged({
    $txtDeskewValue.Text = "$([Math]::Round($sldDeskew.Value))%"
})

# Modified ComboBox population
$cmbCompression.Items.Clear()
@("LZW", "ZIP", "RLE", "NONE", "JPEG", "WEBP") | ForEach-Object {
    $cmbCompression.Items.Add($_)
}

$cmbColorspace.Items.Clear()
$colorspaces | ForEach-Object {
    $cmbColorspace.Items.Add($_)
}

$cmbPSM.Items.Clear()
0..13 | ForEach-Object {
    $cmbPSM.Items.Add("PSM $_")
}

# Show window
$window.ShowDialog() | Out-Null
