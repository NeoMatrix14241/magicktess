Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MagickTessTK OCR - Automated OCR Processing Tool" Height="800" Width="900"
        Background="#001B1B" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="NAPS2 on Steroids" Foreground="#00FF00" FontSize="24" Margin="0,0,0,10"/>
        
        <StackPanel Grid.Row="1" Margin="0,10">
            <TextBlock Foreground="#00FF00" TextWrapping="Wrap" Margin="0,0,0,10">
                The script will only process files as batch and will treat folders as the document where its
                pages will be the tif files, read the ReadMe.txt file for proper documentation
            </TextBlock>
            
            <TextBlock Foreground="#00FF00" Margin="0,10">
                Repository: https://github.com/tesseract-ocr/tesseract
            </TextBlock>
        </StackPanel>

        <GroupBox Grid.Row="2" Header="Folder Settings" Foreground="#00FF00" Margin="0,10">
            <StackPanel>
                <DockPanel Margin="0,5">
                    <TextBlock Text="Input Folder: " Foreground="#00FF00" Width="100"/>
                    <Button Content="Browse" Width="70" DockPanel.Dock="Right" Name="btnInputBrowse"/>
                    <TextBox Name="txtInputFolder" Margin="5,0"/>
                </DockPanel>
                
                <DockPanel Margin="0,5">
                    <TextBlock Text="Output Folder: " Foreground="#00FF00" Width="100"/>
                    <Button Content="Browse" Width="70" DockPanel.Dock="Right" Name="btnOutputBrowse"/>
                    <TextBox Name="txtOutputFolder" Margin="5,0"/>
                </DockPanel>
                
                <DockPanel Margin="0,5">
                    <TextBlock Text="Archive Folder: " Foreground="#00FF00" Width="100"/>
                    <Button Content="Browse" Width="70" DockPanel.Dock="Right" Name="btnArchiveBrowse"/>
                    <TextBox Name="txtArchiveFolder" Margin="5,0"/>
                </DockPanel>
            </StackPanel>
        </GroupBox>

        <TextBlock Grid.Row="3" Foreground="#00FF00" TextWrapping="Wrap" Margin="0,10">
            Folder List Generated:
            &#x0a;• input - [BATCH OCR ONLY] Where your folders with tif files that will be processed for OCR
            &#x0a;• archive - Where your folders in input folder will be moved after OCR
            &#x0a;• output - Where your processed OCR files in pdf format
            &#x0a;• logs - Where the logs are stored for the entire process
            &#x0a;
            &#x0a;Note:
            &#x0a;• press "ctrl + c" in powershell to cancel
        </TextBlock>

        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10">
            <Button Name="btnSave" Content="Save Settings" Width="100" Margin="5,0"/>
            <Button Name="btnStart" Content="Start OCR Process" Width="120"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$btnInputBrowse = $window.FindName("btnInputBrowse")
$btnOutputBrowse = $window.FindName("btnOutputBrowse")
$btnArchiveBrowse = $window.FindName("btnArchiveBrowse")
$txtInputFolder = $window.FindName("txtInputFolder")
$txtOutputFolder = $window.FindName("txtOutputFolder")
$txtArchiveFolder = $window.FindName("txtArchiveFolder")
$btnSave = $window.FindName("btnSave")
$btnStart = $window.FindName("btnStart")

# Load current settings
$iniPath = Join-Path $PSScriptRoot "settings.ini"
if (Test-Path $iniPath) {
    $content = Get-Content $iniPath
    foreach ($line in $content) {
        if ($line -match "InputFolder=(.*)") { $txtInputFolder.Text = $matches[1] }
        if ($line -match "OutputFolder=(.*)") { $txtOutputFolder.Text = $matches[1] }
        if ($line -match "ArchiveFolder=(.*)") { $txtArchiveFolder.Text = $matches[1] }
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

# Save settings
$btnSave.Add_Click({
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
"@
    $content | Set-Content $iniPath
    [System.Windows.MessageBox]::Show("Settings saved successfully!", "Success")
})

# Start OCR Process
$btnStart.Add_Click({
    $startProcess = Join-Path $PSScriptRoot "start_process.bat"
    Start-Process -FilePath $startProcess
})

# Show window
$window.ShowDialog() | Out-Null