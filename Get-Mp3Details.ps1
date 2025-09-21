Clear-Host
write-host "Warming up... Please Wait"
#--------- Set Parameters ----------#

## Set Tempfolder for last folder use test
$TempFolder = "$env:TEMP\MP3Analyze"
## If first run, no folder is set start in
$InitialFolder = "C:\Temp\Sidify\Download-Temp\"
## Log Powershell output to file in same directory
$Global:LogginEnabled = $true ## $true = yes | $false = no
$Prefix = "(SP-RIP-N)"
$debug = $true
#--------- DO Not Edit Below ----------#

# Load stopwatch
$StopWatch = New-Object System.Diagnostics.Stopwatch

#mediaplayer
Add-Type -AssemblyName presentationCore
$mediaPlayer = New-Object system.windows.media.mediaplayer
$Global:PlayerVolume = 0.3 # Default volume
# Find Tempfolder and create if not exist
if(!(Test-Path $TempFolder)){
    New-Item -ItemType Directory $TempFolder
}

# Check if ID3 powershell gallery module is installed. If not, prompt user to install manually
$ID3Module = Get-Module -Name ID3 -ListAvailable
if(!($ID3Module)){
    Write-Warning "The required PowerShell module 'ID3' is not installed."
    Write-Host ""
    Write-Host "To install it, open an elevated (Run as Administrator) PowerShell prompt and run:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "    Install-Module -Name ID3 -Scope AllUsers" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "If you have never used the PowerShell Gallery before, you may need to run:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "After installing, re-run this script." -ForegroundColor Yellow
    Read-Host "Press enter to exit"
    break
} else {
    Import-Module -Name ID3
}

# --- FFmpeg and ffmpeg-normalize checks and user instructions ---
$path = $env:Path -split ";"

# Check for ffmpeg in Source/ffmpeg/bin/ffmpeg.exe (relative to script) or in PATH
$ffmpegPath = $null

# Check relative to script location
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ffmpegLocal = Join-Path $scriptDir "Source\ffmpeg\bin\ffmpeg.exe"
if (Test-Path $ffmpegLocal) {
    $ffmpegPath = $ffmpegLocal
}

# If not found locally, check PATH
if (-not $ffmpegPath) {
    foreach ($p in $path) {
        if (![string]::IsNullOrWhiteSpace($p)) {
            $candidate = Join-Path $p "ffmpeg.exe"
            if (Test-Path $candidate) {
                $ffmpegPath = $candidate
                break
            }
        }
    }
}

if (-not $ffmpegPath) {
    Write-Warning "FFmpeg was not found."
    Write-Host ""
    Write-Host "FFmpeg should be present at:" -ForegroundColor Yellow
    Write-Host "    $ffmpegLocal" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Or in your system PATH. To add it to your PATH environment variable:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "1. Open System Properties (Win+Pause > Advanced system settings > Environment Variables)" -ForegroundColor Yellow
    Write-Host "2. Under 'System variables', select 'Path', then click 'Edit'" -ForegroundColor Yellow
    Write-Host "3. Add the folder containing ffmpeg.exe (e.g. $($scriptDir)\Source\ffmpeg\bin)" -ForegroundColor Cyan
    Write-Host "4. Click OK and restart computer or logoff." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Download FFmpeg: https://ffmpeg.org/download.html" -ForegroundColor Cyan
    Read-Host "Press enter to exit"
    break
}

# Check for ffmpeg-normalize (Python tool)
$ffmpeg_normalize_found = $false
try {
    $ffmpegNormalizeVersion = & python -m ffmpeg_normalize -h 2>$null
    if ($LASTEXITCODE -eq 0 -or $ffmpegNormalizeVersion) {
        $ffmpeg_normalize_found = $true
    }
} catch {
    $ffmpeg_normalize_found = $false
}
if (-not $ffmpeg_normalize_found) {
    Write-Warning "ffmpeg-normalize (Python tool) not found."
    Write-Host ""
    Write-Host "To install it, open a command prompt and run:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "    pip install ffmpeg-normalize" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "You need Python and pip installed. See: https://github.com/slhck/ffmpeg-normalize" -ForegroundColor Yellow
    Read-Host "Press enter to exit"
    break
}

#--------- Begin of Functions ----------#

# Show folder browser dialog and return selected path (modern style)
Function Get-Folder($initialDirectory) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Folder"
    $form.Size = New-Object System.Drawing.Size(500,180)
    $form.StartPosition = 'CenterScreen'
    $form.Topmost = $true
    $form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None

    # Custom title bar
    $titleBar = New-Object System.Windows.Forms.Panel
    $titleBar.Size = New-Object System.Drawing.Size(500, 30)
    $titleBar.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $titleBar.Dock = [System.Windows.Forms.DockStyle]::Top

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Select Folder"
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $titleLabel.AutoSize = $true
    $titleLabel.Location = New-Object System.Drawing.Point(10, 5)
    $titleBar.Controls.Add($titleLabel)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "X"
    $closeButton.ForeColor = [System.Drawing.Color]::White
    $closeButton.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $closeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $closeButton.FlatAppearance.BorderSize = 0
    $closeButton.Size = New-Object System.Drawing.Size(30, 30)
    $closeButton.Location = New-Object System.Drawing.Point(460, 0)
    $closeButton.Add_Click({ $form.Close() })
    $titleBar.Controls.Add($closeButton)
    $form.Controls.Add($titleBar)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,40)
    $label.Size = New-Object System.Drawing.Size(480,40)
    $label.Text = "Please select the folder containing your MP3/MP4 files:"
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $label.ForeColor = [System.Drawing.Color]::FromArgb(255, 228, 181)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,80)
    $textBox.Size = New-Object System.Drawing.Size(350,25)
    $textBox.Text = $initialDirectory
    $textBox.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $textBox.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($textBox)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Location = New-Object System.Drawing.Point(370,78)
    $browseButton.Size = New-Object System.Drawing.Size(100,28)
    $browseButton.Text = "Browse..."
    $browseButton.BackColor = [System.Drawing.Color]::FromArgb(173, 216, 230)
    $browseButton.ForeColor = [System.Drawing.Color]::Black
    $browseButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $browseButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $browseButton.Add_Click({
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowser.SelectedPath = $textBox.Text
        if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $textBox.Text = $FolderBrowser.SelectedPath
        }
    })
    $form.Controls.Add($browseButton)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(120,120)
    $okButton.Size = New-Object System.Drawing.Size(100,32)
    $okButton.Text = "OK"
    $okButton.BackColor = [System.Drawing.Color]::FromArgb(144, 238, 144)
    $okButton.ForeColor = [System.Drawing.Color]::Black
    $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $okButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(260,120)
    $cancelButton.Size = New-Object System.Drawing.Size(100,32)
    $cancelButton.Text = "Cancel"
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(255, 182, 193)
    $cancelButton.ForeColor = [System.Drawing.Color]::Black
    $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and (Test-Path $textBox.Text)) {
        return $textBox.Text
    } else {
        return $null
    }
}

# Show a Yes/No/Stop dialog for user confirmation
function Ask-User($Title,$Message){
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(420,260)
    $form.StartPosition = 'CenterScreen'
    $form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None

    # Custom title bar
    $titleBar = New-Object System.Windows.Forms.Panel
    $titleBar.Size = New-Object System.Drawing.Size(420, 30)
    $titleBar.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $titleBar.Dock = [System.Windows.Forms.DockStyle]::Top

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $Title
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $titleLabel.AutoSize = $true
    $titleLabel.Location = New-Object System.Drawing.Point(10, 5)
    $titleBar.Controls.Add($titleLabel)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "X"
    $closeButton.ForeColor = [System.Drawing.Color]::White
    $closeButton.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $closeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $closeButton.FlatAppearance.BorderSize = 0
    $closeButton.Size = New-Object System.Drawing.Size(30, 30)
    $closeButton.Location = New-Object System.Drawing.Point(380, 0)
    $closeButton.Add_Click({ $form.Close() })
    $titleBar.Controls.Add($closeButton)
    $form.Controls.Add($titleBar)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20,50)
    $label.Size = New-Object System.Drawing.Size(380,100)
    $label.Text = $Message
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $label.ForeColor = [System.Drawing.Color]::FromArgb(255, 228, 181)
    $label.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $label.AutoSize = $false
    $form.Controls.Add($label)

    $GoodButton = New-Object System.Windows.Forms.Button
    $GoodButton.Location = New-Object System.Drawing.Point(40,170)
    $GoodButton.Size = New-Object System.Drawing.Size(100,32)
    $GoodButton.Text = 'Yes'
    $GoodButton.BackColor = [System.Drawing.Color]::FromArgb(144, 238, 144)
    $GoodButton.ForeColor = [System.Drawing.Color]::Black
    $GoodButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $GoodButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $GoodButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $form.AcceptButton = $GoodButton
    $form.Controls.Add($GoodButton)

    $BadButton = New-Object System.Windows.Forms.Button
    $BadButton.Location = New-Object System.Drawing.Point(160,170)
    $BadButton.Size = New-Object System.Drawing.Size(100,32)
    $BadButton.Text = 'No'
    $BadButton.BackColor = [System.Drawing.Color]::FromArgb(255, 182, 193)
    $BadButton.ForeColor = [System.Drawing.Color]::Black
    $BadButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $BadButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $BadButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $form.Controls.Add($BadButton)

    $ReCheckButton = New-Object System.Windows.Forms.Button
    $ReCheckButton.Location = New-Object System.Drawing.Point(280,170)
    $ReCheckButton.Size = New-Object System.Drawing.Size(100,32)
    $ReCheckButton.Text = 'Stop'
    $ReCheckButton.BackColor = [System.Drawing.Color]::FromArgb(221, 160, 221)
    $ReCheckButton.ForeColor = [System.Drawing.Color]::Black
    $ReCheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $ReCheckButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $ReCheckButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $ReCheckButton
    $form.Controls.Add($ReCheckButton)

    $form.Topmost = $true

    $Prop = New-Object System.Windows.Forms.Form -Property @{TopMost = $true }
    $result = $form.ShowDialog($prop)
    switch ($result) {
        'Yes'    { return "yes" }
        'No'     { return "no" }
        'Cancel' { return "cancel" }
        default  { return $null }
    }
}
function Get-UniqueFolderName {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )
    $base = $FolderPath
    $counter = 1
    while (Test-Path $FolderPath) {
        $FolderPath = "$base ($counter)"
        $counter++
    }
    return $FolderPath
}
# Main MP3 review dialog with all options
Function Get-Response($Name, $total, $processed) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'MP3 Player'
    $form.Size = New-Object System.Drawing.Size(780, 400)
    $form.StartPosition = 'CenterScreen'
    $form.KeyPreview = $true
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)

    # Custom title bar
    $titleBar = New-Object System.Windows.Forms.Panel
    $titleBar.Size = New-Object System.Drawing.Size(770, 30)
    $titleBar.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $titleBar.Dock = [System.Windows.Forms.DockStyle]::Top

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "MP3 Player"
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $titleLabel.AutoSize = $true
    $titleLabel.Location = New-Object System.Drawing.Point(10, 5)
    $titleBar.Controls.Add($titleLabel)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "X"
    $closeButton.ForeColor = [System.Drawing.Color]::White
    $closeButton.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $closeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $closeButton.FlatAppearance.BorderSize = 0
    $closeButton.Size = New-Object System.Drawing.Size(30, 30)
    $closeButton.Location = New-Object System.Drawing.Point(720, 0)
    $closeButton.Add_Click({ $form.Close() })
    $titleBar.Controls.Add($closeButton)
    $form.Controls.Add($titleBar)

    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Location = New-Object System.Drawing.Point(10, 40)
    $nameLabel.Size = New-Object System.Drawing.Size(630, 50)
    $nameLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 228, 181)
    $nameLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $nameLabel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $nameLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter 
    $nameLabel.Text = "$($Name -replace '.mp3','')"
    $form.Controls.Add($nameLabel)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 80)
    $label.Size = New-Object System.Drawing.Size(630, 160)
    $label.ForeColor = [System.Drawing.Color]::White
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $label.Text = "
    Choose Good, leave file in place.
    Choose Top, move item to folder Top.
    Choose Bad, move item to folder Bad.

    When Choose Re-Check, play Mp3 Again.

    Processed: $processed / $total
    "
    $form.Controls.Add($label)

    $timeBar = New-Object System.Windows.Forms.TrackBar
    $timeBar.Location = New-Object System.Drawing.Point(10, 240)
    $timeBar.Size = New-Object System.Drawing.Size(730, 45)
    $timeBar.Minimum = 0
    $timeBar.Maximum = 100
    $timeBar.TickFrequency = 1
    $timeBar.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $form.Controls.Add($timeBar)

    # --- Add volume bar ---
    $volumeLabel = New-Object System.Windows.Forms.Label
    $volumeLabel.Text = "Volume"
    $volumeLabel.Location = New-Object System.Drawing.Point(10, 285)
    $volumeLabel.Size = New-Object System.Drawing.Size(60, 20)
    $volumeLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($volumeLabel)

    $volumeBar = New-Object System.Windows.Forms.TrackBar
    $volumeBar.Location = New-Object System.Drawing.Point(70, 280)
    $volumeBar.Size = New-Object System.Drawing.Size(200, 45)
    $volumeBar.Minimum = 0
    $volumeBar.Maximum = 100
    $volumeBar.TickFrequency = 10
    $volumeBar.Value = [math]::Round(($Global:PlayerVolume) * 100)
    $volumeBar.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $form.Controls.Add($volumeBar)

    $mediaPlayer.Volume = $Global:PlayerVolume

    $volumeBar.Add_Scroll({
        $mediaPlayer.Volume = $volumeBar.Value / 100
        $Global:PlayerVolume = $mediaPlayer.Volume
    })

    # --- Move the response buttons down to avoid overlap ---
    $buttonY = 330
    $GoodButton = New-Object System.Windows.Forms.Button
    $GoodButton.Location = New-Object System.Drawing.Point(80, $buttonY)
    $GoodButton.Size = New-Object System.Drawing.Size(140, 50)
    $GoodButton.Text = 'Good (1)'
    $GoodButton.BackColor = [System.Drawing.Color]::FromArgb(173, 216, 230)
    $GoodButton.ForeColor = [System.Drawing.Color]::Black
    $GoodButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $GoodButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $GoodButton.FlatAppearance.BorderSize = 0
    $GoodButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $GoodButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $form.Controls.Add($GoodButton)

    $TopButton = New-Object System.Windows.Forms.Button
    $TopButton.Location = New-Object System.Drawing.Point(240, $buttonY)
    $TopButton.Size = New-Object System.Drawing.Size(140, 50)
    $TopButton.Text = 'Top (2)'
    $TopButton.BackColor = [System.Drawing.Color]::FromArgb(144, 238, 144)
    $TopButton.ForeColor = [System.Drawing.Color]::Black
    $TopButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $TopButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $TopButton.FlatAppearance.BorderSize = 0
    $TopButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $TopButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($TopButton)

    $BadButton = New-Object System.Windows.Forms.Button
    $BadButton.Location = New-Object System.Drawing.Point(400, $buttonY)
    $BadButton.Size = New-Object System.Drawing.Size(140, 50)
    $BadButton.Text = 'Bad (3)'
    $BadButton.BackColor = [System.Drawing.Color]::FromArgb(255, 182, 193)
    $BadButton.ForeColor = [System.Drawing.Color]::Black
    $BadButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $BadButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $BadButton.FlatAppearance.BorderSize = 0
    $BadButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $BadButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $form.Controls.Add($BadButton)

    $ReCheckButton = New-Object System.Windows.Forms.Button
    $ReCheckButton.Location = New-Object System.Drawing.Point(560, $buttonY)
    $ReCheckButton.Size = New-Object System.Drawing.Size(140, 50)
    $ReCheckButton.Text = 'Re-Check (4)'
    $ReCheckButton.BackColor = [System.Drawing.Color]::FromArgb(221, 160, 221)
    $ReCheckButton.ForeColor = [System.Drawing.Color]::Black
    $ReCheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $ReCheckButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $ReCheckButton.FlatAppearance.BorderSize = 0
    $ReCheckButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $ReCheckButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    $form.Controls.Add($ReCheckButton)

    $form.Topmost = $true

    # Timer for updating trackbar position
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
        if ($mediaPlayer.NaturalDuration.HasTimeSpan) {
            $timeBar.Value = [math]::Round(($mediaPlayer.Position.TotalSeconds / $mediaPlayer.NaturalDuration.TimeSpan.TotalSeconds) * 100)
        }
    })
    $timer.Start()

    $timeBar.Add_Scroll({
        if ($mediaPlayer.NaturalDuration.HasTimeSpan) {
            $mediaPlayer.Position = [TimeSpan]::FromSeconds(($timeBar.Value / 100) * $mediaPlayer.NaturalDuration.TimeSpan.TotalSeconds)
        }
    })
    $form.Add_Shown({$form.Activate(); $timeBar.focus()})
    $form.Add_KeyDown({
        switch ($_.KeyCode) {
            'D1' { $GoodButton.PerformClick() }
            'D2' { $TopButton.PerformClick() }
            'D3' { $BadButton.PerformClick() }
            'D4' { $ReCheckButton.PerformClick() }
        }
    })

    $Prop = New-Object System.Windows.Forms.Form -Property @{TopMost = $true}
    $result = $form.ShowDialog($Prop)
    $timer.Stop()
    $mediaPlayer.Stop()
    $mediaPlayer.Close()
    return $result
}

# Play MP3 and show review dialog
Function Start-Mp3($data, $total, $processed) {
    try {
        $mediaPlayer.open($data.Fullname)
        $mediaPlayer.Play()
        $response = Get-Response -Name $data.Name -total $total -processed $processed
        return $response
    } catch {
        Write-Error "Unable to play audio"
    }
}

# Move file to appropriate folder based on user response
Function Check-File($File, $total, $processed) {
    $response = Start-Mp3 -data $File -total $total -processed $processed
    $destination = ""
    $subfolder = ""

    switch ($response) {
        "Yes" { return "Good" } # Leave file in place
        "OK" { $subfolder = "Top" }
        "No" { $subfolder = "Bad" }
        "Retry" { return Check-File -file $File -total $total -processed $processed }
        default {
            Write-Host "You Hit Cancel"
            Read-Host "Press enter to exit"
            break
        }
    }

    if ($subfolder) {
        try {
            $destination = "$MessureFolder\$subfolder\"
            New-Item -ItemType Directory -Path $destination -Force -ea SilentlyContinue | Out-Null
            Move-Item -Path $File.Fullname -Destination $destination -ErrorAction Stop
            return $subfolder
        } catch {
            if ($_.Exception.Message -like "*already exists*") {
                $errorBase = "$MessureFolder\Error\"
                $duplicateBase = "$errorBase\Duplicate\"
                $errorDestination = "$duplicateBase\$subfolder\"

                New-Item -ItemType Directory -Path $errorBase -Force -ea SilentlyContinue | Out-Null
                New-Item -ItemType Directory -Path $duplicateBase -Force -ea SilentlyContinue | Out-Null
                New-Item -ItemType Directory -Path $errorDestination -Force -ea SilentlyContinue | Out-Null

                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
                $extension = [System.IO.Path]::GetExtension($File.Name)
                $counter = 1
                $newFileName = "$baseName ($counter)$extension"
                $newFilePath = Join-Path $errorDestination $newFileName

                while (Test-Path $newFilePath) {
                    $counter++
                    $newFileName = "$baseName ($counter)$extension"
                    $newFilePath = Join-Path $errorDestination $newFileName
                }

                Move-Item -Path $File.Fullname -Destination $newFilePath -Force
                return "Error: Duplicate. Moved To duplicate error folder"
            } else {
                $errorDestination = "$MessureFolder\Error\"
                New-Item -ItemType Directory -Path $errorDestination -Force -ea SilentlyContinue | Out-Null
                Move-Item -Path $File.Fullname -Destination $errorDestination -Force
                return "Error: $($_.Exception.Message)"
            }
        }
    }
}

Function Get-MP3MetaData{
    [CmdletBinding()
    ]
    [Alias()]
    [OutputType([Psobject])]
    Param
    (
        [String] [Parameter(Mandatory=$true, ValueFromPipeline=$true)] $Directory
    )

    Begin
    {
        $shell = New-Object -ComObject "Shell.Application"
    }
    Process
    {

        Foreach($Dir in $Directory)
        {
            $ObjDir = $shell.NameSpace($Dir)
            $Files = Get-ChildItem $Dir| ?{$_.Extension -in '.mp3','.mp4'}

            Foreach($File in $Files)
            {
                $ObjFile = $ObjDir.parsename($File.Name)
                $MetaData = @{}
                $MP3 = ($ObjDir.Items()|?{$_.path -like "*.mp3" -or $_.path -like "*.mp4"})
                $PropertArray = 0,1,2,12,13,14,15,16,17,18,19,20,21,22,27,28,36,220,223
            
                Foreach($item in $PropertArray)
                { 
                    If($ObjDir.GetDetailsOf($ObjFile, $item)) #To avoid empty values
                    {
                        $MetaData[$($ObjDir.GetDetailsOf($MP3,$item))] = $ObjDir.GetDetailsOf($ObjFile, $item)
                    }
                 
                }

                New-Object psobject -Property $MetaData |select *, @{n="Directory";e={$Dir}}, @{n="Fullname";e={Join-Path "$Dir" "$($File.Name)" -Resolve}}, @{n="Extension";e={$File.Extension}}
            }
        }
    }
    End
    {
    }
}

function Start-Normalize($folder){

    $items = Get-ChildItem -Path "$folder" -File -filter "*.mp3"
    $totalitems = $items.count
    $itemstodo = $totalitems

    $waitmessage =  "Normalizing File...Approx wait Time: " +(0..$totalitems| % -Begin {$Total = 0} -Process {$Total += (New-TimeSpan -second 2)} -End {$Total})

    $items|%{
        $filename = $_
        if(-not $debug){
            Clear-Host
        } 
        write-host $waitmessage
        Write-Host "$itemstodo / $totalitems  -  $filename";$itemstodo = ($itemstodo - 1)
        # Use CBR 320k, standard sample rate, do not use -map_metadata (not supported by ffmpeg-normalize)
        ffmpeg-normalize $filename.FullName -of $($folder + "\Normalize") --normalization-type peak --target-level 0 -c:a libmp3lame -b:a 320k -ar 44100 -vn -ext mp3

        # --- Fix: Restore album art if present ---
        $normalizedFile = Join-Path ($folder + "\Normalize") $filename.Name
        $coverTemp = [System.IO.Path]::GetTempFileName() + ".jpg"
        # Extract album art (if any)
        ffmpeg -y -i "$($filename.FullName)" -an -vcodec copy "$coverTemp" 2>$null
        if ((Test-Path $coverTemp) -and ((Get-Item $coverTemp).Length -gt 0)) {
            # Embed album art into normalized file
            $outTemp = [System.IO.Path]::GetTempFileName() + ".mp3"
            ffmpeg -y -i "$normalizedFile" -i "$coverTemp" -map 0:a -map 1 -c copy -id3v2_version 3 "$outTemp" 2>$null
            Move-Item -Force "$outTemp" "$normalizedFile"
            Remove-Item "$coverTemp" -Force
        } else {
            if (Test-Path $coverTemp) { Remove-Item "$coverTemp" -Force }
        }
        # --- end album art fix ---
    }

    return $($folder + "\Normalize")
}

function Remove-Silence($folder){
    New-Item -Path "$($folder + "\Silence\") " -Force -ea SilentlyContinue |Out-Null
    $items = Get-ChildItem -Path "$folder" -File -filter "*.mp3"

    $totalitems = $items.count
    $itemstodo = $totalitems
    $waitmessage = "Removing Silence in MP3..."

    $items|%{
        if(-not $debug){
            Clear-Host
        } 
        $filename = $_ 
        write-host $waitmessage
        Write-Host "$itemstodo / $totalitems  -  $filename"
        $itemstodo = ($itemstodo - 1)
        # Improved silence removal: only at end, less aggressive
        $srcFile = $folder + "\" + $filename.name
        $dstFile = $folder + "\Silence\" + $filename.name
        ffmpeg -i $srcFile -hide_banner -loglevel error -af "silenceremove=start_periods=0:stop_periods=1:stop_duration=2:stop_threshold=-50dB:detection=peak" $dstFile

        # --- Restore album art from source file if present ---
        $coverTemp = [System.IO.Path]::GetTempFileName() + ".jpg"
        ffmpeg -y -i "$srcFile" -an -vcodec copy "$coverTemp" 2>$null
        if ((Test-Path $coverTemp) -and ((Get-Item $coverTemp).Length -gt 0)) {
            $outTemp = [System.IO.Path]::GetTempFileName() + ".mp3"
            ffmpeg -y -i "$dstFile" -i "$coverTemp" -map 0:a -map 1 -c copy -id3v2_version 3 "$outTemp" 2>$null
            Move-Item -Force "$outTemp" "$dstFile"
            Remove-Item "$coverTemp" -Force
        } else {
            if (Test-Path $coverTemp) { Remove-Item "$coverTemp" -Force }
        }
        # --- end album art fix ---
        #ffmpeg -i $($folder + "\"+$filename.name) -y -c:a libmp3lame -b:a 256k -af silenceremove=1:0:-50dB -loglevel warning $($folder + "\Silence\"+$filename.name) 
    }
    
    return $($folder + "\Silence\")
}

function Fix-Id3andFileName ($folder,$Prefix){
    $items = Get-ChildItem -Path "$folder" -File -filter "*.mp3"
    $totalitems = $items.count
    $waitmessage = "Renaming Files and updating ID3Tag... Please Wait"
    $itemstodo = $totalitems
    $items|%{
        if(-not $debug){
            Clear-Host
        } 
        $filename = $_ ;write-host $waitmessage;Write-Host "$itemstodo / $totalitems  -  $filename";$itemstodo = ($itemstodo - 1);`
        Write-Host $Prefix
        $file  = $filename
        $CurrentTag = Get-Id3Tag $file.FullName
        $Artist = $CurrentTag.Artists

        # Strip Title and replace Characters
        $Title = $($CurrentTag.Title).replace("7`"","7 inch").replace("12`"","12 inch").replace("?","").replace("`/"," ").replace("`"","").split('[')[0].split(']')[0]
        $NewName = "$Artist - $Title $Prefix"
        $ext = $file.Extension

        $newfilename = Rename-Item -Path $file.FullName  -NewName $($NewName + $ext ) -PassThru

        Start-Sleep 1
        $tag = @{}
        $tag.Add('Title',($Title + " $Prefix"))
        set-Id3Tag -Path "$($newfilename.FullName)" -Tags $tag 

    }
}

function Retrive-Tag($File,$Artist, $Title){
    $tag = Get-Id3Tag $File
    $tag
    #[xml]$result = Invoke-WebRequest -Uri "https://musicbrainz.org/ws/2/release/?query=`"Going%20to%20Ibiza`"&artist=venga"

    #$result.metadata.'release-list'.release.date
}


# Replace the Set-Ending function with a summary dialog and statistics
function Set-Ending() {
    # Initialize counters
    $BadCount = $BadCount
    $topCount = 0
    $goodCount = 0
    $otherCount = 0
    $Errorcount = 0
    $duplicateCount = 0
    $SkipCount = 0

    # Track time for the whole process
    $scriptStart = $StopWatch.StartTime
    $scriptEnd = Get-Date
    $elapsed = $StopWatch.Elapsed

    # Calculate total song length in seconds using ID3 tag duration (try to find the correct property)
    $TotalSongSeconds = 0
    foreach ($item in $ID3TagData) {
        $durationStr = $null
        foreach ($prop in @('Length','Duur','Dauer','Duração','Durée')) {
            if ($item.PSObject.Properties.Name -contains $prop) {
                $durationStr = $item.$prop
                if ($durationStr) { break }
            }
        }
        if ($durationStr -and $durationStr -match '(\d+):(\d+)(?::(\d+))?') {
            $min = [int]$matches[1]
            $sec = [int]$matches[2]
            $hrs = if ($matches[3]) { [int]$matches[1]; $min = [int]$matches[2]; $sec = [int]$matches[3] } else { 0 }
            $TotalSongSeconds += ($hrs * 3600 + $min * 60 + $sec)
        }
    }

    # Format elapsed time
    $elapsedStr = "{0:00}:{1:00}:{2:00}" -f $elapsed.Hours, $elapsed.Minutes, $elapsed.Seconds

    # Format total song length
    $ts = [TimeSpan]::FromSeconds($TotalSongSeconds)
    $totalSongStr = "{0:00}:{1:00}:{2:00}" -f $ts.Hours, $ts.Minutes, $ts.Seconds

    # Calculate time saved
    $timeSaved = $ts - $elapsed
    if ($timeSaved.TotalSeconds -lt 0) { $timeSaved = [TimeSpan]::Zero }
    $timeSavedStr = "{0:00}:{1:00}:{2:00}" -f $timeSaved.Hours, $timeSaved.Minutes, $timeSaved.Seconds

    # Prepare summary string
    $summary = @"
    Total:     $(($ID3TagData|Measure-Object).count)
    Top:       $topCount
    Good:      $goodCount
    Other:     $otherCount
    Bad:       $BadCount
    Skipped:   $SkipCount
    Duplicate: $duplicateCount
    Errors:    $Errorcount

    Total spent : $elapsedStr (hh:mm:ss)
    Total length: $totalSongStr (hh:mm:ss)
    Time saved  : $timeSavedStr (hh:mm:ss)
"@

    Write-output $summary

    # Show final options dialog: Open Folder, Open Log, Exit, with summary and timing
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $finalForm = New-Object System.Windows.Forms.Form
    $finalForm.Text = "MP3Analyze - Finished"
    $finalForm.Size = New-Object System.Drawing.Size(540, 480)
    $finalForm.StartPosition = 'CenterScreen'
    $finalForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $finalForm.MaximizeBox = $false
    $finalForm.MinimizeBox = $false
    $finalForm.Topmost = $true
    $finalForm.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "All files checked. What do you want to do next?"
    $label.Size = New-Object System.Drawing.Size(520, 30)
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $label.ForeColor = [System.Drawing.Color]::FromArgb(255, 228, 181)
    $finalForm.Controls.Add($label)

    $summaryBox = New-Object System.Windows.Forms.TextBox
    $summaryBox.Multiline = $true
    $summaryBox.ReadOnly = $true
    $summaryBox.ScrollBars = "Vertical"
    $summaryBox.Size = New-Object System.Drawing.Size(520, 260)
    $summaryBox.Location = New-Object System.Drawing.Point(10, 45)
    $summaryBox.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Regular)
    $summaryBox.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $summaryBox.ForeColor = [System.Drawing.Color]::White
    $summaryBox.Text = $summary
    $finalForm.Controls.Add($summaryBox)

    $btnOpenFolder = New-Object System.Windows.Forms.Button
    $btnOpenFolder.Text = "Open Folder"
    $btnOpenFolder.Size = New-Object System.Drawing.Size(140,40)
    $btnOpenFolder.Location = New-Object System.Drawing.Point(30,340)
    $btnOpenFolder.BackColor = [System.Drawing.Color]::FromArgb(173, 216, 230)
    $btnOpenFolder.ForeColor = [System.Drawing.Color]::Black
    $btnOpenFolder.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOpenFolder.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnOpenFolder.Add_Click({
        Start-Process explorer.exe "`"$MessureFolder`""
        $finalForm.Close()
    })
    $finalForm.Controls.Add($btnOpenFolder)

    if ($Global:LogginEnabled -and $LogFile) {
        $btnOpenLog = New-Object System.Windows.Forms.Button
        $btnOpenLog.Text = "Open Log"
        $btnOpenLog.Size = New-Object System.Drawing.Size(140,40)
        $btnOpenLog.Location = New-Object System.Drawing.Point(200,340)
        $btnOpenLog.BackColor = [System.Drawing.Color]::FromArgb(221, 160, 221)
        $btnOpenLog.ForeColor = [System.Drawing.Color]::Black
        $btnOpenLog.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnOpenLog.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $btnOpenLog.Add_Click({
            Start-Process notepad.exe "`"$LogFile`""
            $finalForm.Close()
        })
        $finalForm.Controls.Add($btnOpenLog)
    }

    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Text = "Exit"
    $btnExit.Size = New-Object System.Drawing.Size(140,40)
    $btnExit.Location = New-Object System.Drawing.Point(370,340)
    $btnExit.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $btnExit.ForeColor = [System.Drawing.Color]::Black
    $btnExit.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnExit.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnExit.Add_Click({ $finalForm.Close() })
    $finalForm.Controls.Add($btnExit)

    $finalForm.ShowDialog() | Out-Null

    if($Global:LogginEnabled){
        Stop-Transcript | Out-Null
    }
}

#--------- End of Functions ----------#
#--------- End of Functions ----------#
#--------- End of Functions ----------#








#--------- Start of Process ----------#

# Ask Folder to scan user. Also last choosen path from temp folder
if(Test-Path "$TempFolder\PreviousLocation.json"){
    $LastLocation = Get-Content $TempFolder\PreviousLocation.json |ConvertFrom-Json
    if(Test-Path $LastLocation.FullName){
        $Global:Filepath = Get-Folder -initialDirectory $($LastLocation.FullName)
        get-item -Path $Filepath|ConvertTo-Json| Set-Content -Path "$TempFolder\PreviousLocation.json" -Force
    }else{
       $Global:Filepath = Get-Folder -initialDirectory $InitialFolder
       get-item -Path $Filepath|ConvertTo-Json| Set-Content -Path "$TempFolder\PreviousLocation.json" -Force
    }
}else{
    $Global:Filepath = Get-Folder -initialDirectory $InitialFolder
    get-item -Path $Filepath|ConvertTo-Json| Set-Content -Path "$TempFolder\PreviousLocation.json" -Force
}

# Check if choosen path really exist before continue. If exist check filenames for bad Characters and replace
if(!($Filepath)){
    Write-Warning "No path selected"
    Read-Host "Press enter to exit"
    set-ending
    break
}else{
    #$invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
    #$re = "[{0}]" -f [RegEx]::Escape($invalidChars +"][")
    $re2 = "[{0}]" -f [RegEx]::Escape("][")
    #($Name -replace $re)
    #$RegEx = '[][]'
    $RegEx2 = [string][System.IO.Path]::GetInvalidFileNameChars()
    Get-ChildItem $Filepath -File |%{
        if($_.FullName -match $re2){

            $newname = ($_.FullName.split('[')[0].split(']')[0]+ $_.Extension)
            Write-Warning "Bad File name found $($_.FullName)"
            Write-Warning "replace with $newname"
            $BadFileResponse = Ask-User -Title "Warning Bad File Name" -Message "
                Bad File name found:
                $($_.FullName)

                replace with:
                $newname

                "
            if($BadFileResponse -eq "yes"){
                Move-Item -LiteralPath $_.FullName  $newname
            }elseif ($BadFileResponse -eq "No" -or $BadFileResponse -eq "Cancel"){
                Write-Warning "Fix file and try again"
                set-ending
                break
            }

        }
     }
}

# Set choosen location
Set-Location $Filepath

# Ask begin and end check time
# $CheckTime = Get-CheckTime

# Start Logging to output file in temp logs folder (like fileCheck.ps1)
if($Global:LogginEnabled){
    $LogDir = Join-Path $TempFolder "Logs"
    $CurrentFolderName = Split-Path $Filepath -Leaf
    $DateTimeStamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
    $LogFile = "$LogDir\MP3Analyze-$CurrentFolderName-$DateTimeStamp.log"
    New-Item -ItemType Directory -Path $LogDir -ErrorAction SilentlyContinue | Out-Null
    Start-Transcript -Path $LogFile -Append | Out-Null
}

# Clear screen
if(-not $debug){
    Clear-Host
} 


# Check if Normalize folder exists
$NormFolder = Join-Path $Filepath "Normalize"
$redoNormalize = $true
if (Test-Path $NormFolder) {
    $response = Ask-User -Title "Folder '$NormFolder' already exists." -Message "Redo normalization? (yes/no)"
    if ($response -match '^(y|yes)$') {
        $newName = Get-UniqueFolderName $NormFolder
        Rename-Item -Path $NormFolder -NewName (Split-Path $newName -Leaf)
        $redoNormalize = $true
    } else {
        $redoNormalize = $false
    }
}

if ($redoNormalize) {
    Write-Host "Start Normalize"; Start-Sleep 1
    $NormFolder = Start-Normalize -folder $Filepath
}

# Check if Silence folder exists
$SilenceFolder = Join-Path $NormFolder "Silence"
$redoSilence = $true
if (Test-Path $SilenceFolder) {
    $response = Ask-User -Title "Folder '$SilenceFolder' already exists." -Message " Redo silence removal? (yes/no)"
    if ($response -match '^(y|yes)$') {
        $newName = Get-UniqueFolderName $SilenceFolder
        Rename-Item -Path $SilenceFolder -NewName (Split-Path $newName -Leaf)
        $redoSilence = $true
    } else {
        $redoSilence = $false
    }
}

if ($redoSilence) {
    Write-Host "Start Remove Silence"; Start-Sleep 1
    $SilenceFolder = Remove-Silence -folder $NormFolder
}

# Only run Fix-Id3andFileName if redoSilence is true or SilenceFolder did not exist before
if ($redoSilence) {
    Write-Host "Start Optimize ID3"; Start-Sleep 1
    Fix-Id3andFileName -folder $SilenceFolder -Prefix $Prefix
}

## Determen what options have been run and find right folder to process
$Global:MessureFolder = $SilenceFolder

# if($SilenceFolder){
#     $MessureFolder = $SilenceFolder

# }elseif($NormFolder){
#     $MessureFolder = $NormFolder
# }else{
#     $MessureFolder = $Filepath
# }



# Clear Screen and write text
if(-not $debug){
    Clear-Host
} 
write-host "Loading File... Please Wait"

# Analyse Folder and get ID3 Tag and file atributes
$ID3TagData = Get-MP3MetaData -Directory $MessureFolder

# Check if we found files in the above folder
if (-not $ID3TagData -or ($ID3TagData -is [array] -and $ID3TagData.Count -eq 0)) {
    Write-Warning "No files found"
    Read-Host -Prompt "Press enter to exit"
    # break
}

# Set Counters for reporting
$Total = $ID3TagData.Count
$BadCount = 0

# Set counter total time
Clear-Variable TotalMP3Time -Force -ea SilentlyContinue |Out-Null
$TotalMP3Time = (get-date -Hour 0 -Minute 0 -Second 0 -Millisecond 0)

# Clear the screen once more to be sure
if(-not $debug){
    Clear-Host
} 

# Here we go, start stopwatch. For reporting purphose
$StopWatch.Start()

# Finally Run Through Files (replace with new Check-File logic)
$ID3TagData |% {
    if(-not $debug){
        Clear-Host
    } 
    Write-Host "Count: $total / $($ID3TagData.Count)" -ForegroundColor Green
    $Result = Check-File -file $_
    Write-Host "$Result" -ForegroundColor Cyan
    if($Result -eq "Bad"){
        $BadCount = $BadCount + 1
    }
    $total = $Total - 1
    $TotalMP3Time += $_.length
}

# Stop loggin. Stop Stopwatch. Output Reporting. Open destination
Set-Ending

#### It's a wrap ####