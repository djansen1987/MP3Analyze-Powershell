Clear-Host
Write-Host "Warming up... Please Wait"

#--------- Set Parameters ----------#
# Set Tempfolder for last folder use
$TempFolder = "$env:TEMP\MP3Analyze"
if (!(Test-Path $TempFolder)) {
    New-Item -ItemType Directory -Path $TempFolder -Force | Out-Null
}
$InitialFolder = "D:\Temp\Music\Sidify\" # Default start folder
$RunTimeStamp = (Get-Date).ToString("yyyyMMdd") # Date for log naming
$Global:LogginEnabled = $true # Enable/disable logging
$Global:CheckTime = ""
$Global:Destdir = ""

#--------- Load Required Assemblies ----------#
Add-Type -AssemblyName presentationCore
Add-Type -AssemblyName presentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$mediaPlayer = New-Object system.windows.media.mediaplayer

#--------- Functions ----------#

# Show folder browser dialog and return selected path
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
    $form.Size = New-Object System.Drawing.Size(380,260)
    $form.StartPosition = 'CenterScreen'
    $form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None

    # Custom title bar
    $titleBar = New-Object System.Windows.Forms.Panel
    $titleBar.Size = New-Object System.Drawing.Size(380, 30)
    $titleBar.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $titleBar.Dock = [System.Windows.Forms.DockStyle]::Top

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $Title
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
    $closeButton.Location = New-Object System.Drawing.Point(340, 0)
    $closeButton.Add_Click({ $form.Close() })
    $titleBar.Controls.Add($closeButton)
    $form.Controls.Add($titleBar)

    $GoodButton = New-Object System.Windows.Forms.Button
    $GoodButton.Location = New-Object System.Drawing.Point(75,120)
    $GoodButton.Size = New-Object System.Drawing.Size(75,23)
    $GoodButton.Text = 'Yes'
    $GoodButton.DialogResult = [System.Windows.Forms.DialogResult]::yes
    $form.AcceptButton = $GoodButton
    $form.Controls.Add($GoodButton)

    $BadButton = New-Object System.Windows.Forms.Button
    $BadButton.Location = New-Object System.Drawing.Point(150,120)
    $BadButton.Size = New-Object System.Drawing.Size(75,23)
    $BadButton.Text = 'No'
    $BadButton.DialogResult = [System.Windows.Forms.DialogResult]::no
    $form.AcceptButton = $BadButton
    $form.Controls.Add($BadButton)

    $ReCheckButton = New-Object System.Windows.Forms.Button
    $ReCheckButton.Location = New-Object System.Drawing.Point(225,120)
    $ReCheckButton.Size = New-Object System.Drawing.Size(75,23)
    $ReCheckButton.Text = 'Stop'
    $ReCheckButton.DialogResult = [System.Windows.Forms.DialogResult]::cancel
    $form.CancelButton = $ReCheckButton
    $form.Controls.Add($ReCheckButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,40)
    $label.Size = New-Object System.Drawing.Size(480,180)
    $label.Text = $Message
    $label.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($label)
    $form.Topmost = $true

    $Prop = New-Object System.Windows.Forms.Form -Property @{TopMost = $true }
    $form.ShowDialog($prop)
}

# Ask user where to move sorted files (original or new folder)
function Ask-DestinationFolder {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Choose Destination Folder"
    $form.Size = New-Object System.Drawing.Size(420,180)
    $form.StartPosition = 'CenterScreen'
    $form.Topmost = $true
    $form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None

    # Custom title bar
    $titleBar = New-Object System.Windows.Forms.Panel
    $titleBar.Size = New-Object System.Drawing.Size(420, 30)
    $titleBar.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
    $titleBar.Dock = [System.Windows.Forms.DockStyle]::Top

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Choose Destination Folder"
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
    $closeButton.Location = New-Object System.Drawing.Point(380, 0)
    $closeButton.Add_Click({ $form.Close() })
    $titleBar.Controls.Add($closeButton)
    $form.Controls.Add($titleBar)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,40)
    $label.Size = New-Object System.Drawing.Size(390,40)
    $label.Text = "Where do you want to move the sorted files?"
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $label.ForeColor = [System.Drawing.Color]::FromArgb(255, 228, 181)
    $form.Controls.Add($label)

    $originalButton = New-Object System.Windows.Forms.Button
    $originalButton.Location = New-Object System.Drawing.Point(40,80)
    $originalButton.Size = New-Object System.Drawing.Size(150,40)
    $originalButton.Text = "Use Original Folder"
    $originalButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $originalButton.BackColor = [System.Drawing.Color]::FromArgb(173, 216, 230)
    $originalButton.ForeColor = [System.Drawing.Color]::Black
    $originalButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $originalButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($originalButton)

    $newButton = New-Object System.Windows.Forms.Button
    $newButton.Location = New-Object System.Drawing.Point(220,80)
    $newButton.Size = New-Object System.Drawing.Size(150,40)
    $newButton.Text = "Create New Folder"
    $newButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $newButton.BackColor = [System.Drawing.Color]::FromArgb(221, 160, 221)
    $newButton.ForeColor = [System.Drawing.Color]::Black
    $newButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $newButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($newButton)

    $form.AcceptButton = $originalButton
    $form.CancelButton = $originalButton

    $result = $form.ShowDialog()
    return $result
}

# Main MP3 review dialog with all options
Function Get-Response($Name, $total, $processed) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'MP3 Player'
    $form.Size = New-Object System.Drawing.Size(770, 370)
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
    Choose Good, move item to folder Good.
    Choose Top, move item to folder Top.
    Choose Bad, move item to folder Bad.
    Choose Other, move item to folder Other.

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

    # Add all choice buttons
    $GoodButton = New-Object System.Windows.Forms.Button
    $GoodButton.Location = New-Object System.Drawing.Point(30, 320)
    $GoodButton.Size = New-Object System.Drawing.Size(100, 50)
    $GoodButton.Text = 'Good (1)'
    $GoodButton.BackColor = [System.Drawing.Color]::FromArgb(173, 216, 230)
    $GoodButton.ForeColor = [System.Drawing.Color]::Black
    $GoodButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $GoodButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $GoodButton.FlatAppearance.BorderSize = 0
    $GoodButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $GoodButton.DialogResult = [System.Windows.Forms.DialogResult]::yes
    $form.Controls.Add($GoodButton)

    $TopButton = New-Object System.Windows.Forms.Button
    $TopButton.Location = New-Object System.Drawing.Point(150, 320)
    $TopButton.Size = New-Object System.Drawing.Size(100, 50)
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
    $BadButton.Location = New-Object System.Drawing.Point(270, 320)
    $BadButton.Size = New-Object System.Drawing.Size(100, 50)
    $BadButton.Text = 'Bad (3)'
    $BadButton.BackColor = [System.Drawing.Color]::FromArgb(255, 182, 193)
    $BadButton.ForeColor = [System.Drawing.Color]::Black
    $BadButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $BadButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $BadButton.FlatAppearance.BorderSize = 0
    $BadButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $BadButton.DialogResult = [System.Windows.Forms.DialogResult]::no
    $form.Controls.Add($BadButton)

    $OtherButton = New-Object System.Windows.Forms.Button
    $OtherButton.Location = New-Object System.Drawing.Point(390, 320)
    $OtherButton.Size = New-Object System.Drawing.Size(100, 50)
    $OtherButton.Text = 'Other (4)'
    $OtherButton.BackColor = [System.Drawing.Color]::FromArgb(255, 228, 181)
    $OtherButton.ForeColor = [System.Drawing.Color]::Black
    $OtherButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $OtherButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $OtherButton.FlatAppearance.BorderSize = 0
    $OtherButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $OtherButton.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
    $form.Controls.Add($OtherButton)

    $SkipButton = New-Object System.Windows.Forms.Button
    $SkipButton.Location = New-Object System.Drawing.Point(510, 320)
    $SkipButton.Size = New-Object System.Drawing.Size(100, 50)
    $SkipButton.Text = 'Skip (6)'
    $SkipButton.BackColor = [System.Drawing.Color]::FromArgb(230, 230, 200)
    $SkipButton.ForeColor = [System.Drawing.Color]::Black
    $SkipButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $SkipButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $SkipButton.FlatAppearance.BorderSize = 0
    $SkipButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $SkipButton.DialogResult = [System.Windows.Forms.DialogResult]::Abort
    $form.Controls.Add($SkipButton)

    $ReCheckButton = New-Object System.Windows.Forms.Button
    $ReCheckButton.Location = New-Object System.Drawing.Point(630, 320)
    $ReCheckButton.Size = New-Object System.Drawing.Size(100, 50)
    $ReCheckButton.Text = 'Re-Check (5)'
    $ReCheckButton.BackColor = [System.Drawing.Color]::FromArgb(221, 160, 221)
    $ReCheckButton.ForeColor = [System.Drawing.Color]::Black
    $ReCheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $ReCheckButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $ReCheckButton.FlatAppearance.BorderSize = 0
    $ReCheckButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $ReCheckButton.DialogResult = [System.Windows.Forms.DialogResult]::retry
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
            'D4' { $OtherButton.PerformClick() }
            'D5' { $ReCheckButton.PerformClick() }
            'D6' { $SkipButton.PerformClick() }
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
        "Yes" { $subfolder = "Goed" }
        "OK" { $subfolder = "Top" }
        "Ignore" { $subfolder = "Other" }
        "No" { $subfolder = "Bad" }
        "Retry" { return Check-File -file $File -total $total -processed $processed }
        "Abort" { return "Skip" }
        default {
            Write-Host "You Hit Cancel"
            Read-Host "Press enter to exit"
            break
        }
    }

    if ($response -eq "Abort") {
        return "Skip"
    }

    try {
        $destination = "$Destdir\$subfolder\"
        New-Item -ItemType Directory -Path $destination -Force -ea SilentlyContinue | Out-Null
        Move-Item -Path $File.Fullname -Destination $destination -ErrorAction Stop
        return $subfolder
    } catch {
        if ($_.Exception.Message -like "*already exists*") {
            $errorBase = "$Destdir\Error\"
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
            $errorDestination = "$Destdir\Error\"
            New-Item -ItemType Directory -Path $errorDestination -Force -ea SilentlyContinue | Out-Null
            Move-Item -Path $File.Fullname -Destination $errorDestination -Force
            return "Error: $($_.Exception.Message)"
        }
    }
}

# Get MP3/MP4 metadata for all files in directory
Function Get-MP3MetaData{
    [CmdletBinding()]
    [OutputType([Psobject])]
    Param(
        [String] [Parameter(Mandatory=$true, ValueFromPipeline=$true)] $Directory
    )
    Begin {
        $shell = New-Object -ComObject "Shell.Application"
    }
    Process {
        foreach($Dir in $Directory) {
            $ObjDir = $shell.NameSpace($Dir)
            $Files = Get-ChildItem $Dir | ?{$_.Extension -in '.mp3','.mp4'}
            foreach($File in $Files) {
                $ObjFile = $ObjDir.parsename($File.Name)
                $MetaData = @{}
                $MP3 = ($ObjDir.Items()|?{$_.path -like "*.mp3" -or $_.path -like "*.mp4"})
                $PropertArray = 0,1,2,12,13,14,15,16,17,18,19,20,21,22,27,28,36,220,223
                foreach($item in $PropertArray) { 
                    if($ObjDir.GetDetailsOf($ObjFile, $item)) {
                        $MetaData[$($ObjDir.GetDetailsOf($MP3,$item))] = $ObjDir.GetDetailsOf($ObjFile, $item)
                    }
                }
                New-Object psobject -Property $MetaData | select *, @{n="Directory";e={$Dir}}, @{n="Fullname";e={Join-Path "$Dir" "$($File.Name)" -Resolve}}, @{n="Extension";e={$File.Extension}}
            }
        }
    }
    End {}
}

#--------- Main Script Logic ----------#

# Restore last used folder or ask user for folder
if(Test-Path "$TempFolder\PreviousLocation.json"){
    $LastLocation = Get-Content $TempFolder\PreviousLocation.json | ConvertFrom-Json
    if(Test-Path $LastLocation.FullName){
        $Filepath = Get-Folder -initialDirectory $($LastLocation.FullName)
        Get-Item -Path $Filepath | ConvertTo-Json | Set-Content -Path "$TempFolder\PreviousLocation.json" -Force
    }else{
        $Filepath = Get-Folder -initialDirectory $InitialFolder
        Get-Item -Path $Filepath | ConvertTo-Json | Set-Content -Path "$TempFolder\PreviousLocation.json" -Force
    }
}else{
    $Filepath = Get-Folder -initialDirectory $InitialFolder
    Get-Item -Path $Filepath | ConvertTo-Json | Set-Content -Path "$TempFolder\PreviousLocation.json" -Force
}

# Check for invalid file names and prompt user to fix
if(!($Filepath)){
    Write-Warning "No path selected"
    Read-Host "Press enter to exit"
    break
}else{
    $re2 = "[{0}]" -f [RegEx]::Escape("][")
    Get-ChildItem $Filepath -File | %{
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
                break
            }
        }
    }
}

# Set chosen location and ask for destination folder
Set-Location $Filepath
if(!($org)){
    $destResult = Ask-DestinationFolder
    if ($destResult -eq [System.Windows.Forms.DialogResult]::No) {
        $Global:Destdir = "D:\temp\Music\#MP3-Done$(get-date -Format ddMMyy)"
        if(!(Test-Path $Destdir)){
            New-Item -ItemType Directory $Destdir | Out-Null
        }
    } else {
        $Global:Destdir = $Filepath
    }
}

# Start Logging to output file in temp logs folder
if($Global:LogginEnabled){
    $LogDir = Join-Path $TempFolder "Logs"
    $CurrentFolderName = Split-Path $Filepath -Leaf
    $DateTimeStamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
    $LogFile = "$LogDir\MP3Analyze-$CurrentFolderName-$DateTimeStamp.log"
    New-Item -ItemType Directory -Path $LogDir -ErrorAction SilentlyContinue | Out-Null
    Start-Transcript -Path $LogFile -Append | Out-Null
}

# Initialize counters
$BadCount = 0
$topCount = 0
$goodCount = 0
$otherCount = 0
$Errorcount = 0
$duplicateCount = 0
$SkipCount = 0

# Track time for the whole process
$scriptStart = Get-Date

# Get all MP3/MP4 metadata in folder
$ID3TagData = Get-MP3MetaData -Directory $Filepath
$total = ($ID3TagData | Measure-Object).count

# Calculate total song length in seconds using ID3 tag duration (try to find the correct property)
$TotalSongSeconds = 0
foreach ($item in $ID3TagData) {
    # Try common property names for duration
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

# Process each file and move/categorize based on user input
$ID3TagData | % {
    if(Test-Path $($_.fullname)){
        Write-Host "Count: $total / $($ID3TagData.Count)" -ForegroundColor Green
        $Result = Check-File -file $_ -total $(($ID3TagData|Measure-Object).count) -processed $total
        Write-Host "$Result" -ForegroundColor Cyan
        if($Result -eq "Bad"){
            $BadCount++
        }elseif($Result -eq "Top"){
            $topCount++
        }elseif($Result -eq "good"){
            $goodCount++
        }elseif($Result -eq "other"){
            $otherCount++
        }elseif($Result -eq "Skip"){
            $SkipCount++
        }elseif($Result -like "*Duplicate*"){
            $duplicateCount++
            $Errorcount++
        }elseif($Result -like "Error*"){
            $Errorcount++
        }
        $total = $Total - 1
    }
}

# Calculate elapsed time
$scriptEnd = Get-Date
$elapsed = $scriptEnd - $scriptStart

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
    Start-Process explorer.exe "`"$Filepath`""
    $finalForm.Close()
})
$finalForm.Controls.Add($btnOpenFolder)

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

Stop-Transcript

# Helper function to generate a unique folder name
function Get-UniqueFolderName($baseFolder) {
    $counter = 1
    $newFolder = "${baseFolder}_old"
    while (Test-Path $newFolder) {
        $newFolder = "${baseFolder}_old$counter"
        $counter++
    }
    return $newFolder
}

# Example usage for Silence folder (add this function before you use it)
$SilenceFolder = Join-Path $NormFolder "Silence"
$redoSilence = $true
if (Test-Path $SilenceFolder) {
    $response = Read-Host "Folder '$SilenceFolder' already exists. Redo silence removal? (yes/no)"
    if ($response -match '^(y|yes)$') {
        $newName = Get-UniqueFolderName $SilenceFolder
        Rename-Item -Path $SilenceFolder -NewName (Split-Path $newName -Leaf)
        $redoSilence = $true
    } else {
        $redoSilence = $false
    }
}