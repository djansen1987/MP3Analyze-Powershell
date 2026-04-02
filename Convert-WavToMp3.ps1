#requires -version 5.1

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {    
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell"
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`""
    $psi.UseShellExecute = $true
    [System.Diagnostics.Process]::Start($psi) | Out-Null
    exit 0
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ffmpegPath = Join-Path $scriptRoot "Source\ffmpeg\bin\ffmpeg.exe"

if (-not (Test-Path $ffmpegPath)) {
    $ffmpegCmd = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if ($ffmpegCmd -and $ffmpegCmd.Source) {
        $ffmpegPath = $ffmpegCmd.Source
    }
}

if (-not (Test-Path $ffmpegPath)) {
    [System.Windows.Forms.MessageBox]::Show(
        "FFmpeg niet gevonden. Plaats ffmpeg.exe in Source\\ffmpeg\\bin of zorg dat het in PATH staat.",
        "FFmpeg ontbreekt",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

function Select-Input {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "WAV naar MP3"
    $form.Size = New-Object System.Drawing.Size(420, 180)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Kies een WAV bestand of een hele map met WAV bestanden."
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $form.Controls.Add($label)

    $btnFile = New-Object System.Windows.Forms.Button
    $btnFile.Text = "Bestand..."
    $btnFile.Size = New-Object System.Drawing.Size(120, 30)
    $btnFile.Location = New-Object System.Drawing.Point(10, 50)
    $form.Controls.Add($btnFile)

    $btnFolder = New-Object System.Windows.Forms.Button
    $btnFolder.Text = "Map..."
    $btnFolder.Size = New-Object System.Drawing.Size(120, 30)
    $btnFolder.Location = New-Object System.Drawing.Point(150, 50)
    $form.Controls.Add($btnFolder)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuleren"
    $btnCancel.Size = New-Object System.Drawing.Size(120, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(290, 50)
    $form.Controls.Add($btnCancel)

    $btnFile.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "WAV (*.wav)|*.wav"
        $dialog.Multiselect = $false
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $form.Tag = @{ Type = "File"; Path = $dialog.FileName }
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        }
    })

    $btnFolder.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Kies een map met WAV bestanden"
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $form.Tag = @{ Type = "Folder"; Path = $dialog.SelectedPath }
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        }
    })

    $btnCancel.Add_Click({
        $form.Close()
    })

    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()

    return $form.Tag
}

function Convert-WavFile {
    param(
        [Parameter(Mandatory = $true)][string]$WavPath
    )

    $directory = Split-Path -Parent $WavPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($WavPath)
    $mp3Path = Join-Path $directory ($baseName + ".mp3")

    $arguments = @(
        "-hide_banner",
        "-loglevel", "error",
        "-y",
        "-i", $WavPath,
        "-codec:a", "libmp3lame",
        "-b:a", "320k",
        "-ar", "44100",
        $mp3Path
    )

    & $ffmpegPath @arguments | Out-Null

    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $mp3Path)) {
        Write-Warning "Conversie mislukt: $WavPath"
        return $false
    }

    $originalFolder = Join-Path $directory "original"
    if (-not (Test-Path $originalFolder)) {
        New-Item -ItemType Directory -Path $originalFolder | Out-Null
    }

    $destination = Join-Path $originalFolder ([System.IO.Path]::GetFileName($WavPath))
    Move-Item -Path $WavPath -Destination $destination -Force

    return $true
}

$selection = Select-Input
if (-not $selection) {
    Write-Host "Geannuleerd."
    exit 0
}

$files = @()
if ($selection.Type -eq "File") {
    $files = @($selection.Path)
} elseif ($selection.Type -eq "Folder") {
    $files = Get-ChildItem -Path $selection.Path -Recurse -File -Include *.wav, *.WAV | ForEach-Object { $_.FullName }
}

if (-not $files -or $files.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        "Geen WAV bestanden gevonden.",
        "Niets te doen",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
    exit 0
}

$success = 0
$total = $files.Count
for ($i = 0; $i -lt $total; $i++) {
    $wav = $files[$i]
    $index = $i + 1
    Write-Host "[$index/$total] Converteer: $wav"
    if (Convert-WavFile -WavPath $wav) {
        $success++
    }
}

[System.Windows.Forms.MessageBox]::Show(
    "Klaar. $success van $total bestanden geconverteerd.",
    "Gereed",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null
