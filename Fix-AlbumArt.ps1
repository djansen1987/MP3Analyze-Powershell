function Select-MainFolder {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select the main folder containing the Normalize\Silence subfolder"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }
    return $null
}

function Fix-AlbumArt-InSilenceFolder {
    param(
        [string]$MainFolder
    )
    $normalizeFolder = Join-Path $MainFolder "Normalize"
    $silenceFolder = Join-Path $normalizeFolder "Silence"
    if (!(Test-Path $silenceFolder)) {
        Write-Warning "Silence folder not found: $silenceFolder"
        return
    }
    $mp3s = Get-ChildItem -Path $silenceFolder -Filter *.mp3 -File -Recurse
    foreach ($mp3 in $mp3s) {
        # $srcFile = Join-Path $Main $(($mp3.Name).split(" - ")[0].replace(' (SP-RIP-N)', ''))
        $srcFile = Join-Path $Main $(($mp3.Name).replace(' (SP-RIP-N)', ''))

        $dstFile = $mp3.FullName
        if (!(Test-Path "$srcFile")) {
            $srcFile = (Get-ChildItem $main | where { $_.name -like "$(($mp3.Name).split(" - ")[0])*" }).FullName
            if (!(Test-Path "$srcFile")) {
                Write-Warning "Source file not found for $($srcFile) $($mp3.Name)"
                continue
            }
            # Write-Warning "Source file not found for $($srcFile) $($mp3.Name)"
            # continue
        }
        $coverTemp = [System.IO.Path]::GetTempFileName() + ".jpg"
        ffmpeg -y -i "$srcFile" -an -vcodec copy "$coverTemp" 2>$null
        if ((Test-Path $coverTemp) -and ((Get-Item $coverTemp).Length -gt 0)) {
            $outTemp = [System.IO.Path]::GetTempFileName() + ".mp3"
            ffmpeg -y -i "$dstFile" -i "$coverTemp" -map 0:a -map 1 -c copy -id3v2_version 3 "$outTemp" 2>$null
            Move-Item -Force "$outTemp" "$dstFile"
            Remove-Item "$coverTemp" -Force
            Write-Host "Restored album art for $($mp3.Name)"
        } else {
            if (Test-Path $coverTemp) { Remove-Item "$coverTemp" -Force }
            Write-Host "No album art found in $srcFile"
        }
    }
}

# Example usage:
$main = Select-MainFolder
if ($main) {
    Fix-AlbumArt-InSilenceFolder -MainFolder $main
}
