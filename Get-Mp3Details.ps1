Clear-Host
write-host "Warming up... Please Wait"
#--------- Set Parameters ----------#

## Set Tempfolder for last folder use test
$TempFolder = "$env:TEMP\MP3Analyze"
## If first run, no folder is set start in
$InitialFolder = "C:\Temp\Sidify\Download-Temp\"
## Set current date format and stamp
$RunTimeStamp = $((get-date).ToString("yyyyMMdd"))
## Log Powershell output to file in same directory
$Global:LogginEnabled = $true ## $true = yes | $false = no
$Prefix = "(SP-RIP-N)"

#--------- DO Not Edit Below ----------#

# Initial status
$Done = $false
# Load stopwatch
$StopWatch = New-Object System.Diagnostics.Stopwatch

# Find Tempfolder and create if not exist
if(!(Test-Path $TempFolder)){
    New-Item -ItemType Directory $TempFolder
}

# Check if ID3 powershell gallary module is installed. If not install else import
$ID3Module = Get-Module -Name ID3
if(!($ID3Module)){
    Import-Module -Name ID3
}

$ID3Module = Get-Module -Name ID3
if(!($ID3Module)){
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $IsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if($IsAdmin){
        Install-Module -Name ID3
    }else{
        Write-Warning "Need to be admin to install ID3 Powershell Module"
        Read-Host "Press enter to exit"
        break
    }
}

# Find VLC path
$vlcinstall = Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | % { Get-ItemProperty $_.PsPath } | Select DisplayName,InstallLocation|?{$_.DisplayName -like "*vlc*"}
if(!($vlcinstall)){
    if(Test-Path "C:\Program Files\VideoLAN\VLC\vlc.exe"){
        $vlcPath = "C:\Program Files\VideoLAN\VLC\vlc.exe"
    }elseif(Test-Path "C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"){
        $vlcPath = "C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"
    }else{
        write-host "vlc not found, please install vlc"
        read-host -Prompt "Press enter to exit."
        (New-Object -Com Shell.Application).Open("https://www.videolan.org/vlc/")        
        break
    }
}else{
    $vlcPath = "$($vlcinstall.InstallLocation)\vlc.exe"    
}

#--------- Begin of Functions ----------#

Function Get-Folder($initialDirectory){

            Add-Type -AssemblyName System.Windows.Forms
            $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
                SelectedPath = $initialDirectory; ShowNewFolderButton = $false
            }


        $Prop = New-Object System.Windows.Forms.Form -Property @{TopMost = $true }

        [void]$FolderBrowser.ShowDialog($Prop)  
        Return $FolderBrowser.SelectedPath

}

Function Get-MP3MetaData{
    [CmdletBinding()]
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

Function Show-CurrentSong ($Name,$status,$time) {

    Add-Type -AssemblyName System.Windows.Forms    

    # Build Form
    $objForm = New-Object System.Windows.Forms.Form
    $objForm.Text = $status
    $objForm.Size = New-Object System.Drawing.Size(220,100)

    # Add Label
    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(80,20) 
    $objLabel.Size = New-Object System.Drawing.Size(100,20)
    $objLabel.Text = $Name
    $objForm.Controls.Add($objLabel)
    
    # Show the form
    $objForm.Show()| Out-Null

    # wait 5 seconds
    Start-Sleep -Seconds $time

    # destroy form
    $objForm.Close() | Out-Null
    
}

function Get-Response($Name){
    Add-Type -AssemblyName System.Windows.Forms

    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'How Was the MP3'
    $form.Size = New-Object System.Drawing.Size(380,260)
    $form.StartPosition = 'CenterScreen'

    $GoodButton = New-Object System.Windows.Forms.Button
    $GoodButton.Location = New-Object System.Drawing.Point(75,120)
    $GoodButton.Size = New-Object System.Drawing.Size(75,23)
    $GoodButton.Text = 'Good'
    $GoodButton.DialogResult = [System.Windows.Forms.DialogResult]::yes

    $form.AcceptButton = $GoodButton
    $form.Controls.Add($GoodButton)

    $BadButton = New-Object System.Windows.Forms.Button
    $BadButton.Location = New-Object System.Drawing.Point(150,120)
    $BadButton.Size = New-Object System.Drawing.Size(75,23)
    $BadButton.Text = 'Bad'
    $BadButton.DialogResult = [System.Windows.Forms.DialogResult]::no

    $form.AcceptButton = $BadButton
    $form.Controls.Add($BadButton)

    $ReCheckButton = New-Object System.Windows.Forms.Button
    $ReCheckButton.Location = New-Object System.Drawing.Point(225,120)
    $ReCheckButton.Size = New-Object System.Drawing.Size(75,23)
    $ReCheckButton.Text = 'Re-Check'
    $ReCheckButton.DialogResult = [System.Windows.Forms.DialogResult]::retry

    $form.CancelButton = $ReCheckButton
    $form.Controls.Add($ReCheckButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(480,180)
    $label.Text = "
    How Was $Name ?

    When Choose Good, go to next.
    When Choose Bad, move item to folder Bad and go to next.
    When Choos Re-Check, play Mp3 Again.

    "

    $form.Controls.Add($label)
    $form.Topmost = $true

    $Prop = New-Object System.Windows.Forms.Form -Property @{TopMost = $true }
    $form.ShowDialog($prop)
    #$form.ShowDialog()


}

function Start-Mp3($data){
 

    try{
        $startEnd = ([DateTime]$_.Length).AddSeconds(-$CheckTime).TimeOfDay.TotalSeconds
        Write-host "$($_.name)  (Begin)" -ForegroundColor Cyan
        Start-Process  $vlcPath -ArgumentList "--qt-start-minimized --play-and-exit --qt-notification=0  `"$($_.Fullname)`" --run-time=$CheckTime " -Wait
        Write-host "$($_.name)  (Ending)" -ForegroundColor Cyan
        Start-Process  $vlcPath -ArgumentList "--qt-start-minimized --play-and-exit --qt-notification=0 `"$($_.Fullname)`" --start-time=$startend " -Wait
    }
    catch{}
    
    (Get-Response -Name ($data.name))
}

function Check-File($File){
    $response = Start-Mp3 -data $_

    if($response -eq "Yes"){
        return "Good"
    }

    if($response -eq "No"){
        
        New-Item -ItemType Directory ($_.Directory + "\Bad\") -Force -ea SilentlyContinue|Out-Null
        Move-Item -Path $_.Fullname -Destination ($_.Directory + "\Bad\")
        return "Bad"
    }
    if($response -eq "Retry"){
        Check-File -file $_
    }else{
        write-host "You Hit Cancel"
        read-host "press enter to exit"
        Set-Ending
        break
    }
    
}

function Get-CheckTime(){
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object “System.Windows.Forms.Form”;
    $form.Width = 300;
    $form.Height = 150;
    $form.Text = "Number of second's to check (begin-end)";
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;

    ##############Define text label1
    $textLabel1 = New-Object “System.Windows.Forms.Label”;
    $textLabel1.Left = 25;
    $textLabel1.Top = 10;
    $textLabel1.width = 300
    $textLabel1.Text = "Number of second's to check (begin-end)";

    ############Define text box1 for input
    $textBox1 = New-Object “System.Windows.Forms.TextBox”;
    $textBox1.Left = 25;
    $textBox1.Top = 35;
    $textBox1.width = 200;
    $textBox1.Text = "10";

    $button = New-Object “System.Windows.Forms.Button”;
    $button.Left = 25;
    $button.Top = 70;
    $button.Width = 100;
    $button.Text = “Ok”;
    $button.DialogResult = [System.Windows.Forms.DialogResult]::ok

    $eventHandler = [System.EventHandler]{
    $textBox1.Text;
    $form.Close();};
    $button.Add_Click($eventHandler) ;

    $form.KeyPreview = $True
    $form.Add_KeyDown({
        if ($_.KeyCode -eq "Enter") {
            # if enter, perform click
            $button.PerformClick()
        }
    })
    $form.Add_Shown({$form.Activate(); $textBox1.focus()})
    $form.Controls.Add($textLabel1);
    $form.Controls.Add($textBox1);
    $form.Controls.Add($button);
    #$ret = $form.ShowDialog();

    $Prop = New-Object System.Windows.Forms.Form -Property @{TopMost = $true }
    $show = $form.ShowDialog($prop)
    
    

    If ($show -eq "OK"){
        Return $textBox1.Text
    }
    Else{
        Write-Error "Operation cancelled by user."
        Set-Ending
        break
    }
}

function Ask-User($Title,$Message){
   Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(380,260)
    $form.StartPosition = 'CenterScreen'

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
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(480,180)
    $label.Text = $Message

    $form.Controls.Add($label)
    $form.Topmost = $true

    $Prop = New-Object System.Windows.Forms.Form -Property @{TopMost = $true }
    $form.ShowDialog($prop)
    #$form.ShowDialog()
}

function Start-Normalize($folder){

    $items = Get-ChildItem -Path "$folder" -File -filter "*.mp3"
    $totalitems = $items.count
    $itemstodo = $totalitems

    $waitmessage =  "Normalizing File...Approx wait Time: " +(0..$totalitems| % -Begin {$Total = 0} -Process {$Total += (New-TimeSpan -second 2)} -End {$Total})

    $items|%{
        $filename = $_ 
        Clear-Host
        write-host $waitmessage
        Write-Host "$itemstodo / $totalitems  -  $filename";$itemstodo = ($itemstodo - 1)
        ffmpeg-normalize $filename.FullName -of $($folder + "\Normalize") --normalization-type peak --target-level 0 -c:a libmp3lame -b:a 256k -ext mp3
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
        Clear-Host
        $filename = $_ 
        write-host $waitmessage
        Write-Host "$itemstodo / $totalitems  -  $filename"
        $itemstodo = ($itemstodo - 1)
        ffmpeg -i $($folder + "\"+$filename.name) -y -c:a libmp3lame -b:a 256k -af silenceremove=1:0:-50dB -loglevel warning $($folder + "\Silence\"+$filename.name) 
    }
    
    return $($folder + "\Silence\")
}

function Fix-Id3andFileName ($folder,$Prefix){
    $items = Get-ChildItem -Path "$folder" -File -filter "*.mp3"
    $totalitems = $items.count
    $waitmessage = "Renaming Files and updating ID3Tag... Please Wait"
    $itemstodo = $totalitems
    $items|%{Clear-Host;$filename = $_ ;write-host $waitmessage;Write-Host "$itemstodo / $totalitems  -  $filename";$itemstodo = ($itemstodo - 1);`
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

function Set-Ending(){
    # Were done. Stop The Time!
    $StopWatch.Stop()

    # Some math (proberbly wrong)
    $savedtimecalc = [timespan]::fromseconds( $(( ($TotalMP3Time.TimeOfDay.TotalSeconds) - (($ID3TagData.Count) * ($CheckTime*2)) )) ) 

    # Output some Details/Summary
    
    if($Done){
        Write-Host " "
        Write-Host " "
        Write-Host " "
        Write-Host " "
        Write-Host " "
        Write-Host " "
        Write-Host " --------------- Summary -------------- "
        Write-Host " "
        write-host "Total`t`t`t $($ID3TagData.Count)"
        write-host "Total Bad:`t`t $BadCount"
        write-host "Total Runtime:`t $([string]::Format("`{0:d2}:{1:d2}:{2:d2}",$StopWatch.Elapsed.hours,$StopWatch.Elapsed.minutes,$StopWatch.Elapsed.seconds))"
        write-host "Total playtime:`t $([string]::Format("`{0:d2}:{1:d2}:{2:d2}",$TotalMP3Time.TimeOfDay.hours,$TotalMP3Time.TimeOfDay.minutes,$TotalMP3Time.TimeOfDay.seconds))"
        write-host "Saved Time: `t $([string]::Format("`{0:d2}:{1:d2}:{2:d2}",$savedtimecalc.hours,$savedtimecalc.minutes,$savedtimecalc.seconds))"
        read-host -Prompt "Press enter to exit and open output folder."
        Start-Process explorer $MessureFolder
    }else{
        Write-Warning "Failed or Cancelled"
    }
    if($Global:LogginEnabled){
        stop-Transcript |Out-Null
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
        $Filepath = Get-Folder -initialDirectory $($LastLocation.FullName)
        get-item -Path $Filepath|ConvertTo-Json| Set-Content -Path "$TempFolder\PreviousLocation.json" -Force
    }else{
       $Filepath = Get-Folder -initialDirectory $InitialFolder
       get-item -Path $Filepath|ConvertTo-Json| Set-Content -Path "$TempFolder\PreviousLocation.json" -Force
    }
}else{
    $Filepath = Get-Folder -initialDirectory $InitialFolder
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
            $BadFileResponse = Ask-User -Title "Warning Bad File Name" -Message "                Bad File name found:
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
$CheckTime = Get-CheckTime

# Start Logging to output file in choosen folder
if($Global:LogginEnabled){
    New-Item -ItemType Directory -Path ($Filepath+"\Log\") -ErrorAction SilentlyContinue|Out-Null
    Start-Transcript -Path ($Filepath+"\Log\#1_MP3Analyzed-" + $RunTimeStamp+".log") -Append |Out-Null
}

# Clear screen
Clear-Host

# Ask To do all fixes at once
$FixAllResponse = Ask-User -Title "Normalize Files?" -Message "
Would you like to do all fiexes at once? (Normalize, remove Silence, optimize ID3)
    "
if ($FixAllResponse -eq "Yes"){
    
    Write-Host "Start Normalize";Start-Sleep 1
    $NormFolder = Start-Normalize -folder $Filepath
    
    Write-Host "Start Remove Silence";Start-Sleep 1
    $SilenceFolder = Remove-Silence -folder $NormFolder

    Write-Host "Start Optimize ID3";Start-Sleep 1
    Fix-Id3andFileName -folder $SilenceFolder -Prefix $Prefix

}elseif ($FixAllResponse -eq "No"){


    # Ask if we should Normalize the files. If yes start function
    $NormalizeResponse = Ask-User -Title "Normalize Files?" -Message "
        With this option Mp3Gain will normalize the files to 0dB.

        A separated folder will be created `"Normalized`"
        "
    if ($NormalizeResponse -eq "Yes"){
       $NormFolder = Start-Normalize -folder $Filepath
    }elseif ($NormalizeResponse -eq "No"){
        Write-host "Skipping normalize"
    }elseif ($NormalizeResponse -eq "Cancel"){
        set-ending
        break
    }


    # Ask if we should Remove silence from begin and end. If yes start function
    $RemoveSilenceResponse = Ask-User -Title "Remove Silence?" -Message "
        With this option Silence will be removed from the MP3.

        ffmpeg setting = silenceremove=1:0:-50dB

        Which means:
        Remove from the begin till level is abobe 0 DB
        Remove at the end everything below -50 db
        "
    if ($RemoveSilenceResponse -eq "Yes"){
        if($NormFolder){
            $SilenceFolder = Remove-Silence -folder $NormFolder
        }else{
            $SilenceFolder = Remove-Silence -folder $Filepath
        }
   
    }elseif ($RemoveSilenceResponse -eq "No"){
        Write-host "Skipping Remove Silence"
    }elseif ($RemoveSilenceResponse -eq "Cancel"){
        set-ending
        break
    }


    # Ask if we should fix the Filename and ID3 Tag. If yes start function
    $FixID3TagResponse = Ask-User -Title "Fix ID3 and Filenames?" -Message "
        With This option we will Correct the filename 
        with `"Artist - Title (SP-RIP-N)`".

        Files will be saved in last option sub-folder
        "
    if ($FixID3TagResponse -eq "Yes"){
        if($SilenceFolder){
            Fix-Id3andFileName -folder $SilenceFolder -Prefix $Prefix
        }elseif($NormFolder){
            Fix-Id3andFileName -folder $NormFolder -Prefix $Prefix
        }else{
            Fix-Id3andFileName -folder $Filepath -Prefix $Prefix
        }
    }elseif ($FixID3TagResponse -eq "No"){
        Write-host "Skipping normalize"
    }elseif ($FixID3TagResponse -eq "Cancel"){
        set-ending
        break
    }


}elseif ($FixAllResponse -eq "Cancel"){
    set-ending
    break
}


## Determen what options have been run and find right folder to process
if($SilenceFolder){
    $MessureFolder = $SilenceFolder
}elseif($NormFolder){
    $MessureFolder = $NormFolder
}else{
    $MessureFolder = $Filepath
}



# Clear Screen and write text
Clear-Host
write-host "Loading File... Please Wait"

# Analyse Folder and get ID3 Tag and file atributes
$ID3TagData = Get-MP3MetaData -Directory $MessureFolder

# Check if we found files in the above folder
if (!(($ID3TagData.count) -gt 0)){
    Write-Warning "No files found"
    Read-Host -Prompt "Press enter to exit"
    break
}

# Set Counters for reporting
$Total = $ID3TagData.Count
$BadCount = 0

# Set counter total time
Clear-Variable TotalMP3Time -Force -ea SilentlyContinue |Out-Null
$TotalMP3Time = (get-date -Hour 0 -Minute 0 -Second 0 -Millisecond 0)

# Clear the screen once more to be sure
Clear-Host

# Here we go, start stopwatch. For reporting purphose
$StopWatch.Start()

# Finally Run Through Files
$ID3TagData |% {
    Clear-Host
    write-host "Count: $total / $($ID3TagData.Count)" -ForegroundColor Green
    $Result = Check-File -file $_
    write-host "$Result" -ForegroundColor Cyan
    if($Result -eq "Bad"){
        $BadCount = $BadCount + 1
    }
    $total = $Total - 1
    $TotalMP3Time += $_.length
    
}

# We are at the end of the script. Let ending function know we made it
$Done = $true

# Stop loggin. Stop Stopwatch. Output Reporting. Open destination
Set-Ending

#### It's a wrap ####