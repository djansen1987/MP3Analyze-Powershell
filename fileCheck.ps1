Clear-Host
write-host "Warming up... Please Wait"
#--------- Set Parameters ----------#

## Set Tempfolder for last folder use
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
    
        
        #If ($FolderBrowser -eq "OK"){
        #    Return $FolderBrowser.SelectedPath
        #}
        #Else{
        #    Write-Error "Operation cancelled by user."
        #    break
        #}
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
    $form.Size = New-Object System.Drawing.Size(380,230)
    $form.StartPosition = 'CenterScreen'

    $TopButton = New-Object System.Windows.Forms.Button
    $TopButton.Location = New-Object System.Drawing.Point(0,120)
    $TopButton.Size = New-Object System.Drawing.Size(75,23)
    $TopButton.Text = 'Top'
    $TopButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $form.AcceptButton = $TopButton
    $form.Controls.Add($TopButton)

    $GoodButton = New-Object System.Windows.Forms.Button
    $GoodButton.Location = New-Object System.Drawing.Point(75,120)
    $GoodButton.Size = New-Object System.Drawing.Size(75,23)
    $GoodButton.Text = 'Goed'
    $GoodButton.DialogResult = [System.Windows.Forms.DialogResult]::yes

    $form.AcceptButton = $GoodButton
    $form.Controls.Add($GoodButton)


    $BadButton = New-Object System.Windows.Forms.Button
    $BadButton.Location = New-Object System.Drawing.Point(150,120)
    $BadButton.Size = New-Object System.Drawing.Size(75,23)
    $BadButton.Text = 'Dump'
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

    When Choose Goed, go to next.
    When Choose Top, move item to folder Top and go to next.
    When Choose Dump, move item to folder Dump and go to next.
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
        #Show-CurrentSong -Name ($_.name) -status "Begin" -time 10
#        Start-Process  $vlcPath -ArgumentList " --play-and-exit --qt-notification=0  `"$($_.Fullname)`" --run-time=$CheckTime " -Wait
        Start-Process  $vlcPath -ArgumentList "--qt-start-minimized --play-and-exit --qt-notification=0  `"$($_.Fullname)`" --run-time=$CheckTime " -Wait
        #Show-CurrentSong -Name ($_.name) -status "Ending" -time 10
        #Write-host "$($_.name)  (Ending)" -ForegroundColor Cyan
        #Start-Process  $vlcPath -ArgumentList "--qt-start-minimized --play-and-exit --qt-notification=0 `"$($_.Fullname)`" --start-time=$startend " -Wait
#        Start-Process  $vlcPath -ArgumentList " --play-and-exit --qt-notification=0 `"$($_.Fullname)`" --start-time=$startend " -Wait
    }
    catch{}
    
    (Get-Response -Name ($data.name))
}

function Check-File($File){
    $response = Start-Mp3 -data $_

    if($response -eq "Yes"){
        New-Item -ItemType Directory ($_.Directory + "\Goed\") -Force -ea SilentlyContinue|Out-Null
        Move-Item -Path $_.Fullname -Destination ($_.Directory + "\Goed\")
        return "Goed"
    }

    if($response -eq "OK"){
        New-Item -ItemType Directory ($_.Directory + "\Top\") -Force -ea SilentlyContinue|Out-Null
        Move-Item -Path $_.Fullname -Destination ($_.Directory + "\Top\")
        return "Top"
    }

    if($response -eq "No"){
        
        New-Item -ItemType Directory ($_.Directory + "\Dump\") -Force -ea SilentlyContinue|Out-Null
        Move-Item -Path $_.Fullname -Destination ($_.Directory + "\Dump\")
        return "Dump"
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
# to scan user. Also last choosen path from temp folder
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
            $BadFileResponse = Ask-User -Title "Warning Bad File Name" -Message "Bad File name found:
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

## Determen what options have been run and find right folder to process
   if($SilenceFolder){
        $MessureFolder = $SilenceFolder
   }elseif($NormFolder){
        $MessureFolder = $NormFolder
   }else{
        $MessureFolder = $Filepath
   }

Write-warning "Please wait while folder is scanned"
# Analyse Folder and get ID3 Tag and file atributes
$ID3TagData = Get-MP3MetaData -Directory $MessureFolder

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
    if(Test-Path $($_.fullname)){
        write-host "Count: $total / $($ID3TagData.Count)" -ForegroundColor Green
        $Result = Check-File -file $_
        write-host "$Result" -ForegroundColor Cyan
        if($Result -eq "Dump"){
            $BadCount = $BadCount + 1
        }    
        if($Result -eq "Top"){
            $topCount = $topCount + 1
        }
        if($Result -eq "good"){
            $goodCount = $goodCount + 1
        }
        $total = $Total - 1
        $TotalMP3Time += $_.length
    }

    
}

# We are at the end of the script. Let ending function know we made it
$Done = $true

# Stop loggin. Stop Stopwatch. Output Reporting. Open destination
Set-Ending

#### It's a wrap ####

 
 
                                                                                                       #  Top  -     Goed  -    Dump