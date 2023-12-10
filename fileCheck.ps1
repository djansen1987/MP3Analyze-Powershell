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

#mediaplayer
Add-Type -AssemblyName presentationCore
$mediaPlayer = New-Object system.windows.media.mediaplayer

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

function Get-Response($Name){
    Add-Type -AssemblyName System.Windows.Forms

    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'How Was the MP3'
    $form.Size = New-Object System.Drawing.Size(380,230)
    $form.StartPosition = 'CenterScreen'

    $GoodButton = New-Object System.Windows.Forms.Button
    $GoodButton.Location = New-Object System.Drawing.Point(0,120)
    $GoodButton.Size = New-Object System.Drawing.Size(75,23)
    $GoodButton.Text = 'Goed'
    $GoodButton.DialogResult = [System.Windows.Forms.DialogResult]::yes

    $form.AcceptButton = $GoodButton
    $form.Controls.Add($GoodButton)

    $TopButton = New-Object System.Windows.Forms.Button
    $TopButton.Location = New-Object System.Drawing.Point(75,120)
    $TopButton.Size = New-Object System.Drawing.Size(75,23)
    $TopButton.Text = 'Top'
    $TopButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $form.AcceptButton = $TopButton
    $form.Controls.Add($TopButton)

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

}


function Start-Mp3($data){
 

    try{
        
        $startEnd = ([DateTime]$_.Length).AddSeconds(-([int]$CheckTime +1)).TimeOfDay.TotalSeconds

        Write-host "$($_.name)  (First $CheckTime seconds)" -ForegroundColor Cyan

        $mediaPlayer.open($($_.Fullname))
        #mediaPlayer.open("C:\Sidify-download\Top 40 2021\Normalize\Silence\Donnie, Rene Froger - Bon Gepakt (SP-RIP-N).mp3")
        $mediaPlayer.Position=New-Object System.TimeSpan(0, 0, 0, 0, 0)
        $mediaPlayer.Play()

        Start-Sleep ([int]$CheckTime + 2)
        $mediaPlayer.Pause()
        Start-Sleep -Milliseconds 500

        Write-host "$($_.name)  (Last $CheckTime seconds)" -ForegroundColor Cyan
        $mediaPlayer.Position=New-Object System.TimeSpan(0, 0, 0, $startEnd, 0)
        $mediaPlayer.Play()
        
        Start-Sleep -Seconds ([int]$CheckTime + 1)
        
        $mediaPlayer.Stop()
        $mediaPlayer.Close()

    }
    catch{
    
        Write-Error "unable to play audio"
    }
    
    (Get-Response -Name ($data.name))
}

function Check-File($File){

    $response = Start-Mp3 -data $File

    if($response -eq "Yes"){
        return "Goed"
    }

    if($response -eq "OK"){
        New-Item -ItemType Directory ($File.Directory + "\Top\") -Force -ea SilentlyContinue|Out-Null
        Move-Item -Path $_.Fullname -Destination ($File.Directory + "\Top\")
        return "Top"
    }

    if($response -eq "No"){
        
        New-Item -ItemType Directory ($File.Directory + "\Dump\") -Force -ea SilentlyContinue|Out-Null
        Move-Item -Path $File.Fullname -Destination ($File.Directory + "\Dump\")
        return "Dump"
    }
    if($response -eq "Retry"){
        Check-File -file $File
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
#clear-Host


$total = $Files.Count
$BadCount = 0

# Clear the screen once more to be sure
#Clear-Host
$ID3TagData = Get-MP3MetaData -Directory $Filepath


# Finally Run Through Files
$ID3TagData |% {
    #Clear-Host
    if(Test-Path $($_.fullname)){
        Write-Host "Count: $total / $($ID3TagData.Count)" -ForegroundColor Green
        $Result = Check-File -file $_
        Write-Host "$Result" -ForegroundColor Cyan
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
    }

    
}
                                                                                                       #  Top  -     Goed  -    Dump