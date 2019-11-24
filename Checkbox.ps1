#define a tooltip object
$tooltip1 = New-Object System.Windows.Forms.ToolTip
$ShowHelp={
     #display popup help
    #each value is the name of a control on the form.
    
     Switch ($this.name) {
        "checkbox1" {$tip = "With this option Mp3Gain will normalize the files to 0dB"}
        "checkbox2" {$tip = "With this option Silence will be removed from the MP3. ffmpeg setting = silenceremove=1:0:-50dB. Which means: Remove from the begin till level is abobe 0 DB. Remove at the end everything below -50 db"}
        "checkbox3" {$tip = "With This option we will correct the filename with `"Artist - Title (SP-RIP-N)`". Files will be saved in last option sub-folder"}
     }
     $tooltip1.SetToolTip($this,$tip)
} #end ShowHelp

function checkbox_test{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    
    # Set the size of your form
    $Form = New-Object System.Windows.Forms.Form
    $Form.width = 500
    $Form.height = 270
    $Form.Text = ”Which Checks would you like to perform?”
 
    # Set the font of the text to be used within the form
    $Font = New-Object System.Drawing.Font("Times New Roman",12)
    $Form.Font = $Font
 
    # create your checkbox 
    $checkbox1 = new-object System.Windows.Forms.checkbox
    $checkbox1.Location = new-object System.Drawing.Size(30,30)
    $checkbox1.Size = new-object System.Drawing.Size(250,30)
    $checkbox1.add_MouseHover($ShowHelp)
    $checkbox1.Name = "checkbox1"
    $checkbox1.Text = "Normalize Audio Files"
    $checkbox1.Checked = $true
    $Form.Controls.Add($checkbox1)  
    
    # create your checkbox 
    $checkbox2 = new-object System.Windows.Forms.checkbox
    $checkbox2.Location = new-object System.Drawing.Size(30,60)
    $checkbox2.Size = new-object System.Drawing.Size(250,30)
    $checkbox2.add_MouseHover($ShowHelp)
    $checkbox2.Name = "checkbox2"
    $checkbox2.Text = "Remove Silence"
    $checkbox2.Checked = $true
    $Form.Controls.Add($checkbox2)  
    
    # create your checkbox 
    $checkbox3 = new-object System.Windows.Forms.checkbox
    $checkbox3.Location = new-object System.Drawing.Size(30,90)
    $checkbox3.Size = new-object System.Drawing.Size(250,30)
    $checkbox3.add_MouseHover($ShowHelp)
    $checkbox3.Name = "checkbox3"
    $checkbox3.Text = "Correct Filename"
    $checkbox3.Checked = $true
    $Form.Controls.Add($checkbox3)  
    
    # Add an OK button
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = new-object System.Drawing.Size(130,150)
    $OKButton.Size = new-object System.Drawing.Size(100,40)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$Form.Close()})
    $form.Controls.Add($OKButton)
 
    #Add a cancel button
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = new-object System.Drawing.Size(255,150)
    $CancelButton.Size = new-object System.Drawing.Size(100,40)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({$Form.Close()})
    $form.Controls.Add($CancelButton)
    
    
    ###########  This is the important piece ##############
    #                                                     #
    # Do something when the state of the checkbox changes #
    #######################################################
    #$checkbox1.Add_CheckStateChanged({
    #$OKButton.Enabled = $checkbox1.Checked })
    
    
    # Activate the form
    $Form.Add_Shown({$Form.Activate()})
    [void] $Form.ShowDialog() 
    return $checkbox1.Checked,$checkbox2.Checked,$checkbox3.Checked
}
 
#Call the function
checkbox_test