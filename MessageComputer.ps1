################################################
################################################
########### Message Computer ###################
########## By Sam Robinson   ###################
################################################

#Send a message to a computer that pops up on the users screen



##### Declare Variables #########
$MachineName
$Message
$TimeToDisplay = 300;

################ Send Message Function #####
function send_message{
  $Message = $MessageBox.Text
  $MachineName = $MachineNameBox.Text
  msg /server:$MachineName * /time:$TimeToDisplay "$Message"
  $wshell = New-Object -ComObject WScript.Shell
  $wshell.Popup("$Message Sent", 0, "Done", 4096)
}

######## Build the Gui
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(470,200)

$MainFont = New-Object System.Drawing.Font("Times New Roman",18,[System.Drawing.FontStyle]::Bold)
$Label = New-Object System.Windows.Forms.Label
$Label.Location = New-Object System.Drawing.Size(5,5)
$Label.AutoSize = $True
$Label.Text = "Send Message to Computer"
$Label.Font = $MainFont
$Form.Controls.Add($Label)

###### Text Entry Boxes


$MachineNameBox = New-Object System.Windows.Forms.TextBox
$MachineNameBox.Location = New-Object System.Drawing.Size(15,80)
$MachineNameBox.Size = New-Object System.Drawing.Size(150,150)
$Form.Controls.Add($MachineNameBox)
		
$MachineNameLabel = New-Object System.Windows.Forms.Label
$MachineNameLabel.Location = New-Object System.Drawing.Size(10,55)
$MachineNameLabel.AutoSize = $True
$MachineNameLabel.Font = $Font
$MachineNameLabel.Text = "Machine Name:"
$Form.Controls.Add($MachineNameLabel)



$MessageBox = New-Object System.Windows.Forms.TextBox
$MessageBox.Location = New-Object System.Drawing.Size(255,80)
$MessageBox.Size = New-Object System.Drawing.Size(150,150)
$Form.Controls.Add($MessageBox)
		
$MessageLabel = New-Object System.Windows.Forms.Label
$MessageLabel.Location = New-Object System.Drawing.Size(250,55)
$MessageLabel.AutoSize = $True
$MessageLabel.Font = $Font
$MessageLabel.Text = "Message to Display:"
$Form.Controls.Add($MessageLabel)




##### Buttons

$Button = New-Object System.Windows.Forms.Button 
$Button.Location = New-Object System.Drawing.Size(145,110) 
$Button.Size = New-Object System.Drawing.Size(140,40) 
$Button.Text = "Send Message" 
$Button.Add_Click({send_message}) 
$Form.Controls.Add($Button)

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()