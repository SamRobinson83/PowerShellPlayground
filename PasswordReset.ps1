#################################################################################
#################################################################################
################## Reset Account Passwords ######################################
##################     By Sam Robinson     ######################################
#################################################################################
#################################################################################

############### Global Variables
$FirstName
$LastName
$FullName




############################ Functions

function resetpassword
{
    Import-Module ActiveDirectory
    $FirstName = $FirstNameBox.Text
    $LastName = $LastNameBox.Text
    $secureString = = ConvertTo-SecureString $PasswordBox.Text -AsPlainText -Force
    $FullName = $FirstName + " " + $LastName
    Set-ADAccountPassword -Identity $FullName -NewPassword $secureString
    $wshell = New-Object -ComObject WScript.Shell
    $wshell.Popup("Operation Completed", 0, "Done", 0x1)
}





############ Build the GUI

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form
$objForm.Text = "Create Accounts"
$objForm.Size = New-Object System.Drawing.Size(300,300) 
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
   {$x=$objTextBox.Text;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$ProgramLabel = New-Object System.Windows.Forms.Label
$ProgramLabel.Location = New-Object System.Drawing.Size(10,10)
$ProgramLabel.Size = New-Object System.Drawing.Size(300,25)
$ProgramLabel.Text = "Create Accounts"
$objForm.Controls.Add($ProgramLabel)

###### Left Column

$FirstNameBox = New-Object System.Windows.Forms.TextBox
$FirstNameBox.Location = New-Object System.Drawing.Size(10,80)
$FirstNameBox.Size = New-Object System.Drawing.Size(100,100)
$objForm.Controls.Add($FirstNameBox)

$FirstNameLabel = New-Object System.Windows.Forms.Label
$FirstNameLabel.Location = New-Object System.Drawing.Size(10,63)
$FirstNameLabel.Size = New-Object System.Drawing.Size(100,17)
$FirstNameLabel.Text = "First Name:"
$objForm.Controls.Add($FirstNameLabel)

$LastNameBox = New-Object System.Windows.Forms.TextBox
$LastNameBox.Location = New-Object System.Drawing.Size(10,125)
$LastNameBox.Size = New-Object System.Drawing.Size(100,100)
$objForm.Controls.Add($LastNameBox)

$LastNameLabel = New-Object System.Windows.Forms.Label
$LastNameLabel.Location = New-Object System.Drawing.Size(10,109)
$LastNameLabel.Size = New-Object System.Drawing.Size(100,17)
$LastNameLabel.Text = "Last Name:"
$objForm.Controls.Add($LastNameLabel)

$PasswordBox = New-Object System.Windows.Forms.TextBox
$PasswordBox.Location = New-Object System.Drawing.Size(10,167)
$PasswordBox.Size = New-Object System.Drawing.Size(100,100)
$objForm.Controls.Add($PasswordBox)

$PasswordLabel = New-Object System.Windows.Forms.Label
$PasswordLabel.Location = New-Object System.Drawing.Size(10,150)
$PasswordLabel.Size = New-Object System.Drawing.Size(100,17)
$PasswordLabel.Text = "Password:"
$objForm.Controls.Add($PasswordLabel)



#### Account Creation Button

$CreateAccountButton = New-Object System.Windows.Forms.Button
$CreateAccountButton.Location = New-Object System.Drawing.Size(100,205)
$CreateAccountButton.Size = New-Object System.Drawing.Size(150,23)
$CreateAccountButton.Text = "Reset Password"
$CreateAccountButton.Add_Click({ resetpassword })
$objForm.Controls.Add($CreateAccountButton)



$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

$x