#
 # @CreateTime: Jun 13, 2018 10:33 AM
 # @Author: Sam Robinson
 # @Contact: undefined
# @Last Modified By: Sam Robinson
# @Last Modified Time: Jun 13, 2018 11:26 AM
 # @Description: GUI to quickly modify the "Dial In" Attribute in Active Directory 
#
Import-Module ActiveDirectory

### Variables
$Username




##function
function allowdialin {

    
}

function disabledialin {
    $Username = $UserToChangeBox.Text

    
    
}





###### Build GUI
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form
$objForm.Text = "Change Dial In Status"
$objForm.Size = New-Object System.Drawing.Size(400,250) 
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$x=$objTextBox.Text;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$ProgramLabel = New-Object System.Windows.Forms.Label
$ProgramLabel.Location = New-Object System.Drawing.Size(10,10)
$ProgramLabel.Size = New-Object System.Drawing.Size(300,25)
$ProgramLabel.Text = "Dial In Status Change"
$objForm.Controls.Add($ProgramLabel)
$UserToChangeBox = New-Object System.Windows.Forms.TextBox
$UserToChangeBox.Location = New-Object System.Drawing.Size(10,80)
$UserToChangeBox.Size = New-Object System.Drawing.Size(100,100)
$objForm.Controls.Add($UserToChangeBox)

$UserToChangeLabel = New-Object System.Windows.Forms.Label
$UserToChangeLabel.Location = New-Object System.Drawing.Size(10,63)
$UserToChangeLabel.Size = New-Object System.Drawing.Size(100,17)
$UserToChangeLabel.Text = "User To Change:"
$objForm.Controls.Add($UserToChangeLabel)

$EjectButton = New-Object System.Windows.Forms.Button
$EjectButton.Location = New-Object System.Drawing.Size(10,160)
$EjectButton.Size = New-Object System.Drawing.Size(150,23)
$EjectButton.Text = "Allow Dial In"
$EjectButton.Add_Click({ allowdialin })
$objForm.Controls.Add($EjectButton)

$DisableDialInButton = New-Object System.Windows.Forms.Button
$DisableDialInButton.Location = New-Object System.Drawing.Size(200,160)
$DisableDialInButton.Size = New-Object System.Drawing.Size(150,23)
$DisableDialInButton.Text = "Disable Dial In"
$DisableDialInButton.Add_Click({ disabledialin })
$objForm.Controls.Add($DisableDialInButton)




$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

$x