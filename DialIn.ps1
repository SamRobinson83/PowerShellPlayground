#
 # @CreateTime: Jun 13, 2018 10:33 AM
 # @Author: Sam Robinson
 # @Contact: undefined
# @Last Modified By: Sam Robinson
# @Last Modified Time: Jun 13, 2018 12:58 PM
 # @Description: GUI to quickly modify the "Dial In" Attribute in Active Directory 
#
Import-Module ActiveDirectory

### Variables
$Username




##function
function allowdialin {
    $Username = $UserToChangeBox.Text
    
    set-aduser $Username -replace @{msnpallowdialin=$true}

    $wshell = New-Object -ComObject WScript.Shell
    $wshell.Popup("$Username Dial In Enabled", 0, "Done", 4096)
}

function disabledialin {
    $Username = $UserToChangeBox.Text

    set-aduser $Username -replace @{msnpallowdialin=$false}

    $wshell = New-Object -ComObject WScript.Shell
    $wshell.Popup("$Username Dial In Disabled", 0, "Done", 4096)
    
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

$UsernameLabel = New-Object System.Windows.Forms.Label
$UsernameLabel.Location = New-Object System.Drawing.Size(10,10)
$UsernameLabel.Size = New-Object System.Drawing.Size(300,25)
$UsernameLabel.Text = "Username to Change:"
$objForm.Controls.Add($UsernameLabel)

$UserToChangeBox = New-Object System.Windows.Forms.TextBox
$UserToChangeBox.Location = New-Object System.Drawing.Size(10,35)
$UserToChangeBox.Size = New-Object System.Drawing.Size(100,50)
$objForm.Controls.Add($UserToChangeBox)


$AllowDialInButton = New-Object System.Windows.Forms.Button
$AllowDialInButton.Location = New-Object System.Drawing.Size(10,75)
$AllowDialInButton.Size = New-Object System.Drawing.Size(150,100)
$AllowDialInButton.Text = "Allow Dial In"
$AllowDialInButton.Add_Click({ allowdialin })
$objForm.Controls.Add($AllowDialInButton)

$DisableDialInButton = New-Object System.Windows.Forms.Button
$DisableDialInButton.Location = New-Object System.Drawing.Size(200,75)
$DisableDialInButton.Size = New-Object System.Drawing.Size(150,100)
$DisableDialInButton.Text = "Disable Dial In"
$DisableDialInButton.Add_Click({ disabledialin })
$objForm.Controls.Add($DisableDialInButton)




$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

$x