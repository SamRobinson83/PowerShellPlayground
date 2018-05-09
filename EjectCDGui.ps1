###################################################################
###################################################################
#############  Eject CD GUI #######################################
############# By Sam Robinson #####################################
###################################################################

############ Variables
$ComputerName



######### EJECT Function
function ejectcd{
	$ComputerName = $ComputerToEjectBox.Text
	Invoke-Command -ComputerName $ComputerName -ScriptBlock {
	
	$sh = New-Object -ComObject "Shell.Application"
	$sh.Namespace(17).Items() | 
		Where-Object { $_.Type -eq "CD Drive" } | 
			foreach { $_.InvokeVerb("Eject") }
	}
	$wshell = New-Object -ComObject WScript.Shell
    $wshell.Popup("CD Ejected on $ComputerName", 0, "CD Ejected on $ComputerName", 4096)  
 }
 ###### Build GUI
 [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form
$objForm.Text = "Eject CD"
$objForm.Size = New-Object System.Drawing.Size(300,250) 
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$x=$objTextBox.Text;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$ProgramLabel = New-Object System.Windows.Forms.Label
$ProgramLabel.Location = New-Object System.Drawing.Size(10,10)
$ProgramLabel.Size = New-Object System.Drawing.Size(300,25)
$ProgramLabel.Text = "Eject CD"
$objForm.Controls.Add($ProgramLabel)

$ComputerToEjectBox = New-Object System.Windows.Forms.TextBox
$ComputerToEjectBox.Location = New-Object System.Drawing.Size(10,80)
$ComputerToEjectBox.Size = New-Object System.Drawing.Size(100,100)
$objForm.Controls.Add($ComputerToEjectBox)

$ComputerToEjectLabel = New-Object System.Windows.Forms.Label
$ComputerToEjectLabel.Location = New-Object System.Drawing.Size(10,63)
$ComputerToEjectLabel.Size = New-Object System.Drawing.Size(100,17)
$ComputerToEjectLabel.Text = "Computer to Eject CD:"
$objForm.Controls.Add($ComputerToEjectLabel)

$EjectButton = New-Object System.Windows.Forms.Button
$EjectButton.Location = New-Object System.Drawing.Size(10,160)
$EjectButton.Size = New-Object System.Drawing.Size(150,23)
$EjectButton.Text = "Eject CD"
$EjectButton.Add_Click({ ejectcd })
$objForm.Controls.Add($EjectButton)



$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

$x