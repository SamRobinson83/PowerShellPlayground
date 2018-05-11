################################################################
################################################################
##############   Active Directory Account Creation #############
############## Version 3.0 Created by Sam Robinson #############
############## includes Random Passphrase Generator#############
############## Creates on Prem Exchange Mailboxes  #############
################################################################

### Global Variables
$FirstName
$LastName
$FullName
$DisplayName
$FirstInitial
$upn
$email
$Date = Get-Date
$whoami = whoami
$CreatedBy = $whoami.Substring(12)
$Description = "Account created by $CreatedBy on $Date"
$Department
$PasswordPlainText
$CombineOU
$AccountType
$Domain = "@domain.com" #Domain
$OUPath
$Manager
$EID
$Company = "Company" #Company Name
$Title
$SamAccountName
$HomeDirectory
$HomeDirectoryRoot = "\\server\share\private\" #File Share Location for Home Directory
$HomeDriveLetter = "H:"  #Home Drive Letter
$BaseOU = "OU=Users,DC=domain,DC=com"  # The base OU
$PWPart1
$PWPart2
$PWPart3
$PWPart4
$PWPart5 = Get-Random -Minimum 100 -Maximum 999
$secureString
#Used to conncet to your exchange server to create accounts there
$UserCredential = Get-Credential
$MailboxDB = "Mailbox Database"
$ConnectionURI = "http://exchangeserver/PowerShell/"




##### Functions
function set_notes #Sets the notes to employeed ID
{
		Set-ADUser -Identity $SamAccountName -Replace @{info = $EID}
        $wshell = New-Object -ComObject WScript.Shell
        $wshell.Popup("Account $upn Created.", 0, "Done", 4096)
		$wshell = New-Object -ComObject WScript.Shell
		$wshell.Popup("Password is $PasswordPlainText", 0, "Done", 4096)
}

function create_email #Creates the account using Exchange Powershell
{
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionURI -Authentication Kerberos -Credential $UserCredential
	Import-PSSession $Session -AllowClobber
	Start-Sleep 20  #Sleep to allow the session to import fully
	
	if ($AccountType -eq "No-Email") #If person is "No Email" they are created as an Equipment mailbox because they still need to have training scheduled but do no need full Exchange functionality
	{
		New-Mailbox -UserPrincipalName $upn -Alias $SamAccountName -Database $MailboxDB -Name $DisplayName  –OrganizationalUnit $OUPath -Password $secureString -FirstName $FirstName -LastName $LastName -DisplayName $DisplayName -Equipment
		Start-Sleep 2
		Remove-PSSession $Session
		Set-ADUser $SamAccountName -Description $Description -Manager $Manager -Title $Title -EmployeeID $EID -Company $Company -FullName $FullName -DisplayName $DisplayName
		set_notes
	}
	else
	{
		New-Mailbox -UserPrincipalName $upn -Alias $SamAccountName -Database $MailboxDB -Name $DisplayName  –OrganizationalUnit $OUPath -Password $secureString -FirstName $FirstName -LastName $LastName -DisplayName $DisplayName
		Start-Sleep 2
		Remove-PSSession $Session
		Set-ADUser $SamAccountName -Description $Description -Manager $Manager -Title $Title -EmployeeID $EID -Company $Company -HomeDirectory $HomeDirectory -HomeDrive $HomeDriveLetter -FullName $FullName -DisplayName $DisplayName
		$wshell = New-Object -ComObject WScript.Shell
        $wshell.Popup("Account $upn Created", 0, "Done", 4096)
		$wshell = New-Object -ComObject WScript.Shell
		$wshell.Popup("Password is $PasswordPlainText and copied to clipboard", 0, "Done", 4096)
	}
	
}

function create_passphrase #Creates a random passphrase
{
	
	function pw
	{
		$tmp1 = Get-Content "\server\share\Passphrase\Part1.txt"  #Part 1 of the password, word file with capitalized words
		$tmp2 = Get-Content "\\server\share\Passphrase\Part2.txt" #Part 2 of the password, word file
		$tmp3 = Get-Content "\\server\share\Passphrase\Part3.txt" #Part 3 of the password, punctuation
		$tmp4 = Get-Content "\\server\share\Passphrase\Part4.txt" #Part 4 of the passsword, word file
		
		$MLPart1 = $tmp1.Length -1
		$MLPart2 = $tmp2.Length -1
		$MLPart3 = $tmp3.Length -1
		$MLPart4 = $tmp4.Length -1
		
		$myran1 = Get-Random -Minimum 0 -Maximum $MLPart1
		$myran2 = Get-Random -Minimum 0 -Maximum $MLPart2
		$myran3 = Get-Random -Minimum 0 -Maximum $MLPart3
		$myran4 = Get-Random -Minimum 0 -Maximum $MLPart4
		
		$PWPart1 = $tmp1[$myran1]
		$PWPart2 = $tmp2[$myran2]
		$PWPart3 = $tmp3[$myran3]
		$PWPart4 = $tmp4[$myran4]
		
		$PasswordPlainText = $PWPart1 + $PWPart3 + $PWPart2 + $PWPart3 + $PWPart4 + $PWPart3 + $PWPart5
		
		if ($PasswordPlainText.Length -le 20)  #Due to old software on our network passwords cannot be over 20 characters
		{
			$secureString = ConvertTo-SecureString $PasswordPlainText -AsPlainText -Force
			$PasswordPlainText | clip.exe #Copies the password to your clipboard
			create_email
		}
		else  #If password over 20 characters, it recreates it
		{
			pw
		}
	}
	
    pw
}    
function create_account  #Create Account function
{
			Import-Module ActiveDirectory
            $FirstName = $FirstNameBox.Text
            $LastName = $LastNameBox.Text
            $FullName = $FirstName + " " + $LastName
			$DisplayName = $LastName + ", " + $FirstName
            $AccountType = $AccountTypeBox.Text
            $Manager = $ManagerBox.Text
            $Department = $DepartmentBox.Text
            $SamAccountName = $FirstName.ToLower() + "." + $LastName.ToLower()
            $upn = $SamAccountName + $Domain
            $email = $upn
            $EID = $EIDBox.Text
            $Title = $TitleBox.Text
			$HomeDirectory = $HomeDirectoryRoot + $SamAccountName
            ### Specifies which OU Accounts need to go in based off the Department box
            if ( $Department -eq "Department 1")
			{
                $OUPath = "OU=Department 1," + $BaseOU
			}
			if ( $Department -eq "Department 2")
			{
                $OUPath = "OU=Department 2," + $BaseOU
			}
            
			
			$User = Get-ADUser -LDAPFilter "(sAMAccountName=$SamAccountName)"
            
            #Checks to see if the user already exists
			if ($User -eq $Null)
			{
				create_passphrase
			}
			else
			{
				$wshell = New-Object -ComObject WScript.Shell
				$wshell.Popup("Username Already Exists", 0, "Already Exists", 4096)
			}
			
			
			
}


############ Build the GUI

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form
$objForm.Text = "Create Accounts"
$objForm.Size = New-Object System.Drawing.Size(315,315) 
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
				
		$TitleBox = New-Object System.Windows.Forms.TextBox
		$TitleBox.Location = New-Object System.Drawing.Size(10,212)
		$TitleBox.Size = New-Object System.Drawing.Size(100,100)
		$objForm.Controls.Add($TitleBox)
		
		$TitleLabel = New-Object System.Windows.Forms.Label
		$TitleLabel.Location = New-Object System.Drawing.Size(10,191)
		$TitleLabel.Size = New-Object System.Drawing.Size(100,17)
		$TitleLabel.Text = "Title:"
		$objForm.Controls.Add($TitleLabel)
		
		#### Right Column
		
		$ManagerBox = New-Object System.Windows.Forms.TextBox
		$ManagerBox.Location = New-Object System.Drawing.Size(150,80)
		$ManagerBox.Size = New-Object System.Drawing.Size(100,100)
		$objForm.Controls.Add($ManagerBox)
		
		$ManagerLabel = New-Object System.Windows.Forms.Label
		$ManagerLabel.Location = New-Object System.Drawing.Size(150,63)
		$ManagerLabel.Size = New-Object System.Drawing.Size(100,17)
		$ManagerLabel.Text = "Manager:"
		$objForm.Controls.Add($ManagerLabel)
		
		$DepartmentBox = New-Object System.Windows.Forms.ComboBox
		$DepartmentBox.Location = New-Object System.Drawing.Size(150,125)
		$DepartmentBox.Size = New-Object System.Drawing.Size(140,17)
		[void]$DepartmentBox.Items.Add("Department 1")
		[void]$DepartmentBox.Items.Add("Department 2")
		$objForm.Controls.Add($DepartmentBox)
		
		$DepartmentBoxLabel = New-Object System.Windows.Forms.Label
		$DepartmentBoxLabel.Location = New-Object System.Drawing.Size(150,109)
		$DepartmentBoxLabel.Size = New-Object System.Drawing.Size(110,17)
		$DepartmentBoxLabel.Text = "Choose Department:"
		$objForm.Controls.Add($DepartmentBoxLabel)
		
		$AccountTypeBox = New-Object System.Windows.Forms.ComboBox
		$AccountTypeBox.Location = New-Object System.Drawing.Size(150,167)
		$AccountTypeBox.Size = New-Object System.Drawing.Size(100,17)
		[void]$AccountTypeBox.Items.Add("Email")
		[void]$AccountTypeBox.Items.Add("No-Email")
		$objForm.Controls.Add($AccountTypeBox)
		
		$AccountTypeBoxLabel = New-Object System.Windows.Forms.Label
		$AccountTypeBoxLabel.Location = New-Object System.Drawing.Size(150,151)
		$AccountTypeBoxLabel.Size = New-Object System.Drawing.Size(110,17)
		$AccountTypeBoxLabel.Text = "Account Type:"
		$objForm.Controls.Add($AccountTypeBoxLabel)
		
		$EIDBox = New-Object System.Windows.Forms.TextBox
		$EIDBox.Location = New-Object System.Drawing.Size(150,212)
		$EIDBox.Size = New-Object System.Drawing.Size(100,100)
		$objForm.Controls.Add($EIDBox)
		
		$EIDLabel = New-Object System.Windows.Forms.Label
		$EIDLabel.Location = New-Object System.Drawing.Size(150,191)
		$EIDLabel.Size = New-Object System.Drawing.Size(100,17)
		$EIDLabel.Text = "Employee ID:"
		$objForm.Controls.Add($EIDLabel)
		
		
		
		
		#### Account Creation Button
		
		$CreateAccountButton = New-Object System.Windows.Forms.Button
		$CreateAccountButton.Location = New-Object System.Drawing.Size(75,245)
		$CreateAccountButton.Size = New-Object System.Drawing.Size(150,23)
		$CreateAccountButton.Text = "Create Account"
		$CreateAccountButton.Add_Click({ create_account })
		$objForm.Controls.Add($CreateAccountButton)
		
		
		
		$objForm.Topmost = $True
		
		$objForm.Add_Shown({$objForm.Activate()})
		[void] $objForm.ShowDialog()

$x