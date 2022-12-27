<#
.SYNOPSIS
   Use a scheduled task to call this script upon event trigger.
.DESCRIPTION
   Grabs the info from the Event 8215 and if the user has exceeded 5 file blocks, calls the Network Killer Script that disables their network adapters
.Author
   Sam Robinson
#>

####### Global Variables
$Username
$date = Get-Date -Format "yyyyMMdd"
$CompList = New-Object System.Collections.ArrayList
$AdminCredUsername = "Admin"
$AdminCredPassword =cat C:\Scripts\Vold\Admin.txt | ConvertTo-SecureString
$AdminCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AdminCredUsername,$AdminCredPassword
$Prefix = "po*"


###### Functions

#### Petrificus Totalus triggers when a user trips the Voldemort alert 5 times in a day, it disables their network adapters to prevent further spreading
function petrificus_totalus {
	
	### Disable machine network card for machines user is logged into
	function disablenic { 
		$computers = Get-ADComputer -Filter {Enabled -eq 'true' -and SamAccountName -like $Prefix}
		$CompCount = $Computers.Count
		
		#Start main foreach loop, search processes on all computers
		foreach ($comp in $computers){
			$Computer = $comp.Name
			$Reply = $null
			$Reply = test-connection $Computer -count 1 -quiet
			if($Reply -eq 'True'){
				if($Computer -eq $env:COMPUTERNAME){
					#Get explorer.exe processes without credentials parameter if the query is executed on the localhost
					$proc = gwmi win32_process -ErrorAction SilentlyContinue -computer $Computer -Filter "Name = 'explorer.exe'"
				}
				else{
					#Get explorer.exe processes with credentials for remote hosts
					$proc = gwmi win32_process -ErrorAction SilentlyContinue -Credential $AdminCred -computer $Computer -Filter "Name = 'explorer.exe'"
				}			
					#If $proc is empty return msg else search collection of processes for username
				if([string]::IsNullOrEmpty($proc)){
					if (Test-Path C:\Scripts\BlockUsers\$Username.Failed.txt)
					{
						Add-Content C:\Scripts\BlockUsers\$Username.Failed.txt "Failed to check $Computer!"
					}
					else
					{
						New-Item C:\Scripts\BlockUsers\$Username.Failed.txt 
						Add-Content "Failed to check $Computer!"
					}	
				}
				else{	
					$progress++			
					ForEach ($p in $proc) {				
						$temp = ($p.GetOwner()).User
						if ($temp -eq $Username){
							write-host "$Username is logged on $Computer"
							Add-Content C:\Temp\$Username.$date.txt "$Computer"
							$CompList.Add($Computer)
							([wmiclass]"\\$Computer\root\cimv2:win32_Process").create("cmd /c wmic path win32_networkadapter where ""NetConnectionID LIKE 'Local%'"" call disable")
							([wmiclass]"\\$Computer\root\cimv2:win32_Process").create("cmd /c wmic path win32_networkadapter where ""NetConnectionID LIKE 'Ethernet%'"" call disable")
							([wmiclass]"\\$Computer\root\cimv2:win32_Process").create("cmd /c wmic path win32_networkadapter where ""NetConnectionID LIKE 'Wireless%'"" call disable")
							([wmiclass]"\\$Computer\root\cimv2:win32_Process").create("cmd /c wmic path win32_networkadapter where ""NetConnectionID LIKE 'Wi-Fi%'"" call disable")
						}
					}
				}	
			}
		}	
	}
	disablenic
	
	### Lock User Account
	function lockaccount {
		Disable-ADAccount -Identity $Username
	}
	lockaccount
}

### VoldCheck will grab the last event from the event log and pull the username from the message, if user has triggered the warning 5 times in a day will disable their network adapters
function voldcheck{
	$VoldBlockEvent = Get-EventLog -LogName "Application" -Source "SRMSVC" -Newest 1
	$UsernameTemp = $VoldBlockEvent.Message
	$UsernameTemp1 = $UsernameTemp.Substring(17)
	$Username = $UsernameTemp1.Substring(0, $UsernameTemp1.IndexOf(' '))
	
	if (Test-Path "C:\Scripts\BlockUsers\$Username.$date.txt")
	{
		Add-Content C:\Scripts\BlockUsers\$Username.$date.txt "1"
		$CountTemp = Get-Content C:\Scripts\BlockUsers\$Username.$date.txt | Measure-Object
		$Count = $CountTemp.Count
		if ($Count -gt 5)
		{
			$messageParameters = @{ 
				Subject = "$Username causing Voldemort... executing petrificus totalus" 
				Body = "$Username causing Voldemort Infection"
				From = "from@domain.com" 
				To = "to@vdomain.com"
				SmtpServer = "smtp.server.com" 
			} 
			Send-MailMessage @messageParameters
			### Calls the Network Kill Script
			petrificus_totalus
		}
		else {
			$messageParameters = @{ 
				Subject = "$Username changing file extensions to blocked extensions" 
				Body = "$Username changing file extensions to blocked extensions"
				From = "from@domain.com" 
				To = "to@domain.com"
				SmtpServer = "smtp.server.com"
			} 
			Send-MailMessage @messageParameters
		}
	}
	else
	{
		New-Item C:\Scripts\BlockUsers\$Username.$date.txt -type file
		Add-Content C:\Scripts\BlockUsers\$Username.$date.txt "1"
	}
}
voldcheck