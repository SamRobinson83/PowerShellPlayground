####################################################################################
####################################################################################
################### Account Lockout Alert ##########################################
################### By Sam Robinson       ##########################################
####################################################################################

## Use this script on Domain Controllers, Run as a Scheduled Task Triggered by Event 4740 in the Security Event Log




Import-Module ActiveDirectory
#Grabs the Event from the Event Log
$AccountLockOutEvent = Get-EventLog -LogName "Security" -InstanceID 4740 -Newest 1
#Extracts out the Username and computer from the log
$LockedAccount = $($AccountLockOutEvent.ReplacementStrings[0])
$AccountLockOutEventTime = $AccountLockOutEvent.TimeGenerated
$AccountLockOutEventMessage = $AccountLockOutEvent.Message
$Date = Get-Date -Format "yyyyMMdd-HH-mm"
$AccountLockOutEventMessage | Out-File "C:\LockedAccounts\$LockedAccount.$Date.txt"
Start-Sleep 2 #Brief pause
$Last1 = Get-Content "C:\LockedAccounts\$LockedAccount.$Date.txt" -last 1
$Trimmed1 = $Last1.Trim()
$Computer = $Trimmed1.Substring(22)

#Construct your email and send it
$messageParameters = @{ 
	Subject = "Account Locked Out on $Computer : $LockedAccount" 
	Body = "Account $LockedAccount was locked out on $AccountLockOutEventTime.`n`nEvent Details:`n`n$AccountLockOutEventMessage"
	From = "lockout@domain.com" 
	To = "ITDepartment@domain.com" 
	SmtpServer = "mailserver.domain.com" 
	}
Send-MailMessage @messageParameters
