#######################################################
#######################################################
#############  OU Shutdown ############################
########## By Sam Robinson ############################
#######################################################

# Grabs all computer names in the specified OU and then shuts them down remotely, great for scheduled shutdowns


$Name = Get-ADComputer -Filter * -SearchBase "DC=domain,DC=com" | Select -ExpandProperty Name #Specifies the directory to search, use the FDQN

foreach ($N in $Name) {
	$ComputerName = $N
	Invoke-Command -ComputerName $ComputerName -ScriptBlock { Shutdown -s -f }
}