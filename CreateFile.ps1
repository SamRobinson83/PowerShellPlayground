#################################################################
#################################################################
#############  File Creation ####################################
######## By Sam Robinson     ####################################
#################################################################


#Create a file in given directory that is a junk test file

$Directory = Read-Host "Enter directory where you want to create the file:"
$Filename = Read-Host "Enter name of the file you want to create:"
$SizeMB = Read-Host "Enter the size of the file you want to create in MB:"
$SizeBytes = $SizeMB * 1048576
$FullPath = "$Directory$Filename"
fsutil file createnew $FullPath $SizeBytes