# This script creates basic MSOL user acounts that don't require any special rules.

# Get the folder and select the most recent csv file for import
$filename = (gci "C:\PathToFile\new-users" | sort LastWriteTime | select fullname -last 1) 

# Setting the filename from the folder for use in our user creation function below.
$users = import-csv $filename.fullname

# Sets licenses that are NOT needed for the users. If you want users to be able to access all O365 options, you can delete this and take off the $LO variable in the $users function below.
$LO = New-MsolLicenseOptions -AccountSkuId "ExampleCompany:STANDARDWOFFPACK_IW_FACULTY" -DisabledPlans YAMMER_EDU, KAIZALA_O365_P2

# create new o365 user accounts #
write-verbose "Process starting, will take a few minutes..."

# Creates new users based on the header row in your csv import file. Change the $_. variables to whatever your header rows have listed.
$users | ForEach-object {
    New-MsolUser -UserPrincipalName $_.Email -password $_.password -FirstName $_.FirstName -LastName $_.LastName -DisplayName $_.DisplayName -UsageLocation $_.Country -LicenseAssignment "ExampleCompany:STANDARDWOFFPACK_IW_FACULTY" -LicenseOptions $LO 
} 


write-verbose "Done creating new accounts."

