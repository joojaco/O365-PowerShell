# This script gets all users from a specific domain and blocks their sign-in access.

# Variable for the filename.
$filename = "C:\PathToExportFile\blocksignin.csv"


# This gets all users from the specified domain name, selects their userprincipalname (usually email address), and exports them to the filename you specified.
get-msoluser  -domainname example.com -all | select userprincipalname | export-csv $filename

# This imports the csv file that was created and loops through and blocks signin access for each user on the list.
Import-Csv $filename | ForEach-Object {
$upn = $_."UserPrincipalName"
Set-MsolUser -UserPrincipalName $upn -BlockCredential $true
}