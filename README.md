# Get-UserPublicFolderPermissions
Imports a list of users from csv file, checks for public folders each user has access to, then outputs results to another csv file.

### Example usage:
.\Get-UserPublicFolderPermissions.ps1 -csvFile "c:\temp\userList.csv"

#### You can create the csv file to be used with the script by exporting from Exchange:
##### Create a .csv file with DisplayName of all user mailboxes:
Get-Mailbox | ?{$_.RecipientTypeDetails -eq 'UserMailbox'} | select DisplayName | Export-Csv c:\temp\userList.csv -NoTypeInformation

##### For larger environments, use additional filtering to split the user lists into smaller groups:
Get-Mailbox | ?{$_.RecipientTypeDetails -eq 'UserMailbox' -and $_.OrganizationalUnit -like "\*CORP\*"} | select DisplayName | Export-Csv c:\temp\userList.csv -NoTypeInformation
