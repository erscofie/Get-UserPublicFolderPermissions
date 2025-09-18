##################################################
# Imports a list of users from csv file, checks for public folders each user has access to, then outputs results to another csv file.
# 
#	Example usage:
# 	.\Get-UserPublicFolderPermissions.ps1 -csvFile "c:\temp\userList.csv"
# 
# 	You can create the csv file to be used with the script by exporting from Exchange:
#	Get-Mailbox | ?{$_.RecipientTypeDetails -eq 'UserMailbox'} | select DisplayName | Export-Csv c:\temp\userList.csv -NoTypeInformation
#   
#	
##################################################

[CmdletBinding()]
Param(
   [Parameter(Mandatory=$True)]
   [string]$csvFile
 )

# Import users from CSV
$users = Import-Csv -Path $csvFile

# Get all public folders recursively
$publicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited

# Initialize an array to store results
$results = @()

foreach ($user in $users) {
    $DisplayName = $user.DisplayName
    foreach ($folder in $publicFolders) {
        try {
            $permissions = Get-PublicFolderClientPermission -Identity $folder.Identity -ErrorAction Stop
            foreach ($perm in $permissions) {
                if ($perm.User.DisplayName -eq $DisplayName) {
                    $results += [PSCustomObject]@{
                        DisplayName = $DisplayName
                        FolderPath   = $folder.Identity
                        AccessRights = $perm.AccessRights -join ", "
                    }
                }
            }
        } catch {
            Write-Warning "Failed to get permissions for folder: $($folder.Identity)"
        }
    }
}

# Export results to .csv

$timestamp = Get-Date -format yy-MM-ddTHH-mm-ss
$folderPath = Split-Path -Path $csvFile
$outputFile = $folderPath + '\folderPermissions_' + $timestamp + '.csv'
$results | Export-Csv $outputFile -NoTypeInformation
[console]::WriteLine("Output saved to $outputFile")