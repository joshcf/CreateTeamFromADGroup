# Powershell script that creates Teams based on AD groups

# Load the CSV

# Iterate through Groups in the OU
# For each group:
#   If Group does not exist in O365
#       Create group in O365
#       Add Salamander 365 service principal as owner (e43dbdcd-37c1-4d43-a4d2-36a00903e568)
#       Set O365 group to be hidden from address list
#       Turn off welcome email
#       Add members from AD group as members
#       Add members from AD group as owners
#       Create Team
#       Add AD group/O365 group paring to CSV
#   If group does exist
#       do nothing
# Sort CSV alphabetically
# Write out CSV

#  Load Modules
Import-Module ActiveDirectory
Import-Module -UseWindowsPowerShell AzureAD
#Connect-AzureAD
#Connect-MicrosoftTeams

# Import CSV
$CSV = import-csv .\TeamsGroupSync.csv

$ADGroups = Get-ADGroup -SearchBase "OU=Teams Groups,OU=Groups,OU=_Bedales,DC=bedales,DC=org,DC=uk" -filter {GroupCategory -eq "Security"}

foreach ($ADGroup in $ADGroups) {
    # Extract the O365 group name from the AD group name
    $O365GroupName = $ADGroup.Name -replace "Teams " 
    #Write-Output $O365GroupName
    $O365GroupExists = Get-AzureADGroup -SearchString $O365GroupName

    if ($O365GroupExists -eq $null){
        Write-Output "Nothing Here"
    }else {
        Write-Output "Here is the group:"
        Write-Output $O365GroupExists.DisplayName
    }


}


# Sort CSV data and write
$CSV | Sort-Object -Property ADGroup | Export-Csv -path .\test.csv -NoTypeInformation -UseQuotes Never
