# Powershell script that creates Teams based on AD groups

# Load the CSV

# Iterate through Groups in the OU
# For each group:
#   If Group does not exist in O365
#       Create group in O365
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

# Variables

$ADGroupsOU = ""


#  Load Modules
Import-Module ActiveDirectory
Import-Module -UseWindowsPowerShell AzureAD
#Connect-AzureAD
#Connect-MicrosoftTeams
#Connect-ExchangeOnline

# Import CSV
$CSV = import-csv .\TeamsGroupSync.csv

$ADGroups = Get-ADGroup -SearchBase $ADGroupsOU -filter {GroupCategory -eq "Security"}

foreach ($ADGroup in $ADGroups) {
    # Extract the O365 group name from the AD group name
    $O365GroupName = $ADGroup.Name -replace "Teams " 
    $O365GroupExists = Get-AzureADGroup -SearchString $O365GroupName
    
    # Check if group exists in Azure
    if ($O365GroupExists -eq $null){
        # Group does not exist in Azure
        Write-Output "Creating Office 365 Group"
        
        # Create email string - Strip out special charactars
        $O365GroupEmail = $O365GroupName -replace '[#?.\\\/\ ]','-' -replace "'" -replace '"' -replace '[\(\)\[\]\{\}]'

        # Create O365 Group
        New-AzureADGroup -DisplayName $O365GroupName -MailEnabled $true -MailNickName $O365GroupEmail -Description $ADGroup.Description
        
        $O365Group = Get-AzureADGroup -SearchString $O365GroupName
        
        # Disable Welcome Message
        Set-UnifiedGroup -Identity $O365GroupName -UnifiedGroupWelcomeMessageEnable:$false
        
        # Hide from GAL
        Set-UnifiedGroup -Identity $O365GroupName -HiddenFromAddressListsEnabled $true
        
        # Get list of members of AD group
        # For each member:
        #   Get upn
        #   Find user in Azure AD
        #   Add user to O365 group as member
        #   Add user to O365 AD group as owner

        $Users = Get-ADGroupMember -Identity $ADGroup.Name
        foreach ($User in $Users) {
            $GUID = $User.objectGUID
            $UPN = (Get-ADUser -Identity $GUID).UserPrincipalName
            $AzureADUser = Get-AzureADUser -SearchString $UPN
            Add-AzureADGroupMember -ObjectId $O365Group.ObjectID -RefObjectId $AzureADUser.ObjectID
            Add-AzureADGroupOwner - ObjectId $O365Group.ObjectID -RefObjectId $AzureADUser.ObjectID

        }
        
        
        # Create Basic locked down Team from O365 Group
        #new-team -GroupID $O365Group.ObjectID  -AllowGiphy $false -AllowStickersAndMemes $false -AllowCustomMemes $false -AllowGuestCreateUpdateChannels $false -AllowGuestDeleteChannels $false -AllowCreateUpdateChannels $false -AllowDeleteChannels $false -AllowAddRemoveApps $false -AllowCreateUpdateRemoveTabs $false -AllowCreateUpdateRemoveConnectors $false -ShowInTeamsSearchAndSuggestions $false -AllowOwnerDeleteMessages $true 
        
        # Add new AD group/O365 group pair to CSV
        $newRow = New-Object PSObject -Property @{ ADGroup = $ADGroup.Name ; O365Group = $O365GroupName }
        $CSV += $newRow
        
    }else {
        # Group exists in Azure
        Write-Output "Group $O365GroupExists already exists in AzureAD"
    }


}


# Sort CSV data and write
$CSV | Sort-Object -Property ADGroup | Export-Csv -path .\test.csv -NoTypeInformation -UseQuotes Never
