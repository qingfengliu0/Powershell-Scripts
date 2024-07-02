# Define variables
$groupName = "BCS RHL Support Services"
$groupOU = "OU=Distribution,OU=Groups,OU=Broadmead Care Society,DC=bcs,DC=local"
$groupScope = "Universal"
$groupType = "Distribution"
$aliasName = $groupName -replace '\s+'
#$usersToAdd = @("user1@yourdomain.com", "user2@yourdomain.com") # Add the appropriate users here

# Load the required modules
Import-Module ActiveDirectory
Import-Module ExchangeOnlineManagement

# Create new AD group
New-ADGroup -Name $groupName -SamAccountName $groupName -GroupCategory Distribution -GroupScope $groupScope -Path $groupOU

# Add users to the group
foreach ($user in $usersToAdd) {
    Add-ADGroupMember -Identity $groupName -Members $user
}

# Enable the distribution group in Exchange
Enable-DistributionGroup -Identity $groupName -Alias $aliasName

# Hide the distribution group from the address list
#Set-DistributionGroup -Identity $groupName -HiddenFromAddressListsEnabled $true

# Set delivery management to allow senders inside and outside the organization
Set-DistributionGroup -Identity $groupName -RequireSenderAuthenticationEnabled $false
Get-DistributionGroup -Identity $groupName


# Output a completion message
Write-Host "The distribution group has been created and configured successfully."
Write-Host "You can now log in to the Exchange Admin Center to review the settings if needed."

