# Variables
$fullName = "Nigel Living"

$ouPath = "OU=Misc. Mailboxes,OU=Users,OU=Broadmead Care Society,DC=bcs,DC=local"
$groupOUPath = "OU=Security,OU=Groups,OU=Broadmead Care Society,DC=bcs,DC=local"
$domain = "@bcs.local"
#########above requires changes#########################
$password = ""  # Replace with the password from KeePass
$groupName = "Mailbox - $fullName"
$groupDescription = "Members have full access to the $fullName shared mailbox"
$lastName = ""
$alias = $fullName -replace " ", ""
$techSupportUser = "Tech Support"
$userPrincipalName = "$alias$domain"
$samAccountName = $alias
# Log into Exchange Server (currently Dominator)

# Go to Broadmead Care Society\Users\Misc Mailboxes OU

# Create new account
New-ADUser -Name $fullName -GivenName $fullName -Surname $lastName -UserPrincipalName $userPrincipalName `
           -SamAccountName $samAccountName -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
           -PasswordNeverExpires $true -CannotChangePassword $true -Enabled $true `
           -Path $ouPath -Description "Shared Mailbox for $fullName"

# Create AD groups for giving rights to the mailbox
New-ADGroup -Name $groupName -GroupScope Universal -GroupCategory Security `
            -Path $groupOUPath -Description $groupDescription

# Add appropriate users to group
# Add-ADGroupMember -Identity $groupName -Members <list of members>

# Create a new mailbox for this user
Enable-Mailbox -Identity $userPrincipalName

# If the default email address is not correct:
# Get-Mailbox -Identity $fullName | Set-Mailbox -EmailAddresses @{Add="newemail@domain.com"; Remove="oldemail@domain.com"}

# If the AD groups require full access to the mailbox:
# Open the Exchange Management Shell
Add-MailboxPermission -Identity $fullName -User $groupName -AccessRights FullAccess -AutoMapping:$false

# If the AD group requires Send As access to the mailbox:
Add-ADPermission -Identity $fullName -User $groupName -ExtendedRights "Send As"

# Open up the Exchange Shell and run the following to set the mailbox as shared:
Set-Mailbox -Identity $fullName -Type Shared

# Refresh Exchange Console
# Icon for this mailbox (in list of mailboxes) should change
# This disables the corresponding AD account

# Add MessageCopy parameters so sent mail is copied to shared mailbox
Set-Mailbox -Identity $fullName -MessageCopyForSentAsEnabled $true -MessageCopyForSendOnBehalfEnabled $true

# If (and only if) the AD groups created above will be given access only to specific
# Outlook folders, you must create Exchange Distribution Groups for these AD groups
# New-DistributionGroup -Name "Mailbox${alias}Editor" -Alias "Mailbox${alias}Editor" -Type Security

# If distribution groups were created, hide them from GAL
# Set-DistributionGroup -Identity "Mailbox${alias}Editor" -HiddenFromAddressListsEnabled $true

# If users need access to a specific folder
# Grant Tech Support full access to the mailbox in Exchange Console
#Add-MailboxPermission -Identity $fullName -User $techSupportUser -AccessRights FullAccess

# Log into tech PC as Tech Support and add mailbox to Outlook Profile
# Grant distribution groups appropriate permissions on folders
# Remove Tech Support Full Access
#Remove-MailboxPermission -Identity $fullName -User $techSupportUser -AccessRights FullAccess

# Note: Send As permission is subject to Exchange caching
# So it may take up to 2hr for changes to this to take effect

# Members of AD groups can add the shared mailbox to their Outlook profiles
# However, don't 
