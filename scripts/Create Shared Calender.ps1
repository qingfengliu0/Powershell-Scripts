# Import necessary modules
Import-Module ActiveDirectory
Import-Module ExchangeOnlineManagement

# Define variables
$calenderName = "Family Activities"  # Replace with the actual room name
################################## Section above require change ####################
$EditorgroupName = "C - $calenderName Editor"
$ReviewergroupName = "C - $calenderName Reviewer"
$EditorgroupDescription = "Members have Editor access to $calenderName calendar"
$ReivewergroupDescription = "Members have Reviewer access to $calenderName calendar"
$EditorAlias = "C-$($calenderName)Editor" -replace '\s+'
$reviewerAlias = "C-$($calenderName)Reviewer" -replace '\s+'
$userFirstName = "C - $calenderName"
$CalenderAlias = $calenderName -replace '\s+'
$userLastName = ""
$password = "HomerSimpson99"  # Replace with the actual password
$calenderSMTP = "$CalenderAlias@mountstmary.ca"

# Create access control group in Groups OU
New-ADGroup -Name $EditorgroupName -GroupScope Universal -GroupCategory Security -Path "OU=Security,OU=Groups,OU=Mount St. Mary,DC=mtstmary,DC=local" -Description $EditorgroupDescription
New-ADGroup -Name $ReviewergroupName -GroupScope Universal -GroupCategory Security -Path "OU=Security,OU=Groups,OU=Mount St. Mary,DC=mtstmary,DC=local" -Description $ReivewergroupDescription

# Create Exchange Distribution Group
Enable-DistributionGroup -Identity $EditorgroupName -Alias $EditorAlias
Enable-DistributionGroup -Identity $ReviewergroupName -Alias $reviewerAlias

# Add users who need access to resource to Distribution Group
# Example: Add-DistributionGroupMember -Identity $displayName -Member "User1"

# Create new user account in Misc Mailboxes OU
New-ADUser -GivenName $userFirstName -Surname $userLastName -Name $userFirstName -UserPrincipalName $calenderSMTP `
    -SamAccountName $CalenderAlias -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -PasswordNeverExpires $true `
    -Enabled $true -Path "OU=Misc. Mailboxes,OU=Users,OU=Mount St. Mary,DC=mtstmary,DC=local"

# Create Exchange Mailbox for this user
Enable-Mailbox -Identity $calenderName -PrimarySmtpAddress $calenderSMTP

# Set Delivery Restrictions
Set-Mailbox -Identity $calenderName -AcceptMessagesOnlyFrom "Administrator"

# Update E-mail Addresses
Set-Mailbox -Identity $calenderName -PrimarySmtpAddress "$CalenderAlias@mountstmary.local"
Set-Mailbox -Identity $calenderName -EmailAddressPolicyEnabled $false


#Set the calender mailbox as shared

Set-Mailbox $calenderName -Type:Shared

#Grant groups permissions to calenders
Add-MailboxFolderPermission -Identity "${$calenderName}" -User $EditorgroupName -AccessRights Editor
Add-MailboxFolderPermission -Identity "${$calenderName}" -User $ReviewergroupName -AccessRights Reviewer

#Default: remove Free/Busy time; change to “None”
#Anonymous: confirm is “None” for calende 
#Set-MailboxFolderPermission -Identity "${$calenderName}:\Calendar" -User Default -AccessRights None
#Set-MailboxFolderPermission -Identity "${$calenderName}:\Calendar" -User Anonymous -AccessRights None
#################above need fix ###############################

# Hide resource group from Exchange lists
Set-ADGroup -Identity $EditorgroupName -Add @{msExchHideFromAddressLists=$true}
Set-ADGroup -Identity $ReviewergroupName -Add @{msExchHideFromAddressLists=$true}

