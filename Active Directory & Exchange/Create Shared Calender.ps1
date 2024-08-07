# Import necessary modules
Import-Module ActiveDirectory
Import-Module ExchangeOnlineManagement
#Client OU#
#"OU=Security,OU=Groups,OU=Mount St. Mary,DC=mtstmary,DC=local"
#"OU=Misc. Mailboxes,OU=Users,OU=Mount St. Mary,DC=mtstmary,DC=local"
#mountstmary.ca
# Define variables
$calenderName = "OHS"  # Replace with the actual room name
$SecurityOU = "OU=Security,OU=Groups,OU=Broadmead Care Society,DC=bcs,DC=local"
$miscMailboxOU = "OU=Misc. Mailboxes,OU=Users,OU=Broadmead Care Society,DC=bcs,DC=local"
$CalenderAlias = $calenderName -replace '\s+'
$calenderSMTP = "$CalenderAlias@broadmeadcare.com"
$calenderLocal = "$CalenderAlias@broadmeadcare.local"
################################## Section above require change ####################
$EditorgroupName = "C - $calenderName Editor"
$ReviewergroupName = "C - $calenderName Reviewer"
$EditorgroupDescription = "Members have Editor access to $calenderName calendar"
$ReivewergroupDescription = "Members have Reviewer access to $calenderName calendar"
$EditorAlias = "C-$($calenderName)Editor" -replace '\s+'
$reviewerAlias = "C-$($calenderName)Reviewer" -replace '\s+'
$userFirstName = "C - $calenderName"

$userLastName = ""
$password = ""  # Replace with the actual password

# Create access control group in Groups OU
New-ADGroup -Name $EditorgroupName -GroupScope Universal -GroupCategory Security -Path $SecurityOU -Description $EditorgroupDescription
New-ADGroup -Name $ReviewergroupName -GroupScope Universal -GroupCategory Security -Path $SecurityOU -Description $ReivewergroupDescription

# Create Exchange Distribution Group
Enable-DistributionGroup -Identity $EditorgroupName -Alias $EditorAlias
Enable-DistributionGroup -Identity $ReviewergroupName -Alias $reviewerAlias

# Add users who need access to resource to Distribution Group
# Example: Add-DistributionGroupMember -Identity $displayName -Member "User1"

# Create new user account in Misc Mailboxes OU
New-ADUser -GivenName $userFirstName -Surname $userLastName -Name $userFirstName -UserPrincipalName $calenderSMTP `
    -SamAccountName $CalenderAlias -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -PasswordNeverExpires $true `
    -Enabled $true -Path $miscMailboxOU

# Create Exchange Mailbox for this user
Enable-Mailbox -Identity $calenderName -PrimarySmtpAddress $calenderSMTP

# Set Delivery Restrictions
Set-Mailbox -Identity $calenderName -AcceptMessagesOnlyFrom "Administrator"

# Update E-mail Addresses
Set-Mailbox -Identity $calenderName -PrimarySmtpAddress $calenderLocal
Set-Mailbox -Identity $calenderName -EmailAddressPolicyEnabled $false


#Set the calender mailbox as shared

Set-Mailbox $calenderName -Type:Shared
Add-MailboxPermission -Identity $calenderName -User "tech support" -AccessRights FullAccess -InheritanceType All
#############################below needs fix######################################
#Grant groups permissions to calenders
#Add-MailboxFolderPermission -Identity "${$calenderName}" -User $EditorgroupName -AccessRights Editor
#Add-MailboxFolderPermission -Identity "${$calenderName}" -User $ReviewergroupName -AccessRights Reviewer

#Default: remove Free/Busy time; change to “None”
#Anonymous: confirm is “None” for calende 
#Set-MailboxFolderPermission -Identity "${$calenderName}:\Calendar" -User Default -AccessRights None
#Set-MailboxFolderPermission -Identity "${$calenderName}:\Calendar" -User Anonymous -AccessRights None
#################above need fix ###############################

# Hide resource group from Exchange lists
Set-ADGroup -Identity $EditorgroupName -Add @{msExchHideFromAddressLists=$true}
Set-ADGroup -Identity $ReviewergroupName -Add @{msExchHideFromAddressLists=$true}

