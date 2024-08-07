# Replace 'SourceUser' and 'TargetUser' with the actual usernames
$sourceUser = "CarolineW"
$targetUser = "cspack"

# Get the source user's security identifier (SID)
$sourceSID = (Get-ADUser -Identity $sourceUser).SID

# Get the source user's group memberships
$sourceGroups = (Get-ADUser -Identity $sourceUser -Properties MemberOf).MemberOf

# Set the target user's group memberships to match the source user's memberships
foreach ($group in $sourceGroups) {
    Add-ADGroupMember -Identity $group -Members $targetUser
}

