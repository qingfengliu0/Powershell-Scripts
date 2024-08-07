# Define the source and target group names
$sourceGroup = "BCS TLAB Nurses"
$targetGroup = "User Scripts - Shortcuts - CareRX"
$logFile = "NewGroupMembersLog.txt"

# Get current members of the target group
$existingMembers = Get-ADGroupMember -Identity $targetGroup | Select-Object -ExpandProperty DistinguishedName

# Get members of the source group and add to the target group
Get-ADGroupMember -Identity $sourceGroup | ForEach-Object {
    $member = $_
    $memberDN = $member.DistinguishedName
    if ($existingMembers -notcontains $memberDN) {
        try {
            Add-ADGroupMember -Identity $targetGroup -Members $memberDN
            $user = Get-ADUser -Identity $memberDN -Properties GivenName, Surname
            $firstName = $user.GivenName
            $lastName = $user.Surname
            $fullName = "$firstName $lastName"
            $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Added: $fullName from '$sourceGroup' to '$targetGroup'"
            Add-Content -Path $logFile -Value $logEntry
            Write-Output $logEntry
        } catch {
            Write-Error "Failed to add $fullName to $targetGroup. $_"
        }
    }
}

Write-Output "Group membership update completed."
