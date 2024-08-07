# Define the source and target group names
$groupsList = "groups.txt"
$userslist = "users.txt"
$logFile = "NewGroupMembersLog.txt"




(Get-Content $groupsList).Trim() | ForEach-Object{
    $group = $_
    $existingMembers = Get-ADGroupMember -Identity $group | Select-Object -ExpandProperty samAccountName
    # Get members of the suser lists and add to the target group
    Get-Content $userslist | ForEach-Object {
        $fullname = $_
        $FirstName, $LastName = $_ -split " "
        $user = @(Get-ADUser -Filter "GivenName -eq '$($FirstName)' -and Surname -eq '$($LastName)'" -Properties *)
        $memberDN = $user.SamAccountName
        if ($existingMembers -notcontains $memberDN ) {
            try {
                Add-ADGroupMember -Identity $group -Members $memberDN
                $user = Get-ADUser -Identity $memberDN -Properties GivenName, Surname
                $firstName = $user.GivenName
                $lastName = $user.Surname
                $fullName = "$firstName $lastName"
                $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Added: $fullName to $group"
                Add-Content -Path $logFile -Value $logEntry
                Write-Output $logEntry
            } catch {
                Write-Error "Failed to add $fullName to $group. $_"
            }
        }
    }
}

Write-Output "Group membership update completed."