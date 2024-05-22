# Define function to get Netlogon script and Remote Desktop profile for users
function Get-UserNetlogonAndProfile {
    param (
    [string]$OUPath = "OU=Regular,OU=Users,OU=Luther Court Society,DC=lcs,DC=local"
)

$users = Get-ADUser -Filter * -Property * -SearchBase $OUPath

$result = @()
foreach ($user in $users) {
    $username = $user.SamAccountName
    $scriptPath = $user.ScriptPath
    # Fetch TerminalServicesProfilePath using ADSI with error handling
    $tsProfilePath = $null
    try {
        $adsiUser = [adsi]"LDAP://$($user.DistinguishedName)"
        $tsProfilePath = $adsiUser.psbase.InvokeGet("terminalservicesprofilepath")
    } catch {
        Write-Warning "Failed to retrieve TerminalServicesProfilePath for user: $username"
    }
    
    # Generate expected script name in "FirstInitialLastName.bat" format
    
    $expectedScriptName = "$username.bat"

    # Check if the ScriptPath is in the expected format
    $scriptMatches = $false
    if ($scriptPath -ne $null -and $expectedScriptName -ne "") {
        if ($scriptPath -eq $expectedScriptName) {
            $scriptMatches = $true
        }
    }

    # Check if TerminalServicesProfilePath is set correctly
    $profileMatches = $false
    if ($tsProfilePath -ne $null) {
        $expectedProfilePath = "\\Red\TSProfiles\$username"
        if ($tsProfilePath -eq $expectedProfilePath) {
            $profileMatches = $true
        }
    }

    # Add to result if conditions are not met
    if (-not $scriptMatches -or -not $profileMatches) {
        $result += [pscustomobject]@{
            SamAccountName = $username
            DisplayName = $displayName
            ScriptPath = $scriptPath
            ExpectedScriptName = $expectedScriptName
            ScriptMatches = $scriptMatches
            TerminalServicesProfilePath = $tsProfilePath
            ProfileMatches = $profileMatches
        }
    }
}

return $result
}

#search under the regular ou
$ouPath = "OU=Regular,OU=Users,OU=Luther Court Society,DC=lcs,DC=local"
$usersWithIssues = Get-UserNetlogonAndProfile -OUPath $ouPath

# Export the result to a CSV file
$usersWithIssues | Export-Csv -Path "UsersWithNetlogonAndProfileIssues.csv" -NoTypeInformation

Write-Output "Script completed. Results exported to UsersWithNetlogonAndProfileIssues.csv"
