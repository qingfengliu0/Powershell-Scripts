# Function to delete a profile folder or file using Robocopy 
$UserCredential = Get-Credential
$UserListPath = "Users to Be Deleted.txt"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://POWDERKING.lcs.local/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking
function Remove-ProfileItem {
    param (
        [string]$path
    )

    if (Test-Path $path) {
        try {
            # Use Robocopy to delete the directory or file
            Robocopy Emptyfolder $path  /purge /e /ndl /njh
            Remove-Item -path $path
            Write-Output "Deleted item at path: $path" |Out-File -FilePath UsersDeleteResult.txt -Append
        } catch {
            Write-Output "Error deleting item at path: $path - $_" |Out-File -FilePath UsersDeleteResult.txt -Append
        }
    } else {
        Write-Output "Item not found at path: $path" |Out-File -FilePath UsersDeleteResult.txt -Append
    }
}
#get the username based on the firstname and lastname
# Get the user account
function fetchUser{
    param (
        [string]$userfullname
    )
    # Split each entry into name and username
    $FirstName, $LastName = $userfullname -split " "
    
    $user = Get-ADUser -Filter "GivenName -eq '$($FirstName)' -and Surname -eq '$($LastName)'" -Properties *
    return $user
}

function DeleteMailbox{
    param(
        [string]$userfullname
    )
        Enter-PSSession -Session $Session
        try {
        # Attempt to remove the mailbox
        Remove-Mailbox -Identity $userfullname -Permanent $true -ErrorAction Stop

        # If removal is successful, write a success message to the log file
        Write-Output "$userfullname mailbox has been removed" | Out-File -FilePath UsersDeleteResult.txt -Append
    }
    catch {
        # If an exception occurs during removal, write an error message to the log file
        Write-Output "Error removing $userfullname mailbox: $_" | Out-File -FilePath UsersDeleteResult.txt -Append
    }
    Exit-PSSession
}
#If the user is disabled then deleted the folders else say the user is still active !



if (Test-Path $UserListPath) {
    # Read the list of names and usernames from the file
    $UserList = Get-Content $UserListPath

    # Loop through each entry and disable the corresponding user account
    foreach ($user in $UserList) {
        write-host "the user is $user"
        $userObject = fetchUser -userfullname $user
        # get the paths
        $personalDrivePath = $userObject.HomeDirectory
        $romaingProfile = $userObject.ProfilePath
        $roamingProfilePathV2 = "$($userObject.ProfilePath).V2"
        $roamingProfilePathV6 = "$($userObject.ProfilePath).V6"
        $lldpuser = [adsi]"LDAP://$userObject"
        $lldpusertspath = $lldpuser.psbase.InvokeGet(“terminalservicesprofilepath”)
        $tsProfilePathV2 = "$($lldpusertspath).V2"
        $tsProfilePathV6 = "$($lldpusertspath).V6"
        #the logon script could be pointed to the wrong name but still working, so pointing to the correct one. 
        $netlogonScriptPath = "$($userObject.samAccountName).bat"
        $netlogonScriptPath = "\\red\NETLOGON\$netlogonScriptPath"
        if ($userObject.Enabled -eq $false) {
            Write-Host "$username is disabled."|Out-File -FilePath UsersDeleteResult.txt -Append
            # Delete the user's personal drive
            Remove-ProfileItem -path $personalDrivePath

            # Delete the roaming profile folder for V2 format
            Remove-ProfileItem -path $roamingProfilePathV2

            # Delete the roaming profile folder for V6 format
            Remove-ProfileItem -path $roamingProfilePathV6

            # Delete the terminal server profile folder for V2 format
            Remove-ProfileItem -path $tsProfilePathV2

            # Delete the terminal server profile folder for V6 format
            Remove-ProfileItem -path $tsProfilePathV6

            # Check if the user's logon script exists and delete it
            Remove-ProfileItem -path $netlogonScriptPath

            #Remove email account
            DeleteMailbox -userfullname $user

            #Remove Profile on Terminal Server
	        Start-Process -FilePath 'delprof2.exe' -ArgumentList "/c:Phoenix /p /id:$($userObject.samAccountName)"

        } elseif ($userObject.Enabled -eq $true) {
            Write-Host "user stil enabled" |Out-File -FilePath UsersDeleteResult.txt -Append
        }else{
		Write-Host "user not found" |Out-File -FilePath UsersDeleteResult.txt -Append
		}
    }
}

