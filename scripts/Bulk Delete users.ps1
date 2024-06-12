# Function to delete a profile folder or file using Robocopy 
$UserCredential = Get-Credential
$fileServer = "Empress"
$UserListPath = "Users to Be Deleted.csv"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://POWDERKING.lcs.local/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking

function PresentUserList {
    param (
        [array]$users
    )

    if ($users.Count -eq 0) {
        Log-Action "No users found."
        return
    }

    $formattedUsers = $users | ForEach-Object {
        [PSCustomObject]@{
            SamAccountName = $_.SamAccountName
            GivenName = $_.GivenName
            Surname = $_.Surname
        }
    }

    Write-Host "List of users found:"
    $formattedUsers | Format-Table -Property SamAccountName, GivenName, Surname -AutoSize | Out-String | Write-Host
}
#remove the shorcut if it still there
function RemoveShortcut {
    param(
        [string]$permissionUserDrivePath,
        [string]$shortcutName
    )
    
    # Attempt to remove the shortcut from the user's desktop
    try {
        # Define the path to the shortcut on the user's desktop
        $targetPath = "$permissionUserDrivePath\Desktop"
        $shortcutPath = Join-Path -Path $targetPath -ChildPath "$shortcutName.lnk"
        
        # Check if the shortcut exists
        if (Test-Path $shortcutPath) {
            # Remove the shortcut
            Remove-Item -Path $shortcutPath -Force
            Log-Action "Removed shortcut $shortcutName from $permissionUserDrivePath's desktop"
        } else {
            Log-Action "Shortcut $shortcutName does not exist on $permissionUserDrivePath's desktop"
        }
    } catch {
        Log-Action "Failed to remove shortcut $shortcutName from $permissionUserDrivePath's desktop : $_"
    }
}
#find the user to be delted
function fetchUser {
    param (
        [string]$userfullname,
        [String]$typeOfUser
    )
    
    # Split each entry into FirstName and LastName
    $FirstName, $LastName = $userfullname -split " "
    
    
        # Function to handle duplicates and prompt user to choose
        function HandleDuplicates {
        param (
            [array]$users
        )

        if ($users.Count -eq 0) {
            Log-Action "No users found."
            return $null
        }

        # Present the list of users
        PresentUserList -users $users

        Write-Host "Multiple users found. Please choose the correct $typeofUser by entering the corresponding SamAccountName:"
    
        $selectedIndex = Read-Host "Enter the SamAccountName of the $typeofUser you want to select"
        $selectedUser = $users | Where-Object { $_.SamAccountName -eq $selectedIndex }
        
        if ($selectedUser) {
            $ADUser = Get-ADUser -Identity $selectedUser -Properties *
            return $ADUser
        } else {
            Log-Action "Invalid selection. No user found with SamAccountName '$selectedIndex'."
            return $null
        }
    }

    try {
        $user = $null

        # Try finding by FirstName and LastName
        if ($LastName -and $Fistname) {
            $Firstname = $Firstname.trim()
            $Lastname = $Lastname.trim()
            $user = @(Get-ADUser -Filter "GivenName -eq '$($FirstName)' -and Surname -eq '$($LastName)'" -Properties *)
            $count = $user.Count
            Log-Action "Found $count user(s) with the given name: $userfullname"
            if ($count -eq 1) {
                return $user
            } elseif ($count -gt 1) {
                return HandleDuplicates -users $user
            }
        }

        # If not found, try finding by FirstName only
        if (-not $user) {
            $Firstname = $Firstname.trim()
            $user = @(Get-ADUser -Filter "GivenName -eq '$($FirstName)'" -Properties *)
            $count = $user.Count
            Log-Action "Found $count user(s) with the first name: $FirstName"
            if ($count -eq 1) {
                return $user
            } elseif ($count -gt 1) {
                return HandleDuplicates -users $user
            }
        }

        # If still not found, try finding by LastName only
        if (-not $user) {
            $LastName = $LastName.trim()
            $user = @(Get-ADUser -Filter "Surname -eq '$($LastName)'" -Properties *)

            $count = $user.Count
            Log-Action "Found $count user(s) with the last name: $LastName"
            if ($count -eq 1) {
                return $user
            } elseif ($count -gt 1) {
                return HandleDuplicates -users $user
            }
        }

        # If no user found at all
        if (-not $user) {
            Log-Action "No user found with the given name(s): $userfullname."
            return $null
        }
    } catch {
        Log-Action "An error occurred while trying to retrieve the user: $_"
        return $null
    }
}
function Log-Action {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFilePath -Value "$timestamp - $message"
}
function Remove-ProfileItem {
    param (
        [string]$path
    )

    if (Test-Path $path) {
        try {
            # Use Robocopy to delete the directory or file
            Robocopy Emptyfolder $path  /purge /e /ndl /njh
            Remove-Item -path $path
            Log-Action "Deleted item at path: $path" |Out-File -FilePath UsersDeleteResult.txt -Append
        } catch {
            Log-Action "Error deleting item at path: $path - $_" |Out-File -FilePath UsersDeleteResult.txt -Append
        }
    } else {
        Log-Action "Item not found at path: $path" |Out-File -FilePath UsersDeleteResult.txt -Append
    }
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
        Log-Action "$userfullname mailbox has been removed" | Out-File -FilePath UsersDeleteResult.txt -Append
    }
    catch {
        # If an exception occurs during removal, write an error message to the log file
        Log-Action "Error removing $userfullname mailbox: $_" | Out-File -FilePath UsersDeleteResult.txt -Append
    }
    Exit-PSSession
}

#disconnect the exchange session
function Disconnect-ExchangeOnline {
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.PSSession]$Session
    )

    try {
        Remove-PSSession -Session $Session
        Log-Action "Disconnected from Exchange Online"
    } catch {
        Log-Action "Failed to disconnect from Exchange Online: $_"
    }
}

function DeleteProfile{
    param
    Start-Process -FilePath 'delprof2.exe' -ArgumentList "/c:Phoenix /p /id:$($userObject.samAccountName)"
}

#main method
if (Test-Path $UserListPath) {
    # Read the list of names and usernames from the file
    $UserList = Get-Content $UserListPath
    # Loop through each entry and disable the corresponding user account
    foreach ($user in $UserList) {
        if ($user.DeletionApproval -eq "Yes"){
        #if the Approval for Deletion is true
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
        $netlogonScriptPath = "\\$fileServer\NETLOGON\$netlogonScriptPath"
        #if the the users is currently disabled 
        if ($userObject.Enabled -eq $false) {
            Log-Action "$username is disabled."|Out-File -FilePath UsersDeleteResult.txt -Append
            # Delete the user's personal drive
            if (user.Archive -eq "Yes"){
                #Archive the mailbox

                New-MailboxExportRequest -Mailbox $userObject.mail.toString() -FilePath "\\Cecelia\c$\temp\$($userobject.samAccountName).pst" -Name "$($userobject.samAccountName)"
                #Remove email account
                DeleteMailbox -userfullname $userObject.mail.toString()
            else {
                #Remove email account
                #DeleteMailbox -userfullname $userObject.mail.toString()
            }
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

        } elseif ($userObject.Enabled -eq $true) {
            Log-Action "user stil enabled" |Out-File -FilePath UsersDeleteResult.txt -Append
        }else{
		Log-Action "user not found" |Out-File -FilePath UsersDeleteResult.txt -Append
		}
    }
    }else{
        Log-Action ("the user $user did not get a approval for deletion yet, move to the next one")
    }
}

