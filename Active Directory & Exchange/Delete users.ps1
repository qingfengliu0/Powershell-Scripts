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

function RemoveShortcut {
    param(
        [string]$permissionUserDrivePath,
        [string]$shortcutName
    )
    
    try {
        $targetPath = "$permissionUserDrivePath\Desktop"
        $shortcutPath = Join-Path -Path $targetPath -ChildPath "$shortcutName.lnk"
        
        if (Test-Path $shortcutPath) {
            Remove-Item -Path $shortcutPath -Force
            Log-Action "Removed shortcut $shortcutName from $permissionUserDrivePath's desktop"
        } else {
            Log-Action "Shortcut $shortcutName does not exist on $permissionUserDrivePath's desktop"
        }
    } catch {
        Log-Action "Failed to remove shortcut $shortcutName from $permissionUserDrivePath's desktop : $_"
    }
}

function fetchUser {
    param (
        [string]$userfullname,
        [string]$typeOfUser
    )
    
    # Split each entry into FirstName and LastName
    $FirstName, $LastName = $userfullname -split " "
    $FirstName = $FirstName.trim()
    $Lastname = $Lastname.trim()
    # Function to handle duplicates and prompt user to choose
    function HandleDuplicates {
        param (
            [array]$users
        )

        if ($users.Count -eq 0) {
            Write-Output "No users found."
            return $null
        }

        # Present the list of users
        PresentUserList -users $users

        Write-Host "Multiple users found. Please choose the correct $typeOfUser by entering the corresponding SamAccountName:"
    
        $selectedIndex = Read-Host "Enter the SamAccountName of the $typeOfUser you want to select"
        $selectedUser = $users | Where-Object { $_.SamAccountName -eq $selectedIndex }
        
        if ($selectedUser) {
            $ADUser = Get-ADUser -Identity $selectedUser -Properties *
            return $ADUser
        } else {
            Write-Output "Invalid selection. No user found with SamAccountName '$selectedIndex'."
            return $null
        }
    }

    try {
        $user = $null

        # Try finding by FirstName and LastName
        if ($LastName -and $FirstName) {
            $FirstName = $FirstName.trim()
            $LastName = $LastName.trim()
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
            $FirstName = $FirstName.trim()
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
    $logMessage = "$timestamp - $message"
    Add-Content -Path $logFilePath -Value $logMessage
    Write-Host $logMessage
}

function Remove-ProfileItem {
    param (
        [string]$path
    )

    if (Test-Path $path) {
        try {
             Write-Host "Starting to purge the directory: $path"
            
            # Use Robocopy to purge the directory
            Robocopy $Emptyfolder $path /purge /e /ndl /njh *> $null
            Write-Host "Purged the directory using Robocopy: $path"
            # Remove the directory and its contents recursively
            Remove-Item -Path $path -Recurse -Force
            Write-Host "Successfully removed the directory: $path"
            # Log successful deletion
            Log-Action "Deleted item at path: $path"
        } catch {
            # Log any errors that occur
            Log-Action "Error deleting item at path: $path - $_"
        }
    } else {
        # Log if the item was not found
        Log-Action "Item not found at path: $path"
    }
}



function DeleteMailbox {
    param(
        [string]$userfullname
    )
    Enter-PSSession -Session $Session
    try {
        Remove-Mailbox -Identity $userfullname -Permanent $true -ErrorAction Stop
        Log-Action "$userfullname mailbox has been removed" | Out-File -FilePath UsersDeleteResult.txt -Append
    } catch {
        Log-Action "Error removing $userfullname mailbox: $_" | Out-File -FilePath UsersDeleteResult.txt -Append
    }
    Exit-PSSession
}

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

function DeleteProfile {
    param (
        [string]$samAccountName,
        [array]$computers
    )

    foreach ($computer in $computers) {
        try {
            Start-Process -FilePath 'delprof2.exe' -ArgumentList "/c:$computer /p /id:$samAccountName" -Wait
            Log-Action "Deleted profile for $samAccountName on $computer"
        } catch {
            Log-Action "Failed to delete profile for $samAccountName on $computer : $_"
        }
    }
}
function Remove-MailboxBySmtp {
    param (
        [string]$smtpAddress
    )

    try {
        # Check if the SMTP address is not empty
        if (-not [string]::IsNullOrWhiteSpace($smtpAddress)) {
            # Confirm the SMTP address to be deleted
            Write-Host "Attempting to remove mailbox for: $smtpAddress"

            # Remove the mailbox with confirmation
            Remove-Mailbox -Identity $smtpAddress -Confirm:$true

            Write-Host "The mailbox for $smtpAddress has been successfully removed."
        } else {
            Write-Host "The provided SMTP address is empty or invalid."
        }
    } catch {
        Write-Host "An error occurred while attempting to remove the mailbox for $smtpAddress."
        Write-Host "Error details: $_"
    }
}

# Get Exchange admin credentials
$UserCredential = Get-Credential
$logFilePath = "userdelete.log"
# Connect to Exchange Online
$Emptyfolder = "EmptyFolder"
# ---------------define the following variable for different client #-----------------
$exchangeserver =  "http://cecelia.devonprop.local/PowerShell/"
$fileserver = "Hall"
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeserver -Authentication Kerberos -Credential $UserCredential
    Import-PSSession $Session -DisableNameChecking
    Log-Action "Connected to Exchange Online"
} catch {
    Log-Action "Failed to connect to Exchange Online: $_"
    exit
}

do {
    $userfullname = Read-Host "Enter the full name of the user to be deleted (Firstname Lastname) or type 'exit' to quit"
    if ($userfullname -eq 'exit') { break }
    $confirmation = "no"
    $userObject = fetchUser -userfullname $userfullname -typeOfUser "user to be deleted"
    if ($userObject) {
        Write-Host "You selected user: $($userObject.GivenName) $($userObject.Surname)"
        $confirmation = Read-Host "Is this correct? (yes/no)"
     } else {
            Log-Action "Can't find user in the Active Directory: $userfullname"
            continue
     }

    if ($userObject -ne $null -and $confirmation -eq "yes") {
        $personalDrivePath = $userObject.HomeDirectory
        $roamingProfilePath = $userObject.ProfilePath
        $roamingProfilePathV2 = "$roamingProfilePath.V2"
        $roamingProfilePathV6 = "$roamingProfilePath.V6"
        $lldpuser = [adsi]"LDAP://$($userObject.DistinguishedName)"
        $lldpusertspath = $lldpuser.psbase.InvokeGet("terminalservicesprofilepath")
        $tsProfilePathV2 = "$lldpusertspath.V2"
        $tsProfilePathV6 = "$lldpusertspath.V6"
        $netlogonScriptPath = "\\$fileServer\NETLOGON\$($userObject.samAccountName).bat"

        if ($userObject.Enabled -eq $false) {
            Log-Action "$($userObject.SamAccountName) is disabled."

            Remove-ProfileItem -path $personalDrivePath
            Remove-ProfileItem -path $roamingProfilePathV2
            Remove-ProfileItem -path $roamingProfilePathV6
            Remove-ProfileItem -path $tsProfilePathV2
            Remove-ProfileItem -path $tsProfilePathV6
            Remove-MailboxBySmtp -smtpAddress $userObject.mail.toString() -Confirm:$true
            Remove-ProfileItem -path $netlogonScriptPath  
            $computers = @()
        do {
            $computer = Read-Host "Enter the name of a computer the user has access to (type 'done' when finished)"
            if ($computer -ne 'done') {
                $computers += $computer
            }
        } while ($computer -ne 'done')
        
            DeleteProfile -samAccountName $userObject.samAccountName -computers $computers
            Log-Action "Completed deletion tasks for user $($userObject.GivenName) $($userObject.Surname)."
        } elseif ($userObject.Enabled -eq $true) {
            Log-Action "User $($userObject.GivenName) $($userObject.Surname) is still enabled, skipping deletion."
        } elseif ($confirmation -eq "no"){
        
            Log-Action "You refuse delete the found user $($userObject.GivenName) $($userObject.Surname)"
        }else {
            Log-Action "User $($userObject.GivenName) $($userObject.Surname) not found."
            }
        }
    } while ($true)
    
    # Disconnect from Exchange Online
    Disconnect-ExchangeOnline -Session $Session
    Write-Host "Processing completed. Log file created at $logFilePath"
