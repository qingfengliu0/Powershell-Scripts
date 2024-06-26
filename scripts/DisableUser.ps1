# Load required modules
#Import-Module ImportExcel
Import-Module ActiveDirectory
#Import-Module ExchangeOnlineManagement

# Define the log file path
$logFilePath = "userdisable.log"
$Session = $null

# Function to disconnect from Exchange Online
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

function PresentUserList {
    param (
        [array]$users
    )

    if ($users.Count -eq 0) {
        Write-Output "No users found."
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

function fetchUser {
    param (
        [string]$userfullname,
        [string]$typeOfUser
    )
    
    # Split each entry into FirstName and LastName
    $FirstName, $LastName = $userfullname -split " "
    
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

        Write-Host "users found. Please choose the correct $typeOfUser by entering the corresponding SamAccountName:"
    
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

    function FindUser {
        param (
            [string]$filter
        )
        return @(Get-ADUser -Filter $filter -Properties *)
    }

    try {
        $user = $null
        $count = 0

        switch -regex ($true) {
            # Try finding by FirstName and LastName
            { $LastName -and $FirstName } {
                $FirstName = $FirstName.trim()
                $LastName = $LastName.trim()
                $user = FindUser -filter "GivenName -eq '$FirstName' -and Surname -eq '$LastName'"
                $count = $user.Count
                Log-Action "Found $count user(s) with the given name: $userfullname"
                if ($count -ge 1) { return HandleDuplicates -users $user }
                continue
            }

            # If not found, try finding by FirstName only
            { -not $user } {
                $FirstName = $FirstName.trim()
                $user = FindUser -filter "GivenName -eq '$FirstName'"
                $count = $user.Count
                Log-Action "Found $count user(s) with the first name: $FirstName"
                if ($count -ge 1) { return HandleDuplicates -users $user }
                continue
            }

            # If still not found, try finding by LastName only
            { -not $user } {
                $LastName = $LastName.trim()
                $user = FindUser -filter "Surname -eq '$LastName'"
                $count = $user.Count
                Log-Action "Found $count user(s) with the last name: $LastName"
                if ($count -ge 1) { return HandleDuplicates -users $user }
                continue
            }

            # If still not found, try finding by DisplayName
            { -not $user } {
                $user = FindUser -filter "DisplayName -like '*$userfullname*'"
                $count = $user.Count
                Log-Action "Found $count user(s) with the display name like: $userfullname"
                if ($count -ge 1) { return HandleDuplicates -users $user }
                continue
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


# Function to log actions
function Log-Action {
    param (
        [string]$message
    )
    write-Host $message
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFilePath -Value "$timestamp - $message"
}

# Function to handle email forwarding
function ForwardEmail {
    param (
        [string]$userfullname, 
        [string]$forwardingUserName
    )

    try {
        $forwardingSmtpEmailAddress = (Get-Mailbox -Identity $forwardingUserName).PrimarySmtpAddress.ToString()
        Log-Action "Convert to email address for $forwardingUserName to $forwardingSmtpEmailAddress"
    } catch {
        Log-Action "Failed to convert $forwardingUserName to a domain email address"
    }

    try {
        Set-Mailbox -Identity $userfullname -ForwardingSMTPAddress $forwardingSmtpEmailAddress -DeliverToMailboxAndForward $true
        Log-Action "Set email forwarding for $userfullname to $forwardingSmtpEmailAddress"
    } catch {
        Log-Action "Failed to set email forwarding for $userfullname : $_"
    }
}

function GrantFullPermission {
    param (
        [string]$userfullname, 
        [string]$EmailPermissionUser
    )

    try {
        Add-MailboxPermission -Identity $userfullname -User $EmailPermissionUser -AccessRights FullAccess -InheritanceType All
        Log-Action "Granted $EmailPermissionUser full access to $userfullname mailbox"
    } catch {
        Log-Action "Failed to grant mailbox permissions for $userfullname : $_"
    }
}

function SetOOOMessage {
    param (
        [string]$userfullname, 
        [string]$InternaloooMessage,
        [string]$ExternaloooMessage
    )

    try {
        $InternaloooMessage = $InternaloooMessage.replace("\\n", "`n")
        $ExternaloooMessage = $ExternaloooMessage.replace("\\n", "`n")
        $InternaloooMessage = '<pre>' + $InternaloooMessage + '</pre>'
        $ExternaloooMessage = '<pre>' + $ExternaloooMessage + '</pre>'
        Set-MailboxAutoReplyConfiguration -Identity $userfullname -AutoReplyState Enabled -InternalMessage $InternaloooMessage -ExternalMessage $ExternaloooMessage
        Log-Action "Set Out-of-Office message for $userfullname"
    } catch {
        Log-Action "Failed to set Out-of-Office message for $userfullname : $_"
    }
}

function Remove-UserFromDistributionLists {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserIdentity
    )

    $groups = Get-ADUser -Identity $UserIdentity -Properties MemberOf | Select-Object -ExpandProperty MemberOf
    $DisplayName = Get-ADUser -Identity $UserIdentity -Properties DisplayName | Select-Object -ExpandProperty DisplayName

    if (-not $groups) {
        Log-Action "$UserIdentity is not a member of any groups."
        return
    }
    $count = 0
    try {
        foreach ($groupDN in $groups) {
            $group = Get-ADGroup -Identity $groupDN -Properties GroupCategory
            if ($group.GroupCategory -eq "Distribution") {
                $count++
                try {
                    Remove-ADGroupMember -Identity $group -Members $UserIdentity -Confirm:$false
                    Log-Action "Removed user $DisplayName from distribution list $($group.Name)."
                } catch {
                    Log-Action "Failed to remove user $DisplayName from distribution list $($group.Name): $_"
                }
            }
        }
    } catch {
        Log-Action "An error occurred: $_"
    }
    if ($count -eq 0){
         Log-Action "$UserIdentity is not a member of any groups."
    }
}

function ProvideFolderAccess {
    param (
        [string]$FolderPermissionUser, 
        [string]$PersonalDrivePath
    )

    try {
        $permissions = "Modify"
        $accessType = "Allow"

        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($FolderPermissionUser, $permissions, $accessType)
        $acl = Get-Acl $PersonalDrivePath
        $acl.SetAccessRule($accessRule)
        Set-Acl -Path $PersonalDrivePath -AclObject $acl

    } catch {
        Log-Action "Failed to grant folder access for $FolderPermissionUser : $_"
    }
}

function AddShortcut {
    param (
        [string]$PermissionUserDrivePath,
        [string]$PersonalDrivePath
    )

    try {
        $targetPath = "$PermissionUserDrivePath\Desktop"
        $shortcutPath = Join-Path -Path $targetPath -ChildPath "$userfullname.lnk"
        $wshShell = New-Object -ComObject WScript.Shell
        $shortcut = $wshShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $PersonalDrivePath
        $shortcut.Save()

        Log-Action "Created folder shortcut on $PermissionUserDrivePath's desktop"
    } catch {
        Log-Action "Failed to create folder shortcut for $PermissionUserDrivePath : $_"
    }
}

# Function to remove first and last occurrence of double quotes
function Remove-FirstAndLastQuotes {
    param (
        [string]$inputString
    )
    if ($inputString.StartsWith('"') -and $inputString.EndsWith('"')) {
        return $inputString.Substring(1, $inputString.Length - 2)
    }
    return $inputString
}

# Function to get user input
function Get-UserInput {
    param (
        [string]$Prompt
    )
    Read-Host -Prompt $Prompt
}

# Get Exchange admin credentials
$UserCredential = Get-Credential

# Connect to Exchange Online
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://cecelia.devonprop.local/PowerShell/ -Authentication Kerberos -Credential $UserCredential
    Import-PSSession $Session -DisableNameChecking
    Log-Action "Connected to Exchange Online"
} catch {
    Log-Action "Failed to connect to Exchange Online: $_"
    exit
}

# Loop to get user input and perform actions
do {
    $userfullname = Get-UserInput -Prompt "Enter the full name of the user to be disabled (or type 'exit' to quit)"
    if ($userfullname -eq 'exit') { break }

    $userObject = fetchUser -userfullname $userfullname -typeOfUser "User to Be Disabled"
    $ConfirmedDisable = Get-UserInput -Prompt "Confirm disable $($userObject.SamAccountName), Type Y for yes?"
    if ($userObject -eq $null -or $ConfirmedDisable -ne "Y") {
        Log-Action "Can't find user in the Active Directory: $userfullname"
        continue
    } else {
        $personalDrivePath = $userObject.HomeDirectory
        # Remove from DLs
        Remove-UserFromDistributionLists -UserIdentity $userObject.SamAccountName
        # Disable the user and hide from address list
        try {
            Disable-ADAccount -Identity $userObject.SamAccountName
            Set-ADUser -Identity $userObject.SamAccountName -Replace @{msExchHideFromAddressLists=$true} -ErrorAction Stop
            Log-Action "Disabled user account for $userfullname"
        } catch {
            Log-Action "Failed to disable user account for $userfullname : $_"
        }
        
        #forwarding
        do {
            $forwardingUserName = Get-UserInput -Prompt "Enter the username to forward emails to (leave blank if not required)"
            
            if (-not $forwardingUserName) {
                Log-Action "Forwarding not required for $userfullname"
                break
            }
        
            $fowarduserObject = fetchUser -userfullname $forwardingUserName -typeOfUser "user to forward email to"
            
            if ($fowarduserObject) {
                ForwardEmail -userfullname $userObject.mail.toString() -forwardingUserName $fowarduserObject.mail.toString()
                $validUser = $true
            } else {
                Write-Host "User not found. Please try again."
                $validUser = $false
            }
        } while (-not $validUser)
        

        # Grant full permission

        do {
            $EmailPermissionUser = Get-UserInput -Prompt "Enter the username to grant full mailbox permissions to (leave blank if not required)"
            
            if (-not $EmailPermissionUser) {
                Log-Action "Email permission not required for $userfullname"
                break
            }
            
            $EmailPermissionUserObject = fetchUser -userfullname $EmailPermissionUser -typeOfUser "user to have full email permission"
        
            if ($EmailPermissionUserObject) {
                GrantFullPermission -userfullname $userObject.mail.toString() -EmailPermissionUser $EmailPermissionUserObject.mail.toString()
                $validUser = $true
            } else {
                Write-Host "User not found. Please try again."
                $validUser = $false
            }
        } while (-not $validUser)
        

        # Set OOO message
        $InternaloooMessage = Get-UserInput -Prompt "Enter the internal Out-of-Office message (leave blank if not required, use \\n as newline)"
        $ExternaloooMessage = Get-UserInput -Prompt "Enter the external Out-of-Office message (leave blank if not required, use \\n as newline)"
        if ($InternaloooMessage -or $ExternaloooMessage) {
            $InternaloooMessage = Remove-FirstAndLastQuotes -inputString $InternaloooMessage
            $ExternaloooMessage = Remove-FirstAndLastQuotes -inputString $ExternaloooMessage
            SetOOOMessage -userfullname $userObject.mail.ToString() -InternaloooMessage $InternaloooMessage -ExternaloooMessage $ExternaloooMessage
        } else {
            Log-Action "Out-of-Office message not required for $userfullname"
        }

        do {
            $FolderPermissionUser = Get-UserInput -Prompt "Enter the username to grant folder permissions to (leave blank if not required)"
            
            if (-not $FolderPermissionUser) {
                Log-Action "Folder permission not required for $userfullname"
                break
            }
        
            if ($personalDrivePath) {
                $permissionUserObject = fetchUser -userfullname $FolderPermissionUser -typeOfUser "user to have modify folder permission"
                
                if ($permissionUserObject) {
                    $permissionUserDrivePath = $permissionUserObject.HomeDirectory
                    # Provide folder access to disable user's personal drive path
                    ProvideFolderAccess -FolderPermissionUser $permissionUserObject.samAccountName -PersonalDrivePath $personalDrivePath
                    # Add the shortcut to permission user's desktop pointing to the disabled user
                    AddShortcut -PermissionUserDrivePath $permissionUserDrivePath -PersonalDrivePath $personalDrivePath
                    $validUser = $true
                } else {
                    Write-Host "User not found. Please try again."
                    $validUser = $false
                }
            } else {
                Log-Action "$userfullname home directory does not exist"
                $validUser = $true
            }
        } while (-not $validUser)
        
    }
} while ($true)

# Call the function to disconnect and log the action
Disconnect-ExchangeOnline -Session $Session
Write-Host "Processing completed. Log file created at $logFilePath"
