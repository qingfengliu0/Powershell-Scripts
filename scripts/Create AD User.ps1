$logFilePath = "usercreation.log"
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
function Get-RandomWord {
    $words = @("Sapphire", "Emerald", "Crimson", "Amethyst", "Beryl", "Citrine", "Diamond", "Eclipse", "Fuchsia", "Garnet")
    return $words | Get-Random
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
            Write-Output "No users found."
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
            Write-Output "Invalid selection. No user found with SamAccountName '$selectedIndex'."
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
# Function to log actions
function Log-Action {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFilePath -Value "$timestamp - $message"
}
function Copy-ADUserAsTemplate {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TemplateUserName,

        [Parameter(Mandatory = $true)]
        [string]$NewUserName,

        [Parameter(Mandatory = $true)]
        [string]$NewUserPassword,

        [string]$NewUserUPN = "$NewUserName@domain.com",

        [string]$NewUserOU = (Get-ADUser -Identity $TemplateUserName).DistinguishedName
    )

    # Get the template user object
    $templateUser = Get-ADUser -Identity $TemplateUserName -Properties *

    # Modify properties for the new user
    $newUserInstance = $templateUser.PSObject.Copy()
    $newUserInstance.Name = $NewUserName
    $newUserInstance.SamAccountName = $NewUserName
    $newUserInstance.UserPrincipalName = $NewUserUPN
    $newUserInstance.AccountPassword = (ConvertTo-SecureString $NewUserPassword -AsPlainText -Force)
    $newUserInstance.Enabled = $true
    $newUserInstance.DistinguishedName = $NewUserOU

    # Create the new user
    New-ADUser -Instance $newUserInstance

    # Get the group memberships of the template user
    $templateUserGroups = Get-ADUser -Identity $TemplateUserName -Properties MemberOf | Select-Object -ExpandProperty MemberOf

    # Add the new user to each of the template user's groups
    foreach ($group in $templateUserGroups) {
        Add-ADGroupMember -Identity $group -Members $NewUserName
    }
}

# Function to create a new user on the domain
function New-DomainUser {
    param (
        [string]$firstName,
        [string]$lastName,
        [string]$copyFromFirstName,
        [string]$copyFromLastName
    )

    $copyFromUserFullName = "$copyFromFirstName $copyFromLastName"

    # Confirm the copy from user
    $copyFromUser = fetchuser -userfullname $copyFromUserFullName -typeOfUser "Copy from Users"
    if (-not $copyFromUser) {
        Write-Host "User not found. Please try again."
        return
    }

    # Create the account name and password
    $accountName = "$($firstName.Substring(0, 1))$lastName"
    $randomNumbers = -join ((48..57) | Get-Random -Count 2 | ForEach-Object { [char]$_ })
    $randomWord = Get-RandomWord
    $password = "$($firstName.Substring(0, 1).ToLower())$($lastName.Substring(0, 1).ToLower())$randomNumbers$randomWord"
    write-host "the password is $password"
    # Create the new user
    $newuser = New-ADUser -Name "$firstName $lastName" -GivenName $firstName -Surname $lastName -SamAccountName $accountName -UserPrincipalName "$accountName@domain.com" -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -Enabled $true -PasswordNeverExpires $true -DisplayName "$firstName $lastName"

    # Set additional properties
    Set-ADUser -Identity $accountName -Description $copyFromUser.Description

    # Ask for remote access
    $remoteAccess = Read-Host "Does the user need remote access? (yes/no)"
    if ($remoteAccess -eq "yes") {
        Remove-ADGroupMember -Identity "Deny TS+RDS" -Members $accountName
        Add-ADGroupMember -Identity "Remote Desktop Gateway Users" -Members $accountName
        Add-ADGroupMember -Identity "Terminal Services Users" -Members $accountName
    }

    # Modify security on \\Empress\Users\<username>
    $userFolder = $newuser.HomeDirectory
    $acl = Get-Acl $userFolder
    $acl.Access | Where-Object { $_.IdentityReference -like "*Administrators" } | ForEach-Object { $acl.RemoveAccessRule($_) }
    Set-Acl -Path $userFolder -AclObject $acl

    # Open Exchange Management Console and create mailbox
    New-Mailbox -Alias $accountName -Name "$firstName $lastName" -UserPrincipalName "$accountName@domain.com" -SamAccountName $accountName -FirstName $firstName -LastName $lastName -Password (ConvertTo-SecureString $password -AsPlainText -Force)

    # Edit mailbox settings
    Set-Mailbox -Identity "$firstName $lastName" -PrimarySmtpAddress "$($firstName[0])$lastName@devonproperties.com" -EmailAddresses @{Add="$($firstName[0])$lastName@devonprop.local", "$($firstName[0])$lastName@devonprop.com"} -DisplayName "$firstName $lastName - Dan 2024-02-28"
    Set-Mailbox -Identity "$firstName $lastName" -RetentionPolicy "Standard Retention Policy"

    # Mailbox delegation for Property Managers
    $isPropertyManager = Read-Host "Is the new user a Property Manager? (yes/no)"
    if ($isPropertyManager -eq "yes") {
        Add-MailboxPermission -Identity "$firstName $lastName" -User "Alex" -AccessRights FullAccess -InheritanceType All
        Add-RecipientPermission -Identity "$firstName $lastName" -Trustee "Alex" -AccessRights SendAs
    }
    # Export username, password, and email address
    $userInfo = @(
        @{
            Username  = $accountName
            Password  = $password
            Email     = "$($firstName[0])$lastName@devonproperties.com"
        }
    )

    $userInfo | Export-Csv -Path "$accountName-Info.csv" -NoTypeInformation
     # Export the list of groups the user has access to
    Get-ADUser -Identity $accountName -Properties MemberOf | Select-Object -ExpandProperty MemberOf | Get-ADGroup | Select-Object Name | Export-Csv -Path "$accountName-Groups.csv" -NoTypeInformation
}
   


# Main script
$firstName = Read-Host "Enter the first name of the employee"
$lastName = Read-Host "Enter the last name of the employee"
$copyFromFirstName = Read-Host "Enter the first name of the user to copy from"
$copyFromLastName = Read-Host "Enter the last name of the user to copy from"

#get exchange admin credentials
$UserCredential = Get-Credential

# Connect to Exchange Online

 try {
      $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://cecelia.devonprop.local/PowerShell/  -Authentication Kerberos -Credential $UserCredential
      Import-PSSession $Session -DisableNameChecking
      Log-Action "Connected to Exchange Online"
    } catch {
      Log-Action "Failed to connect to Exchange Online: $_"
        exit
}
# Create the new user

New-DomainUser -firstName $firstName.trim() -lastName $lastName.trim() -copyFromFirstName $copyFromFirstName -copyFromLastName $copyFromLastName
# Call the function to disconnect and log the action
Disconnect-ExchangeOnline -Session $Session
Write-Host "Processing completed. Log file created at $logFilePath"
