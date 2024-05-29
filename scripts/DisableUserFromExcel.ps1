
# Load required modules
Import-Module ImportExcel
Import-Module ActiveDirectory
#Import-Module ExchangeOnlineManagement

# Define the path to the Excel file and log file
$excelFilePath = "userToBeDisabled.xlsx"
$logFilePath = "userdisable.log"

#function to disconnect from exchange
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

function fetchUser{
    param (
        [string]$userfullname
    )
    # Split each entry into name and username
    $FirstName, $LastName = $userfullname -split " "
    #handle when lastname is empty like "testuser2"
    try {
        if ($LastName -eq $null) {
            $user = Get-ADUser -Filter "GivenName -eq '$($FirstName)'" -Properties *
        } else {
            $user = Get-ADUser -Filter "GivenName -eq '$($FirstName)' -and Surname -eq '$($LastName)'" -Properties *
        }
        
        if ($user -eq $null) {
            Write-Warning "No user found with the given name(s)."
            Disconnect-ExchangeOnline -Session $Session
            throw "$user cannot be found, check the source"
        } else {
            return $user
        }
    } catch {
        Write-Error "An error occurred while trying to retrieve the user: $_"
        Disconnect-ExchangeOnline -Session $Session
        throw "$user cannot be found, check the source"
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

#functiontoHandleForwarding email
function ForwardEmail{
    param(
        [string] $userfullname, 
        [string] $forwardingUserName
    )
    #convert forwardingusername to smtp address
    try{
        $forwardingSmtpEmailAddress = (Get-Mailbox -Identity $forwardingUserName).PrimarySmtpAddress.ToString()
        Log-Action "Convert to email address for $forwardingUserName to $forwardingSmtpEmailAddress"
    }Catch{
        Log-Action "failed to convert the $forwardingUserName to a domain email address "
    }
    # Forward emails
    try {
        Set-Mailbox -Identity $userfullname -ForwardingSMTPAddress $forwardingSmtpEmailAddress -DeliverToMailboxAndForward $true
        Log-Action "Set email forwarding for $userfullname to $forwardingEmailAddress"
    } catch {
        Log-Action "Failed to set email forwarding for $userfullname : $_"
    }

}

function GrantFullPermission{
    param(
        [string] $userfullname, 
        [string] $EmailPermissionUser
    )
    # Set email permissions
    try {
        Add-MailboxPermission -Identity $userfullname -User $EmailPermissionUser -AccessRights FullAccess -InheritanceType All
        Log-Action "Granted $EmailPermissionUser full access to $userfullname mailbox"
    } catch {
        Log-Action "Failed to grant mailbox permissions for $userfullname : $_"
    }
}

function setOOOMessage{
    param(
        [string] $userfullname, 
        [String] $oooMessage
        
    )
    # Configure Out-of-Office message
    try {
        Set-MailboxAutoReplyConfiguration -Identity $userfullname -AutoReplyState Enabled -InternalMessage $oooMessage -ExternalMessage $oooMessage
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

    try {
        # Iterate over each group and check if it is a distribution list (DL)
        foreach ($groupDN in $groups) {
            $group = Get-ADGroup -Identity $groupDN -Properties GroupCategory

            # Check if the group is a distribution list
            if ($group.GroupCategory -eq "Distribution") {
                try {
                    # Remove the user from the distribution list
                    Remove-ADGroupMember -Identity $group -Members $user -Confirm:$false
                    Log-Action "Removed user $UserIdentity from distribution list $($group.Name)."
                } catch {
                    Log-Action "Failed to remove user $UserIdentity from distribution list $($group.Name): $_"
                }
            }
        }
    } catch {
        # Handle any errors
        Log-Action "An error occurred: $_"
    }
}




#function to provide modify access
function provideFolderAccess{
    param(
    [String] $FolderPermissionUser, 
    [String] $personalDrivePath
    )

    try {
        # Define the user and the permissions
        $permissions = "Modify"
        $accessType = "Allow"

        # Create a new FileSystemAccessRule object
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($folderPermissionUser, $permissions, $accessType)

        #get the acl and set the acl at the user drive path
        $acl = Get-Acl $personalDrivePath
        $acl.SetAccessRule($accessRule)
        Set-Acl -Path $personalDrivePath -AclObject $acl

    } catch {
        Log-Action "Failed to grant folder access for $username : $_"
    }
    
}

function AddShorcut{

    param(
        [string]$permissionUserDrivePath,
        [string]$personalDrivePath
    )
    # Place a shortcut on the user's desktop
    try {
        #create a shorcut on requester's desktop with name of the disableuser.ink
        $targetPath = "$permissionUserDrivePath\Desktop"
        $shortcutPath = Join-Path -Path $targetPath -ChildPath "$userfullname.lnk"
        #point the shorcut to disable user's home folder
        $wshShell = New-Object -ComObject WScript.Shell
        $shortcut = $wshShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $personalDrivePath  # The folder you want to link to
        $shortcut.Save()

        Log-Action "Created folder shortcut on $folderPermissionUser's desktop"
    } catch {
        Log-Action "Failed to create folder shortcut for $folderPermissionUser : $_"
    }
}

# Read the Excel file
$users = Import-Excel -Path $excelFilePath
# Connect to Exchange Online
$UserCredential = Get-Credential
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://cecelia.devonprop.local/PowerShell/ -Authentication Kerberos -Credential $UserCredential
    Import-PSSession $Session -DisableNameChecking
    Log-Action "Connected to Exchange Online"
} catch {
    Log-Action "Failed to connect to Exchange Online: $_"
    exit
}

foreach ($user in $users) {
    $userfullname = $user.Name
    $userObject = fetchUser -userfullname $user.Name
    $personalDrivePath = $userObject.HomeDirectory
    

    #forward email
    if($user.forwardingUserName -eq $null){
        Log-Action "Forwarding Not required for $userfullname"
    }else{
        $forwardingUserName = $user.forwardingUserName
        fetchuser -userfullname $user.forwardingUserName
        ForwardEmail -userfullname $userfullname -forwardingUserName $user.forwardingUserName
    }
    
    #grant full permission
    if($EmailPermissionUser -eq $null){ 
        Log-Action "Email Permission Not required for $userfullname"
    }else{
        $EmailPermissionUser = $user.EmailPermissionUser
        fetchUser -userfullname $user.EmailPermissionUser
        GrantFullPermission -userfullname $userfullname -EmailPermissionUser $EmailPermissionUser
    }

    #setup ooomessage 
    if($ooMessage -eq $Null){
        Log-Action "ooo message not required for $userfullname"
    }else{
        $oooMessage = $user.OOOMessage
        setOOOMessage -userfullname $userfullname -oooMessage $oooMessage
    }
    
    #setup homefolder access
    if($folderPermissionUser -eq $null){
        Log-Action "folder permission not required for $userfullname"
    }else{
        $folderPermissionUser = $user.FolderPermissionUser
        $permissionUserObject = fetchUser -userfullname $folderPermissionUser
        $permissionUserDrivePath = $permissionUserObject.HomeDirectory
        # Provide folder access to disable user's personal drive path
        provideFolderAccess -folderPermissionUser $folderPermissionUser -PersonalDrivePath $personalDrivePath
        #add the shorcut to permission user's desktop point to the disable user
        AddShorcut -permissionUserDrivePath $permissionUserDrivePath -PersonalDrivePath $personalDrivePath
    }

    #Remove from DLs
    Remove-UserFromDistributionLists -UserIdentity $userObject
    # Disable the user and hide from address list
    try {
        Disable-ADAccount -Identity $userObject
	    Set-ADUser -identity $userObject -Replace @{msExchHideFromAddressLists=$true} -ErrorAction Stop
        Log-Action "Disabled user account for $userfullname"
    } catch {
        Log-Action "Failed to disable user account for $userfullname : $_"
    }
}
# Call the function to disconnect and log the action

Disconnect-ExchangeOnline -Session $Session
Write-Host "Processing completed. Log file created at $logFilePath"
