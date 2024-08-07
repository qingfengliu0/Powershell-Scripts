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
function Get-UserInput {
    param (
        [string]$Prompt
    )
    Read-Host -Prompt $Prompt
}

function fetchUser {
    param (
        [string]$userfullname,
        [string]$typeOfUser
    )
    
    # Split the entry into LastName and FirstName
    $FirstName, $Lastname = $userfullname -split " "
    
    # Trim any whitespace from the names
    $FirstName = $FirstName.Trim()
    $LastName = $LastName.Trim()
    
    # Function to handle duplicates and prompt user to choose
    function HandleDuplicates {
        param (
            [array]$users
        )

        if ($users.Count -eq 0) {
            Write-Host "No users found."
            return $null
        }

        # Present the list of users
        PresentUserList -users $users

        Write-Host "Multiple users found. Please choose the correct $typeOfUser by entering the corresponding SamAccountName:"
    
        $selectedIndex = Read-Host "Enter the SamAccountName of the $typeOfUser you want to select"
        $selectedUser = $users | Where-Object { $_.SamAccountName -eq $selectedIndex }
        
        if ($selectedUser) {
            $ADUser = Get-ADUser -Identity $selectedUser.SamAccountName -Properties *
            return $ADUser
        } else {
            Write-Host "Invalid selection. No user found with SamAccountName '$selectedIndex'."
            return $null
        }
    }

    try {
        $user = $null

        # Try finding by FirstName and LastName
        if ($LastName -and $FirstName) {
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
            $user = @(Get-ADUser -Filter "GivenName -eq '$($FirstName)'" -Properties *)
            $count = $user.Count
            Log-Action "Found $count user(s) with the first name: $FirstName"
            PresentUserList -users $user
            if ($count -eq 1) {
                return $user
            } elseif ($count -gt 1) {
                return HandleDuplicates -users $user
            }
        }

        # If still not found, try finding by LastName only
        if (-not $user) {
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

function Log-Action {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Add-Content -Path $logFilePath -Value $logMessage
    Write-Host $logMessage
}

function Copy-UserHomeDirectory {
    param (
        [string]$sourceDir,
        [string]$destinationDir
    )
    
    if (-Not (Test-Path -Path $sourceDir -PathType Container)) {
        Write-Host "Source directory does not exist."
        exit 1
    }

    if (-Not (Test-Path -Path $destinationDir -PathType Container)) {
        New-Item -ItemType Directory -Path $destinationDir
    }

    Copy-Item -Path $sourceDir -Destination $destinationDir -Recurse -Force

    Log-Action "Home directory copied successfully from $sourceDir to $destinationDir."
}

function Archive-Emails {
    param (
        [string]$username,
        [string]$destinationDir
    )
    
    $userObject = Get-ADUser -Identity $username -Properties mail, Surname, GivenName
    $userEmail = $userObject.mail.ToString()
    $dateString = (Get-Date).ToString("yyyyMMdd")
    $pstFileName = "$username$dateString.pst"
    $tempPstPath = "\\Cecelia\c$\temp\$pstFileName"
    $destinationPstPath = Join-Path -Path $destinationDir -ChildPath $pstFileName
    
    # Create a mailbox export request
    New-MailboxExportRequest -Mailbox $userEmail -FilePath $tempPstPath -Name "$username$dateString"
    
    # Wait for the export request to complete
    while (($exportStatus = Get-MailboxExportRequest -Name "$username$dateString").Status -ne "Completed") {
        Write-Host "Waiting for mailbox export to complete..."
        Start-Sleep -Seconds 30
    }
    
    # Move the PST file to the archive location
    Move-Item -Path $tempPstPath -Destination $destinationPstPath -Force
    
    Log-Action "Emails archived successfully to $destinationPstPath."
}

$archiveBaseDir = "\\Empress\Payroll\Former Employees\Other"

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

do {
    $userfullname = Get-UserInput -Prompt "Enter the full name of the user to be archived (Firstname Lastname) or type 'exit' to quit"
    if ($userfullname -eq 'exit') { break }

    $userObject = fetchUser -userfullname $userfullname -typeOfUser "archive user"
    if ($userObject -eq $null) {
        Log-Action "Can't find user in the Active Directory: $userfullname"
        continue
    } elseif ($userObject.Enabled -eq $true) {
        Write-Host "The selected user account is not disabled. Please enter a different user."
        continue   
    } else {
        Write-Host "The selected user account is confirmed as Disabled."
        
        $archiveDir = Join-Path -Path $archiveBaseDir -ChildPath "$($userObject.Surname), $($userObject.GivenName)"
        
        try {
            Copy-UserHomeDirectory -sourceDir $userObject.HomeDirectory -destinationDir $archiveDir
            Archive-Emails -username $userObject.SamAccountName -destinationDir $archiveDir
            Invoke-Item -Path $archiveDir
            Log-Action "Successfully archived user data for $($userObject.GivenName) $($userObject.Surname)."
        } catch {
            Log-Action "An error occurred while archiving data for user $($userObject.GivenName) $($userObject.Surname): $_"
        }
    }
} while ($true)

        
# Call the function to disconnect and log the action
Disconnect-ExchangeOnline -Session $Session
Write-Host "Processing completed. Log file created at $logFilePath"
