$UserCredential = Get-Credential
Connect-ExchangeOnline -UserPrincipalName youradmin@domain.com -Credential $UserCredential

# Parameters (replace with actual values)
$NewFullName = "Lorie Doolan"
$oldUsername = "LHarvey"
# Split the full name into parts
$nameParts = $NewFullName.Split(" ")
# Get the first initial and last name
$Firstname = $nameParts[0]
$firstInitial = $nameParts[0][0]
$lastName = $nameParts[1]
$name = "$lastName, $Firstname"
# Combine them to form the username
$NewUsername = "$firstInitial$lastName"
$Domain = "pacificahousing.ca"
$UserDN = "OU=Regular,OU=Users,OU=Pacifica Housing,DC=pacifica,DC=local"

# 1. Reset Password
#Set-ADAccountPassword -Identity $UserDN -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "Welcome2Pacifica" -Force)

# 2. Change Name in AD
$User = Get-ADUser -Identity $OldUsername

# Rename user and update attributes
Set-ADUser -Identity $User -SamAccountName $NewUsername -GivenName $Firstname -Surname $lastName -UserPrincipalName "$NewUsername@pacifica.local" 
$User = Get-ADUser -Identity $NewUsername
# 3. Update Terminal Server Profile Path (assuming TSProfiles are on \\Leo\TSProfiles)
$TSProfilePath = "\\Leo\TSProfiles\$NewUsername"
Set-ADUser -Identity $User -ProfilePath $TSProfilePath

# 4. Update User Profile Path (assuming profiles are on \\Leo\Profiles)
$ProfilePath = "\\Leo\Profiles\$NewUsername"
Set-ADUser -Identity $User -ProfilePath $ProfilePath

# 5. Update Home Folder Path (assuming home folders are on \\Leo\Users)
$HomeFolderPath = "\\Leo\Users\$NewUsername"
Set-ADUser -Identity $User -HomeDirectory $HomeFolderPath

# 6. Update Exchange Mailbox
$Mailbox = Get-Mailbox -Identity $OldUsername
Set-Mailbox -Identity $Mailbox -Alias $NewUsername
Set-Mailbox -Identity $Mailbox -PrimarySmtpAddress "$NewUsername@$Domain"

# Add new email addresses and retain old ones
$OldEmailAddress = "$OldUsername@$Domain"
$NewEmailAddress = "$NewUsername@$Domain"
Set-Mailbox -Identity $Mailbox -EmailAddresses @{Add = $NewEmailAddress}

# 7. Delete Existing Profiles from PCs (requires running on each PC)
# List of PCs the user has logged into
$PCList = @("PC1", "PC2", "PC3")
foreach ($PC in $PCList) {
    Invoke-Command -ComputerName $PC -ScriptBlock {
        param($OldUsername)
        Start-Process -FilePath "C:\Utilities\DelProf2.exe" -ArgumentList "/c:$env:COMPUTERNAME /p /id:$OldUsername" -Wait -NoNewWindow
    } -ArgumentList $OldUsername
}

# Close Exchange Online connection
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "User renaming process completed successfully."
