#Bulk remove user bat file
$UserlistPath = "userscriptremove.txt"
function Remove-ProfileItem {
    param (
        [string]$path
    )

    if (Test-Path $path) {
        try {
            # Use Robocopy to delete the directory or file
            Robocopy Emptyfolder $path  /purge /e
            Remove-Item -path $path
            Write-Host "Deleted item at path: $path" |Out-File -FilePath UsersDeleteResult.txt -Append
        } catch {
            Write-Host "Error deleting item at path: $path - $_" |Out-File -FilePath UsersDeleteResult.txt -Append
        }
    } else {
        Write-Host "Item not found at path: $path" |Out-File -FilePath UsersDeleteResult.txt -Append
    }
}

function fetchUser{
    param (
        [string]$userfullname
    )
    # Split each entry into name and username
    $FirstName, $LastName = $userfullname -split " "
    
    $user = Get-ADUser -Filter "GivenName -eq '$($FirstName)' -and Surname -eq '$($LastName)'" -Properties *
    return $user
}

if (Test-Path $UserListPath) {
    # Read the list of names and usernames from the file
    $UserList = Get-Content $UserListPath
    foreach ($user in $UserList) {
    $userObject = fetchUser -userfullname $user
    # what the netlogon should be
    $LogonScriptName = "$($userObject.samAccountName).bat"

    # what is the actual netlogon script 
    $ActualLogonScriptPath = $userObject.ScriptPath
    $ActualLogonScriptName = ($ActualLogonScriptPath -split "\\")[-1]
    #if and only if they the name remove the script otherwise leave it and write down the actual path 
        if($ActualLogonScriptName -eq $LogonScriptName){
            Remove-ProfileItem -path $ActualLogonScriptPath
            Write-Output "$user netlogon removed from path $ActualLogonScriptPath" |Out-File -FilePath UsersDeleteResult.txt -Append
        }else{
            write-output "the path $ActualLogonScriptPath does not look correct, $user's logon script not removed "|Out-File -FilePath UsersDeleteResult.txt -Append
        }
    }
}