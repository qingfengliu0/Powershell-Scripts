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


if (Test-Path $UserListPath) {
    # Read the list of names and usernames from the file
    $UserList = Get-Content $UserListPath
    foreach ($fullname in $UserList) {
        # Split the full name into an array of words
        $nameParts = $fullname -split " "
        # Check if the input has at least two parts (first name and last name)
        if ($nameParts.Length -ge 2) {
            # Extract the first initial of the first name
            $firstInitial = $nameParts[0][0]

            # Extract the last name (assuming it's the last part of the input)
            $lastName = $nameParts[-1]

            # Combine into the desired format
            $netlogonfilename = "$firstInitial$lastName.bat"
            $nelogonfilepath 

           

            Write-Output "Created file: $outputFileName"
        } else {
            Write-Output "Please enter a full name with at least a first name and a last name."
        }
   
    }      
}