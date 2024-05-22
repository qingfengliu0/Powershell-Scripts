# Define a function to test ping to a URL
function Test-Ping {
    param(
        [string]$url
    )

    # Use Test-Connection cmdlet to ping the URL
    $pingResult = Test-Connection -ComputerName $url -Count 4 -ErrorAction SilentlyContinue

    # Check if the ping was successful
    if ($pingResult) {
        Write-Output "$url is reachable. Average ping time: $($pingResult.ResponseTime) ms"
    }
    else {
        Write-Output "$url is unreachable."
    }
}

# Main script
$file = "urls.txt"  # Change this to the name of your text file containing URLs

# Read the URLs from the text file
$urls = Get-Content $file

# Loop through each URL and test ping
foreach ($url in $urls) {
    $url = $url.Trim()  # Remove any leading or trailing whitespace

    if ($url -ne "") {  # Skip empty lines
        Test-Ping $url
    }
}