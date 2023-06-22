# Define the list of servers to monitor and patch
$servers = "Server1", "Server2", "Server3"

# Loop through each server
foreach ($server in $servers) {
    Write-Host "Monitoring availability and applying security patches for server: $server"

    # Check server availability
    $pingResult = Test-Connection -ComputerName $server -Count 1 -Quiet

    if ($pingResult) {
        Write-Host "Server $server is available. Proceeding with patching..."

        # Get the list of available security patches
        $patches = Get-HotFix -ComputerName $server | Where-Object { $_.Description -like "*Security Update*" }

        if ($patches) {
            Write-Host "Found security patches on server $server."

            # Install security patches
            Write-Host "Installing security patches on server $server..."
            $patches | ForEach-Object {
                Write-Host "Installing patch: $($_.HotFixID)"
                Install-HotFix -ComputerName $server -Id $_.HotFixID -Confirm:$false
            }

            Write-Host "Security patches installed successfully on server $server."
        }
        else {
            Write-Host "No security patches found on server $server."
        }
    }
    else {
        Write-Host "Server $server is not available. Skipping patching."
    }

    # Display completion message for the server
    Write-Host "Availability monitoring and patching completed for server: $server"
    Write-Host "-----------------------------------------------"
}
