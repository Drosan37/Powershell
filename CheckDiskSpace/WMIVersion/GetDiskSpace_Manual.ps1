# Define the file path where server names are stored
$serverListFile = "F:\SQLDiscovery\CheckDiskSpace\Servers.txt"  # Update this path to your actual file location
$csvOutputFile = "F:\SQLDiscovery\CheckDiskSpace\Outlist.csv"

# Read server names from the file
$servers = Get-Content $serverListFile

# Create an empty array to store the results
$diskInfo = @()

# Set culture for avoid problems with decimal
$culture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")

# Loop through each server and retrieve disk space for mounted volumes
foreach ($server in $servers) {
    $volumes = Get-WMIObject Win32_Volume -ComputerName $server | Where-Object { $_.DriveType -eq 3 -and $_.Name -match "^[C F-Z]:\\(?!Store98|Store99|TempDB).*" }

    foreach ($volume in $volumes) {
        # Create custom object as row of array
        $objArrRow = [PSCustomObject]@{            
            ServerName     = $server            
            MountPoint     = $volume.Name
            TotalSizeGB    = "{0}" -f ($volume.Capacity / 1GB).ToString("0.00",$culture)
            FreeSpaceGB    = "{0}" -f ($volume.FreeSpace / 1GB).ToString("0.00",$culture)
            FreePercentage = "{0}" -f (($volume.FreeSpace / $volume.Capacity) * 100).ToString("0.00",$culture) 
        }

        # Add row to array
        $diskInfo += $objArrRow
    }
}

# Sort the results by Free Percentage (ascending) and format output
$diskInfo | Sort-Object ServerName, {[double]($_.FreePercentage -replace '%', '')} | Format-Table -AutoSize

# Export to CSV file
$diskInfo | Export-Csv -Path $csvOutputFile -NoTypeInformation

# Display confirmation message
Write-Host "Disk space report saved to: $csvOutputFile"
