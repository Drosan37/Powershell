# Define the file path where server names are stored
$serverListFile = "F:\SQLDiscovery\CheckDiskSpace\ServersTest.txt"  # Update this path to your actual file location
$csvOutputFile = "F:\SQLDiscovery\CheckDiskSpace\Outlist.csv"

# Initialize variables	
$intQueryTimeout = 600;

# Read server names from the file
$servers = Get-Content $serverListFile

# Create an empty array to store the results
$diskInfo = @()

# Set culture for avoid problems with decimal
$culture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")

# Define query for retrieve disks spaces
$strQueryDisksSpaces = ("
    ;WITH DBFiles AS
    (
        SELECT DISTINCT
            LEFT(physical_name COLLATE SQL_Latin1_General_CP1_CI_AS, 1) AS DriveLetter
        FROM sys.database_files
        UNION
        SELECT DISTINCT
            LEFT(physical_name COLLATE SQL_Latin1_General_CP1_CI_AS, 1)
        FROM sys.master_files
    )
    SELECT
        vs.volume_mount_point AS MountPoint,
        CAST(vs.total_bytes / 1024.0 / 1024 / 1024 AS DECIMAL(10,2)) AS TotalSizeGB,
        CAST(vs.available_bytes / 1024.0 / 1024 / 1024 AS DECIMAL(10,2)) AS FreeSpaceGB,
        CAST((vs.total_bytes - vs.available_bytes) / 1024.0 / 1024 / 1024 AS DECIMAL(10,2)) AS UsedGB,
        CAST(vs.available_bytes * 100.0 / vs.total_bytes AS DECIMAL(5,2)) AS FreePercentage
    FROM sys.master_files f
    CROSS APPLY sys.dm_os_volume_stats(f.database_id, f.file_id) vs
    WHERE LEFT(vs.volume_mount_point, 1) IN (SELECT DriveLetter FROM DBFiles)
    AND vs.volume_mount_point not like '%Store9_%'
    GROUP BY
        vs.volume_mount_point,
        vs.total_bytes,
        vs.available_bytes
    ORDER BY MountPoint;
")

# Loop through each server and retrieve disk space for mounted volumes
foreach ($instance in $servers) 
{
    try		
    {		
	    # Execute query to retrieve server list	
	    $arrDiskSpaceInfo = Invoke-Sqlcmd $strQueryDisksSpaces -ServerInstance $instance -Database "master" -QueryTimeout $intQueryTimeout -ErrorAction 'Stop';	
    }		
    catch		
    {		
	    # Set error message	
	    $strErrorMsg = "ERROR on Execute query for retrieve info:`r`n$($_.Exception.Message)"; 	

	    # Write Error Log	
	    Write-Host $strErrorMsg;	
    }

}

# Sort the results by Free Percentage (ascending) and format output
$arrDiskSpaceInfo | Sort-Object ServerName, {[double]($_.FreePercentage -replace '%', '')} | Format-Table -AutoSize

# Export to CSV file
$arrDiskSpaceInfo | Export-Csv -Path $csvOutputFile -NoTypeInformation

# Display confirmation message
Write-Host "Disk space report saved to: $csvOutputFile"
