# Script for gather information about disk size on all PROD servers
Import-Module SqlServer -Version 21.1.18209			
			
# ## Function to write error log
function WriteErrorLog
{
	# Function params
	param ( [string] $StringToWrite )

	# Write to file
	Add-Content -path $strErrorLogFile -value $StringToWrite
}	

# Define the file path where server names are stored
$strErrorLogFile = "F:\SQLDiscovery\CheckDiskSpace\Logs\ErrorCheckDiskSpace.log"

# Set date now
$strDateNow = Get-Date -format "yyyyMMdd"

# Set errorlog file name
$strErrorLogFile = "{0}{1}_{2}" -f $strErrorLogFile.Substring(0,$strErrorLogFile.LastIndexOf('\')+1),$strDateNow,$strErrorLogFile.Substring($strErrorLogFile.LastIndexOf('\')+1);

# Initialize variables	
$intQueryTimeout = 600;	

# Define query for retrieve server list
$strQueryServerList = ("    
    SELECT DISTINCT
        ServerName
      , ADDomain
	  , PortNumber
	  , InstanceName
    FROM CHK.GetServerList
    --WHERE ADDomain = 'ADGR' 
")

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

# Create an empty array to store the results
$diskInfo = @()

# Set culture for avoid problems with decimal
$culture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")

# Defile string for query insert, start with Truncate for remove old data
$queryInsert = "TRUNCATE TABLE [CHK].[DiskSpace];"

try		
{		
	# Execute query to retrieve server list	
	$arrServerList = Invoke-Sqlcmd $strQueryServerList -ServerInstance "IDERADB,45000" -Database "SqlServerMap" -QueryTimeout $intQueryTimeout -ErrorAction 'Stop';	
}		
catch		
{		
	# Set error message	
	$strErrorMsg = "ERROR on Execute query for retrieve server List:`r`n$($_.Exception.Message)"; 	
		
	# Write Error Log	
	WriteErrorLog $strErrorMsg;	
}	

# Loop through each server and retrieve disk space for mounted volumes
foreach ($server in $arrServerList) {
    # Set string with servername and addomain    
    $serverFQDN = ("{0}.{1}.net" -f $server.ServerName,$server.ADDomain)

    # Print server where try to connect 
    Write-Host $serverFQDN -ForegroundColor Yellow

    # Set string for instance connection
    $strInstance = ("{0},{1}" -f $serverFQDN,$server.PortNumber)
    
    try		
    {		
	    # Execute query to retrieve server list	
	    $arrDiskSpaceInfo = Invoke-Sqlcmd $strQueryDisksSpaces -ServerInstance $strInstance -Database "master" -QueryTimeout $intQueryTimeout -ErrorAction 'Stop';	

        # Cycle for result array
        foreach($objArrRow in $arrDiskSpaceInfo)
        {
             # Add query insert to string
            $queryInsert += "
            INSERT INTO [CHK].[DiskSpace] ([ADDomain],[ServerName],[MountPoint],[TotalSizeGB],[FreeSpaceGB],[FreePercentage])
                VALUES ('{0}','{1}','{2}',{3},{4},{5});
            " -f $server.ADDomain, $server.ServerName, $objArrRow.MountPoint, $objArrRow.TotalSizeGB, $objArrRow.FreeSpaceGB, $objArrRow.FreePercentage 
        }
    }		
    catch		
    {		
	    # Set error message	
	    $strErrorMsg = "ERROR on Execute query for retrieve info:`r`n$($_.Exception.Message)"; 	

	    # Write Error Log	
	    WriteErrorLog $strErrorMsg;	
    }	      
}

try		
{		
	# Execute query to retrieve server list	
	Invoke-Sqlcmd -Query $queryInsert -ServerInstance "IDERADB,45000" -Database "DBAdmin" -QueryTimeout $intQueryTimeout -ErrorAction 'Stop';	
}		
catch		
{		
	# Set error message	
	$strErrorMsg = "ERROR on Execute query for retrieve server List:`r`n$($_.Exception.Message)"; 	
		
	# Write Error Log	
	WriteErrorLog $strErrorMsg;	
}

# Display confirmation message
Write-Host "Disk space report saved to DBAdmin Database on IDERADB"
