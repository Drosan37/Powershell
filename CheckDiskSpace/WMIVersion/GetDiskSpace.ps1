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
$csvOutputFile = "F:\SQLDiscovery\CheckDiskSpace\Outlist.csv"
$strErrorLogFile = "F:\SQLDiscovery\CheckDiskSpace\Logs\{0}_ErrorCheckDiskSpace.log"

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
    WHERE ADDomain = 'ADGR' 
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
        
    $serverFQDN = ("{0}.{1}.net" -f $server.ServerName,$server.ADDomain)
    Write-host $serverFQDN

    $volumes = Get-WMIObject Win32_Volume -ComputerName $serverFQDN | Where-Object { $_.DriveType -eq 3 -and $_.Name -match "^[C F-Z]:\\(?!Store98|Store99|TempDB).*" }    
    
    foreach ($volume in $volumes) {
        # Create custom object as row of array
        $objArrRow = [PSCustomObject]@{
            ADDomain       = $server.ADDomain
            ServerName     = $server.ServerName            
            MountPoint     = $volume.Name
            TotalSizeGB    = "{0}" -f ($volume.Capacity / 1GB).ToString("0.00",$culture)
            FreeSpaceGB    = "{0}" -f ($volume.FreeSpace / 1GB).ToString("0.00",$culture)
            FreePercentage = "{0}" -f (($volume.FreeSpace / $volume.Capacity) * 100).ToString("0.00",$culture) 
        }

        # Add row to array
        $diskInfo += $objArrRow
        
        # Add query insert to string
        $queryInsert += "
        INSERT INTO [CHK].[DiskSpace] ([ADDomain],[ServerName],[MountPoint],[TotalSizeGB],[FreeSpaceGB],[FreePercentage])
            VALUES ('{0}','{1}','{2}',{3},{4},{5});
        " -f $objArrRow.ADDomain, $objArrRow.ServerName, $objArrRow.MountPoint, $objArrRow.TotalSizeGB, $objArrRow.FreeSpaceGB, $objArrRow.FreePercentage      
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
