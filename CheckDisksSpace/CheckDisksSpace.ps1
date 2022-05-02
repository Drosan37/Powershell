#################################
## Script for check disk space ##
#################################
# Input Parameter
param ( 
      [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
      [ValidateSet('File','Console')]
      [string] $OutputType = "Console"
    , [string] $IncludeVolumes = $false 
)

# ## Function to write error log
function WriteErrorLog
{
	# Function params
	param ( [string] $StringToWrite )

	# Write to file
	Add-Content -path $strErrorLogFile -value $StringToWrite
}

# ## Function for write output file
function WriteOutputFile
{
	# Function params
	param ( [string[]] $arrParams, $Initialize = $false )
	
	# Initialize variables
	$strWriteOutput = "";
		
	# Cycle for array params
	foreach($strArg in $arrParams)
	{		
		# Add parameters to string for output
		$strWriteOutput += "{0}{1}" -f $strArg,$strSeparator
	}
		
	# Remove last pipe
	$strWriteOutput = $strWriteOutput.substring(0,$strWriteOutput.length-1)
	
    # Check if need to reinitialize file
    if(($Initialize -eq $true) -and (Test-Path -path $strOutputFile))
    {
        # Write to file but reinitialize
        Clear-Content -path $strOutputFile
    }

	# Write to file
	Add-Content -path $strOutputFile -value $strWriteOutput

	# DEBUG
	Write-Debug $strWriteOutput 
}

# -- Custom Variables --
$strSourceDir = $strSourceDir = Split-Path -Path $MyInvocation.MyCommand.Path;
$strSourceCsvFile = "${strSourceDir}\Source\ServerList.csv";
$strOutputFile = "${strSourceDir}\Output\DisksSpace.csv";
$strErrorLogFile = "${strSourceDir}\Logs\ErrorCheck.log";
$strSeparator = ";";
$intQueryTimeout = 20;
# -- End Custom Variables --

# Initialize variables
$strDateNow=Get-Date -format "yyyyMMdd"
$strServerName = "";
$strErrorMsg = "";

# Check param for volumes
if($IncludeVolumes)
{
    # Filter for get only disk
    $strFilter = "DriveType = 3 Or DriveType = 4";
}
else
{
    # Get all
    $strFilter = ""
}

#Debug
Write-Debug $outputType;

# Rewrite variable for add date to file name
$strOutputFile = "{0}{1}_{2}" -f $strOutputFile.Substring(0,$strOutputFile.LastIndexOf('\')+1),$strDateNow,$strOutputFile.Substring($strOutputFile.LastIndexOf('\')+1);
$strErrorLogFile = "{0}{1}_{2}" -f $strErrorLogFile.Substring(0,$strErrorLogFile.LastIndexOf('\')+1),$strDateNow,$strErrorLogFile.Substring($strErrorLogFile.LastIndexOf('\')+1);

if($OutputType -eq "Console")
{
    # Command for retrieve size and free space about server volume	
    Write-Host "#########################" 
				"### Check disks space ###" 
				"#########################"

    # Print out header row
    Write-Host "ServerName - DiskName - Size(GB) - FreeSpace(GB)"
}
else
{
    # Call function for write output in file (write header)
    WriteOutputFile "ServerName","DiskName","Size(GB)","FreeSpace(GB)" -Initialize $true;
}

# Read csv file and save into array
$arrCsv = Import-Csv $strSourceCsvFile

# Cycle of array for get values
foreach($strRow in $arrCsv)
{
	# Get server name
	$strServerName = $strRow.ServerName
			
	#DEBUG
	Write-Debug $strServerName	
	
	try
	{			
		# Execute WMI query for check Volumes
		$objWmi = Get-WmiObject -Namespace "root\cimv2" -ComputerName $strServerName -Class Win32_LogicalDisk -filter $strFilter | 
			select Name, @{LABEL='Size(GB)';EXPRESSION={"{0:N2}" -f ($_.Size/1GB)} } , @{LABEL='FreeSpace(GB)';EXPRESSION={"{0:N2}" -f ($_.FreeSpace/1GB)} } | Sort-Object Name	

		#DEBUG
		Write-Debug $objWmi.Count
		
		#Cycle for output
		foreach($strDisk in $objWmi)
		{
			if($outputType -eq "Console") 
				{ Write-Host ("{0} - {1} - {2} - {3}" -f $strServername, $strDisk.Name, $strDisk."Size(GB)", $strDisk."FreeSpace(GB)") }
			else
				{ WriteOutputFile $strServerName, $strDisk.Name, $strDisk."Size(GB)", $strDisk."FreeSpace(GB)" }
		}
	}
	catch
	{
		# Set error message
		$strErrorMsg = "ERROR on Execute command for server ${strServerName}:`r`n$($_.Exception.Message)"; 
	
		# Print out to Console
		#Write-Host $strErrorMsg;
		
		# Write Error Log
		WriteErrorLog $strErrorMsg;
	}
}	