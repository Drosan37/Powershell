############################################
## Script for check entries on Hosts file ##
############################################
# ## Function for write output file
function WriteOutputFile
{
	# Function params
	param ( [string[]] $arrParams )
	
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

	# Write to file
	Add-Content -path $strOutputFile -value $strWriteOutput
 	
	# DEBUG
	#Write-Host $strWriteOutput 
}

# -- Custom Variables --
$strSourceDir = Split-Path -Path $MyInvocation.MyCommand.Path
$strSourceCsvFile = "${strSourceDir}\Check\ServerList.csv";
$strOutputFile = "${strSourceDir}\Check\HostsFiles.csv";
$strSeparator = ",";
$strPathHosts = "\\{0}\c$\Windows\System32\drivers\etc\hosts"
$strRegex = "^[\s\t]*[0-9]{1,4}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}"
$strServerName = "localhost"
# -- End Custom Variables --

# Initialize variables
$strDateNow=Get-Date -format "yyyyMMdd"
$strServerName = "";
$strErrorMsg = "";
$intFlagEntry = 0;

# Rewrite variable for add date to file name
$strOutputFile = "{0}{1}_{2}" -f $strOutputFile.Substring(0,$strOutputFile.LastIndexOf('\')+1),$strDateNow,$strOutputFile.Substring($strOutputFile.LastIndexOf('\')+1);

# Call function for write output in file (write header)
WriteOutputFile "ServerName","LineValue";

# Read csv file and save into array
$arrCsv = Import-Csv $strSourceCsvFile

# Cycle of array for get values
foreach($strRow in $arrCsv)
{
	# Get server name
	$strServerName = $strRow.ServerName

    # Initialize flag for entry
    $intFlagEntry = 0
			
	#DEBUG
	Write-Debug $strServerName	
	
	try
	{
        # Replace string with server name		
        $strServerPath = ($strPathHosts -f $strServerName)  

        # Cycle for each line into file hosts
        foreach($line in Get-Content $strServerPath) {
            if($line -match $strRegex){
                # Work here
                Write-Host ("{0} - {1}" -f $strServerName, $line)
                WriteOutputFile $strServerName,$line;

                # Set flag for indicate there is an entry row
                $intFlagEntry = 1
            }
        } 

        # If 0, write empy row for server
        if($intFlagEntry -eq 0)
        {
            WriteOutputFile $strServerName,"NoEntry";
        }                
	}
	catch
	{
		# Set error message
		$strErrorMsg = "ERROR on Execute command for server ${strServerName}:`r`n$($_.Exception.Message)"; 
	
		# Print out to Console
		Write-Host $strErrorMsg;
	}
}	