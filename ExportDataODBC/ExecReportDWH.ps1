#####################################
## Powershell - Report DWH Utility ##
#####################################
# Script for read data from Vertica DWH
param (
      [string]$ReportName = 'NONE'
    , [String[]]$QueryParams
)

# Import Module
Import-Module .\Modules\Module-Database.psm1 -Force
Import-Module .\Modules\Module-Xml.psm1 -Force
Import-Module .\Modules\Module-Log.psm1 -Force

# Define variables
$strRootPath = Split-Path -Path $MyInvocation.MyCommand.Path
$strQueriesPath = "$strRootPath\Queries\"
$strOutputPath = "$strRootPath\Output\"
$strLogPath = "$strRootPath\Logs\"
$intQueryTimeout = 20;
$strDateNow = Get-Date -format "yyyyMMdd"
$strServerName = "";
$strErrorMsg = "";

# DEBUG
Write-Debug $strRootPath

# Read Xml file config
$objXmlConf = ReadAllXmlFile -FilePath "$strRootPath\Config\Config.xml"

# Set variable after read config Xml
$strLogFileName = ("{0}_{1}" -f $strDateNow, $objXmlConf.ConfigFile.LogInfo.FileName)
$strLogFile = ("{0}\{1}" -f $strLogPath, $strLogFileName)

# Init file log
InitFileLog -PathLog $strLogPath -FileName $strLogFileName

try
{
    # Open connection with source database
    $objSqlConnSource = OpenConnection -DsnName $objXmlConf.ConfigFile.SourceInstance.DSN 
}
catch
{
    # Write Log
    WriteErrorLog -FullPathLogFile $strLogFile -StringToWrite ("{0}: {1}" -f $strInstance, $_.Exception.Message)
}

# Cycle for all Type tag
foreach($objReport in $objXmlConf.ConfigFile.Reports.ChildNodes)
{
    # DEBUG
    Write-Debug ("{0} - {1}" -f $objReport.Query, $objReport.FileOutput) 

    # Check report name for execute
    if($objReport.Name -eq $ReportName)
    {
        try
        {
            # Define sql command text (raw for get in one row)
            $strSqlCommandText = Get-Content ("$strRootPath\Queries\{0}" -f $objReport.Query) -Raw
            
            # Check if number of param expected is equal than passed params (array count)
            if($QueryParams.Count -eq $objReport.ParamsNumber)
            {
                # Replace placeholder into query text
                $strSqlCommandText = $strSqlCommandText -f $QueryParams
            }
            else
            {
                # Force an error
                throw "Different number of params"
            }            

            # DEBUG
            Write-Debug $strSqlCommandText
        
            # Execute query for gather data
            $objDataTable = ExecuteReader -SqlConnection $objSqlConnSource -CommandText $strSqlCommandText

            # Export data to CSV file
            $objDataTable | Export-Csv ("$strOutputPath\{0}" -f $objReport.FileOutput) -NoTypeInformation -Force
            
        }
        catch
        {
            # DEBUG 
            Write-Debug $_.Exception

            # Write Log
            WriteErrorLog -FullPathLogFile $strLogFile -StringToWrite ("Report {0}: {1}" -f $ReportName, $_.Exception.Message)
        }
    }
}

# Close connection with source database
CloseConnection $objSqlConnSource    

# If log is empty, delete it
if([string]::IsNullOrWhiteSpace((Get-content ("{0}\{1}" -f $strLogPath, $strLogFileName)))) { Remove-Item -Path ("{0}\{1}" -f $strLogPath, $strLogFileName) }