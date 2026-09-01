#############################
### Script for test connection string  ###
#############################
# -- Input params --
param (
	  [string]$QueryFile = "DefaultTestQuery.sql"
	, [string]$ServerName = "T_LSTNR_EDP.adgrtest.net"
	, [string]$PortNumber = "35040"
	, [string]$InstanceName = ""
	, [string]$DatabaseName = "model"
	, [string]$Encrypt = "no"
	, [string]$ApplicationName = "CheckConString"
	, [ValidateSet('SQLClient','OLEDB','ODBC')]
	  [string]$ConnectionProviderType = "SQLClient"
)	
# -- End Input params --

# Read file SQL for query
$strQueryText = Get-Content $QueryFile;

# Check if use port number or instance name
if ($InstanceName -ne "") { $InstancePortValue = ("\{0}" -f $InstanceName); }
else { $InstancePortValue = (",{0}" -f $PortNumber); }

# Check connection provider type
switch($ConnectionProviderType)
{
	#SQLNCLI
	"SQLClient"
	{
		# Initialze objects
		$objSqlConnection = New-Object System.Data.SqlClient.SqlConnection
		$objSqlCmd = New-Object System.Data.SqlClient.SqlCommand
		$objSqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		
		# Set connection string
		$objSqlConnection.ConnectionString = "Data Source={0}{1};Initial Catalog={2};Trusted_Connection=yes;Encrypt={3};Application Name={4}" -f $ServerName,$InstancePortValue,$DatabaseName,$Encrypt,$ApplicationName
		
		# Verbose
		Write-Verbose $objSqlConnection.ConnectionString; 

		# Set CommandText		
		$objSqlCmd.CommandText = $strQueryText
		
		# Set object connection for execute statement
		$objSqlCmd.Connection = $objSqlConnection
		
		# Set object command for get data
		$objSqlAdapter.SelectCommand = $objSqlCmd
	}
	
	#OLEDB
	"OLEDB"
	{
		# Initialze objects
		$objSqlConnection = New-Object System.Data.OleDb.OleDbConnection
		$objSqlCmd = New-Object System.Data.OleDb.OleDbCommand
		$objSqlAdapter = New-Object System.Data.OleDb.OleDbDataAdapter
		
		# Set connection string
		$objSqlConnection.ConnectionString = "Data Source={0}{1};Initial Catalog={2};Provider=SQLNCLI11.1;Integrated Security=SSPI;Auto Translate=False;Encrypt={3};Application Name={4}" -f $ServerName,$InstancePortValue,$DatabaseName,$Encrypt,$ApplicationName
		
		# Verbose
		Write-Verbose $objSqlConnection.ConnectionString; 

		# Set CommandText		
		$objSqlCmd.CommandText = $strQueryText
		
		# Set object connection for execute statement
		$objSqlCmd.Connection = $objSqlConnection
		
		# Set object command for get data
		$objSqlAdapter.SelectCommand = $objSqlCmd	
	}
	
	#ODBC
	"ODBC"
	{
		# Initialze objects
		$objSqlConnection = New-Object System.Data.Odbc.OdbcConnection
		$objSqlCmd = New-Object System.Data.Odbc.OdbcCommand
		$objSqlAdapter = New-Object System.Data.Odbc.OdbcDataAdapter
		
		# Set connection string (use placeholder for driver because there is issue with escape characters)
		$objSqlConnection.ConnectionString = "Driver={5};Server={0}{1};Database={2};Trusted_Connection=yes;Encrypt={3};Application Name={4}" -f $ServerName,$InstancePortValue,$DatabaseName,$Encrypt,$ApplicationName,'{SQL Server Native Client 11.0}'
		
		# Verbose
		Write-Verbose $objSqlConnection.ConnectionString; 

		# Set CommandText		
		$objSqlCmd.CommandText = $strQueryText
		
		# Set object connection for execute statement
		$objSqlCmd.Connection = $objSqlConnection
		
		# Set object command for get data
		$objSqlAdapter.SelectCommand = $objSqlCmd	
	}
}

try
{
	# Initialize Dataset object
	$objDataSet = New-Object System.Data.DataSet

	# Fill dataset with result query (out-null for don't print 1 value as output Fill method)
	$objSqlAdapter.Fill($objDataSet) | Out-Null;

	# Print out the columns
	foreach($objDataColumn in $objDataSet.Tables[0].Columns) { $strHeader += ("{0}|" -f $objDataColumn.ColumnName); }
	
	# Remove last pipe
	Write-Host $strHeader.substring(0,$strHeader.length-1);
	
	# cycle for each row
	foreach($objDataRow in $objDataSet.Tables[0].Rows)
	{	
		# Cycle for each columns and compose row string
		foreach($objDataColumn in $objDataSet.Tables[0].Columns) { $strRow += ("{0}|" -f $objDataRow[$objDataColumn.ColumnName]); }
		
		# Print out row without last pipe
		Write-Host $strRow.substring(0,$strRow.length-1);
	}

	# Print out success message
	Write-Host "Statement executed successfully!";
}
catch
{
	# Print out error message
	Write-Host ("[ERROR]: {0}" -f $_.Exception.Message);
}
