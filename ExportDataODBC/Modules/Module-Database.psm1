##########################################################
## Powershell Module for Database Functions for ODBC ##
##########################################################
function OpenConnection {
    param(
        [Parameter(Mandatory)]
        [string]$DsnName
    )

    # Set connection string (trusted auth)
    $strConnString = 'DSN={0}' -f $DsnName

    # Create object connection
    $objSqlConnection = New-Object System.Data.Odbc.OdbcConnection $strConnString
    
    try
    {    
        # Open connection
        $objSqlConnection.Open()    

        # Return SqlConnection
        return $objSqlConnection
    }
    catch [System.Data.SqlClient.SqlException] 
    {
        # DEBUG
        Write-Debug $_.Exception

        # Throw Exception
        throw
    }
}

function CloseConnection {
    param(
        [Parameter(Mandatory)]
        [System.Data.Odbc.OdbcConnection]$SqlConnection
    )

    # Close connection (object passed) and dispose
    $SqlConnection.Close()
    $SqlConnection.Dispose()

    # Force Garbage Collector
    [System.GC]::Collect()    
}

function ExecuteReader {
    param(
        [Parameter(Mandatory)]
        [System.Data.Odbc.OdbcConnection]$SqlConnection,

        [Parameter(Mandatory)]
        [string]$CommandText
    )
    
    try {		
        # Create object SqlCommand		
		$objSqlcommand = New-Object System.Data.Odbc.OdbcCommand($CommandText,$SqlConnection)
        $objSqlcommand.CommandTimeout = 600
        		
        # Create object SqlAdapter		
		$objSqlAdapter = New-Object System.Data.Odbc.OdbcDataAdapter $objSqlcommand
		
        # Create object Dataset    
        $objDataset = New-Object System.Data.DataSet
		
        # Execute query and fill dataset (out-null avoid record number print)
        $objSqlAdapter.Fill($objDataset) | Out-Null
		
        # Return DataTable
		return $objDataset.Tables 				
    } 
    catch [System.Data.Odbc.OdbcException] {
        # DEBUG
        Write-Debug $_.Exception.Message
    
        # Throw Exception
        throw
    }
    finally {
        
    }
}

function ExecuteBulk {
    param(
        [Parameter(Mandatory)]
        [System.Data.Odbc.OdbcConnection]$SqlConnection,

        [Parameter(Mandatory)]
        [System.Data.DataTable]$DataTable,

        [Parameter(Mandatory)]
        [string]$DestTableName
    )

    try {		
        # Create object SqlBulk
		$objSqlBulkCopy = New-Object System.Data.Odbc.OdbcBulkCopy $SqlConnection

        # Set destination table
        $objSqlBulkCopy.DestinationTableName = $DestTableName
        
        # Execute bulk insert, with datatable data
        $objSqlBulkCopy.WriteToServer($DataTable)
    } catch [System.Data.Odbc.OdbcException] {
        # DEBUG
        Write-Debug $_.Exception.Message

        # Throw Exception
        throw
    } finally {

    }
}

function ExecuteNonQuery {
    param(
        [Parameter(Mandatory)]
        [System.Data.Odbc.OdbcConnection]$SqlConnection,

        [Parameter(Mandatory)]
        [string]$CommandText
    )
    
    try {		
        # Create object SqlCommand		
		$objSqlcommand = New-Object System.Data.Odbc.Odbccommand($CommandText,$SqlConnection)
		
        # Execute query 
        $objSqlcommand.ExecuteNonQuery() | Out-nUll			
    } catch [System.Data.Odbc.OdbcException] {
        # DEBUG
        Write-Debug $_.Exception.Message

        # Throw Exception
        throw
    } finally {
        
    }
}

# Export module
Export-ModuleMember -Function OpenConnection, CloseConnection
Export-ModuleMember -Function ExecuteReader, ExecuteBulk, ExecuteNonQuery