#########################################
## Powershell Module for Log Functions ##
#########################################
function InitFileLog
{
    # Function params
	param ( 
        [Parameter(Mandatory)]
        [string] $PathLog,    

        [Parameter(Mandatory)]
        [string] $FileName
    )

    try
    {
        # Create a new file
        New-Item -Path $PathLog -Name $FileName -ErrorAction Stop | Out-Null
    }
    catch [System.IO.IOException]
    {
        # Esiste, perciò non lo creo
    }
}

function WriteErrorLog
{
	# Function params
	param ( 
        [Parameter(Mandatory)]
        [string] $FullPathLogFile,    

        [Parameter(Mandatory)]    
        [string] $StringToWrite 
    )

	# Write to file
	Add-Content -Path $FullPathLogFile -Value $StringToWrite
}

# Export module
Export-ModuleMember -Function WriteErrorLog, InitFileLog 