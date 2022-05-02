#########################################
## Powershell Module for Xml Functions ##
#########################################
## Function for read all xml file
function ReadAllXmlFile {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    # Read file and get Xml object
    [xml]$objXmlFile = Get-Content -Path $FilePath

    # Return xml object
    return $objXmlFile
}

# Export module
Export-ModuleMember -Function ReadAllXmlFile