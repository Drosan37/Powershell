# Script for create SQL Group in specific OU
param
(
      [Parameter(Mandatory)]
      [ValidateSet("CREATE", "DELETE")]
      $CommandType
    , [Parameter(Mandatory)]
      $DomainName
    , [Parameter(Mandatory)]
      $OrganizationUnit
    , [Parameter(Mandatory)]
      $GroupName
    , $Description
)

# Function for split domain to LDAP String
Function GetLDAPDomain
(
    $DomainNameToSplit
)
{
    # Initialize variables
    $strRetLDAPDomain = ''
    $arrSplitString = $DomainNameToSplit.Split('.')

    # Cycle for each .
    foreach($strTemp in $arrSplitString)
    {
        # Add DC string for each part of domain
        $strRetLDAPDomain += ",DC=$strTemp"
    }

    # Remove first comma
    $strRetLDAPDomain = $strRetLDAPDomain.Substring(1,$strRetLDAPDomain.Length-1) 
    
    # Return string
    return $strRetLDAPDomain       
}

Function GetLDAPOrgUnit
(
    $PathToSplit
)
{
    # Initialize variables
    $strRetLDAPOrgUnit = ''
    $arrSplitStrng = $PathToSplit.Split('\')

    # Cycle for each .
    foreach($strTemp in $arrSplitStrng)
    {
        # Add DC string for each part of domain
        $strRetLDAPOrgUnit = ",OU=$strTemp" + $strRetLDAPOrgUnit
    }

    # Remove first comma
    $strRetLDAPOrgUnit = $strRetLDAPOrgUnit.Substring(1,$strRetLDAPOrgUnit.Length-1) 
    
    # Return string
    return $strRetLDAPOrgUnit       
}

# Initialize variables
# Call function for return splitted domain (LDAP string)
$strLDAPDomain = GetLDAPDomain($DomainName)

# Call function for return splitted organization unit (LDAP string)
$strSQLGroupsOU = GetLDAPOrgUnit($OrganizationUnit)

# Check commad type
if ($CommandType -eq 'DELETE')
{
    # Add group to AD
    Get-ADGroup -Filter "Name -like '$GroupName'" -SearchBase "$strSQLGroupsOU,$strLDAPDomain" | Remove-ADGroup -Confirm:$False
}
else
{
    # Add group to AD
    New-ADGroup -Path "$strSQLGroupsOU,$strLDAPDomain" -Name $GroupName -GroupScope Global -GroupCategory Security -Description $Description
}