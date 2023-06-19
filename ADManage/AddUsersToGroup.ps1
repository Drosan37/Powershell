# Script for add or remove member to given group
param
(
      [Parameter(Mandatory)]
      [ValidateSet("ADD", "REMOVE")]
      $CommandType    
    , [Parameter(Mandatory)]
      $GroupName 
    , [Parameter(Mandatory)]
      [string[]] $UsersList
)

# Initialize variables

# Check command type
if($CommandType -eq 'REMOVE')
{
    # Remove users from group
    Remove-ADGroupMember -Identity $GroupName -Members $UsersList -Confirm:$False
}
else
{
   # Add users to group
   Add-ADGroupMember -Identity $GroupName -Members $UsersList
}