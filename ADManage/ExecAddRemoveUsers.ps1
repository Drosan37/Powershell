# Define array for users list to add or remove from/to group
$arrUsers = @(
     "user"    
 )
 
 # Call command
.\AddUsersToGroup.ps1 -CommandType REMOVE -GroupName SQL_INT_TESTADD_RO -UsersList $arrUsers