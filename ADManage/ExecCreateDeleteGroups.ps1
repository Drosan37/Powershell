# Create
.\ManageADGroups.ps1 -CommandType CREATE -DomainName dgigroup.local -GroupName SQL_INT_TESTADD_RO -Description "Gruppo ReadOnly" -OrganizationUnit "SQL\Security Groups"
#.\ManageADGroups.ps1 -CommandType CREATE -DomainName dgigroup.local -GroupName SQL_INT_NAME_RW -Description "Gruppo ReadWrite" -OrganizationUnit "SQL\Security Groups"
#.\ManageADGroups.ps1 -CommandType CREATE -DomainName dgigroup.local -GroupName SQL_INT_NAME_OW -Description "Gruppo DBOwner" -OrganizationUnit "SQL\Security Groups"


# Delete
#.\ManageADGroups.ps1 -CommandType DELETE -DomainName dgigroup.local -GroupName SQL_INT_NAME_RO -OrganizationUnit "SQL\Security Groups"
#.\ManageADGroups.ps1 -CommandType DELETE -DomainName dgigroup.local -GroupName SQL_INT_NAME_RW -OrganizationUnit "SQL\Security Groups"
#.\ManageADGroups.ps1 -CommandType DELETE -DomainName dgigroup.local -GroupName SQL_INT_NAME_OW -OrganizationUnit "SQL\Security Groups"


