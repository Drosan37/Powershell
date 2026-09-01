<#
.SYNOPSIS
    Add SQL Servers to Central Management Server (CMS) with support for nested groups
.DESCRIPTION
    Script to register SQL Server instances in CMS with full support for nested/hierarchical groups
.EXAMPLE
    # Simple group: Production
    # Nested group: Production\WebServers
    # Deep nesting: Production\Region1\WebServers
#>

# ===========================
# CONFIGURATION - Edit these values
# ===========================
$CMSServerInstance = "IDERADB.adgr.net,45000"  # Your CMS server name

# SYNC MODE: Set to $true to remove servers not in the list (synchronize CMS with list)
# Set to $false to only add/update without removing anything
$SyncMode = $true  # Change to $true to enable sync/cleanup

# Define servers with nested groups using backslash separator
# RegisteredName: Display name shown in SSMS (optional, defaults to ServerName)
# ServerName: Actual SQL Server connection string

# Script for gather information about AG Status on all servers
Import-Module SqlServer -Version 21.1.18209			

# Define query for retrieve server list
$strQueryServerList = ("    
    SELECT * FROM CHK.GetServerList  
") 

# Set culture for avoid problems with decimal
$culture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")

try		
{		
	# Execute query to retrieve server list	
	$arrServerList = Invoke-Sqlcmd $strQueryServerList -ServerInstance "IDERADB,45000" -Database "SqlServerMap" -QueryTimeout 600 -ErrorAction 'Stop';	
}		
catch		
{
    Write-Host "ERROR on Execute query for retrieve server List:`r`n$($_.Exception.Message)";  		
}

# Create empty array
$ServersToAdd = @()

# Loop through each server and retrieve disk space for mounted volumes
foreach ($server in $arrServerList) {
    # Define variable for each server
    $strADDomain = $server.ADDomain
    $strServerName = ("{0}.{1}.net,{2}" -f $server.ServerName, $server.ADDomain.toLower(),$server.PortNumber)
    $strInstanceName = $server.InstanceName    
    $strRegisteredServer = ("{0}.{1}.net" -f $server.ServerName, $server.ADDomain.toLower())
    $strEnvironment = $server.InstanceName.Substring(0,4)
    
    $ServersToAdd += @{RegisteredName = $strRegisteredServer; ServerName = $strServerName; GroupPath = "INSTANCES\$strADDomain\$strEnvironment\$strInstanceName"; Description = ""}        
}

#$ServersToAdd = @()


# Define query for retrieve listener list
$strQueryListenerList = ("     
    SELECT DISTINCT * FROM.[RO].[ExcelListener] 
") 

try		
{		
	# Execute query to retrieve server list	
	$arrListenerList = Invoke-Sqlcmd $strQueryListenerList -ServerInstance "IDERADB,45000" -Database "SqlServerMap" -QueryTimeout 600 -ErrorAction 'Stop';	
}		
catch		
{
    Write-Host "ERROR on Execute query for retrieve listener List:`r`n$($_.Exception.Message)";  		
}

# Loop through each listener
foreach ($server in $arrListenerList) {
    # Define variable for each listener
    $strADDomain = $server.ADDomain
    $strServerName = ("{0},{1}" -f $server.InternalName,$server.Port)
    $strInstanceName = $server.InstanceName    
    $strRegisteredServer = $server.InternalName
    $strEnvironment = $server.InstanceName.Substring(0,4)
    
    $ServersToAdd += @{RegisteredName = $strRegisteredServer; ServerName = $strServerName; GroupPath = "LISTENERS\$strADDomain\$strEnvironment\$strInstanceName"; Description = ""} 
        
}

# ===========================
# SCRIPT START
# ===========================

# Load SQL Server assemblies
Write-Host "Loading SQL Server Management Objects..." -ForegroundColor Cyan
try {
    if (Get-Module -ListAvailable -Name SqlServer) {
        Import-Module SqlServer -ErrorAction Stop
        Write-Host "SqlServer module loaded successfully" -ForegroundColor Green
    }
    else {
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Management.RegisteredServers") | Out-Null
        Write-Host "SMO assemblies loaded successfully" -ForegroundColor Green
    }
}
catch {
    Write-Error "Failed to load SQL Server components. Please install SQL Server Management Studio or SqlServer PowerShell module."
    Write-Host "Install module with: Install-Module -Name SqlServer" -ForegroundColor Yellow
    exit 1
}

# Connect to CMS
Write-Host "`nConnecting to CMS: $CMSServerInstance" -ForegroundColor Cyan
try {
    $smoServer = New-Object Microsoft.SqlServer.Management.Smo.Server($CMSServerInstance)
    $cmsStore = New-Object Microsoft.SqlServer.Management.RegisteredServers.RegisteredServersStore($smoServer.ConnectionContext.SqlConnectionObject)
    $rootGroup = $cmsStore.DatabaseEngineServerGroup
    Write-Host "Connected successfully!" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to CMS: $_"
    exit 1
}

# Function to create nested groups recursively
function Get-OrCreateNestedGroup {
    param(
        [string]$GroupPath,
        [Microsoft.SqlServer.Management.RegisteredServers.ServerGroup]$ParentGroup = $rootGroup
    )
    
    # Split the path by backslash
    $groupNames = $GroupPath -split '\\'
    
    $currentGroup = $ParentGroup
    $currentPath = ""
    
    foreach ($groupName in $groupNames) {
        if ($currentPath -eq "") {
            $currentPath = $groupName
        } else {
            $currentPath = "$currentPath\$groupName"
        }
        
        # Try to get existing group
        $nextGroup = $currentGroup.ServerGroups[$groupName]
        
        if ($null -eq $nextGroup) {
            Write-Host "  Creating group: $currentPath" -ForegroundColor Yellow
            $nextGroup = New-Object Microsoft.SqlServer.Management.RegisteredServers.ServerGroup($currentGroup, $groupName)
            $nextGroup.Create()
        }
        
        $currentGroup = $nextGroup
    }
    
    return $currentGroup
}

# Function to display group hierarchy
function Show-GroupHierarchy {
    param(
        [Microsoft.SqlServer.Management.RegisteredServers.ServerGroup]$Group,
        [int]$Indent = 0
    )
    
    $prefix = "  " * $Indent
    Write-Host "$prefix📁 $($Group.Name)" -ForegroundColor Cyan
    
    # Show servers in this group
    foreach ($server in $Group.RegisteredServers) {
        Write-Host "$prefix  🖥️  $($server.ServerName)" -ForegroundColor White
    }
    
    # Recursively show subgroups
    foreach ($subGroup in $Group.ServerGroups) {
        Show-GroupHierarchy -Group $subGroup -Indent ($Indent + 1)
    }
}

# Add servers
Write-Host "`nAdding/Updating servers in CMS..." -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

$successCount = 0
$skipCount = 0
$errorCount = 0
$updatedCount = 0

# Build a hashtable of servers from the list for quick lookup
$serverList = @{}
foreach ($server in $ServersToAdd) {
    $regName = if ($server.RegisteredName) { $server.RegisteredName } else { $server.ServerName }
    $key = "$($server.GroupPath)|$regName"
    $serverList[$key] = $server
}

foreach ($server in $ServersToAdd) {
    # RegisteredName is what shows in SSMS, ServerName is the actual connection string
    $registeredName = if ($server.RegisteredName) { $server.RegisteredName } else { $server.ServerName }
    $serverName = $server.ServerName
    $groupPath = $server.GroupPath
    $description = $server.Description
    
    Write-Host "Processing: $registeredName" -ForegroundColor White
    Write-Host "  Display Name: $registeredName" -ForegroundColor Gray
    Write-Host "  Connection String: $serverName" -ForegroundColor Gray
    Write-Host "  Group Path: $groupPath" -ForegroundColor Gray
    
    try {
        # Get or create the nested group structure
        $targetGroup = Get-OrCreateNestedGroup -GroupPath $groupPath
        
        # Check if server already exists in this group (by registered name)
        $existingServer = $targetGroup.RegisteredServers[$registeredName]
        
        if ($existingServer) {
            # Check if we need to update it
            $needsUpdate = $false
            
            if ($existingServer.ServerName -ne $serverName) {
                Write-Host "  🔄 Connection string changed: '$($existingServer.ServerName)' → '$serverName'" -ForegroundColor Yellow
                $existingServer.ServerName = $serverName
                $needsUpdate = $true
            }
            
            if ($existingServer.Description -ne $description) {
                Write-Host "  🔄 Description changed" -ForegroundColor Yellow
                $existingServer.Description = $description
                $needsUpdate = $true
            }
            
            if ($needsUpdate) {
                $existingServer.Alter()
                Write-Host "  ✓ Successfully updated" -ForegroundColor Green
                $updatedCount++
            }
            else {
                Write-Host "  ⚠️  Server already exists (no changes needed)" -ForegroundColor Yellow
                $skipCount++
            }
        }
        else {
            # Create new registered server
            # First parameter (name) is what displays in SSMS
            # ServerName property is the actual connection string
            $registeredServer = New-Object Microsoft.SqlServer.Management.RegisteredServers.RegisteredServer($targetGroup, $registeredName)
            $registeredServer.ServerName = $serverName
            $registeredServer.Description = $description
            
            # Save to CMS
            $registeredServer.Create()
            
            Write-Host "  ✓ Successfully added" -ForegroundColor Green
            $successCount++
        }
    }
    catch {
        Write-Host "  ✗ Failed: $_" -ForegroundColor Red
        $errorCount++
    }
    
    Write-Host ""
}

# SYNC MODE: Remove servers that are not in the list
if ($SyncMode) {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "SYNC MODE: Removing servers not in list..." -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    $removedCount = 0
    
    function Remove-ServersNotInList {
        param(
            [Microsoft.SqlServer.Management.RegisteredServers.ServerGroup]$Group,
            [string]$CurrentPath = ""
        )
        
        $groupPath = if ($CurrentPath -eq "") { $Group.Name } else { "$CurrentPath\$($Group.Name)" }
        
        # Remove servers in this group that aren't in the list
        $serversToRemove = @()
        foreach ($regServer in $Group.RegisteredServers) {
            $key = "$groupPath|$($regServer.Name)"
            
            if (-not $serverList.ContainsKey($key)) {
                $serversToRemove += $regServer
            }
        }
        
        foreach ($serverToRemove in $serversToRemove) {
            try {
                Write-Host "Removing: $($serverToRemove.Name)" -ForegroundColor Red
                Write-Host "  From Group: $groupPath" -ForegroundColor Gray
                Write-Host "  Connection: $($serverToRemove.ServerName)" -ForegroundColor Gray
                $serverToRemove.Drop()
                Write-Host "  ✓ Removed successfully" -ForegroundColor Green
                $script:removedCount++
            }
            catch {
                Write-Host "  ✗ Failed to remove: $_" -ForegroundColor Red
            }
            Write-Host ""
        }
        
        # Recursively process subgroups
        foreach ($subGroup in $Group.ServerGroups) {
            Remove-ServersNotInList -Group $subGroup -CurrentPath $groupPath
        }
    }
    
    # Start from root group, but skip the root itself in the path
    foreach ($topLevelGroup in $rootGroup.ServerGroups) {
        Remove-ServersNotInList -Group $topLevelGroup -CurrentPath ""
    }
    
    Write-Host "Removed $removedCount server(s) not in the list`n" -ForegroundColor Magenta
    
    # Remove empty groups that don't match the list
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "SYNC MODE: Removing empty groups not in list..." -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    $removedGroupCount = 0
    
    # Build a set of all group paths that should exist (from the server list)
    $validGroupPaths = @{}
    foreach ($server in $ServersToAdd) {
        $path = $server.GroupPath
        $validGroupPaths[$path] = $true
        
        # Also add all parent paths
        $parts = $path -split '\\'
        for ($i = 1; $i -lt $parts.Length; $i++) {
            $parentPath = ($parts[0..($i-1)] -join '\')
            $validGroupPaths[$parentPath] = $true
        }
    }
    
    function Remove-EmptyGroups {
        param(
            [Microsoft.SqlServer.Management.RegisteredServers.ServerGroup]$Group,
            [string]$CurrentPath = ""
        )
        
        $groupPath = if ($CurrentPath -eq "") { $Group.Name } else { "$CurrentPath\$($Group.Name)" }
        
        # First, recursively process all subgroups (bottom-up)
        $subGroupsToRemove = @()
        foreach ($subGroup in $Group.ServerGroups) {
            Remove-EmptyGroups -Group $subGroup -CurrentPath $groupPath
            
            # After processing children, check if this subgroup is now empty and not in valid paths
            if ($subGroup.ServerGroups.Count -eq 0 -and $subGroup.RegisteredServers.Count -eq 0) {
                $subGroupFullPath = "$groupPath\$($subGroup.Name)"
                if (-not $validGroupPaths.ContainsKey($subGroupFullPath)) {
                    $subGroupsToRemove += $subGroup
                }
            }
        }
        
        # Remove empty subgroups
        foreach ($groupToRemove in $subGroupsToRemove) {
            try {
                $groupFullPath = "$groupPath\$($groupToRemove.Name)"
                Write-Host "Removing empty group: $groupFullPath" -ForegroundColor Red
                $groupToRemove.Drop()
                Write-Host "  ✓ Group removed successfully" -ForegroundColor Green
                $script:removedGroupCount++
            }
            catch {
                Write-Host "  ✗ Failed to remove group: $_" -ForegroundColor Red
            }
            Write-Host ""
        }
    }
    
    # Start from root group
    foreach ($topLevelGroup in $rootGroup.ServerGroups) {
        Remove-EmptyGroups -Group $topLevelGroup -CurrentPath ""
        
        # Check if top-level group itself should be removed
        if ($topLevelGroup.ServerGroups.Count -eq 0 -and $topLevelGroup.RegisteredServers.Count -eq 0) {
            if (-not $validGroupPaths.ContainsKey($topLevelGroup.Name)) {
                try {
                    Write-Host "Removing empty top-level group: $($topLevelGroup.Name)" -ForegroundColor Red
                    $topLevelGroup.Drop()
                    Write-Host "  ✓ Group removed successfully" -ForegroundColor Green
                    $removedGroupCount++
                }
                catch {
                    Write-Host "  ✗ Failed to remove group: $_" -ForegroundColor Red
                }
                Write-Host ""
            }
        }
    }
    
    Write-Host "Removed $removedGroupCount empty group(s) not in the list`n" -ForegroundColor Magenta
}

# Display final hierarchy
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "CMS GROUP HIERARCHY" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Show-GroupHierarchy -Group $rootGroup

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Successfully added: $successCount" -ForegroundColor Green
Write-Host "Successfully updated: $updatedCount" -ForegroundColor Green
Write-Host "Already existed (no changes): $skipCount" -ForegroundColor Yellow
if ($SyncMode) {
    Write-Host "Removed servers (not in list): $removedCount" -ForegroundColor Red
    Write-Host "Removed groups (empty/not in list): $removedGroupCount" -ForegroundColor Red
}
Write-Host "Errors: $errorCount" -ForegroundColor Red
Write-Host "`nDone! Open SSMS and refresh Registered Servers to see the hierarchy." -ForegroundColor Cyan

if (-not $SyncMode) {
    Write-Host "`nNote: Sync Mode is OFF. To remove servers/groups not in the list, set `$SyncMode = `$true" -ForegroundColor Yellow
}

<#
==============================================================================
SYNC MODE EXAMPLES
==============================================================================

SYNC MODE keeps your CMS synchronized with your list:
- Adds new servers from the list
- Updates existing servers if connection string or description changed
- REMOVES servers that are NOT in the list anymore

==============================================================================

Example 1: Basic Sync Mode Usage
------------------------------------------------------------------------------
# Step 1: Define your CURRENT server list (this is your source of truth)
$ServersToAdd = @(
    @{RegisteredName = "PROD-01"; ServerName = "sql-prod-01.domain.com"; GroupPath = "Production"; Description = "Prod 1"},
    @{RegisteredName = "PROD-02"; ServerName = "sql-prod-02.domain.com"; GroupPath = "Production"; Description = "Prod 2"}
)

# Step 2: Enable Sync Mode
$SyncMode = $true

# Step 3: Run the script
# Result: 
# - PROD-01 and PROD-02 will be in CMS
# - Any other servers in the "Production" group will be REMOVED

==============================================================================

Example 2: Update Connection String
------------------------------------------------------------------------------
# Your server moved to a new host
# Old CMS entry: PROD-01 → sql-old.domain.com
# New list:
$ServersToAdd = @(
    @{RegisteredName = "PROD-01"; ServerName = "sql-new.domain.com"; GroupPath = "Production"; Description = "Prod 1"}
)

$SyncMode = $true

# Result: PROD-01 will be updated to point to sql-new.domain.com

==============================================================================

Example 3: Import from CSV and Sync
------------------------------------------------------------------------------
# CSV File (servers.csv):
RegisteredName,ServerName,GroupPath,Description
PROD-Web-01,sql-web-01.domain.com,Production\WebServers,Web 1
PROD-Web-02,sql-web-02.domain.com,Production\WebServers,Web 2
PROD-App-01,sql-app-01.domain.com,Production\AppServers,App 1

# Script:
$ServersToAdd = Import-Csv "C:\Temp\servers.csv"
$SyncMode = $true

# Result: 
# - Only these 3 servers will exist in CMS
# - Any servers previously in Production\WebServers or Production\AppServers 
#   that aren't in the CSV will be removed

==============================================================================

Example 4: Decommission Servers
------------------------------------------------------------------------------
# You had 5 servers, now decommissioning 2 of them

# Old CMS had:
# - PROD-01, PROD-02, PROD-03, PROD-04, PROD-05

# New list (removed PROD-04 and PROD-05):
$ServersToAdd = @(
    @{RegisteredName = "PROD-01"; ServerName = "sql-prod-01.domain.com"; GroupPath = "Production"; Description = "Prod 1"},
    @{RegisteredName = "PROD-02"; ServerName = "sql-prod-02.domain.com"; GroupPath = "Production"; Description = "Prod 2"},
    @{RegisteredName = "PROD-03"; ServerName = "sql-prod-03.domain.com"; GroupPath = "Production"; Description = "Prod 3"}
)

$SyncMode = $true

# Result: PROD-04 and PROD-05 will be removed from CMS

==============================================================================

Example 5: Safe Mode (Add/Update Only - No Removal)
------------------------------------------------------------------------------
# If you want to add/update WITHOUT removing anything:

$ServersToAdd = @(
    @{RegisteredName = "PROD-NEW"; ServerName = "sql-new.domain.com"; GroupPath = "Production"; Description = "New Server"}
)

$SyncMode = $false  # Keep this false

# Result:
# - PROD-NEW will be added
# - All existing servers remain untouched

==============================================================================

Example 6: Reorganize Server Groups
------------------------------------------------------------------------------
# Move servers from one group to another

# Old structure:
# Development\SQL-DEV-01

# New structure (move to Development\Testing):
$ServersToAdd = @(
    @{RegisteredName = "SQL-DEV-01"; ServerName = "sql-dev-01.domain.com"; GroupPath = "Development\Testing"; Description = "Dev Server"}
)

$SyncMode = $true

# Result:
# - SQL-DEV-01 will be removed from "Development" group
# - SQL-DEV-01 will be added to "Development\Testing" group
# NOTE: Moving is actually a remove + add operation

==============================================================================

BEST PRACTICES:
------------------------------------------------------------------------------
1. Keep your server list in a CSV or database as the source of truth
2. Run with $SyncMode = $false first to see what would be added/updated
3. Review the output carefully before enabling $SyncMode = $true
4. Back up your CMS before first sync (export from SSMS)
5. Use version control for your CSV/list file
6. Schedule the script to run regularly for automatic synchronization

SAFETY TIPS:
------------------------------------------------------------------------------
- Start with $SyncMode = $false to test
- The script shows exactly what will be removed before doing it
- Removed servers are only removed from CMS (the actual SQL Server is not affected)
- You can always re-add servers by putting them back in the list

==============================================================================
#>
