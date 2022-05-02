#######################################
## Script for check password Expired ##
#######################################
param ( 
      [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
      [ValidateSet('Expired','Expiration')]
      [string]$CommandType
    , [string]$UserName = "*"
    , [int]$DaysPeriod = 15    
)

# Choose command by command type
switch($CommandType)
{
    "Expired"
    {
        # Get user list with password expired
        $usersList = Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} –Properties "UserPrincipalName", "DisplayName", "msDS-UserPasswordExpiryTimeComputed" |
        Select-Object -Property "UserPrincipalName", "Displayname",@{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} | 
        Sort-Object -Property "ExpiryDate" | Where-Object "ExpiryDate" -lt (Get-Date).ToString("MM/dd/yyyy") 

        break
    }

    "Expiration"
    {
        # Get user list with password expiration (last 15 days)
        $usersList =  Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} –Properties "UserPrincipalName", "DisplayName", "msDS-UserPasswordExpiryTimeComputed" |
        Select-Object -Property "UserPrincipalName", "Displayname",@{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} | 
        Sort-Object -Property "ExpiryDate" | Where-Object "ExpiryDate" -gt (Get-Date).ToString("MM/dd/yyyy") | Where-Object "ExpiryDate" -lt ((Get-Date).AddDays($DaysPeriod)).ToString("MM/dd/yyyy") 
        
        break
    }
}

# Filter by username
$usersList = $usersList | Where-Object "Displayname" -like $UserName  

# Output in table format
$usersList | Format-Table 