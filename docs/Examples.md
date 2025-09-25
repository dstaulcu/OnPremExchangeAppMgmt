# Exchange Add-In Management Framework - Examples

This directory contains practical examples and usage scenarios for the Exchange Add-In Management Framework.

## Quick Start Examples

### 1. Development Mode (Using Mock Servers)

```powershell
# Run with mock servers for testing
.\src\ExchangeAddInManager.ps1 -UseMockServers -Verbose

# Run in What-If mode to see what would happen
.\src\ExchangeAddInManager.ps1 -UseMockServers -WhatIf

# Test specific scenarios
.\tests\TestScenarios.ps1 -TestSuite Basic -GenerateTestReport
```

### 2. Production Mode

```powershell
# Run against real Exchange and Active Directory
.\src\ExchangeAddInManager.ps1 -ExchangeServer "exchange.contoso.com" -Domain "contoso.com"

# Run with custom configuration
.\src\ExchangeAddInManager.ps1 -ExchangeServer "exchange.contoso.com" -Domain "contoso.com" -GroupPattern "addin-*"
```

## Active Directory Group Setup

### Creating Add-In Groups

```powershell
# Create groups following the naming convention
New-ADGroup -Name "app-exchangeaddin-salesforce-prod" -GroupScope Global -GroupCategory Security -Description "https://appexchange.salesforce.com/manifest.xml" -Path "OU=Exchange Add-Ins,DC=contoso,DC=com"

New-ADGroup -Name "app-exchangeaddin-docusign-test" -GroupScope Global -GroupCategory Security -Description "https://apps.docusign.com/manifest.xml" -Path "OU=Exchange Add-Ins,DC=contoso,DC=com"

New-ADGroup -Name "app-exchangeaddin-teams-dev" -GroupScope Global -GroupCategory Security -Description "https://teams.microsoft.com/dev/manifest.xml" -Path "OU=Exchange Add-Ins,DC=contoso,DC=com"
```

### Adding Users to Groups

```powershell
# Add users to add-in groups
Add-ADGroupMember -Identity "app-exchangeaddin-salesforce-prod" -Members "jdoe", "asmith"
Add-ADGroupMember -Identity "app-exchangeaddin-docusign-test" -Members "bwilson"
```

## Scheduled Task Setup

### Create Windows Scheduled Task

```powershell
# Create scheduled task to run every 15 minutes
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File `"C:\ExchangeAddInMgmt\src\ExchangeAddInManager.ps1`" -ExchangeServer `"exchange.contoso.com`" -Domain `"contoso.com`""

$trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 15) -RepetitionDuration (New-TimeSpan -Days 365) -At (Get-Date) -Once

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

$principal = New-ScheduledTaskPrincipal -UserID "CONTOSO\ExchangeAddInService" -LogonType ServiceAccount

Register-ScheduledTask -TaskName "ExchangeAddInManagement" -Action $action -Trigger $trigger -Settings $settings -Principal $principal
```

## Monitoring and Troubleshooting

### Check Framework Status

```powershell
# View recent logs
Get-Content .\logs\info-*.log | Select-Object -Last 50

# Check for errors
Get-Content .\logs\error-*.log | Where-Object { $_ -like "*$(Get-Date -Format 'yyyy-MM-dd')*" }

# View current state
Get-Content .\data\state.json | ConvertFrom-Json | Format-Table
```

### Manual Operations

```powershell
# Manually test group membership
$mockADPath = ".\src\Mock\MockActiveDirectory.psm1"
Import-Module $mockADPath -Force

Get-ADGroup -Filter "Name -like 'app-exchangeaddin-*'" -Properties Description
Get-ADGroupMember -Identity "app-exchangeaddin-salesforce-prod"

# Manually test add-in operations
$mockExchangePath = ".\src\Mock\MockExchange.psm1"
Import-Module $mockExchangePath -Force

Get-App -Mailbox "john.doe@contoso.com"
New-App -Mailbox "test.user@contoso.com" -Url "https://test.com/manifest.xml" -OrganizationApp
```

## Common Scenarios

### Scenario 1: New Add-In Deployment

1. **Create manifest and test in dev environment**
2. **Create AD group for the add-in**
3. **Add test users to the group**
4. **Run framework to validate deployment**
5. **Monitor logs for successful installation**

### Scenario 2: User Onboarding

1. **Add new user to appropriate add-in groups in AD**
2. **Framework will automatically detect membership change on next run**
3. **Add-ins will be installed for the new user**
4. **Verify installation in user's Outlook**

### Scenario 3: User Offboarding

1. **Remove user from add-in groups in AD**
2. **Framework will detect removal and uninstall add-ins**
3. **Verify removal from user's mailbox**

### Scenario 4: Add-In Retirement

1. **Remove all users from the add-in group**
2. **Framework will uninstall from all users**
3. **Archive or delete the AD group**
4. **Clean up any remaining state data**

## Performance Optimization

### Large Environment Considerations

```powershell
# For environments with many groups/users, consider:
# 1. Running less frequently (30-60 minutes)
# 2. Filtering to specific OUs
# 3. Running during off-peak hours

# Example with OU filtering
$users = Get-ADUser -SearchBase "OU=Sales,DC=contoso,DC=com" -Filter * -Properties mail
```

### Batch Processing

```powershell
# Process users in batches to avoid Exchange throttling
$batchSize = 10
$allUsers = @("user1@contoso.com", "user2@contoso.com", "user3@contoso.com")

for ($i = 0; $i -lt $allUsers.Count; $i += $batchSize) {
    $batch = $allUsers[$i..([math]::Min($i + $batchSize - 1, $allUsers.Count - 1))]
    
    foreach ($user in $batch) {
        # Process user
        Write-Host "Processing: $user"
    }
    
    # Pause between batches
    Start-Sleep -Seconds 5
}
```

## Security Considerations

### Service Account Permissions

The service account running the framework needs:

- **Active Directory**: Read permissions on user and group objects
- **Exchange**: App management permissions (New-App, Remove-App, Get-App cmdlets)
- **File System**: Read/write access to state and log directories

```powershell
# Example Exchange RBAC role assignment
New-ManagementRoleAssignment -Role "Org Marketplace Apps" -User "CONTOSO\ExchangeAddInService"
```

### Audit and Compliance

```powershell
# Generate compliance report
$stateData = Get-Content .\data\state.json | ConvertFrom-Json
$report = $stateData | Select-Object GroupName, AddInName, Environment, @{Name="UserCount"; Expression={$_.CurrentMembers.Count}}, LastUpdated

$report | Export-Csv -Path ".\reports\addin-compliance-$(Get-Date -Format 'yyyyMMdd').csv" -NoTypeInformation
```

## Integration Examples

### Email Notifications

```powershell
# Add to framework for error notifications
function Send-NotificationEmail {
    param([string]$Subject, [string]$Body)
    
    $emailParams = @{
        To = "admin@contoso.com"
        From = "exchangeaddins@contoso.com"
        Subject = $Subject
        Body = $Body
        SmtpServer = "smtp.contoso.com"
    }
    
    Send-MailMessage @emailParams
}

# Usage in framework
try {
    # Framework operations
}
catch {
    Send-NotificationEmail -Subject "Exchange Add-In Management Error" -Body $_.Exception.Message
}
```

### SIEM Integration

```powershell
# Log structured data for SIEM consumption
$logEntry = @{
    Timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
    Source = "ExchangeAddInManager"
    EventType = "AddInInstalled"
    User = $userEmail
    AddIn = $config.AddInName
    Environment = $config.Environment
    Success = $true
} | ConvertTo-Json -Compress

Add-Content -Path ".\logs\siem.log" -Value $logEntry
```