# Quick Start Guide - Exchange Add-In Management Framework

## Prerequisites

- PowerShell 5.1 or later
- Windows Server with Exchange Management Tools (production)
- Active Directory PowerShell module (production)
- Appropriate permissions for AD group reading and Exchange add-in management

## Development Setup (5 minutes)

1. **Clone or download the framework**
2. **Open PowerShell and navigate to the project directory**
3. **Run the framework with mock servers**:

```powershell
.\src\ExchangeAddInManager.ps1 -UseMockServers -Verbose
```

You should see output similar to:
```
Exchange Add-In Management Framework
====================================
Mode: Development (Mock)
What-If: False

Discovering add-in groups with pattern: app-exchangeaddin-*
Found 4 matching groups
  - app-exchangeaddin-salesforce-prod: 3 members
  - app-exchangeaddin-docusign-test: 2 members
  - app-exchangeaddin-teams-dev: 1 members
  - app-exchangeaddin-onenote-prod: 3 members

Processing add-in changes...
Processing app-exchangeaddin-salesforce-prod...
  No changes detected
...

Execution Summary:
=================
Groups found: 4
Users to add add-ins: 0
Users to remove add-ins: 0
Add-ins installed: 0
Add-ins removed: 0
Errors encountered: 0

Exchange Add-In Management Process Completed Successfully
```

## Testing the Framework (2 minutes)

Run the comprehensive test suite:

```powershell
.\tests\TestScenarios.ps1 -TestSuite All -GenerateTestReport
```

All tests should pass, indicating the framework is working correctly.

## Simulating Changes (3 minutes)

1. **Import the mock modules** to simulate AD changes:

```powershell
Import-Module .\src\Mock\MockActiveDirectory.psm1 -Force
Import-Module .\src\Mock\MockExchange.psm1 -Force
```

2. **Add a new user and simulate group membership**:

```powershell
# Add new user
Add-MockADUser -SamAccountName "testuser" -Name "Test User" -Mail "test.user@contoso.com"

# Add user to Salesforce add-in group
Add-MockADGroupMember -GroupName "app-exchangeaddin-salesforce-prod" -UserSamAccountName "testuser"
```

3. **Run the framework again** to see it detect and process the change:

```powershell
.\src\ExchangeAddInManager.ps1 -UseMockServers -Verbose
```

You should now see:
```
Processing app-exchangeaddin-salesforce-prod...
  Users to add (1): test.user@contoso.com
  Installing add-in 'salesforce' for user: test.user@contoso.com
Successfully installed 'Salesforce Lightning for Outlook' for mailbox 'test.user@contoso.com'
```

## Production Deployment

### Step 1: Create Service Account
Create a dedicated service account with necessary permissions:
- Active Directory: Read access to user and group objects
- Exchange: App management permissions (Org Marketplace Apps role)

### Step 2: Set Up Active Directory Groups
Create groups following the naming convention:

```powershell
New-ADGroup -Name "app-exchangeaddin-salesforce-prod" -GroupScope Global -GroupCategory Security -Description "https://appexchange.salesforce.com/manifest.xml" -Path "OU=Exchange Add-Ins,DC=contoso,DC=com"
```

### Step 3: Deploy Framework
1. Copy framework files to production server
2. Update configuration in `config\production.json`
3. Test with production parameters:

```powershell
.\src\ExchangeAddInManager.ps1 -ExchangeServer "exchange.contoso.com" -Domain "contoso.com" -WhatIf
```

### Step 4: Schedule Execution
Create a Windows Scheduled Task:

```powershell
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File `"C:\ExchangeAddInMgmt\src\ExchangeAddInManager.ps1`" -ExchangeServer `"exchange.contoso.com`" -Domain `"contoso.com`""
$trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 15) -RepetitionDuration (New-TimeSpan -Days 365) -At (Get-Date) -Once
Register-ScheduledTask -TaskName "ExchangeAddInManagement" -Action $action -Trigger $trigger
```

## Next Steps

- Review the [Examples documentation](Examples.md) for advanced scenarios
- Set up monitoring and alerting for the scheduled task
- Configure email notifications for errors
- Establish a process for managing manifest URLs and add-in lifecycle

## Troubleshooting

**Framework doesn't find any groups:**
- Verify the group naming pattern matches `app-exchangeaddin-*`
- Check that groups have manifest URLs in their Description field

**Add-in installation fails:**
- Verify manifest URL is accessible
- Check Exchange permissions for the service account
- Review Exchange server connectivity

**State file issues:**
- Ensure the data directory is writable
- Check file permissions on state.json

For more detailed troubleshooting, see the log files in the `logs\` directory.