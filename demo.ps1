# Run the Exchange Add-In Management Framework in development mode

Write-Host "Starting Exchange Add-In Management Framework Demo..." -ForegroundColor Cyan

# Change to the project directory
$projectRoot = Split-Path $PSScriptRoot

# Run basic functionality test
Write-Host "`n=== Running Basic Tests ===" -ForegroundColor Yellow
& "$projectRoot\tests\TestScenarios.ps1" -TestSuite Basic

Write-Host "`n=== Running Framework with Mock Servers ===" -ForegroundColor Yellow
& "$projectRoot\src\ExchangeAddInManager.ps1" -UseMockServers -Verbose

Write-Host "`n=== Simulating User Addition ===" -ForegroundColor Yellow

# Import mock modules to simulate changes
Import-Module "$projectRoot\src\Mock\MockActiveDirectory.psm1" -Force
Import-Module "$projectRoot\src\Mock\MockExchange.psm1" -Force

# Add a new user
Add-MockADUser -SamAccountName "demouser" -Name "Demo User" -Mail "demo.user@contoso.com"
Add-MockADGroupMember -GroupName "app-exchangeaddin-salesforce-prod" -UserSamAccountName "demouser"

Write-Host "Added demo.user@contoso.com to Salesforce add-in group"

Write-Host "`n=== Running Framework Again to Process Changes ===" -ForegroundColor Yellow
& "$projectRoot\src\ExchangeAddInManager.ps1" -UseMockServers -Verbose

Write-Host "`n=== Demo Complete ===" -ForegroundColor Green
Write-Host "Check the logs directory for detailed execution logs."
Write-Host "Check the data directory for state persistence."

# Show current state
if (Test-Path "$projectRoot\data\state.json") {
    Write-Host "`n=== Current State ===" -ForegroundColor Cyan
    Get-Content "$projectRoot\data\state.json" | ConvertFrom-Json | Format-Table -AutoSize
}