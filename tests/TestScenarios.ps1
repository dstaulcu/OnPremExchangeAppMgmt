#Requires -Version 5.1

<#
.SYNOPSIS
    Test scenarios and validation for the Exchange Add-In Management Framework
.DESCRIPTION
    Provides comprehensive test scenarios to validate the framework functionality
    including edge cases and error conditions
#>

param(
    [ValidateSet('Basic', 'EdgeCases', 'Performance', 'All')]
    [string]$TestSuite = 'All',
    
    [switch]$GenerateTestReport
)

# Import required modules
$mockADPath = Join-Path $PSScriptRoot "..\src\Mock\MockActiveDirectory.psm1"
$mockExchangePath = Join-Path $PSScriptRoot "..\src\Mock\MockExchange.psm1"

Import-Module $mockADPath -Force
Import-Module $mockExchangePath -Force

#region Test Infrastructure

class TestResult {
    [string]$TestName
    [bool]$Passed
    [string]$Message
    [datetime]$ExecutedAt
    [timespan]$Duration
    
    TestResult([string]$name, [bool]$passed, [string]$message, [timespan]$duration) {
        $this.TestName = $name
        $this.Passed = $passed
        $this.Message = $message
        $this.ExecutedAt = Get-Date
        $this.Duration = $duration
    }
}

class TestRunner {
    [System.Collections.ArrayList]$Results
    [int]$PassedCount
    [int]$FailedCount
    
    TestRunner() {
        $this.Results = @()
        $this.PassedCount = 0
        $this.FailedCount = 0
    }
    
    [TestResult] RunTest([string]$testName, [scriptblock]$testScript) {
        Write-Host "Running test: $testName" -ForegroundColor Yellow
        
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        
        try {
            $testScript.Invoke()
            $stopwatch.Stop()
            
            $result = [TestResult]::new($testName, $true, "Test passed", $stopwatch.Elapsed)
            $this.PassedCount++
            
            Write-Host "  ✅ PASSED" -ForegroundColor Green
        }
        catch {
            $stopwatch.Stop()
            
            $result = [TestResult]::new($testName, $false, $_.Exception.Message, $stopwatch.Elapsed)
            $this.FailedCount++
            
            Write-Host "  ❌ FAILED: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        $this.Results.Add($result) | Out-Null
        return $result
    }
    
    [void] ShowSummary() {
        Write-Host "`n" + "="*50 -ForegroundColor Cyan
        Write-Host "Test Summary" -ForegroundColor Cyan
        Write-Host "="*50 -ForegroundColor Cyan
        Write-Host "Total Tests: $($this.Results.Count)" -ForegroundColor White
        Write-Host "Passed: $($this.PassedCount)" -ForegroundColor Green
        Write-Host "Failed: $($this.FailedCount)" -ForegroundColor Red
        Write-Host "Success Rate: $(if($this.Results.Count -gt 0) { [math]::Round(($this.PassedCount / $this.Results.Count) * 100, 2) } else { 0 })%" -ForegroundColor White
        Write-Host "Total Duration: $([math]::Round(($this.Results | Measure-Object -Property Duration -Sum).Sum.TotalSeconds, 2)) seconds" -ForegroundColor White
    }
    
    [void] GenerateReport([string]$outputPath) {
        $report = @{
            TestRun = @{
                ExecutedAt = Get-Date
                TotalTests = $this.Results.Count
                PassedTests = $this.PassedCount
                FailedTests = $this.FailedCount
                SuccessRate = if($this.Results.Count -gt 0) { [math]::Round(($this.PassedCount / $this.Results.Count) * 100, 2) } else { 0 }
                TotalDuration = ($this.Results | Measure-Object -Property Duration -Sum).Sum
            }
            TestResults = $this.Results | ForEach-Object {
                @{
                    TestName = $_.TestName
                    Passed = $_.Passed
                    Message = $_.Message
                    ExecutedAt = $_.ExecutedAt
                    DurationSeconds = $_.Duration.TotalSeconds
                }
            }
        }
        
        $report | ConvertTo-Json -Depth 3 | Set-Content $outputPath -Encoding UTF8
        Write-Host "Test report generated: $outputPath" -ForegroundColor Green
    }
}

#endregion

#region Test Scenarios

function Test-MockADBasicFunctionality {
    # Test basic AD group discovery
    $groups = Get-ADGroup -Filter "Name -like 'app-exchangeaddin-*'" -Properties Description
    if ($groups.Count -lt 3) {
        throw "Expected at least 3 add-in groups, found $($groups.Count)"
    }
    
    # Test group membership retrieval
    $members = Get-ADGroupMember -Identity "app-exchangeaddin-salesforce-prod"
    if ($members.Count -eq 0) {
        throw "Expected members in salesforce-prod group"
    }
    
    # Test user email resolution
    $user = Get-ADUser -Identity "jdoe" -Properties mail
    if ([string]::IsNullOrEmpty($user.Mail)) {
        throw "User should have email address"
    }
    
    Write-Host "    Mock AD functionality validated" -ForegroundColor Gray
}

function Test-MockExchangeBasicFunctionality {
    # Test app installation
    $beforeCount = (Get-MockInstalledApps)["test.user@contoso.com"]
    if ($null -eq $beforeCount) { $beforeCount = @{} }
    
    New-App -Mailbox "test.user@contoso.com" -Url "https://appexchange.salesforce.com/manifest.xml" -OrganizationApp
    
    $afterCount = (Get-MockInstalledApps)["test.user@contoso.com"]
    if ($afterCount.Count -le $beforeCount.Count) {
        throw "App installation should increase installed app count"
    }
    
    # Test app listing
    $installedApps = Get-App -Mailbox "test.user@contoso.com"
    if ($installedApps.Count -eq 0) {
        throw "Should find installed apps for user"
    }
    
    # Test app removal
    Remove-App -Mailbox "test.user@contoso.com" -Identity "salesforce-outlook-addin"
    
    Write-Host "    Mock Exchange functionality validated" -ForegroundColor Gray
}

function Test-AddInGroupPattern {
    $validGroups = @(
        "app-exchangeaddin-salesforce-prod",
        "app-exchangeaddin-docusign-test", 
        "app-exchangeaddin-teams-dev"
    )
    
    foreach ($groupName in $validGroups) {
        if ($groupName -notmatch "app-exchangeaddin-(.+)-(.+)") {
            throw "Group name should match pattern: $groupName"
        }
        
        $addInName = $matches[1]
        $environment = $matches[2]
        
        if ([string]::IsNullOrEmpty($addInName) -or [string]::IsNullOrEmpty($environment)) {
            throw "Should extract add-in name and environment from: $groupName"
        }
    }
    
    Write-Host "    Group naming pattern validation completed" -ForegroundColor Gray
}

function Test-StateManagement {
    $tempStateFile = Join-Path $env:TEMP "test-state.json"
    
    # Create test state data
    $testState = @(
        [PSCustomObject]@{
            GroupName = "app-exchangeaddin-test-prod"
            AddInName = "test"
            Environment = "prod"
            ManifestUrl = "https://test.com/manifest.xml"
            CurrentMembers = @("user1@test.com", "user2@test.com")
            LastUpdated = Get-Date
        }
    )
    
    # Save state
    $testState | ConvertTo-Json -Depth 3 | Set-Content $tempStateFile -Encoding UTF8
    
    # Load and validate state
    $loadedState = Get-Content $tempStateFile -Raw | ConvertFrom-Json
    
    if ($loadedState.Count -ne 1) {
        throw "Should load exactly one state entry"
    }
    
    if ($loadedState[0].GroupName -ne "app-exchangeaddin-test-prod") {
        throw "State data should be preserved correctly"
    }
    
    # Cleanup
    Remove-Item $tempStateFile -Force -ErrorAction SilentlyContinue
    
    Write-Host "    State management validation completed" -ForegroundColor Gray
}

function Test-UserMembershipChanges {
    # Reset mock data to known state
    Reset-MockADData
    Reset-MockExchangeData
    
    # Get initial members
    $initialMembers = Get-ADGroupMember -Identity "app-exchangeaddin-salesforce-prod" |
                     ForEach-Object { (Get-ADUser -Identity $_.SamAccountName -Properties mail).Mail }
    
    # Add a new member
    Add-MockADUser -SamAccountName "newuser" -Name "New User" -Mail "new.user@contoso.com"
    Add-MockADGroupMember -GroupName "app-exchangeaddin-salesforce-prod" -UserSamAccountName "newuser"
    
    # Get updated members
    $updatedMembers = Get-ADGroupMember -Identity "app-exchangeaddin-salesforce-prod" |
                     ForEach-Object { (Get-ADUser -Identity $_.SamAccountName -Properties mail).Mail }
    
    if ($updatedMembers.Count -ne ($initialMembers.Count + 1)) {
        throw "Member count should increase by 1"
    }
    
    if ("new.user@contoso.com" -notin $updatedMembers) {
        throw "New user should be in updated member list"
    }
    
    # Remove a member
    Remove-MockADGroupMember -GroupName "app-exchangeaddin-salesforce-prod" -UserSamAccountName "newuser"
    
    $finalMembers = Get-ADGroupMember -Identity "app-exchangeaddin-salesforce-prod" |
                   ForEach-Object { (Get-ADUser -Identity $_.SamAccountName -Properties mail).Mail }
    
    if ("new.user@contoso.com" -in $finalMembers) {
        throw "Removed user should not be in final member list"
    }
    
    Write-Host "    User membership change validation completed" -ForegroundColor Gray
}

function Test-FullWorkflowSimulation {
    # Reset to clean state
    Reset-MockADData
    Reset-MockExchangeData
    
    # Create temporary state file
    $tempStateFile = Join-Path $env:TEMP "workflow-test-state.json"
    $tempLogPath = Join-Path $env:TEMP "workflow-test-logs"
    
    try {
        # Run the framework in WhatIf mode
        $frameworkScript = Join-Path $PSScriptRoot "..\src\ExchangeAddInManager.ps1"
        
        & $frameworkScript -UseMockServers -WhatIf -StateFilePath $tempStateFile -LogPath $tempLogPath -Verbose
        
        # Verify state file was created
        if (!(Test-Path $tempStateFile)) {
            throw "State file should be created after framework execution"
        }
        
        # Verify log directory was created
        if (!(Test-Path $tempLogPath)) {
            throw "Log directory should be created after framework execution"
        }
        
        Write-Host "    Full workflow simulation completed successfully" -ForegroundColor Gray
    }
    finally {
        # Cleanup
        Remove-Item $tempStateFile -Force -ErrorAction SilentlyContinue
        Remove-Item $tempLogPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Test-ErrorHandling {
    # Test invalid group access
    try {
        Get-ADGroupMember -Identity "nonexistent-group"
        throw "Should have thrown an error for nonexistent group"
    }
    catch {
        if ($_.Exception.Message -notlike "*not found*") {
            throw "Should get appropriate error message for missing group"
        }
    }
    
    # Test invalid user access
    try {
        Get-ADUser -Identity "nonexistent-user" -Properties mail
        throw "Should have thrown an error for nonexistent user"
    }
    catch {
        if ($_.Exception.Message -notlike "*not found*") {
            throw "Should get appropriate error message for missing user"
        }
    }
    
    # Test invalid app installation
    try {
        New-App -Mailbox "test@test.com" -Url "https://invalid-manifest-url.com/manifest.xml" -OrganizationApp
        throw "Should have thrown an error for invalid manifest URL"
    }
    catch {
        if ($_.Exception.Message -notlike "*not found*") {
            throw "Should get appropriate error message for invalid manifest"
        }
    }
    
    Write-Host "    Error handling validation completed" -ForegroundColor Gray
}

function Test-PerformanceWithLargeDataset {
    # Reset and create larger dataset
    Reset-MockADData
    Reset-MockExchangeData
    
    # Add many users
    for ($i = 1; $i -le 100; $i++) {
        Add-MockADUser -SamAccountName "perfuser$i" -Name "Performance User $i" -Mail "perfuser$i@contoso.com"
    }
    
    # Add users to groups
    for ($i = 1; $i -le 50; $i++) {
        Add-MockADGroupMember -GroupName "app-exchangeaddin-salesforce-prod" -UserSamAccountName "perfuser$i"
    }
    
    # Measure performance of group member retrieval
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    $members = Get-ADGroupMember -Identity "app-exchangeaddin-salesforce-prod"
    $emailMembers = @()
    foreach ($member in $members) {
        $user = Get-ADUser -Identity $member.SamAccountName -Properties mail
        if ($user.mail) {
            $emailMembers += $user.mail
        }
    }
    
    $stopwatch.Stop()
    
    if ($emailMembers.Count -lt 50) {
        throw "Should have at least 50 email members"
    }
    
    if ($stopwatch.ElapsedMilliseconds -gt 5000) { # 5 second threshold
        throw "Performance test took too long: $($stopwatch.ElapsedMilliseconds)ms"
    }
    
    Write-Host "    Performance test completed in $($stopwatch.ElapsedMilliseconds)ms with $($emailMembers.Count) members" -ForegroundColor Gray
}

#endregion

#region Main Execution

function Invoke-TestSuite {
    param([string]$Suite)
    
    $runner = [TestRunner]::new()
    
    Write-Host "Exchange Add-In Management Framework - Test Suite: $Suite" -ForegroundColor Cyan
    Write-Host "="*60 -ForegroundColor Cyan
    
    if ($Suite -in @('Basic', 'All')) {
        $runner.RunTest("Mock AD Basic Functionality", { Test-MockADBasicFunctionality })
        $runner.RunTest("Mock Exchange Basic Functionality", { Test-MockExchangeBasicFunctionality })
        $runner.RunTest("Add-In Group Pattern Validation", { Test-AddInGroupPattern })
        $runner.RunTest("State Management", { Test-StateManagement })
    }
    
    if ($Suite -in @('EdgeCases', 'All')) {
        $runner.RunTest("User Membership Changes", { Test-UserMembershipChanges })
        $runner.RunTest("Error Handling", { Test-ErrorHandling })
        $runner.RunTest("Full Workflow Simulation", { Test-FullWorkflowSimulation })
    }
    
    if ($Suite -in @('Performance', 'All')) {
        $runner.RunTest("Performance with Large Dataset", { Test-PerformanceWithLargeDataset })
    }
    
    $runner.ShowSummary()
    
    if ($GenerateTestReport) {
        $reportPath = Join-Path $PSScriptRoot "..\data\test-report-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
        $runner.GenerateReport($reportPath)
    }
    
    return $runner
}

# Execute tests
$testRunner = Invoke-TestSuite -Suite $TestSuite

if ($testRunner.FailedCount -gt 0) {
    exit 1
}

#endregion