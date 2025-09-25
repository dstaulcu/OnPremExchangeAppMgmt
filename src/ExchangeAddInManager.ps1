#Requires -Version 5.1

<#
.SYNOPSIS
    Exchange Add-In Management Framework
.DESCRIPTION
    Manages Exchange add-ins in on-premises environments by monitoring Active Directory 
    group membership and orchestrating add-in installations/removals via Exchange PowerShell cmdlets
.EXAMPLE
    # Development mode with mock servers
    .\ExchangeAddInManager.ps1 -UseMockServers -Verbose
.EXAMPLE
    # Production mode
    .\ExchangeAddInManager.ps1 -ExchangeServer "exchange.contoso.com" -Domain "contoso.com"
#>

[CmdletBinding()]
param(
    [Parameter(ParameterSetName='Production', Mandatory=$true)]
    [string]$ExchangeServer,
    
    [Parameter(ParameterSetName='Production', Mandatory=$true)]
    [string]$Domain,
    
    [Parameter(ParameterSetName='Mock')]
    [switch]$UseMockServers,
    
    [string]$GroupPattern = "app-exchangeaddin-*",
    
    [string]$StateFilePath = ".\data\state.json",
    
    [string]$LogPath = ".\logs",
    
    [switch]$WhatIf
)

# Import required modules based on mode
if ($UseMockServers) {
    Write-Verbose "Loading mock servers for development..."
    $mockADPath = Join-Path $PSScriptRoot "Mock\MockActiveDirectory.psm1"
    $mockExchangePath = Join-Path $PSScriptRoot "Mock\MockExchange.psm1"
    
    if (Test-Path $mockADPath) {
        Import-Module $mockADPath -Force
    } else {
        throw "Mock Active Directory module not found at: $mockADPath"
    }
    
    if (Test-Path $mockExchangePath) {
        Import-Module $mockExchangePath -Force
    } else {
        throw "Mock Exchange module not found at: $mockExchangePath"
    }
    
    # Set default values for mock mode
    $Domain = "contoso.com"
    $ExchangeServer = "mock-exchange.contoso.com"
} else {
    # Production mode - verify required modules
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw "ActiveDirectory module is required for production mode. Please install RSAT-AD-PowerShell feature."
    }
    Import-Module ActiveDirectory
}

#region Configuration Classes

class AddInConfig {
    [string]$GroupName
    [string]$AddInName
    [string]$Environment
    [string]$ManifestUrl
    [string[]]$CurrentMembers
    [string[]]$PreviousMembers
    
    AddInConfig() {
        $this.CurrentMembers = @()
        $this.PreviousMembers = @()
    }
    
    [string[]] GetUsersToAdd() {
        return $this.CurrentMembers | Where-Object { $_ -notin $this.PreviousMembers }
    }
    
    [string[]] GetUsersToRemove() {
        return $this.PreviousMembers | Where-Object { $_ -notin $this.CurrentMembers }
    }
}

#endregion

#region Main Management Class

class ExchangeAddInManager {
    [string]$ExchangeServer
    [string]$Domain
    [string]$GroupPattern
    [string]$StateFilePath
    [string]$LogPath
    [bool]$UseMockServers
    [bool]$WhatIf
    [System.Collections.ArrayList]$AddInConfigs
    [hashtable]$Statistics

    ExchangeAddInManager([string]$exchangeServer, [string]$domain, [string]$groupPattern, 
                        [string]$stateFile, [string]$logPath, [bool]$useMockServers, [bool]$whatIf) {
        $this.ExchangeServer = $exchangeServer
        $this.Domain = $domain
        $this.GroupPattern = $groupPattern
        $this.StateFilePath = $stateFile
        $this.LogPath = $logPath
        $this.UseMockServers = $useMockServers
        $this.WhatIf = $whatIf
        $this.AddInConfigs = @()
        $this.Statistics = @{
            GroupsFound = 0
            UsersToAdd = 0
            UsersToRemove = 0
            AddInsInstalled = 0
            AddInsRemoved = 0
            Errors = 0
        }
        
        # Ensure directories exist
        $this.EnsureDirectories()
    }

    [void] EnsureDirectories() {
        $stateDir = Split-Path $this.StateFilePath -Parent
        if (![string]::IsNullOrEmpty($stateDir) -and !(Test-Path $stateDir)) {
            New-Item -Path $stateDir -ItemType Directory -Force | Out-Null
            Write-Verbose "Created state directory: $stateDir"
        }
        
        if (!(Test-Path $this.LogPath)) {
            New-Item -Path $this.LogPath -ItemType Directory -Force | Out-Null
            Write-Verbose "Created log directory: $($this.LogPath)"
        }
    }

    [void] DiscoverAddInGroups() {
        Write-Host "Discovering add-in groups with pattern: $($this.GroupPattern)" -ForegroundColor Cyan
        $this.LogInfo("DiscoverAddInGroups", "Starting group discovery with pattern: $($this.GroupPattern)")
        
        try {
            $serverParam = @{}
            if (!$this.UseMockServers) {
                $serverParam['Server'] = $this.Domain
            }
            
            $groups = Get-ADGroup -Filter "Name -like '$($this.GroupPattern)'" @serverParam -Properties Description
            $this.Statistics.GroupsFound = $groups.Count
            
            Write-Host "Found $($groups.Count) matching groups" -ForegroundColor Green
            
            foreach ($group in $groups) {
                if ($group.Name -match "app-exchangeaddin-(.+)-(.+)") {
                    $addInName = $matches[1]
                    $environment = $matches[2]
                    
                    # Get manifest URL from group description
                    $manifestUrl = $group.Description
                    if ([string]::IsNullOrEmpty($manifestUrl)) {
                        Write-Warning "No manifest URL found in description for group: $($group.Name)"
                        $this.LogWarning("DiscoverAddInGroups", "No manifest URL for group: $($group.Name)")
                        continue
                    }
                    
                    $config = [AddInConfig]::new()
                    $config.GroupName = $group.Name
                    $config.AddInName = $addInName
                    $config.Environment = $environment
                    $config.ManifestUrl = $manifestUrl
                    $config.CurrentMembers = $this.GetGroupMembers($group.Name)
                    
                    $this.AddInConfigs.Add($config) | Out-Null
                    
                    Write-Host "  - $($group.Name): $($config.CurrentMembers.Count) members" -ForegroundColor Gray
                    $this.LogInfo("DiscoverAddInGroups", "Found group $($group.Name) with $($config.CurrentMembers.Count) members")
                } else {
                    Write-Verbose "Group name doesn't match expected pattern: $($group.Name)"
                }
            }
        }
        catch {
            $errorMsg = "Failed to discover add-in groups: $($_.Exception.Message)"
            Write-Error $errorMsg
            $this.LogError("DiscoverAddInGroups", $errorMsg)
            $this.Statistics.Errors++
        }
    }

    [string[]] GetGroupMembers([string]$groupName) {
        try {
            $serverParam = @{}
            if (!$this.UseMockServers) {
                $serverParam['Server'] = $this.Domain
            }
            
            $members = Get-ADGroupMember -Identity $groupName @serverParam | 
                       Where-Object { $_.objectClass -eq 'user' }
            
            $emailAddresses = @()
            
            foreach ($member in $members) {
                $user = Get-ADUser -Identity $member.SamAccountName -Properties mail @serverParam
                if ($user.mail) {
                    $emailAddresses += $user.mail
                } else {
                    Write-Warning "No email address found for user: $($member.SamAccountName)"
                    $this.LogWarning("GetGroupMembers", "No email for user: $($member.SamAccountName)")
                }
            }
            
            return $emailAddresses
        }
        catch {
            $errorMsg = "Failed to get group members for '$groupName': $($_.Exception.Message)"
            Write-Error $errorMsg
            $this.LogError("GetGroupMembers", $errorMsg)
            $this.Statistics.Errors++
            return @()
        }
    }

    [void] LoadPreviousState() {
        Write-Host "Loading previous state..." -ForegroundColor Cyan
        
        if (Test-Path $this.StateFilePath) {
            try {
                $stateData = Get-Content $this.StateFilePath -Raw | ConvertFrom-Json
                
                foreach ($config in $this.AddInConfigs) {
                    $previousConfig = $stateData | Where-Object { $_.GroupName -eq $config.GroupName }
                    if ($previousConfig) {
                        $config.PreviousMembers = $previousConfig.CurrentMembers
                        Write-Verbose "Loaded previous state for $($config.GroupName): $($config.PreviousMembers.Count) members"
                    }
                }
                
                $this.LogInfo("LoadPreviousState", "Successfully loaded previous state")
            }
            catch {
                $errorMsg = "Failed to load previous state: $($_.Exception.Message)"
                Write-Warning $errorMsg
                $this.LogWarning("LoadPreviousState", $errorMsg)
                
                # Initialize empty previous state
                foreach ($config in $this.AddInConfigs) {
                    $config.PreviousMembers = @()
                }
            }
        } else {
            Write-Host "No previous state file found - treating as initial run" -ForegroundColor Yellow
            $this.LogInfo("LoadPreviousState", "No previous state file found")
        }
    }

    [void] SaveCurrentState() {
        Write-Host "Saving current state..." -ForegroundColor Cyan
        
        try {
            $stateData = $this.AddInConfigs | ForEach-Object {
                [PSCustomObject]@{
                    GroupName = $_.GroupName
                    AddInName = $_.AddInName
                    Environment = $_.Environment
                    ManifestUrl = $_.ManifestUrl
                    CurrentMembers = $_.CurrentMembers
                    LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
            
            $stateData | ConvertTo-Json -Depth 3 | Set-Content $this.StateFilePath -Encoding UTF8
            $this.LogInfo("SaveCurrentState", "Successfully saved current state")
        }
        catch {
            $errorMsg = "Failed to save state: $($_.Exception.Message)"
            Write-Error $errorMsg
            $this.LogError("SaveCurrentState", $errorMsg)
            $this.Statistics.Errors++
        }
    }

    [void] ProcessAddInChanges() {
        Write-Host "Processing add-in changes..." -ForegroundColor Cyan
        
        foreach ($config in $this.AddInConfigs) {
            Write-Host "Processing $($config.GroupName)..." -ForegroundColor White
            
            $usersToAdd = $config.GetUsersToAdd()
            $usersToRemove = $config.GetUsersToRemove()
            
            $this.Statistics.UsersToAdd += $usersToAdd.Count
            $this.Statistics.UsersToRemove += $usersToRemove.Count
            
            if ($usersToAdd.Count -eq 0 -and $usersToRemove.Count -eq 0) {
                Write-Host "  No changes detected" -ForegroundColor Gray
                continue
            }
            
            Write-Host "  Users to add ($($usersToAdd.Count)): $($usersToAdd -join ', ')" -ForegroundColor Green
            Write-Host "  Users to remove ($($usersToRemove.Count)): $($usersToRemove -join ', ')" -ForegroundColor Red
            
            foreach ($user in $usersToAdd) {
                $this.InstallAddIn($user, $config)
            }
            
            foreach ($user in $usersToRemove) {
                $this.RemoveAddIn($user, $config)
            }
        }
    }

    [void] InstallAddIn([string]$userEmail, [AddInConfig]$config) {
        $action = "Installing add-in '$($config.AddInName)' for user: $userEmail"
        Write-Host "  $action" -ForegroundColor Green
        
        if ($this.WhatIf) {
            Write-Host "    [WHATIF] Would install add-in" -ForegroundColor Magenta
            $this.LogInfo("InstallAddIn", "[WHATIF] $action")
            return
        }
        
        try {
            if ($this.UseMockServers) {
                # Use mock Exchange cmdlet
                New-App -Mailbox $userEmail -Url $config.ManifestUrl -OrganizationApp | Out-Null
            } else {
                # Use real Exchange cmdlet via remote session
                $this.EnsureExchangeConnection()
                
                Invoke-Command -ComputerName $this.ExchangeServer -ScriptBlock {
                    param($UserEmail, $ManifestUrl)
                    New-App -OrganizationApp -Mailbox $UserEmail -Url $ManifestUrl
                } -ArgumentList $userEmail, $config.ManifestUrl | Out-Null
            }
            
            $this.Statistics.AddInsInstalled++
            $this.LogInfo("InstallAddIn", "Successfully installed $($config.AddInName) for $userEmail")
        }
        catch {
            $errorMsg = "Failed to install add-in for $userEmail`: $($_.Exception.Message)"
            Write-Error "    $errorMsg"
            $this.LogError("InstallAddIn", $errorMsg)
            $this.Statistics.Errors++
        }
    }

    [void] RemoveAddIn([string]$userEmail, [AddInConfig]$config) {
        $action = "Removing add-in '$($config.AddInName)' from user: $userEmail"
        Write-Host "  $action" -ForegroundColor Red
        
        if ($this.WhatIf) {
            Write-Host "    [WHATIF] Would remove add-in" -ForegroundColor Magenta
            $this.LogInfo("RemoveAddIn", "[WHATIF] $action")
            return
        }
        
        try {
            # First, find the installed app
            $installedApp = $null
            
            if ($this.UseMockServers) {
                $installedApp = Get-App -Mailbox $userEmail | 
                               Where-Object { $_.DisplayName -like "*$($config.AddInName)*" }
            } else {
                $this.EnsureExchangeConnection()
                
                $installedApp = Invoke-Command -ComputerName $this.ExchangeServer -ScriptBlock {
                    param($UserEmail, $AddInName)
                    Get-App -Mailbox $UserEmail | Where-Object { $_.DisplayName -like "*$AddInName*" }
                } -ArgumentList $userEmail, $config.AddInName
            }
            
            if ($installedApp) {
                if ($this.UseMockServers) {
                    Remove-App -Mailbox $userEmail -Identity $installedApp.Identity
                } else {
                    Invoke-Command -ComputerName $this.ExchangeServer -ScriptBlock {
                        param($UserEmail, $AppIdentity)
                        Remove-App -Mailbox $UserEmail -Identity $AppIdentity
                    } -ArgumentList $userEmail, $installedApp.Identity
                }
                
                $this.Statistics.AddInsRemoved++
                $this.LogInfo("RemoveAddIn", "Successfully removed $($config.AddInName) from $userEmail")
            } else {
                Write-Warning "    Add-in not found for removal: $($config.AddInName) for user $userEmail"
                $this.LogWarning("RemoveAddIn", "Add-in not found: $($config.AddInName) for $userEmail")
            }
        }
        catch {
            $errorMsg = "Failed to remove add-in from $userEmail`: $($_.Exception.Message)"
            Write-Error "    $errorMsg"
            $this.LogError("RemoveAddIn", $errorMsg)
            $this.Statistics.Errors++
        }
    }

    [void] EnsureExchangeConnection() {
        if ($this.UseMockServers) {
            return # No connection needed for mock mode
        }
        
        # In production, you would establish a remote PowerShell session to Exchange
        # This is a placeholder for the actual implementation
        Write-Verbose "Exchange connection would be established here for server: $($this.ExchangeServer)"
    }

    [void] ShowStatistics() {
        Write-Host "`nExecution Summary:" -ForegroundColor Cyan
        Write-Host "=================" -ForegroundColor Cyan
        Write-Host "Groups found: $($this.Statistics.GroupsFound)" -ForegroundColor White
        Write-Host "Users to add add-ins: $($this.Statistics.UsersToAdd)" -ForegroundColor White
        Write-Host "Users to remove add-ins: $($this.Statistics.UsersToRemove)" -ForegroundColor White
        Write-Host "Add-ins installed: $($this.Statistics.AddInsInstalled)" -ForegroundColor Green
        Write-Host "Add-ins removed: $($this.Statistics.AddInsRemoved)" -ForegroundColor Red
        Write-Host "Errors encountered: $($this.Statistics.Errors)" -ForegroundColor $(if($this.Statistics.Errors -eq 0) { "Green" } else { "Red" })
    }

    # Logging methods
    [void] LogInfo([string]$operation, [string]$message) {
        $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [INFO] $operation`: $message"
        $logFile = Join-Path $this.LogPath "info-$(Get-Date -Format 'yyyy-MM-dd').log"
        Add-Content -Path $logFile -Value $logEntry -Encoding UTF8
    }

    [void] LogWarning([string]$operation, [string]$message) {
        $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [WARNING] $operation`: $message"
        $logFile = Join-Path $this.LogPath "warning-$(Get-Date -Format 'yyyy-MM-dd').log"
        Add-Content -Path $logFile -Value $logEntry -Encoding UTF8
    }

    [void] LogError([string]$operation, [string]$message) {
        $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [ERROR] $operation`: $message"
        $logFile = Join-Path $this.LogPath "error-$(Get-Date -Format 'yyyy-MM-dd').log"
        Add-Content -Path $logFile -Value $logEntry -Encoding UTF8
    }
}

#endregion

#region Main Execution

try {
    Write-Host "Exchange Add-In Management Framework" -ForegroundColor Cyan
    Write-Host "====================================" -ForegroundColor Cyan
    Write-Host "Mode: $(if($UseMockServers) { 'Development (Mock)' } else { 'Production' })" -ForegroundColor Yellow
    Write-Host "What-If: $WhatIf" -ForegroundColor Yellow
    Write-Host ""
    
    # Create manager instance
    $manager = [ExchangeAddInManager]::new(
        $ExchangeServer, 
        $Domain, 
        $GroupPattern, 
        $StateFilePath, 
        $LogPath, 
        $UseMockServers.IsPresent, 
        $WhatIf.IsPresent
    )
    
    # Execute management process
    $manager.DiscoverAddInGroups()
    $manager.LoadPreviousState()
    $manager.ProcessAddInChanges()
    $manager.SaveCurrentState()
    $manager.ShowStatistics()
    
    Write-Host "`nExchange Add-In Management Process Completed Successfully" -ForegroundColor Green
}
catch {
    Write-Error "Management process failed: $($_.Exception.Message)"
    Write-Error "Stack Trace: $($_.ScriptStackTrace)"
    exit 1
}

#endregion