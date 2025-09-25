#Requires -Version 5.1

<#
.SYNOPSIS
    Mock Exchange Management Shell module for development and testing
.DESCRIPTION
    Provides mock implementations of Exchange cmdlets for testing
    the Exchange Add-In Management Framework without requiring actual Exchange infrastructure
#>

# Module-level variables to store mock data
$script:MockInstalledApps = @{}
$script:MockAppDefinitions = @{}

#region Mock Data Initialization

function Initialize-MockExchangeData {
    <#
    .SYNOPSIS
        Initializes mock Exchange data with app definitions and installations
    #>
    
    # Clear existing data
    $script:MockInstalledApps = @{}
    $script:MockAppDefinitions = @{}
    
    # Define mock app manifests
    $script:MockAppDefinitions = @{
        "https://appexchange.salesforce.com/manifest.xml" = @{
            DisplayName = "Salesforce Lightning for Outlook"
            AppId = "salesforce-outlook-addin"
            Version = "1.2.3"
            Publisher = "Salesforce.com"
            Description = "Integrate Salesforce with Outlook"
            ManifestUrl = "https://appexchange.salesforce.com/manifest.xml"
        }
        "https://apps.docusign.com/manifest.xml" = @{
            DisplayName = "DocuSign for Outlook"
            AppId = "docusign-outlook-addin"
            Version = "2.1.0"
            Publisher = "DocuSign Inc."
            Description = "Send documents for signature from Outlook"
            ManifestUrl = "https://apps.docusign.com/manifest.xml"
        }
        "https://teams.microsoft.com/dev/manifest.xml" = @{
            DisplayName = "Microsoft Teams Meeting Add-in"
            AppId = "teams-meeting-addin"
            Version = "1.0.15"
            Publisher = "Microsoft Corporation"
            Description = "Schedule Teams meetings from Outlook"
            ManifestUrl = "https://teams.microsoft.com/dev/manifest.xml"
        }
        "https://onenote.com/addin/manifest.xml" = @{
            DisplayName = "OneNote Web Clipper"
            AppId = "onenote-web-clipper"
            Version = "3.4.1"
            Publisher = "Microsoft Corporation"
            Description = "Save emails and content to OneNote"
            ManifestUrl = "https://onenote.com/addin/manifest.xml"
        }
    }
    
    # Initialize with some pre-installed apps for testing
    $script:MockInstalledApps = @{
        "john.doe@contoso.com" = @{
            "salesforce-outlook-addin" = @{
                Identity = "salesforce-outlook-addin"
                DisplayName = "Salesforce Lightning for Outlook"
                AppId = "salesforce-outlook-addin"
                Mailbox = "john.doe@contoso.com"
                Enabled = $true
                InstallDate = (Get-Date).AddDays(-30)
                ManifestUrl = "https://appexchange.salesforce.com/manifest.xml"
            }
        }
        "alice.smith@contoso.com" = @{
            "salesforce-outlook-addin" = @{
                Identity = "salesforce-outlook-addin"
                DisplayName = "Salesforce Lightning for Outlook" 
                AppId = "salesforce-outlook-addin"
                Mailbox = "alice.smith@contoso.com"
                Enabled = $true
                InstallDate = (Get-Date).AddDays(-15)
                ManifestUrl = "https://appexchange.salesforce.com/manifest.xml"
            }
            "docusign-outlook-addin" = @{
                Identity = "docusign-outlook-addin"
                DisplayName = "DocuSign for Outlook"
                AppId = "docusign-outlook-addin" 
                Mailbox = "alice.smith@contoso.com"
                Enabled = $true
                InstallDate = (Get-Date).AddDays(-7)
                ManifestUrl = "https://apps.docusign.com/manifest.xml"
            }
        }
    }
    
    Write-Verbose "Mock Exchange data initialized with $($script:MockAppDefinitions.Count) app definitions"
}

#endregion

#region Mock Exchange Cmdlets

function New-App {
    <#
    .SYNOPSIS
        Mock implementation of New-App cmdlet for installing Exchange add-ins
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory=$true)]
        [string]$Url,
        
        [switch]$OrganizationApp,
        
        [switch]$PrivateCatalog,
        
        [string]$Enabled = "True"
    )
    
    if ([string]::IsNullOrEmpty($script:MockAppDefinitions)) {
        Initialize-MockExchangeData
    }
    
    # Check if we have a definition for this manifest URL
    if (-not $script:MockAppDefinitions.ContainsKey($Url)) {
        throw "App manifest not found at URL: $Url"
    }
    
    $appDef = $script:MockAppDefinitions[$Url]
    
    # Initialize user's app collection if needed
    if (-not $script:MockInstalledApps.ContainsKey($Mailbox)) {
        $script:MockInstalledApps[$Mailbox] = @{}
    }
    
    # Check if app is already installed
    if ($script:MockInstalledApps[$Mailbox].ContainsKey($appDef.AppId)) {
        Write-Warning "App '$($appDef.DisplayName)' is already installed for mailbox '$Mailbox'"
        return
    }
    
    # Create new app installation
    $appInstallation = @{
        Identity = $appDef.AppId
        DisplayName = $appDef.DisplayName
        AppId = $appDef.AppId
        Mailbox = $Mailbox
        Enabled = [System.Convert]::ToBoolean($Enabled)
        InstallDate = Get-Date
        ManifestUrl = $Url
        Version = $appDef.Version
        Publisher = $appDef.Publisher
        Description = $appDef.Description
    }
    
    # Install the app
    $script:MockInstalledApps[$Mailbox][$appDef.AppId] = $appInstallation
    
    Write-Host "Successfully installed '$($appDef.DisplayName)' for mailbox '$Mailbox'" -ForegroundColor Green
    
    # Return the installed app object
    return [PSCustomObject]$appInstallation
}

function Remove-App {
    <#
    .SYNOPSIS
        Mock implementation of Remove-App cmdlet for uninstalling Exchange add-ins
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory=$true)]
        [string]$Identity
    )
    
    if ([string]::IsNullOrEmpty($script:MockInstalledApps)) {
        Initialize-MockExchangeData
    }
    
    # Check if user has any apps installed
    if (-not $script:MockInstalledApps.ContainsKey($Mailbox)) {
        Write-Warning "No apps found for mailbox '$Mailbox'"
        return
    }
    
    # Check if specific app is installed
    if (-not $script:MockInstalledApps[$Mailbox].ContainsKey($Identity)) {
        Write-Warning "App with identity '$Identity' not found for mailbox '$Mailbox'"
        return
    }
    
    $appName = $script:MockInstalledApps[$Mailbox][$Identity].DisplayName
    
    # Remove the app
    $script:MockInstalledApps[$Mailbox].Remove($Identity)
    
    Write-Host "Successfully removed '$appName' from mailbox '$Mailbox'" -ForegroundColor Green
}

function Get-App {
    <#
    .SYNOPSIS
        Mock implementation of Get-App cmdlet for listing installed Exchange add-ins
    #>
    [CmdletBinding()]
    param(
        [string]$Mailbox,
        
        [string]$Identity,
        
        [switch]$OrganizationApp
    )
    
    if ([string]::IsNullOrEmpty($script:MockInstalledApps)) {
        Initialize-MockExchangeData
    }
    
    $results = @()
    
    if ($Mailbox) {
        # Get apps for specific mailbox
        if ($script:MockInstalledApps.ContainsKey($Mailbox)) {
            $userApps = $script:MockInstalledApps[$Mailbox]
            
            if ($Identity) {
                # Get specific app for specific user
                if ($userApps.ContainsKey($Identity)) {
                    $results += [PSCustomObject]$userApps[$Identity]
                }
            } else {
                # Get all apps for specific user
                foreach ($app in $userApps.Values) {
                    $results += [PSCustomObject]$app
                }
            }
        }
    } else {
        # Get apps for all users
        foreach ($mailbox in $script:MockInstalledApps.Keys) {
            $userApps = $script:MockInstalledApps[$mailbox]
            
            if ($Identity) {
                # Get specific app for all users who have it
                if ($userApps.ContainsKey($Identity)) {
                    $results += [PSCustomObject]$userApps[$Identity]
                }
            } else {
                # Get all apps for all users
                foreach ($app in $userApps.Values) {
                    $results += [PSCustomObject]$app
                }
            }
        }
    }
    
    return $results
}

function Test-AppManifest {
    <#
    .SYNOPSIS
        Mock function to test if an app manifest URL is valid
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ManifestUrl
    )
    
    if ([string]::IsNullOrEmpty($script:MockAppDefinitions)) {
        Initialize-MockExchangeData
    }
    
    return $script:MockAppDefinitions.ContainsKey($ManifestUrl)
}

#endregion

#region Mock Data Management Functions

function Add-MockAppDefinition {
    <#
    .SYNOPSIS
        Adds a new app definition to mock Exchange data
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ManifestUrl,
        
        [Parameter(Mandatory=$true)]
        [string]$DisplayName,
        
        [Parameter(Mandatory=$true)]
        [string]$AppId,
        
        [string]$Version = "1.0.0",
        
        [string]$Publisher = "Unknown Publisher",
        
        [string]$Description = ""
    )
    
    $script:MockAppDefinitions[$ManifestUrl] = @{
        DisplayName = $DisplayName
        AppId = $AppId
        Version = $Version
        Publisher = $Publisher
        Description = $Description
        ManifestUrl = $ManifestUrl
    }
    
    Write-Verbose "Added mock app definition: $DisplayName"
}

function Get-MockInstalledApps {
    <#
    .SYNOPSIS
        Returns all currently installed apps for debugging
    #>
    return $script:MockInstalledApps
}

function Get-MockAppDefinitions {
    <#
    .SYNOPSIS
        Returns all available app definitions for debugging
    #>
    return $script:MockAppDefinitions
}

function Reset-MockExchangeData {
    <#
    .SYNOPSIS
        Resets all mock Exchange data to initial state
    #>
    Initialize-MockExchangeData
    Write-Verbose "Mock Exchange data reset to initial state"
}

function Set-MockAppInstallation {
    <#
    .SYNOPSIS
        Manually sets an app installation state for testing scenarios
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory=$true)]
        [string]$AppId,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$AppData
    )
    
    if (-not $script:MockInstalledApps.ContainsKey($Mailbox)) {
        $script:MockInstalledApps[$Mailbox] = @{}
    }
    
    $script:MockInstalledApps[$Mailbox][$AppId] = $AppData
    Write-Verbose "Set mock app installation: $AppId for $Mailbox"
}

#endregion

# Initialize mock data when module is imported
Initialize-MockExchangeData

# Export public functions
Export-ModuleMember -Function @(
    'New-App',
    'Remove-App',
    'Get-App',
    'Test-AppManifest',
    'Add-MockAppDefinition',
    'Get-MockInstalledApps',
    'Get-MockAppDefinitions',
    'Reset-MockExchangeData',
    'Set-MockAppInstallation'
)