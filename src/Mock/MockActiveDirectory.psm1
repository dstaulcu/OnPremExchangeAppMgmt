#Requires -Version 5.1

<#
.SYNOPSIS
    Mock Active Directory module for development and testing
.DESCRIPTION
    Provides mock implementations of ActiveDirectory cmdlets for testing
    the Exchange Add-In Management Framework without requiring actual AD infrastructure
#>

# Module-level variables to store mock data
$script:MockUsers = @()
$script:MockGroups = @()
$script:MockGroupMemberships = @{}

#region Mock Data Initialization

function Initialize-MockADData {
    <#
    .SYNOPSIS
        Initializes mock AD data with realistic test users and groups
    #>
    
    # Clear existing data
    $script:MockUsers = @()
    $script:MockGroups = @()
    $script:MockGroupMemberships = @{}
    
    # Create mock users
    $script:MockUsers = @(
        @{
            SamAccountName = "jdoe"
            Name = "John Doe"
            Mail = "john.doe@contoso.com"
            ObjectClass = "user"
            DistinguishedName = "CN=John Doe,OU=Users,DC=contoso,DC=com"
        },
        @{
            SamAccountName = "asmith" 
            Name = "Alice Smith"
            Mail = "alice.smith@contoso.com"
            ObjectClass = "user"
            DistinguishedName = "CN=Alice Smith,OU=Users,DC=contoso,DC=com"
        },
        @{
            SamAccountName = "bwilson"
            Name = "Bob Wilson" 
            Mail = "bob.wilson@contoso.com"
            ObjectClass = "user"
            DistinguishedName = "CN=Bob Wilson,OU=Users,DC=contoso,DC=com"
        },
        @{
            SamAccountName = "cjohnson"
            Name = "Carol Johnson"
            Mail = "carol.johnson@contoso.com"
            ObjectClass = "user"
            DistinguishedName = "CN=Carol Johnson,OU=Users,DC=contoso,DC=com"
        },
        @{
            SamAccountName = "dlee"
            Name = "David Lee"
            Mail = "david.lee@contoso.com"
            ObjectClass = "user"
            DistinguishedName = "CN=David Lee,OU=Users,DC=contoso,DC=com"
        }
    )
    
    # Create mock add-in groups
    $script:MockGroups = @(
        @{
            Name = "app-exchangeaddin-salesforce-prod"
            Description = "https://appexchange.salesforce.com/manifest.xml"
            ObjectClass = "group"
            DistinguishedName = "CN=app-exchangeaddin-salesforce-prod,OU=Groups,DC=contoso,DC=com"
            SamAccountName = "app-exchangeaddin-salesforce-prod"
        },
        @{
            Name = "app-exchangeaddin-docusign-test"
            Description = "https://apps.docusign.com/manifest.xml"
            ObjectClass = "group"
            DistinguishedName = "CN=app-exchangeaddin-docusign-test,OU=Groups,DC=contoso,DC=com"
            SamAccountName = "app-exchangeaddin-docusign-test"
        },
        @{
            Name = "app-exchangeaddin-teams-dev"
            Description = "https://teams.microsoft.com/dev/manifest.xml"
            ObjectClass = "group"
            DistinguishedName = "CN=app-exchangeaddin-teams-dev,OU=Groups,DC=contoso,DC=com"
            SamAccountName = "app-exchangeaddin-teams-dev"
        },
        @{
            Name = "app-exchangeaddin-onenote-prod"
            Description = "https://onenote.com/addin/manifest.xml"
            ObjectClass = "group"
            DistinguishedName = "CN=app-exchangeaddin-onenote-prod,OU=Groups,DC=contoso,DC=com"
            SamAccountName = "app-exchangeaddin-onenote-prod"
        },
        @{
            Name = "regular-group"
            Description = "A regular AD group not related to add-ins"
            ObjectClass = "group"
            DistinguishedName = "CN=regular-group,OU=Groups,DC=contoso,DC=com"
            SamAccountName = "regular-group"
        }
    )
    
    # Create mock group memberships
    $script:MockGroupMemberships = @{
        "app-exchangeaddin-salesforce-prod" = @("jdoe", "asmith", "bwilson")
        "app-exchangeaddin-docusign-test" = @("asmith", "cjohnson")
        "app-exchangeaddin-teams-dev" = @("dlee")
        "app-exchangeaddin-onenote-prod" = @("jdoe", "cjohnson", "dlee")
        "regular-group" = @("jdoe", "asmith")
    }
    
    Write-Verbose "Mock AD data initialized with $($script:MockUsers.Count) users and $($script:MockGroups.Count) groups"
}

#endregion

#region Mock ActiveDirectory Cmdlets

function Get-ADGroup {
    <#
    .SYNOPSIS
        Mock implementation of Get-ADGroup cmdlet
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Filter,
        
        [string]$Server,
        
        [string[]]$Properties = @()
    )
    
    if ([string]::IsNullOrEmpty($script:MockGroups)) {
        Initialize-MockADData
    }
    
    # Parse the filter (simplified - only handles Name -like patterns)
    if ($Filter -match "Name -like '(.+)'") {
        $namePattern = $matches[1]
        $namePattern = $namePattern.Replace('*', '.*')
        
        $matchingGroups = $script:MockGroups | Where-Object { 
            $_.Name -match $namePattern 
        }
    } else {
        Write-Warning "Mock Get-ADGroup only supports 'Name -like' filters"
        return @()
    }
    
    # Return groups with requested properties
    foreach ($group in $matchingGroups) {
        $result = [PSCustomObject]@{
            Name = $group.Name
            ObjectClass = $group.ObjectClass
            DistinguishedName = $group.DistinguishedName
            SamAccountName = $group.SamAccountName
        }
        
        if ($Properties -contains "Description") {
            $result | Add-Member -NotePropertyName "Description" -NotePropertyValue $group.Description
        }
        
        Write-Output $result
    }
}

function Get-ADGroupMember {
    <#
    .SYNOPSIS
        Mock implementation of Get-ADGroupMember cmdlet
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Identity,
        
        [string]$Server
    )
    
    if ([string]::IsNullOrEmpty($script:MockGroupMemberships)) {
        Initialize-MockADData
    }
    
    if ($script:MockGroupMemberships.ContainsKey($Identity)) {
        $memberSamAccountNames = $script:MockGroupMemberships[$Identity]
        
        foreach ($samAccountName in $memberSamAccountNames) {
            $user = $script:MockUsers | Where-Object { $_.SamAccountName -eq $samAccountName }
            if ($user) {
                Write-Output ([PSCustomObject]@{
                    SamAccountName = $user.SamAccountName
                    Name = $user.Name
                    ObjectClass = $user.ObjectClass
                    DistinguishedName = $user.DistinguishedName
                })
            }
        }
    } else {
        Write-Warning "Group '$Identity' not found in mock data"
    }
}

function Get-ADUser {
    <#
    .SYNOPSIS
        Mock implementation of Get-ADUser cmdlet
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Identity,
        
        [string[]]$Properties = @(),
        
        [string]$Server
    )
    
    if ([string]::IsNullOrEmpty($script:MockUsers)) {
        Initialize-MockADData
    }
    
    $user = $script:MockUsers | Where-Object { 
        $_.SamAccountName -eq $Identity -or $_.Name -eq $Identity 
    }
    
    if ($user) {
        $result = [PSCustomObject]@{
            SamAccountName = $user.SamAccountName
            Name = $user.Name
            ObjectClass = $user.ObjectClass
            DistinguishedName = $user.DistinguishedName
        }
        
        if ($Properties -contains "mail") {
            $result | Add-Member -NotePropertyName "Mail" -NotePropertyValue $user.Mail
        }
        
        Write-Output $result
    } else {
        Write-Warning "User '$Identity' not found in mock data"
    }
}

#endregion

#region Mock Data Management Functions

function Add-MockADUser {
    <#
    .SYNOPSIS
        Adds a user to the mock AD data
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName,
        
        [Parameter(Mandatory=$true)]
        [string]$Name,
        
        [Parameter(Mandatory=$true)]
        [string]$Mail
    )
    
    $script:MockUsers += @{
        SamAccountName = $SamAccountName
        Name = $Name
        Mail = $Mail
        ObjectClass = "user"
        DistinguishedName = "CN=$Name,OU=Users,DC=contoso,DC=com"
    }
    
    Write-Verbose "Added mock user: $SamAccountName ($Mail)"
}

function Add-MockADGroupMember {
    <#
    .SYNOPSIS
        Adds a user to a mock AD group
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        
        [Parameter(Mandatory=$true)]
        [string]$UserSamAccountName
    )
    
    if (-not $script:MockGroupMemberships.ContainsKey($GroupName)) {
        $script:MockGroupMemberships[$GroupName] = @()
    }
    
    if ($UserSamAccountName -notin $script:MockGroupMemberships[$GroupName]) {
        $script:MockGroupMemberships[$GroupName] += $UserSamAccountName
        Write-Verbose "Added $UserSamAccountName to group $GroupName"
    } else {
        Write-Warning "$UserSamAccountName is already a member of $GroupName"
    }
}

function Remove-MockADGroupMember {
    <#
    .SYNOPSIS
        Removes a user from a mock AD group
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        
        [Parameter(Mandatory=$true)]
        [string]$UserSamAccountName
    )
    
    if ($script:MockGroupMemberships.ContainsKey($GroupName)) {
        $script:MockGroupMemberships[$GroupName] = $script:MockGroupMemberships[$GroupName] | 
            Where-Object { $_ -ne $UserSamAccountName }
        Write-Verbose "Removed $UserSamAccountName from group $GroupName"
    } else {
        Write-Warning "Group $GroupName not found"
    }
}

function Get-MockADGroupMemberships {
    <#
    .SYNOPSIS
        Returns all current group memberships for debugging
    #>
    return $script:MockGroupMemberships
}

function Reset-MockADData {
    <#
    .SYNOPSIS
        Resets all mock AD data to initial state
    #>
    Initialize-MockADData
    Write-Verbose "Mock AD data reset to initial state"
}

#endregion

# Initialize mock data when module is imported
Initialize-MockADData

# Export public functions
Export-ModuleMember -Function @(
    'Get-ADGroup',
    'Get-ADGroupMember', 
    'Get-ADUser',
    'Add-MockADUser',
    'Add-MockADGroupMember',
    'Remove-MockADGroupMember',
    'Get-MockADGroupMemberships',
    'Reset-MockADData'
)