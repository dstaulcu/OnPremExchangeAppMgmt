# Exchange Add-In Management Framework

A PowerShell-based framework for managing Exchange add-ins in on-premises environments through Active Directory group membership.

## Overview

This framework addresses the challenge of managing Exchange add-ins when Outlook Web Access is disabled and Office 365 Admin Center is not available. It uses Active Directory groups to control add-in installations and leverages Exchange PowerShell cmdlets to orchestrate add-in management.

## Architecture

- **Core Framework**: Main add-in management logic (`src/`)
- **Mock Servers**: Development-time simulators for AD and Exchange (`src/Mock/`)
- **Configuration**: Environment-specific settings (`config/`)
- **Data**: State files and test data (`data/`)
- **Tests**: Unit and integration tests (`tests/`)
- **Logs**: Runtime logs and audit trails (`logs/`)

## Quick Start

### Development Environment
```powershell
# Load mock servers for development
Import-Module .\src\Mock\MockActiveDirectory.psm1
Import-Module .\src\Mock\MockExchange.psm1

# Run the framework with mock data
.\src\ExchangeAddInManager.ps1 -UseMockServers
```

### Production Environment
```powershell
# Run against real AD and Exchange
.\src\ExchangeAddInManager.ps1 -ExchangeServer "exchange.domain.com" -Domain "domain.com"
```

## Active Directory Group Naming Convention

Groups should follow this pattern:
```
app-exchangeaddin-{add-in-name}-{environment}
```

Examples:
- `app-exchangeaddin-salesforce-prod`
- `app-exchangeaddin-docusign-test`
- `app-exchangeaddin-teams-dev`

## Features

- ✅ Automatic discovery of add-in groups
- ✅ User membership tracking and state management
- ✅ Automated add-in installation/removal
- ✅ Comprehensive logging and auditing
- ✅ Mock servers for development
- ✅ Environment separation (dev/test/prod)

## Requirements

- PowerShell 5.1 or later
- ActiveDirectory module (production)
- Exchange Management Shell access (production)
- Appropriate permissions for AD group reading and Exchange add-in management

## License

MIT License