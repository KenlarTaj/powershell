# PowerShell

PowerShell scripts for Windows and Microsoft 365 management.

## Overview

This repository contains PowerShell scripts designed to help automate and manage Windows environments and Microsoft 365 resources. All scripts are written to be clear, reusable, and follow best practices for PowerShell scripting.

## Scripts

| Script                     | Description                                                        |
|----------------------------|------------------------------------------------------------------|
| [Get-MSIInfo.ps1](#get-msiinfops1)            | Retrieve installed MSI/AppX apps or MSI file properties.         |
| [Get-Resources.ps1](#get-resourcesps1)        | Gather system information, performance, and session/user data.   |
| [Get-LicensesByService.ps1](#get-licensesbyserviceps1)   | List Microsoft 365 licenses that include a specific service.     |
| [Get-ServicesByLicense.ps1](#get-servicesbylicenseps1)   | Display services included in a specific Microsoft 365 license SKU. |
---

### Get-MSIInfo.ps1

Retrieves information about installed MSI and AppX applications or displays properties of a specified MSI file.

#### Features

- Search for installed applications (MSI/AppX) by name or GUID.
- List properties of a given MSI file.
- Filter results by installation context (system or user).
- Optionally export results to a CSV file.
- Compatible with PowerShell 5.1+ and Windows 10/11.

#### Running as Administrator vs. Standard User

- As a standard user, the script can access user-installed applications (HKCU) and AppX packages for the current user.
- As administrator, the script can access system-installed applications (HKLM) and AppX packages for all users, providing a more complete inventory.
- Some AppX queries (especially with -AllUsers) may require administrative privileges.

#### Usage Examples

    .\Get-MSIInfo.ps1 "Office"
    .\Get-MSIInfo.ps1 -f "C:\Installers\app.msi"
    .\Get-MSIInfo.ps1 -s "Adobe"
    .\Get-MSIInfo.ps1 -u
    .\Get-MSIInfo.ps1 notepad -ExportCsv ".\output.csv"

---

### Get-Resources.ps1

Gathers system and user session information including performance metrics, OS data, and logon details. Useful for diagnostics, reporting, or remote asset inspection.

#### Features

- Enumerates all active user sessions on the local system.
- Retrieves CPU usage, total/free memory, OS version, and uptime.
- Extracts logon time and session states per user.
- Supports running locally and remotely (if PowerShell remoting is enabled).
- Fully compatible with domain-joined and standalone environments.

#### Sample Output Fields

- ComputerName
- UserName
- SessionID
- State
- LogonTime
- OSVersion
- Uptime
- CPUUsage
- MemoryFreeMB
- TotalRAM_MB

#### Usage Examples

    .\Get-Resources.ps1
    Get-Content ".\computers.txt" | ForEach-Object { .\Get-Resources.ps1 -ComputerName $_ }
    .\Get-Resources.ps1 | Export-Csv ".\resources_report.csv" -NoTypeInformation"

---

### Get-LicensesByService.ps1

Lists all Microsoft 365 license SKUs that include a specific service (e.g., Exchange, Teams, etc.).

#### Features

- Query Microsoft Graph to find which licenses include a specified service.
- Displays friendly names for both licenses and service plans.
- Helps understand licensing dependencies and service availability across SKUs.
- Requires Microsoft Graph PowerShell SDK (`Microsoft.Graph.Users`, `Microsoft.Graph.Identity.DirectoryManagement`).

#### Usage Examples

    .\Get-LicensesByService.ps1 -ServiceName "EXCHANGE_S_FOUNDATION"
    .\Get-LicensesByService.ps1 -ServiceName "TEAMS1"
    .\Get-LicensesByService.ps1 -ServiceName "PowerBI"

#### Output Fields

- SkuDisplayName
- SkuPartNumber
- ServicePlanName
- ServicePlanId
- ServiceStatus

---

### Get=ServicesByLicense.ps1

Lists all service plans included in a given Microsoft 365 license SKU.

#### Features

- Displays internal and friendly names of services included in a license.
- Allows querying by SKU name or SKU part number (e.g., ENTERPRISEPACK, M365_BUSINESS_PREMIUM).
- Useful for comparing license types or validating service availability.

#### Usage Examples

    .\Get-ServicesByLicense.ps1 -License "ENTERPRISEPACK"
    .\Get-ServicesByLicense.ps1 -License "M365_BUSINESS_PREMIUM"
    .\Get-ServicesByLicense.ps1 -License "EMS" | Export-Csv ".\EMS_services.csv" -NoTypeInformation

#### Output Fields

- ServicePlanName
- ServicePlanId
- ServiceStatus
- LicenseDisplayName
- LicensePartNumber

---

## Requirements

- PowerShell 5.1 or later

---

## License

Distributed under the GPL License. See LICENSE file for details.

---

## Author

Kenlar Taj  
GitHub: https://github.com/KenlarTaj
