# PowerShell

PowerShell scripts for Windows and Microsoft 365 management.

## Overview

This repository contains PowerShell scripts designed to help automate and manage Windows environments and Microsoft 365 resources. All scripts are written to be clear, reusable, and follow best practices for PowerShell scripting.

## Scripts

- Get-MSIInfo.ps1                  - Retrieve installed MSI/AppX apps or MSI file properties.
- Get-Resources.ps1                - Gather system information, performance, and session/user data.
- GetLicensesByService.ps1         - List Microsoft 365 licenses that include a specific service.
- GetServicesByLicense.ps1         - Display services included in a specific Microsoft 365 license SKU.
- Register-PnPAppAndCredential.ps1 - Register an Azure AD app and store credentials for PnP.PowerShell automation.

---

### Get-MSIInfo.ps1

Retrieves information about installed MSI and AppX applications or displays properties of a specified MSI file.

#### Features

- Search for installed applications (MSI/AppX) by name or GUID.
- List properties of a given MSI file.
- Filter results by installation context (system or user).
- Optionally export results to a CSV file.
- Compatible with PowerShell 5.1+ and Windows 10/11.

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

#### Usage Examples

    .\Get-Resources.ps1
    Get-Content ".\computers.txt" | ForEach-Object { .\Get-Resources.ps1 -ComputerName $_ }
    .\Get-Resources.ps1 | Export-Csv ".\resources_report.csv" -NoTypeInformation

---

### GetLicensesByService.ps1

Lists all Microsoft 365 license SKUs that include a specific service (e.g., Exchange, Teams, etc.).

#### Features

- Query Microsoft Graph to find which licenses include a specified service.
- Displays friendly names for both licenses and service plans.
- Helps understand licensing dependencies and service availability across SKUs.
- Requires Microsoft Graph PowerShell SDK.

#### Usage Examples

    .\GetLicensesByService.ps1 -ServiceName "EXCHANGE_S_FOUNDATION"
    .\GetLicensesByService.ps1 -ServiceName "TEAMS1"
    .\GetLicensesByService.ps1 -ServiceName "PowerBI"

---

### GetServicesByLicense.ps1

Lists all service plans included in a given Microsoft 365 license SKU.

#### Features

- Displays internal and friendly names of services included in a license.
- Allows querying by SKU name or part number.
- Useful for comparing license types or validating service availability.

#### Usage Examples

    .\GetServicesByLicense.ps1 -License "ENTERPRISEPACK"
    .\GetServicesByLicense.ps1 -License "M365_BUSINESS_PREMIUM"
    .\GetServicesByLicense.ps1 -License "EMS" | Export-Csv ".\EMS_services.csv" -NoTypeInformation

---

### Register-PnPAppAndCredential.ps1

Registers an Azure AD application and stores its client secret securely in the Windows Credential Manager for use with PnP.PowerShell automation.

#### Features

- Automates Azure AD app registration using Microsoft Graph.
- Creates and stores a client secret securely for later authentication.
- Supports force-replacing an existing app with the same name.
- Verifies successful creation and storage of credentials.
- Great for automating unattended SharePoint/PnP scripts.

#### Requirements

- PowerShell 7.4 or later (due to PnP.PowerShell compatibility).
- Required modules: `Microsoft.Graph.Applications`, `Microsoft.Graph.Identity.DirectoryManagement`, `Microsoft.PowerShell.SecretManagement` (if storing credentials securely).
- Permissions to register apps in Azure AD.

#### Usage Examples

    # Register default app in contoso tenant
    .\Register-PnPAppAndCredential.ps1 -TenantName "contoso"

    # Register app with custom name and force replacement if it exists
    .\Register-PnPAppAndCredential.ps1 -AppName "MyPnPApp" -TenantName "contoso" -ForceReplace

---

## Requirements

- PowerShell 5.1 or later (some scripts require 7.4+)
- Microsoft Graph PowerShell SDK modules:
  - Microsoft.Graph.Users
  - Microsoft.Graph.Identity.DirectoryManagement
  - Microsoft.Graph.Applications
- For remote system queries:
  - WinRM enabled
  - Local admin permissions may be required
- For app registration:
  - Azure AD permissions to register applications

---

## License

Distributed under the GPL License. See LICENSE file for details.

---

## Author

Author :Kenlar Taj  
GitHub : https://github.com/KenlarTaj
E-mail : kenlartaj@gmail.com