# PowerShell

PowerShell scripts for Windows and Microsoft 365 management.

## Overview

This repository contains PowerShell scripts designed to help automate and manage Windows environments and Microsoft 365 resources. All scripts are written to be clear, reusable, and follow best practices for PowerShell scripting.

## Scripts

- Get-MSIInfo.ps1 — Retrieve installed MSI/AppX apps or MSI file properties.
- Get-Resources.ps1 — Gather system information, performance, and session/user data.

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

- As a standard user, the script can only access user-installed applications (HKCU) and AppX packages for the current user.
- As administrator, the script can access both system-installed applications (HKLM) and AppX packages for all users, providing a more complete inventory.
- Some AppX queries (especially with -AllUsers) may require administrative privileges.

#### Usage Examples

    # List all installed applications with "Office" in the name
    .\Get-MSIInfo.ps1 "Office"

    # Display properties of a specified MSI file
    .\Get-MSIInfo.ps1 -f "C:\Installers\app.msi"

    # List all system-installed applications with "Adobe" in the name
    .\Get-MSIInfo.ps1 -s "Adobe"

    # List all user-installed applications
    .\Get-MSIInfo.ps1 -u

    # Export results to a specific CSV file
    .\Get-MSIInfo.ps1 notepad -ExportCsv ".\output.csv"

---

### Get-Resources.ps1

Gathers information about the system’s performance, session, hardware, and logged-in users. Useful for diagnostics, automation, or reporting across environments.

#### Features

- Collect detailed session info: user, logon time, session state.
- Retrieve hardware and OS data: CPU, RAM, OS version, uptime.
- Show resource usage: CPU load, available/free memory.
- Works on both remote and local machines (when WinRM is enabled).
- Output can be used for compliance reports or monitoring systems.

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
- TotalRAM

#### Usage Examples

    # Run locally to display resource info
    .\Get-Resources.ps1

    # Run for a list of computers from a file
    Get-Content ".\computers.txt" | ForEach-Object { .\Get-Resources.ps1 -ComputerName $_ }

    # Export output to CSV
    .\Get-Resources.ps1 | Export-Csv ".\resources_report.csv" -NoTypeInformation

---

## Requirements

- PowerShell 5.1 or later
- For Get-Resources.ps1, remote systems require:
  - WinRM enabled
  - Admin permissions if accessing sessions remotely

---

## License

Distributed under the GPL License. See LICENSE file for details.

---

## Author

Kenlar Taj
GitHub: https://github.com/KenlarTaj
