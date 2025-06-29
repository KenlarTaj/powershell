<#
.SYNOPSIS
    Retrieves information about installed MSI and AppX applications or displays properties of a specified MSI file.

.DESCRIPTION
    This script allows you to:
      - Search for installed applications (MSI/AppX) by name or GUID.
      - List properties of a given MSI file.
      - Filter results by installation context (system or user).
      - Optionally export results to a CSV file.

    It is compatible with PowerShell 5.1+ and Windows 10/11.

    **Running as Administrator vs. Standard User:**
      - When running as a standard user, the script can only access user-installed applications (HKCU) and AppX packages for the current user.
      - When running as administrator, the script can access both system-installed applications (HKLM) and AppX packages for all users, providing a more complete inventory.

.PARAMETER SearchTerm
    The application name, GUID, or MSI file path to search for.

.PARAMETER f
    Switch. Indicates that the SearchTerm is a path to an MSI file. The script will display its internal properties.

.PARAMETER s
    Switch. Search only for system-installed applications (HKLM).

.PARAMETER u
    Switch. Search only for user-installed applications (HKCU).

.PARAMETER ExportCsv
    String. If specified, saves the results to the given CSV file path.

.EXAMPLE
    .\Get-MSIInfo.ps1 "Office"
    Lists all installed applications with "Office" in the name.

.EXAMPLE
    .\Get-MSIInfo.ps1 -f "C:\Installers\app.msi"
    Displays properties of the specified MSI file.

.EXAMPLE
    .\Get-MSIInfo.ps1 -s "Adobe"
    Lists all system-installed applications with "Adobe" in the name.

.EXAMPLE
    .\Get-MSIInfo.ps1 -u
    Lists all user-installed applications.

.EXAMPLE
    .\Get-MSIInfo.ps1 -s "Office" -ExportCsv "C:\Temp\office_apps.csv"
    Lists all system-installed applications with "Office" in the name and exports the results to a CSV file.

.INPUTS
    [String] The SearchTerm can be a string (name, GUID) or a file path.

.OUTPUTS
    [PSCustomObject] Returns a list of application or MSI property objects.

.NOTES
    Author      : Kenlar Taj
    Version     : 1.0
    License     : GNU General Public License v3.0
    Date        : 2025-06-29
    Tested on   : Windows 11 Pro 26100.4349 with PowerShell 5.1.26100.4202 and PowerShell 7.5.2
    Repository  : https://github.com/KenlarTaj/powershel
    Contact     : kenlartaj@gmail.com
    Requirements: Run as user with access to registry and AppX packages.
    Limitations : AppX search may require administrative privileges for all users.
#>

param (
    [Parameter(Position = 0)]
    [string]$SearchTerm,
    [switch]$f,
    [switch]$s,
    [switch]$u,
    [Parameter(Position = 1)]
    [string]$ExportCsv
)

# Searches the registry and AppX packages for installed applications matching the search term.
function Get-InstalledAppByGuidOrName {
    param (
        [string]$SearchTerm,
        [hashtable[]]$Sources
    )

    $results = @()

    foreach ($source in $Sources) {
        $keyPath = $source.Path
        $context = $source.Context
        # Enumerate registry keys for installed applications
        Get-ChildItem $keyPath -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                $item = Get-ItemProperty $_.PsPath -ErrorAction Stop
                $match = $true

                # Filter by GUID or DisplayName if a search term is provided
                if ($SearchTerm) {
                    $match = ($item.PSChildName -like "*$SearchTerm*" -or $item.DisplayName -like "*$SearchTerm*")
                }

                if ($match) {
                    $results += [PSCustomObject]@{
                        PSChildName          = $item.PSChildName
                        PackageFullName      = $null
                        DisplayName          = $item.DisplayName
                        DisplayVersion       = $item.DisplayVersion
                        Publisher            = $item.Publisher
                        InstallDate          = $item.InstallDate
                        InstallLocation      = $item.InstallLocation
                        UninstallString      = $item.UninstallString
                        QuietUninstallString = $item.QuietUninstallString
                        Source               = $_.PsPath
                        InstallContext       = $context
                    }
                }
            } catch {
                # Ignore registry keys that cannot be read
                Write-Error "Error: $_"
            }
        }
    }

    # Search for AppX packages if a search term is provided
    if ($SearchTerm) {
        try {
            $appxList = Get-AppXPackage "*$SearchTerm*" -AllUsers
        } catch {
            $appxList = Get-AppXPackage "*$SearchTerm*"
        }
    } else {
        try {
            $appxList = Get-AppXPackage -AllUsers
        } catch {
            $appxList = Get-AppXPackage
        }
    }

    # Collect AppX package information
    $appxList | ForEach-Object {
        $row = [ordered]@{
            GUID                 = $null
            PackageFullName      = $_.PackageFullName
            DisplayName          = $_.Name
            DisplayVersion       = $_.Version
            Publisher            = $_.Publisher
            InstallDate          = if (Test-Path $_.InstallLocation -ErrorAction SilentlyContinue) { (Get-Item $_.InstallLocation).CreationTime.ToString("yyyyMMdd") } else { "N/A" }
            InstallLocation      = if (Test-Path $_.InstallLocation -ErrorAction SilentlyContinue) { $_.InstallLocation } else { "N/A" }
            UninstallString      = "Remove-AppxPackage -Package $($_.PackageFullName)"
            QuietUninstallString = "N/A"
            Source               = "AppX Package"
            InstallContext       = $null
        }

        # Add user information if available
        if ($_.PackageUserInformation) {
            $i = 1
            foreach ($userInfo in $_.PackageUserInformation) {
                try {
                    $sid = $userInfo.UserSecurityId.Sid
                    $userName = (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount]).Value
                } catch {
                    $userName = $sid
                }
                $row["User$i"] = $userName
                $i++
            }
        }
        $results += New-Object PSObject -Property $row
    }
    return $results | Sort-Object InstallContext, DisplayName
}

# Reads and displays properties from a specified MSI file.
function Get-MsiProperty {
    param (
        [string]$MsiPath
    )

    $currentDir = Get-Location

    # Resolve relative path if needed
    if ($MsiPath -like '.\*') {
        $MsiPath = Join-Path -Path $currentDir -ChildPath $MsiPath.Substring(2)
    }

    # Check if the file exists before proceeding
    if (-not (Test-Path "$MsiPath")) {
        throw "File not found: $MsiPath"
    }

    try {
        # Create Windows Installer COM object and open the MSI database in read-only mode
        $installer = New-Object -ComObject WindowsInstaller.Installer -ErrorAction Stop
        $database = $installer.OpenDatabase($MsiPath, 0)

        # Query the Property table for all properties and values
        $view = $database.OpenView("SELECT ``Property``, ``Value`` FROM ``Property``")
        $view.Execute()

        $results = @()
        while ($record = $view.Fetch()) {
            $results += [PSCustomObject]@{
                Property = $record.StringData(1)
                Value    = $record.StringData(2)
            }
        }
        return $results
    } catch {
        throw "Error reading MSI file: $_"
    }
}

# Main function: decides which operation to perform based on parameters
function Main {
    param (
        [string]$SearchTerm,
        [switch]$f,
        [switch]$s,
        [switch]$u,
        [string]$ExportCsv
    )
    # If -f is specified, show MSI file properties
    if ($f) {
        if ($SearchTerm) {
            $results = Get-MsiProperty -MsiPath $SearchTerm
        } else {
            throw "You must provide an MSI path with the -f parameter."
        }
    } else {
        # Determine registry sources based on -s and -u switches
        if ($s -and -not($u)) {
            $registrySources = @(
                @{ Path = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"; Context = "System" },
                @{ Path = "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"; Context = "System" }
            )
        } elseif ($u -and -not($s)) {
            $registrySources = @(
                @{ Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"; Context = "User" },
                @{ Path = "HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"; Context = "User" }
            )
        } else {
            $registrySources = @(
                @{ Path = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"; Context = "System" },
                @{ Path = "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"; Context = "System" },
                @{ Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"; Context = "User" },
                @{ Path = "HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"; Context = "User" }
            )
        }
        # Search and display installed applications
        $results = Get-InstalledAppByGuidOrName -SearchTerm $SearchTerm -Sources $registrySources
        if ($results.Count -eq 0) {
            Write-Warning "No entries found for '$SearchTerm'."
        } else {
            if ($ExportCsv) {
                if ($ExportCsv -is [string] -and $ExportCsv.Trim() -ne "") {
                    $csvPath = $ExportCsv
                } else {
                    $csvPath = ".\MSIInfoResults.csv" }
                try {
                    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                    Write-Host "`nResults exported to $csvPath" -ForegroundColor Cyan
                } catch {
                    Write-Error "Failed to export results to CSV: $_"
                }
            }
            return $results
        }
    }
}

# Script entry point
try {
    Main -SearchTerm $SearchTerm -f:$f -s:$s -u:$u -ExportCsv:$ExportCsv
} catch {
    Write-Error $_
}
