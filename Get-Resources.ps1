<#
.SYNOPSIS
    Retrieves and displays detailed system information for the local Windows computer.

.DESCRIPTION
    This script collects and displays key system information, including:
      - Computer name, OS version, build, and installation date
      - Last boot time, current local time, and user session info
      - CPU and memory usage
      - Disk usage for all local drives

    Output is formatted for easy reading in the console. Optionally, results can be exported to a CSV file.

.PARAMETER ExportCsv
    String. If specified, saves the results to the given CSV file path.

.EXAMPLE
    .\Get-Resources.ps1
    Displays system information in the console.

.EXAMPLE
    .\Get-Resources.ps1 -ExportCsv "C:\Temp\systeminfo.csv"
    Exports system information to the specified CSV file.

.OUTPUTS
    [PSCustomObject] Returns a list of system information objects.

.NOTES
    Author      : Kenlar Taj
    Version     : 1.0
    License     : GNU General Public License v3.0
    Date        : 2025-06-29
    Tested on   : Windows 11 Pro 26100.4349 with PowerShell 5.1.26100.4202 and PowerShell 7.5.2
    Repository  : https://github.com/KenlarTaj/powershel
    Contact     : kenlartaj@email.com
    Requirements: PowerShell 5.1+ on Windows 10/11
#>

param (
    [Parameter(Position = 0)]
    [string]$ExportCsv
)

function Format-DateWithTimeZone {
    param([datetime]$dt)
    $tz = [System.TimeZoneInfo]::Local
    $offset = $tz.GetUtcOffset($dt)
    $sign = if ($offset.TotalMinutes -ge 0) { "+" } else { "-" }
    $offsetString = "{0}{1:00}:{2:00}" -f $sign, [math]::Abs($offset.Hours), [math]::Abs($offset.Minutes)
    return "$dt (UTC$offsetString)"
}

function Convert-WmiDateWithTimeZone {
    param([string]$wmiDate)
    $dt = [Management.ManagementDateTimeConverter]::ToDateTime($wmiDate)
    return Format-DateWithTimeZone $dt
}

function Get-SystemInfo {
    $osInfo = $null
    try {
        $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    } catch {
        Write-Warning "Failed to retrieve OS Info using CIM method: $($_.Exception.Message)"
    }
    if ($null -eq $osInfo) {
        try {
            $osInfo = Get-WmiObject Win32_OperatingSystem -ErrorAction Stop
        } catch {
            throw "Failed to retrieve OS Info using WMI method: $($_.Exception.Message)"
        }
    }
    $compName = $env:COMPUTERNAME
    $caption = $osInfo.Caption
    $localTime = Get-Date

    $reg = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    $buildNumber = $reg.CurrentBuildNumber
    $ubr = $reg.UBR
    $fullBuild = "$buildNumber.$ubr"

    $buildMap = @{
        '19041' = '2004' #Win10
        '19042' = '20H2'
        '19043' = '21H1'
        '19044' = '21H2'
        '19045' = '22H2'
        '22000' = '21H2' #Win11
        '22621' = '22H2'
        '22631' = '23H2'
        '26100' = '24H2'
    }
    $buildName = $buildMap[$buildNumber]
    if (-not $buildName) { $buildName = "unknown build" }

    if ($osInfo.InstallDate -is [datetime]) {
        $installDateFormatted    = Format-DateWithTimeZone $osInfo.InstallDate
        $lastBootUpTimeFormatted = Format-DateWithTimeZone $osInfo.LastBootUpTime
    } else {
        $installDateFormatted    = Convert-WmiDateWithTimeZone $osInfo.InstallDate
        $lastBootUpTimeFormatted = Convert-WmiDateWithTimeZone $osInfo.LastBootUpTime
    }
    $localTimeFormatted = Format-DateWithTimeZone $localTime

    $totalMem = [math]::Round($osInfo.TotalVisibleMemorySize/1MB,2)
    $freeMem  = [math]::Round($osInfo.FreePhysicalMemory/1MB,2)
    $usedMem  = [math]::Round($totalMem - $freeMem,2)

    $cpuUsage = $null
    try {
        $cimData = Get-CimInstance Win32_Processor -ErrorAction Stop
        $cpuUsage = [math]::Round(($cimData | Measure-Object -Property LoadPercentage -Average).Average, 2)
    }
    catch {
        Write-Warning "Failed to retrieve CPU usage using CIM method: $($_.Exception.Message)"
    }

    if ($null -eq $cpuUsage) {
        try {
            $wmiData = Get-WmiObject Win32_Processor -ErrorAction Stop
            $cpuUsage = [math]::Round(($wmiData | Measure-Object -Property LoadPercentage -Average).Average, 2)
        }
        catch {
            Write-Warning "Failed to retrieve CPU usage using WMI method: $($_.Exception.Message)"
            $cpuUsage = $null
        }
    }

    try {
        $userInfo = quser 2>&1 | ForEach-Object { $_ -replace '\s{2,}', ',' } | ConvertFrom-Csv
        $lastLogin = $userInfo.'LOGON TIME' -join '; '
        $userName = $userInfo.'USERNAME' -join '; '
        $sessionId = $userInfo.'ID' -join '; '
    } catch {
        $lastLogin = $null
        $userName = $null
        $sessionId = $null
    }

    $sysInfo = [ordered]@{
        ComputerName        = $compName
        OS                  = "$caption ($buildName)"
        Build               = $fullBuild
        LocalTime           = $localTimeFormatted
        InstallDate         = $installDateFormatted
        LastBootUpTime      = $lastBootUpTimeFormatted
        LastLogin           = $lastLogin
        UserName            = $userName
        SessionID           = $sessionId
        CPUUsagePercent     = $cpuUsage
        TotalMemoryGB       = $totalMem
        FreeMemoryGB        = $freeMem
        UsedMemoryGB        = $usedMem
    }
    $i = 1

    $disks = $null
    try {
        $disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
    } catch {
        Write-Warning "Failed to retrieve disk information using CIM method: $($_.Exception.Message)"
    }
    if ($null -eq $disks) {
        try {
            $disks = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
        } catch {
            Write-Warning "Failed to retrieve disk information using WMI method: $($_.Exception.Message)"
        }
    }

    foreach ($disk in $disks) {
        $sysInfo["Vol${i}_DeviceID"]   = $disk.DeviceID
        $sysInfo["Vol${i}_VolumeName"] = $disk.VolumeName
        $sysInfo["Vol${i}_SizeGB"]     = "{0:N1}" -f ($disk.Size/1GB)
        $sysInfo["Vol${i}_FreeGB"]     = "{0:N1}" -f ($disk.FreeSpace/1GB)
        $sysInfo["Vol${i}_UsedGB"]     = "{0:N1}" -f ( ($disk.Size - $disk.FreeSpace)/1GB )
        $i++
    }

    return  [PSCustomObject]$sysInfo
}

function Main {
    param([string]$ExportCsv)
    $results = Get-SystemInfo

    if ($ExportCsv) {
        $csvPath = if ($ExportCsv.Trim()) { $ExportCsv } else { ".\SystemInfoResults.csv" }
        try {
            $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "`nResults exported to $csvPath" -ForegroundColor Cyan
        } catch {
            Write-Error "Failed to export results to CSV: $_"
        }
    }
    return $results
}

try {
    Main -ExportCsv:$ExportCsv
} catch {
    Write-Error $_
}