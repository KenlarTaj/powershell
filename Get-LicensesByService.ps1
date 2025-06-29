<#
.SYNOPSIS
    Compares Microsoft 365/Azure licenses by included services, based on a search term.

.DESCRIPTION
    Downloads the official Microsoft CSV reference for license service plans, filters by a search term (e.g., "Defender"),
    and builds a comparative matrix showing which licenses include the searched services.
    The result is displayed in the console and exported to a CSV file.

.PARAMETER SearchTerm
    The term to search for in the included service plans (e.g., "Defender").

.PARAMETER ExportCsv
    String. If specified, saves the results to the given CSV file path. If used without a value, defaults to '.\comparativo-byservice.csv'.

.EXAMPLE
    .\GetLicensesByService.ps1 Defender
    Shows a comparative table of licenses that include services matching "Defender".

.EXAMPLE
    .\GetLicensesByService.ps1 Teams -ExportCsv "C:\Temp\teams_licenses.csv"
    Exports the comparison for "Teams" services to a specific CSV file.

.NOTES
    Author      : Kenlar Taj
    Version     : 1.0
    License     : GNU General Public License v3.0
    Date        : 2025-06-29
    Tested on   : Windows 11 Pro 26100.4349 with PowerShell 5.1.26100.4202 and PowerShell 7.5.2
    Repository  : https://github.com/KenlarTaj/powershel
    Contact     : kenlartaj@gmail.com
    Requirements: PowerShell 5.1+ on Windows 10/11
#>

param (
    [Parameter(Position = 0, Mandatory = $true)]
    [string]$SearchTerm,
    [Parameter(Position = 1)]
    [string]$ExportCsv
)

function Get-LicenseComparisonByService {
    param (
        [string]$Term
    )

    $csvUrl = "https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference"
    try {
        $response = Invoke-WebRequest -Uri $csvUrl -UseBasicParsing -ErrorAction Stop
        $csvLink = ($response.Links | Where-Object { $_.href -like "*.csv" }).href
    } catch {
        throw "Failed to retrieve CSV link from Microsoft Docs: $($_.Exception.Message)"
    }

    $tempCsvPath = "$env:TEMP\licensing-service-plan-reference.csv"
    try {
        Invoke-WebRequest -Uri $csvLink -OutFile $tempCsvPath -ErrorAction Stop
    } catch {
        throw "Failed to download the CSV file: $($_.Exception.Message)"
    }

    try {
        $services = Import-Csv -Path $tempCsvPath -ErrorAction Stop
    } catch {
        throw "Failed to import the CSV file: $($_.Exception.Message)"
    }

    $matchingRows = $services | Where-Object {
        $_.Service_Plans_Included_Friendly_Names -match "(?i)$Term"
    }

    if (-not $matchingRows) {
        Write-Warning "No services found matching '$Term'."
        return $null
    }

    $matchedServices = $matchingRows.Service_Plans_Included_Friendly_Names | Sort-Object -Unique
    $licenses = $matchingRows.Product_Display_Name | Sort-Object -Unique

    $licenseServiceMap = @{}
    foreach ($row in $matchingRows) {
        if (-not $licenseServiceMap.ContainsKey($row.Product_Display_Name)) {
            $licenseServiceMap[$row.Product_Display_Name] = @()
        }
        $licenseServiceMap[$row.Product_Display_Name] += $row.Service_Plans_Included_Friendly_Names
    }

    $comparisonScreen = foreach ($license in $licenses) {
        $row = [ordered]@{ License = $license }
        foreach ($service in $matchedServices) {
            $hasService = $licenseServiceMap[$license] -contains $service
            $row[$service] = if ($hasService) { '✅' } else { '❌' }
        }
        [PSCustomObject]$row
    }

    $comparison = foreach ($license in $licenses) {
        $row = [ordered]@{ License = $license }
        foreach ($service in $matchedServices) {
            $hasService = $licenseServiceMap[$license] -contains $service
            $row[$service] = if ($hasService) { '1' } else { '0' }
        }
        [PSCustomObject]$row
    }

    return @{
        Screen = $comparisonScreen
        Csv    = $comparison
    }
}

function Main {
    param (
        [string]$SearchTerm,
        [string]$ExportCsv
    )

    $result = Get-LicenseComparisonByService -Term $SearchTerm
    if ($null -eq $result) { return }

    $comparisonScreen = $result.Screen
    $comparison = $result.Csv

    if ($ExportCsv) {
        $csvPath = if ($ExportCsv.Trim()) { $ExportCsv } else { ".\comparativo-byservice.csv" }
        try {
            $comparison | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "`nExported to: $csvPath" -ForegroundColor Cyan
        } catch {
            Write-Error "Failed to export results to CSV: $_"
        }
    }

    $comparisonScreen | Format-Table -AutoSize
}

try {
    Main -SearchTerm $SearchTerm -ExportCsv $ExportCsv
} catch {
    Write-Error $_
}