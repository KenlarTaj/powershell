<#
.SYNOPSIS
    Compares Microsoft 365/Azure services by license, based on one or more license search terms.

.DESCRIPTION
    Downloads the official Microsoft CSV reference for license service plans, filters by license name(s) (e.g., "Defender", "Intune"),
    and builds a comparative matrix showing which services are included in the searched licenses.
    The result is displayed in the console and can be exported to a CSV file.

.PARAMETER SearchTerm
    One or more terms to search for in the license names (e.g., "Defender", "Intune").

.PARAMETER ExportCsv
    String. If specified, saves the results to the given CSV file path. If used without a value, defaults to '.\comparativo-bylicense.csv'.

.EXAMPLE
    .\GetServicesByLicense.ps1 Defender Intune
    Shows a comparative table of services included in licenses matching "Defender" or "Intune".

.EXAMPLE
    .\GetServicesByLicense.ps1 E5 -ExportCsv "C:\Temp\e5_services.csv"
    Exports the comparison for "E5" licenses to a specific CSV file.

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

function Get-ServiceComparisonByLicense {
    param (
        [string[]]$Terms
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

    # Filtrar licenças com base nos termos (case-insensitive)
    $matchingRows = foreach ($term in $Terms) {
        $services | Where-Object { $_.Product_Display_Name -match "(?i)$term" }
    }

    # Remover duplicatas
    $matchingRows = $matchingRows | Sort-Object Product_Display_Name, Service_Plans_Included_Friendly_Names -Unique

    if (-not $matchingRows) {
        Write-Warning "No licenses found matching '$($Terms -join ', ')'."
        return $null
    }

    $licenses = $matchingRows.Product_Display_Name | Sort-Object -Unique
    $matchedServices = $matchingRows.Service_Plans_Included_Friendly_Names | Sort-Object -Unique

    # Criar mapa: licença → serviços
    $licenseServiceMap = @{}
    foreach ($row in $matchingRows) {
        if (-not $licenseServiceMap.ContainsKey($row.Product_Display_Name)) {
            $licenseServiceMap[$row.Product_Display_Name] = @()
        }
        $licenseServiceMap[$row.Product_Display_Name] += $row.Service_Plans_Included_Friendly_Names
    }

    # Construir tabela comparativa para tela
    $comparisonScreen = foreach ($service in $matchedServices) {
        $row = [ordered]@{ Service = $service }
        foreach ($license in $licenses) {
            $hasService = $licenseServiceMap[$license] -contains $service
            $row[$license] = if ($hasService) { '✅' } else { '❌' }
        }
        [PSCustomObject]$row
    }

    # Construir tabela comparativa para CSV
    $comparison = foreach ($service in $matchedServices) {
        $row = [ordered]@{ Service = $service }
        foreach ($license in $licenses) {
            $hasService = $licenseServiceMap[$license] -contains $service
            $row[$license] = if ($hasService) { '1' } else { '0' }
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
        [string[]]$SearchTerm,
        [string]$ExportCsv
    )

    $result = Get-ServiceComparisonByLicense -Terms $SearchTerm
    if ($null -eq $result) { return }

    $comparisonScreen = $result.Screen
    $comparison = $result.Csv

    if ($ExportCsv) {
        $csvPath = if ($ExportCsv.Trim()) { $ExportCsv } else { ".\comparativo-bylicense.csv" }
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