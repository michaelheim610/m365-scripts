#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Zeigt alle Lizenzen im Tenant mit lesbarem Produktnamen, Verbrauch und freiem Bestand.

.DESCRIPTION
    Microsoft Graph liefert Lizenzen nur mit technischem Namen, zum Beispiel
    SPE_E3. Das Skript lädt die offizielle Zuordnungsliste von Microsoft und
    ersetzt den technischen Namen durch den Produktnamen, zum Beispiel
    Microsoft 365 E3. Dazu kommen gekaufte, verwendete und freie Lizenzen.

    Kann die Zuordnungsliste nicht geladen werden, läuft das Skript weiter und
    zeigt nur den technischen Namen.

    Das Skript liest nur. Es ändert nichts im Tenant.

.PARAMETER OnlyWithFreeUnits
    Gibt nur Lizenzen aus, bei denen noch Einheiten frei sind.

.PARAMETER ExportPath
    Optionaler Pfad für einen CSV-Export.

.EXAMPLE
    .\Get-LicenseReport.ps1

    Zeigt alle Lizenzen des Tenants.

.EXAMPLE
    .\Get-LicenseReport.ps1 -OnlyWithFreeUnits -ExportPath .\FreieLizenzen.csv

    Zeigt nur Lizenzen mit freiem Bestand und speichert sie als CSV.

.NOTES
    Autor          Michael Heim, https://michaelheim.one
    Voraussetzung  Modul Microsoft.Graph.Authentication
    Berechtigung   Organization.Read.All
#>
[CmdletBinding()]
param(
    [switch]$OnlyWithFreeUnits,

    [string]$ExportPath
)

# Offizielle Liste "Product names and service plan identifiers for licensing".
$mappingUrl = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'

if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes 'Organization.Read.All' -NoWelcome
}

# Lädt die Zuordnung von SKU-ID zu Produktname als Hashtable.
function Get-ProductNameMap {
    param([Parameter(Mandatory)][string]$Url)

    $map = @{}
    $tempFile = [System.IO.Path]::GetTempFileName()
    try {
        Invoke-WebRequest -Uri $Url -OutFile $tempFile -UseBasicParsing
        foreach ($row in Import-Csv -Path $tempFile) {
            if ($row.GUID -and -not $map.ContainsKey($row.GUID)) {
                $map[$row.GUID] = $row.Product_Display_Name
            }
        }
    }
    catch {
        Write-Warning "Die Produktnamen konnten nicht geladen werden. Es wird der technische Name angezeigt. ($_)"
    }
    finally {
        Remove-Item -Path $tempFile -ErrorAction SilentlyContinue
    }
    return $map
}

$productNames = Get-ProductNameMap -Url $mappingUrl
$skus = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus').value

$report = foreach ($sku in $skus) {
    $productName = $productNames[$sku.skuId]
    if (-not $productName) { $productName = $sku.skuPartNumber }

    [PSCustomObject]@{
        Produkt       = $productName
        SkuPartNumber = $sku.skuPartNumber
        Status        = $sku.capabilityStatus
        Gesamt        = $sku.prepaidUnits.enabled
        Verwendet     = $sku.consumedUnits
        Frei          = $sku.prepaidUnits.enabled - $sku.consumedUnits
    }
}

if ($OnlyWithFreeUnits) {
    $report = $report | Where-Object { $_.Frei -gt 0 }
}

$report = $report | Sort-Object Produkt

if ($ExportPath) {
    $report | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "Export gespeichert unter $ExportPath"
}

$report
