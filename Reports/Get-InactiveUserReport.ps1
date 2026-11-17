#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Findet lizenzierte Benutzer ohne Aktivität in Exchange, Teams, OneDrive und SharePoint.

.DESCRIPTION
    Lädt den Microsoft 365 Aktivitätsbericht über Microsoft Graph und markiert
    alle Konten, die im gewählten Zeitraum in keinem der vier Dienste aktiv
    waren. Das ist ein guter Startpunkt für Lizenzbereinigung und für die Frage,
    welche Konten noch gebraucht werden.

    Wenn im Admin Center die Anonymisierung der Berichte aktiv ist, stehen statt
    Benutzernamen nur Kennungen im Ergebnis. Die Einstellung findest du unter
    Einstellungen, Einstellungen der Organisation, Berichte.

    Das Skript liest nur. Es ändert nichts im Tenant.

.PARAMETER Period
    Auswertungszeitraum des Berichts. Möglich sind D7, D30, D90 und D180.

.PARAMETER OnlyInactive
    Gibt nur die inaktiven Konten aus.

.PARAMETER ExportPath
    Optionaler Pfad für einen CSV-Export.

.EXAMPLE
    .\Get-InactiveUserReport.ps1 -Period D90 -OnlyInactive

    Zeigt alle Konten, die in den letzten 90 Tagen in keinem Dienst aktiv waren.

.NOTES
    Autor          Michael Heim, https://michaelheim.one
    Voraussetzung  Modul Microsoft.Graph.Authentication
    Berechtigung   Reports.Read.All
#>
[CmdletBinding()]
param(
    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$Period = 'D30',

    [switch]$OnlyInactive,

    [string]$ExportPath
)

if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes 'Reports.Read.All' -NoWelcome
}

$days = [int]$Period.Substring(1)
$cutoffDate = (Get-Date).Date.AddDays(-$days)

# Graph liefert den Bericht als CSV-Datei. Sie wird in eine temporäre Datei
# geladen, eingelesen und danach wieder gelöscht.
$tempFile = [System.IO.Path]::GetTempFileName()
try {
    $uri = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='$Period')"
    Invoke-MgGraphRequest -Method GET -Uri $uri -OutputFilePath $tempFile
    $activity = Import-Csv -Path $tempFile
}
finally {
    Remove-Item -Path $tempFile -ErrorAction SilentlyContinue
}

# Ein leeres Datum bedeutet, dass es im Dienst noch nie Aktivität gab.
function Test-ActiveSince {
    param([string]$LastActivity, [datetime]$Since)

    $date = $LastActivity -as [datetime]
    return ($null -ne $date -and $date -ge $Since)
}

$report = foreach ($user in $activity) {
    # Gelöschte Konten stehen noch im Bericht, sind aber nicht mehr relevant.
    if ($user.'Is Deleted' -eq 'True') { continue }

    $activeAnywhere =
        (Test-ActiveSince $user.'Exchange Last Activity Date' $cutoffDate) -or
        (Test-ActiveSince $user.'Teams Last Activity Date' $cutoffDate) -or
        (Test-ActiveSince $user.'OneDrive Last Activity Date' $cutoffDate) -or
        (Test-ActiveSince $user.'SharePoint Last Activity Date' $cutoffDate)

    [PSCustomObject]@{
        Benutzer   = $user.'User Principal Name'
        Lizenzen   = $user.'Assigned Products'
        Exchange   = $user.'Exchange Last Activity Date'
        Teams      = $user.'Teams Last Activity Date'
        OneDrive   = $user.'OneDrive Last Activity Date'
        SharePoint = $user.'SharePoint Last Activity Date'
        Inaktiv    = -not $activeAnywhere
    }
}

if ($OnlyInactive) {
    $report = $report | Where-Object Inaktiv
}

$report = $report | Sort-Object Benutzer

if ($ExportPath) {
    $report | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "Export gespeichert unter $ExportPath"
}

$inactiveCount = @($report | Where-Object Inaktiv).Count
Write-Host "$inactiveCount Konten ohne Aktivität in den letzten $days Tagen."

$report
