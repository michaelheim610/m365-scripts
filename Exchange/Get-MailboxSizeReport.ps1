#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Erstellt eine Größenübersicht aller Postfächer in Exchange Online.

.DESCRIPTION
    Liest alle Postfächer der gewählten Typen aus, ermittelt Größe und Anzahl
    der Elemente und gibt das Ergebnis sortiert nach Größe aus. Am Ende steht
    eine Gesamtsumme.

    Das Skript liest nur. Es ändert nichts im Tenant.

.PARAMETER RecipientTypeDetails
    Postfachtypen, die ausgewertet werden. Standard sind Benutzerpostfächer
    und freigegebene Postfächer.

.PARAMETER Top
    Gibt nur die größten n Postfächer aus. 0 bedeutet alle.

.PARAMETER ExportPath
    Optionaler Pfad für einen CSV-Export, zum Beispiel .\Postfaecher.csv

.EXAMPLE
    .\Get-MailboxSizeReport.ps1

    Zeigt alle Benutzer- und freigegebenen Postfächer mit Größe an.

.EXAMPLE
    .\Get-MailboxSizeReport.ps1 -Top 20 -ExportPath .\Top20.csv

    Zeigt die 20 größten Postfächer und speichert sie als CSV.

.NOTES
    Autor          Michael Heim, https://michaelheim.one
    Voraussetzung  Modul ExchangeOnlineManagement
    Berechtigung   Eine Rolle mit Leserechten auf Empfänger, z.B. View-Only Recipients
#>
[CmdletBinding()]
param(
    [string[]]$RecipientTypeDetails = @('UserMailbox', 'SharedMailbox'),

    [ValidateRange(0, [int]::MaxValue)]
    [int]$Top = 0,

    [string]$ExportPath
)

# Wandelt die Größenangabe von Exchange Online in GB um.
# Exchange liefert einen Text wie "1.2 GB (1,288,490,188 bytes)".
# Die Bytes in der Klammer sind unabhängig von Einheit und Sprache.
function ConvertTo-GigaByte {
    param([object]$Size)

    if ("$Size" -match '\(([\d.,\s]+)bytes\)') {
        $bytes = [double]($Matches[1] -replace '[^\d]', '')
        return [math]::Round($bytes / 1GB, 2)
    }
    return 0
}

# Nur verbinden, wenn noch keine Sitzung besteht.
if (-not (Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
}

Write-Verbose 'Lese Postfächer aus ...'
$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $RecipientTypeDetails

$report = foreach ($mailbox in $mailboxes) {
    $statistics = Get-MailboxStatistics -Identity $mailbox.ExchangeGuid.ToString()

    [PSCustomObject]@{
        DisplayName        = $mailbox.DisplayName
        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
        Typ                = $mailbox.RecipientTypeDetails
        Elemente           = $statistics.ItemCount
        GroesseGB          = ConvertTo-GigaByte -Size $statistics.TotalItemSize
        LetzteAnmeldung    = $statistics.LastLogonTime
    }
}

$report = $report | Sort-Object GroesseGB -Descending
if ($Top -gt 0) {
    $report = $report | Select-Object -First $Top
}

if ($ExportPath) {
    $report | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "Export gespeichert unter $ExportPath"
}

$totalItems = ($report | Measure-Object -Property Elemente -Sum).Sum
$totalSize  = ($report | Measure-Object -Property GroesseGB -Sum).Sum
Write-Host ("{0} Postfächer, {1:N0} Elemente, {2:N2} GB gesamt" -f @($report).Count, $totalItems, $totalSize)

$report
