#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Erstellt einen HTML-Bericht über alle Conditional Access Richtlinien im Tenant.

.DESCRIPTION
    Der Bericht beantwortet drei Fragen, die in fast jedem Tenant-Review
    auftauchen.

    1. Welche Richtlinien gibt es, in welchem Status, für wen und mit welchen
       Kontrollen?
    2. Welche Named Locations sind definiert und welche gelten als vertrauenswürdig?
    3. Welche Konten sind aus Richtlinien ausgeschlossen, und sehen sie nach
       Break Glass Accounts aus?

    Benutzer, Gruppen und Rollen werden mit Namen statt mit ID angezeigt.

    Das Skript liest nur. Es ändert nichts im Tenant.

.PARAMETER OutputPath
    Pfad der HTML-Datei. Standard ist ConditionalAccess-Report.html im aktuellen Ordner.

.PARAMETER BreakGlassPattern
    Regulärer Ausdruck, an dem Break Glass Accounts am Namen erkannt werden.

.PARAMETER Open
    Öffnet den Bericht nach dem Erstellen im Standardbrowser.

.EXAMPLE
    .\Get-ConditionalAccessReport.ps1 -Open

    Erstellt den Bericht und öffnet ihn im Browser.

.EXAMPLE
    .\Get-ConditionalAccessReport.ps1 -BreakGlassPattern 'bg-|notfall'

    Erkennt Break Glass Accounts an einem eigenen Namensmuster.

.NOTES
    Autor          Michael Heim, https://michaelheim.one
    Voraussetzung  Modul Microsoft.Graph.Authentication
    Berechtigung   Policy.Read.All, Directory.Read.All
    Mehr dazu      https://michaelheim.one/blog/ca-deep-dive-1-dein-digitaler-tuersteher/
#>
[CmdletBinding()]
param(
    [string]$OutputPath = '.\ConditionalAccess-Report.html',

    [string]$BreakGlassPattern = 'breakglass|break-glass|emergency|notfall',

    [switch]$Open
)

if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes 'Policy.Read.All', 'Directory.Read.All' -NoWelcome
}

# Ruft eine Graph-Liste vollständig ab und folgt dabei den Folgeseiten.
function Get-GraphCollection {
    param([Parameter(Mandatory)][string]$Uri)

    while ($Uri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $Uri
        $response.value
        $Uri = $response.'@odata.nextLink'
    }
}

# Macht Text sicher für die Ausgabe in HTML.
function ConvertTo-HtmlText {
    param([object]$Value)
    return [System.Net.WebUtility]::HtmlEncode("$Value")
}

# Übersetzt IDs von Benutzern, Gruppen und Rollen in Namen.
# Jede ID wird nur einmal bei Graph abgefragt.
$script:nameCache = @{}
function Resolve-DirectoryName {
    param([string[]]$Id)

    $names = foreach ($entry in $Id) {
        if (-not $entry) { continue }

        # Feste Werte wie All oder GuestsOrExternalUsers sind keine IDs.
        if ($entry -notmatch '^[0-9a-f]{8}-') { $entry; continue }

        if (-not $script:nameCache.ContainsKey($entry)) {
            try {
                $object = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryObjects/$entry"
                $script:nameCache[$entry] = if ($object.displayName) { $object.displayName } else { $entry }
            }
            catch {
                $script:nameCache[$entry] = "$entry (nicht gefunden)"
            }
        }
        $script:nameCache[$entry]
    }
    return ($names -join ', ')
}

# Rollenvorlagen haben eine eigene ID, die über die Vorlagenliste aufgelöst wird.
$roleTemplates = @{}
foreach ($template in Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/directoryRoleTemplates?$select=id,displayName') {
    $roleTemplates[$template.id] = $template.displayName
}

$policies  = @(Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies')
$locations = @(Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations')

$stateLabels = @{
    enabled                           = 'Aktiv'
    disabled                          = 'Aus'
    enabledForReportingButNotEnforced = 'Nur Bericht'
}

$html = [System.Text.StringBuilder]::new()
[void]$html.Append(@"
<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<title>Conditional Access Bericht</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; margin: 2rem; color: #1f2933; }
  h1 { color: #e65100; }
  h2 { margin-top: 2.5rem; border-bottom: 2px solid #ff6d00; padding-bottom: .3rem; }
  h3 { margin-bottom: .4rem; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 1.5rem; }
  th, td { border: 1px solid #d2d6dc; padding: .5rem .7rem; text-align: left; vertical-align: top; }
  th { background: #f4f5f7; width: 14rem; }
  .state { display: inline-block; padding: .1rem .6rem; border-radius: 1rem; font-size: .85rem; }
  .enabled { background: #d1fae5; }
  .disabled { background: #e5e7eb; }
  .enabledForReportingButNotEnforced { background: #fef3c7; }
  .hint { color: #6b7280; }
</style>
</head>
<body>
<h1>Conditional Access Bericht</h1>
<p class="hint">Erstellt am $(Get-Date -Format 'dd.MM.yyyy HH:mm') mit Get-ConditionalAccessReport.ps1</p>
"@)

# Abschnitt 1. Richtlinien
[void]$html.Append("<h2>1. Richtlinien ($($policies.Count))</h2>")

foreach ($policy in ($policies | Sort-Object displayName)) {
    $users = $policy.conditions.users
    $roleNames = @($users.includeRoles | ForEach-Object { if ($roleTemplates[$_]) { $roleTemplates[$_] } else { $_ } }) -join ', '
    $grant = if ($policy.grantControls) {
        "$($policy.grantControls.operator) $(@($policy.grantControls.builtInControls) -join ', ')"
    } else { 'Keine' }
    $strength = $policy.grantControls.authenticationStrength.displayName
    $session = if ($policy.sessionControls) {
        @($policy.sessionControls.Keys | Where-Object { $policy.sessionControls[$_] }) -join ', '
    } else { 'Keine' }
    $stateLabel = if ($stateLabels[$policy.state]) { $stateLabels[$policy.state] } else { $policy.state }

    [void]$html.Append(@"
<h3>$(ConvertTo-HtmlText $policy.displayName)</h3>
<table>
  <tr><th>Status</th><td><span class="state $(ConvertTo-HtmlText $policy.state)">$(ConvertTo-HtmlText $stateLabel)</span></td></tr>
  <tr><th>Benutzer eingeschlossen</th><td>$(ConvertTo-HtmlText (Resolve-DirectoryName $users.includeUsers))</td></tr>
  <tr><th>Benutzer ausgeschlossen</th><td>$(ConvertTo-HtmlText (Resolve-DirectoryName $users.excludeUsers))</td></tr>
  <tr><th>Gruppen eingeschlossen</th><td>$(ConvertTo-HtmlText (Resolve-DirectoryName $users.includeGroups))</td></tr>
  <tr><th>Gruppen ausgeschlossen</th><td>$(ConvertTo-HtmlText (Resolve-DirectoryName $users.excludeGroups))</td></tr>
  <tr><th>Rollen eingeschlossen</th><td>$(ConvertTo-HtmlText $roleNames)</td></tr>
  <tr><th>Apps</th><td>$(ConvertTo-HtmlText (@($policy.conditions.applications.includeApplications) -join ', '))</td></tr>
  <tr><th>Kontrollen</th><td>$(ConvertTo-HtmlText $grant)</td></tr>
  <tr><th>Authentication Strength</th><td>$(ConvertTo-HtmlText $strength)</td></tr>
  <tr><th>Sitzung</th><td>$(ConvertTo-HtmlText $session)</td></tr>
</table>
"@)
}

# Abschnitt 2. Named Locations
[void]$html.Append("<h2>2. Named Locations ($($locations.Count))</h2>")

if ($locations.Count -eq 0) {
    [void]$html.Append('<p class="hint">Keine Named Locations definiert.</p>')
}
foreach ($location in ($locations | Sort-Object displayName)) {
    if ($location.'@odata.type' -eq '#microsoft.graph.ipNamedLocation') {
        $type = 'IP-Bereiche'
        $detail = @($location.ipRanges | ForEach-Object { $_.cidrAddress }) -join ', '
        $trusted = if ($location.isTrusted) { 'Ja' } else { 'Nein' }
    }
    else {
        $type = 'Länder'
        $detail = @($location.countriesAndRegions) -join ', '
        $trusted = 'Nicht zutreffend'
    }

    [void]$html.Append(@"
<h3>$(ConvertTo-HtmlText $location.displayName)</h3>
<table>
  <tr><th>Typ</th><td>$type</td></tr>
  <tr><th>Inhalt</th><td>$(ConvertTo-HtmlText $detail)</td></tr>
  <tr><th>Vertrauenswürdig</th><td>$trusted</td></tr>
</table>
"@)
}

# Abschnitt 3. Ausgeschlossene Konten
# Jeder Ausschluss ist eine Lücke in der Richtlinie. Sie sollte bewusst
# gesetzt sein, zum Beispiel für Break Glass Accounts.
$excludedIds = @($policies | ForEach-Object { $_.conditions.users.excludeUsers } | Where-Object { $_ -match '^[0-9a-f]{8}-' } | Sort-Object -Unique)

[void]$html.Append("<h2>3. Ausgeschlossene Konten ($($excludedIds.Count))</h2>")

if ($excludedIds.Count -eq 0) {
    [void]$html.Append('<p class="hint">Kein Konto ist direkt aus einer Richtlinie ausgeschlossen. Prüfe, ob Break Glass Accounts über eine Gruppe ausgeschlossen sind.</p>')
}
else {
    [void]$html.Append('<table><tr><th>Name</th><th>Anmeldename</th><th>Sieht nach Break Glass aus</th><th>Ausgeschlossen in</th></tr>')
    foreach ($id in $excludedIds) {
        try {
            $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$id`?`$select=displayName,userPrincipalName"
            $name = $user.displayName
            $upn = $user.userPrincipalName
        }
        catch {
            $name = "$id (nicht gefunden)"
            $upn = ''
        }
        $isBreakGlass = if ("$name $upn" -match $BreakGlassPattern) { 'Ja' } else { 'Nein' }
        $usedIn = @($policies | Where-Object { $_.conditions.users.excludeUsers -contains $id } | ForEach-Object { $_.displayName }) -join ', '

        [void]$html.Append("<tr><td>$(ConvertTo-HtmlText $name)</td><td>$(ConvertTo-HtmlText $upn)</td><td>$isBreakGlass</td><td>$(ConvertTo-HtmlText $usedIn)</td></tr>")
    }
    [void]$html.Append('</table>')
}

[void]$html.Append('</body></html>')

$html.ToString() | Out-File -FilePath $OutputPath -Encoding utf8
Write-Host "Bericht gespeichert unter $OutputPath"

if ($Open) {
    Invoke-Item -Path $OutputPath
}
