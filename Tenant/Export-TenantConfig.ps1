#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Sichert zentrale Sicherheitseinstellungen des Tenants als JSON-Dateien.

.DESCRIPTION
    Schreibt die wichtigsten Konfigurationen in einen Ordner mit Zeitstempel.
    Wer das Skript regelmäßig laufen lässt, kann zwei Stände vergleichen und
    sieht, was sich wann geändert hat. Das hilft nach Vorfällen, bei Audits und
    vor größeren Umbauten.

    Aus Microsoft Graph werden immer gesichert
      - Conditional Access Richtlinien und Named Locations
      - Verzeichnisrollen mit ihren Mitgliedern
      - Standardberechtigungen der Benutzer (Authorization Policy)
      - Richtlinie der Authentifizierungsmethoden

    Optional kommen dazu
      - Transportregeln aus Exchange Online
      - Messaging- und Meeting-Richtlinien aus Microsoft Teams

    Das Skript liest nur. Es ändert nichts im Tenant.

.PARAMETER OutputPath
    Basisordner für die Sicherung. Darin entsteht je Lauf ein Unterordner mit Zeitstempel.

.PARAMETER IncludeExchange
    Sichert zusätzlich die Transportregeln. Benötigt das Modul ExchangeOnlineManagement.

.PARAMETER IncludeTeams
    Sichert zusätzlich Teams-Richtlinien. Benötigt das Modul MicrosoftTeams.

.EXAMPLE
    .\Export-TenantConfig.ps1

    Sichert die Graph-Einstellungen nach .\M365-Config\<Zeitstempel>.

.EXAMPLE
    .\Export-TenantConfig.ps1 -OutputPath D:\Backup\M365 -IncludeExchange -IncludeTeams

    Sichert zusätzlich Exchange und Teams in einen eigenen Ordner.

.NOTES
    Autor          Michael Heim, https://michaelheim.one
    Voraussetzung  Modul Microsoft.Graph.Authentication, optional
                   ExchangeOnlineManagement und MicrosoftTeams
    Berechtigung   Policy.Read.All, RoleManagement.Read.Directory, Directory.Read.All
#>
[CmdletBinding()]
param(
    [string]$OutputPath = '.\M365-Config',

    [switch]$IncludeExchange,

    [switch]$IncludeTeams
)

$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$targetFolder = Join-Path -Path $OutputPath -ChildPath $timestamp
New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null

# Ruft eine Graph-Liste vollständig ab und folgt dabei den Folgeseiten.
function Get-GraphCollection {
    param([Parameter(Mandatory)][string]$Uri)

    while ($Uri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $Uri
        $response.value
        $Uri = $response.'@odata.nextLink'
    }
}

# Schreibt ein Objekt als JSON in den Zielordner.
function Save-Config {
    param(
        [Parameter(Mandatory)][string]$Name,
        [object]$Data
    )

    $file = Join-Path -Path $targetFolder -ChildPath "$Name.json"
    ConvertTo-Json -InputObject $Data -Depth 20 | Out-File -FilePath $file -Encoding utf8
    Write-Host "Gesichert  $Name"
}

# Microsoft Graph
if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes 'Policy.Read.All', 'RoleManagement.Read.Directory', 'Directory.Read.All' -NoWelcome
}

Save-Config -Name 'ConditionalAccess-Policies' -Data @(Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies')
Save-Config -Name 'ConditionalAccess-NamedLocations' -Data @(Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations')
Save-Config -Name 'AuthorizationPolicy' -Data (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/authorizationPolicy')
Save-Config -Name 'AuthenticationMethodsPolicy' -Data (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy')

$roles = foreach ($role in Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/directoryRoles?$select=id,displayName') {
    $members = @(Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)/members?`$select=id,displayName,userPrincipalName")
    [PSCustomObject]@{
        Rolle      = $role.displayName
        Mitglieder = @($members | ForEach-Object {
            [PSCustomObject]@{
                Id   = $_.id
                Name = $_.displayName
                Upn  = $_.userPrincipalName
                Typ  = $_.'@odata.type'
            }
        })
    }
}
Save-Config -Name 'DirectoryRoles' -Data @($roles)

# Exchange Online, optional
if ($IncludeExchange) {
    if (-not (Get-ConnectionInformation)) {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    Save-Config -Name 'Exchange-TransportRules' -Data @(Get-TransportRule)
}

# Microsoft Teams, optional
if ($IncludeTeams) {
    Connect-MicrosoftTeams | Out-Null
    Save-Config -Name 'Teams-MessagingPolicies' -Data @(Get-CsTeamsMessagingPolicy)
    Save-Config -Name 'Teams-MeetingPolicies' -Data @(Get-CsTeamsMeetingPolicy)
}

Write-Host "Fertig. Sicherung liegt in $targetFolder"
