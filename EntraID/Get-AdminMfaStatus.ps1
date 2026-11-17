#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Zeigt alle Konten mit aktiver Administratorrolle und ihre registrierten MFA-Methoden.

.DESCRIPTION
    Liest die Mitglieder aller aktivierten Verzeichnisrollen aus und prüft je
    Konto, welche Anmeldemethoden registriert sind. Das Ergebnis zeigt auf einen
    Blick, welche Admins keine MFA-Methode oder keine phishing-resistente
    Methode haben.

    Ein Passwort allein zählt nicht als MFA. Als phishing-resistent gelten
    FIDO2 (Passkeys), Windows Hello for Business und Platform Credential
    (macOS Platform SSO).

    Erfasst werden aktive Rollenzuweisungen. Berechtigte Zuweisungen in PIM, die
    gerade nicht aktiviert sind, tauchen nicht auf.

    Das Skript liest nur. Es ändert nichts im Tenant.

.PARAMETER OnlyWithoutMfa
    Gibt nur Konten ohne registrierte MFA-Methode aus.

.PARAMETER ExportPath
    Optionaler Pfad für einen CSV-Export.

.EXAMPLE
    .\Get-AdminMfaStatus.ps1

    Listet alle Admins mit Rollen und Methoden auf.

.EXAMPLE
    .\Get-AdminMfaStatus.ps1 -OnlyWithoutMfa

    Zeigt nur die Admins, bei denen keine MFA-Methode registriert ist.

.NOTES
    Autor          Michael Heim, https://michaelheim.one
    Voraussetzung  Modul Microsoft.Graph.Authentication
    Berechtigung   RoleManagement.Read.Directory, User.Read.All,
                   UserAuthenticationMethod.Read.All
#>
[CmdletBinding()]
param(
    [switch]$OnlyWithoutMfa,

    [string]$ExportPath
)

$requiredScopes = @(
    'RoleManagement.Read.Directory'
    'User.Read.All'
    'UserAuthenticationMethod.Read.All'
)

# Methoden, die als zweiter Faktor zählen. Schlüssel ist der Graph-Typ,
# Wert ist die Anzeige im Bericht.
$mfaMethods = @{
    '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'  = 'Authenticator'
    '#microsoft.graph.fido2AuthenticationMethod'                   = 'FIDO2'
    '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' = 'Windows Hello'
    '#microsoft.graph.platformCredentialAuthenticationMethod'      = 'Platform Credential'
    '#microsoft.graph.softwareOathAuthenticationMethod'            = 'Software OATH'
    '#microsoft.graph.phoneAuthenticationMethod'                   = 'Telefon'
}

$phishingResistant = @('FIDO2', 'Windows Hello', 'Platform Credential')

# Ruft eine Graph-Liste vollständig ab und folgt dabei den Folgeseiten.
function Get-GraphCollection {
    param([Parameter(Mandatory)][string]$Uri)

    while ($Uri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $Uri
        $response.value
        $Uri = $response.'@odata.nextLink'
    }
}

if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
}

# Schritt 1. Alle aktivierten Rollen und ihre Mitglieder einsammeln.
$admins = @{}
$roles = Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/directoryRoles?$select=id,displayName'

foreach ($role in $roles) {
    $members = Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)/members?`$select=id,displayName,userPrincipalName,accountEnabled"

    foreach ($member in $members) {
        # Gruppen und Dienstprinzipale überspringen, hier geht es um Benutzer.
        if ($member.'@odata.type' -ne '#microsoft.graph.user') { continue }

        if (-not $admins.ContainsKey($member.id)) {
            $admins[$member.id] = [PSCustomObject]@{
                User  = $member
                Roles = [System.Collections.Generic.List[string]]::new()
            }
        }
        $admins[$member.id].Roles.Add($role.displayName)
    }
}

Write-Verbose "$($admins.Count) Konten mit aktiver Administratorrolle gefunden."

# Schritt 2. Je Konto die registrierten Methoden auswerten.
$report = foreach ($admin in $admins.Values) {
    $methods = Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/users/$($admin.User.id)/authentication/methods"

    $registered = @(
        $methods |
            ForEach-Object { $mfaMethods[$_.'@odata.type'] } |
            Where-Object { $_ } |
            Sort-Object -Unique
    )

    [PSCustomObject]@{
        DisplayName        = $admin.User.displayName
        UserPrincipalName  = $admin.User.userPrincipalName
        Aktiv              = $admin.User.accountEnabled
        Rollen             = ($admin.Roles | Sort-Object) -join ', '
        MfaMethoden        = $registered -join ', '
        MfaRegistriert     = $registered.Count -gt 0
        PhishingResistent  = [bool]($registered | Where-Object { $_ -in $phishingResistant })
    }
}

if ($OnlyWithoutMfa) {
    $report = $report | Where-Object { -not $_.MfaRegistriert }
}

$report = $report | Sort-Object MfaRegistriert, DisplayName

if ($ExportPath) {
    $report | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "Export gespeichert unter $ExportPath"
}

$report
