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

    Rollen, die über eine Gruppe zugewiesen sind, werden aufgelöst. Die
    Mitglieder der Gruppe erscheinen im Bericht, auch aus verschachtelten
    Gruppen. In der Spalte Rollen steht dann die Gruppe dahinter.

    Erfasst werden aktive, tenantweite Rollenzuweisungen. Berechtigte
    Zuweisungen in PIM, die gerade nicht aktiviert sind, und Zuweisungen auf
    einzelne Verwaltungseinheiten tauchen nicht auf.

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
                   GroupMember.Read.All, UserAuthenticationMethod.Read.All
#>
[CmdletBinding()]
param(
    [switch]$OnlyWithoutMfa,

    [string]$ExportPath
)

$requiredScopes = @(
    'RoleManagement.Read.Directory'
    'User.Read.All'
    'GroupMember.Read.All'
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

# Merkt ein Konto mit einer Rolle vor. Mehrere Rollen je Konto werden gesammelt.
function Add-AdminEntry {
    param($User, [string]$RoleLabel)

    if (-not $admins.ContainsKey($User.id)) {
        $admins[$User.id] = [PSCustomObject]@{
            User  = $User
            Roles = [System.Collections.Generic.List[string]]::new()
        }
    }
    if ($RoleLabel -notin $admins[$User.id].Roles) {
        $admins[$User.id].Roles.Add($RoleLabel)
    }
}

# Schritt 1. Alle aktivierten Rollen und ihre Mitglieder einsammeln.
$admins = @{}
$userSelect = 'id,displayName,userPrincipalName,accountEnabled'
$roles = Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/directoryRoles?$select=id,displayName'

foreach ($role in $roles) {
    $members = Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)/members"

    foreach ($member in $members) {
        switch ($member.'@odata.type') {
            '#microsoft.graph.user' {
                Add-AdminEntry -User $member -RoleLabel $role.displayName
            }
            '#microsoft.graph.group' {
                # Rolle hängt an einer Gruppe. Deren Mitglieder sind genauso Admins,
                # deshalb werden sie aufgelöst, verschachtelte Gruppen inklusive.
                try {
                    $groupUsers = @(Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/groups/$($member.id)/transitiveMembers/microsoft.graph.user?`$select=$userSelect")
                }
                catch {
                    # Lieber laut scheitern als Admins stillschweigend auslassen.
                    throw "Die Mitglieder der Gruppe '$($member.displayName)' mit der Rolle '$($role.displayName)' konnten nicht gelesen werden. Der Bericht wäre unvollständig. Prüfe die Berechtigung GroupMember.Read.All. $($_.Exception.Message)"
                }
                foreach ($groupUser in $groupUsers) {
                    Add-AdminEntry -User $groupUser -RoleLabel "$($role.displayName) (über Gruppe $($member.displayName))"
                }
            }
            # Dienstprinzipale haben keine MFA und werden hier nicht bewertet.
        }
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
