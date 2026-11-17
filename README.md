# M365 Scripts

PowerShell-Skripte aus meinem Alltag als Microsoft 365 Consultant. Kein Framework, keine Abhängigkeiten außer den offiziellen Modulen. Jedes Skript beantwortet eine Frage, die in fast jedem Tenant auftaucht.

Alle Skripte in diesem Repo **lesen nur**. Sie ändern nichts im Tenant.

## Skripte

| Skript | Frage | Berechtigung |
|---|---|---|
| [EntraID/Get-AdminMfaStatus.ps1](EntraID/Get-AdminMfaStatus.ps1) | Welche Admins haben keine oder keine phishing-resistente MFA? | RoleManagement.Read.Directory, User.Read.All, GroupMember.Read.All, UserAuthenticationMethod.Read.All |
| [EntraID/Get-ConditionalAccessReport.ps1](EntraID/Get-ConditionalAccessReport.ps1) | Wie sehen meine Conditional Access Richtlinien aus, und wer ist ausgeschlossen? | Policy.Read.All, Directory.Read.All |
| [Exchange/Get-MailboxSizeReport.ps1](Exchange/Get-MailboxSizeReport.ps1) | Wie groß sind meine Postfächer? | View-Only Recipients |
| [Licensing/Get-LicenseReport.ps1](Licensing/Get-LicenseReport.ps1) | Welche Lizenzen habe ich, und wie viele sind frei? | Organization.Read.All |
| [Reports/Get-InactiveUserReport.ps1](Reports/Get-InactiveUserReport.ps1) | Welche lizenzierten Konten nutzt niemand mehr? | Reports.Read.All |
| [Tenant/Export-TenantConfig.ps1](Tenant/Export-TenantConfig.ps1) | Wie sichere ich meine Sicherheitseinstellungen, um Änderungen nachzuvollziehen? | Policy.Read.All, RoleManagement.Read.Directory, Directory.Read.All |

## Voraussetzungen

- PowerShell 7 oder neuer
- Je nach Skript eines dieser Module

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module MicrosoftTeams -Scope CurrentUser
```

## Benutzung

Jedes Skript bringt seine Hilfe mit.

```powershell
Get-Help .\EntraID\Get-AdminMfaStatus.ps1 -Full
```

Beim ersten Aufruf öffnet sich die Anmeldung. Angefragt werden nur die Leserechte, die das jeweilige Skript braucht.

```powershell
# Admins ohne MFA
.\EntraID\Get-AdminMfaStatus.ps1 -OnlyWithoutMfa

# Conditional Access als HTML-Bericht
.\EntraID\Get-ConditionalAccessReport.ps1 -Open

# Konten ohne Aktivität in 90 Tagen als CSV
.\Reports\Get-InactiveUserReport.ps1 -Period D90 -OnlyInactive -ExportPath .\Inaktiv.csv
```

Die Skripte geben Objekte zurück. Du kannst das Ergebnis also weiterverarbeiten.

```powershell
.\Licensing\Get-LicenseReport.ps1 | Where-Object Frei -gt 10 | Format-Table
```

## Hinweise

- Teste jedes Skript zuerst in einem Test-Tenant oder mit einem Konto, das nur Leserechte hat.
- Die Skripte enthalten keine Zugangsdaten. Angemeldet wird interaktiv.
- Fehler gefunden oder eine Idee? Ich freue mich über ein Issue.

## Mehr dazu

Hintergründe zu den Themen stehen in meinem Blog.

- [Conditional Access Deep Dive](https://michaelheim.one/blog/ca-deep-dive-1-dein-digitaler-tuersteher/)
- [Entra ID Security Split](https://michaelheim.one/blog/entra-id-security-the-core/)
- [Copilot Readiness](https://michaelheim.one/blog/copilot-readiness-1-copilot-sieht-alles/)

Train Hard. Work Smart.

[michaelheim.one](https://michaelheim.one) · [LinkedIn](https://www.linkedin.com/in/michaelheim/)
