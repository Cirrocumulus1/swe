# Produktionsbereitstellung

## Sicherheitsmodell

Die 85 Projektnamen sind für alle 25 Mitarbeitenden bestimmt. Der Client ruft `/api/projects` für die Liste und `/api/sync-email` für den Push auf. Die Projektliste bleibt anonym lesbar; der Sync-Endpunkt prüft ein Outlook-SSO-Token auf Signatur, Tenant, App-ID, Scope und Übereinstimmung der Benutzer-Mailbox. Graph-App-Zugangsdaten bleiben ausschließlich in Function-App-Einstellungen beziehungsweise einem unterstützten Key-Vault-Setup.

`GET /api/projects` ist anonym erreichbar und liefert die Projektliste an jeden Aufrufer, der seine URL kennt. Das ist beabsichtigt, weil die Liste ohnehin für alle Mitarbeitenden sichtbar sein soll. `POST /api/sync-email` akzeptiert nur ein geprüftes Outlook-SSO-Token und ist kein beliebiger Graph-Proxy. Trotzdem muss der Teamspace-Dienstbenutzer auf die nötigen Leserechte beschränkt sein. Falls die Projektliste nicht öffentlich aus dem Internet abrufbar sein soll, braucht der Endpunkt eine Zugangsbeschränkung auf Netzwerk-/Hostingebene.

Die Function cached die Teamspace-Antwort 15 Minuten im Arbeitsspeicher. Der Browser cached sie vier Stunden in `localStorage`; manuelles Neuladen umgeht den Browsercache, nicht den serverseitigen 15-Minuten-Cache.

## Sofortmaßnahme: Teamspace-Zugang

Die zuvor im Client enthaltenen Zugangsdaten als offengelegt behandeln: im Teamspace rotieren oder sperren, Berechtigungen prüfen und Audit-Logs kontrollieren. Neue, lesebeschränkte Zugangsdaten ausschließlich als Azure Function App Settings beziehungsweise in einem unterstützten Key-Vault-Setup hinterlegen. Nicht in Manifest, Clientcode, Buildvariablen oder Git einchecken. `api/local.settings.json` ist ignoriert und darf nicht committet werden.

## Azure-Konfiguration

Die Function benötigt diese App Settings:

- `TEAMSPACE_API_BASE_URL`
- `TEAMSPACE_USERNAME`
- `TEAMSPACE_PASSWORD`
- `TENANT_ID` (Microsoft Entra Tenant-ID)
- `CLIENT_ID` (Entra-App-ID, identisch mit der SSO-App-ID im Manifest)
- `CLIENT_SECRET` (nur in Function-App-Einstellungen oder Key Vault)
- `ADDIN_APP_ID_URI` (optional; Standard: `api://<CLIENT_ID>`, muss zur Manifest-Ressource passen)

Die Entra-App muss die delegierte API-Berechtigung `access_as_user` bereitstellen. Für den serverseitigen Graph-Zugriff sind Application-Berechtigungen `Mail.Read` und `Sites.ReadWrite.All` mit Admin-Consent erforderlich. Da `Mail.Read` als App-Berechtigung grundsätzlich tenantweit lesen kann, die Mailbox-Zugriffe zusätzlich per Exchange Application RBAC auf die Mitarbeitenden begrenzen. Die Site-Auflösung sucht mit `GET /v1.0/sites?search=<teamspaceId>` und prüft die Website-Beschreibung; dafür benötigt die App tenantweite Site-Suchberechtigungen. `Sites.Selected` allein reicht für diese Suche nicht aus. Wenn tenantweite Site-Suche nicht genehmigt wird, muss die Architektur stattdessen eine explizite Site-ID-Zuordnung verwenden.

Azure Static Web Apps kann die statischen Dateien aus `dist/` und die Functions aus `api/` gemeinsam bereitstellen. Alternativ eine dedizierte Azure Function App hinter derselben Add-in-Domain/API-Route verwenden. `/api` soll same-origin bleiben, damit keine CORS-Freigabe nötig ist. Server-/Front-Door-Limits können zusätzlich vor übermäßigem Abruf schützen; der Endpoint selbst nimmt keine URL- oder Filterparameter vom Aufrufer entgegen.

Für lokale Entwicklung Azure Functions Core Tools v4 installieren, Azurite bereitstellen und `api/local.settings.example.json` als `api/local.settings.json` kopieren. Dort lokale Function-Einstellungen eintragen. Die drei Dienste separat starten:

```powershell
npm run start:storage
npm run start:api
npm run dev-server
```

Der Dev-Server leitet `/api` an `localhost:7071` weiter. `npm start` startet ebenfalls nur den Dev-Server und öffnet keine Outlook-Desktop-Anwendung.

## New Outlook / OWA-Sideload

Dieses Projekt verwendet das Unified-JSON-Manifest. Das Outlook-Dialogfeld **Add from File** ist der XML-Manifest-Weg; für das JSON-Manifest wird ein Microsoft 365-App-Paket benötigt. `npm run sideload:owa` erstellt ein temporäres ZIP aus `manifest.json` und den referenzierten Icons, registriert es mit `atk install` und öffnet danach Outlook im Web. Dabei wird `OUTLOOK.EXE` nicht gestartet. Das CLI benötigt eine angemeldete Microsoft-365-Testidentität; bei Aufforderung zuerst `npx --yes -p @microsoft/m365agentstoolkit-cli atk auth login m365` ausführen.

Alternativ in VS Code **Tasks: Run Task → Sideload: New Outlook / OWA** ausführen. Für den reinen Browserstart ohne erneutes Registrieren die Launch-Konfiguration **Outlook Web (New Outlook / OWA)** verwenden. Das Outlook-Webkonto muss zur Entra-Konfiguration passen; für den SSO-Test in der Inbox dieses Kontos arbeiten.

Microsoft beschreibt das Unified-Manifest-Sideloading unter [Sideload Office Add-ins that use the unified manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest). Die Registrierung ist kontobezogen und kann nach dem Test mit `atk uninstall --mode title-id --title-id <TitleId> --interactive false` entfernt werden; `TitleId` wird beim Installieren ausgegeben.

## Produktionsmanifest und Build

`production-build-variables.example` enthält nicht geheime Buildwerte. Beispieldomain und Organisationsangaben durch echte HTTPS-Werte ersetzen. Die ältere `.env.production.example` ist ein nicht mehr verwendeter SSO-Entwurf; Webpack liest sie nicht.

Der Produktionsmanifest-Build setzt `ADDIN_MANIFEST_ID` und `ENTRA_CLIENT_ID` zusätzlich zur Add-in-Domain voraus. `ADDIN_MANIFEST_ID` ist die unveränderte Microsoft-365-Paket-ID der bestehenden Produktion; der Workflow setzt sie auf `a860814c-e238-4e66-b96a-5ac1f4dae627`, damit das neue Paket die veröffentlichte App aktualisiert. Der GitHub-Actions-Workflow liest die nicht geheime Entra-App-ID aus der Repository-Variable `ENTRA_CLIENT_ID`; diese Variable muss vor dem Push eingerichtet werden. Die Function-App-Einstellung `ADDIN_APP_ID_URI` muss mit der Ressource im Manifest übereinstimmen, falls statt des Standards `api://<CLIENT_ID>` eine eigene App-ID-URI verwendet wird. Für lokale SSO-Tests `webApplicationInfo.id` und `webApplicationInfo.resource` in `manifest.json` auf dieselbe App-Registrierung setzen.

In PowerShell:

```powershell
$env:ADDIN_BASE_URL = "https://<echte-add-in-domain>/"
$env:ADDIN_MANIFEST_ID = "<bestehende-Microsoft-365-Paket-ID>"
$env:ENTRA_CLIENT_ID = "<Entra-App-ID>"
$env:DEVELOPER_NAME = "<Organisation>"
$env:DEVELOPER_WEBSITE_URL = "https://<Organisationswebsite>/"
$env:PRIVACY_URL = "https://<Datenschutzseite>/"
$env:TERMS_OF_USE_URL = "https://<Nutzungsbedingungen>/"
npm run validate
npm run build
npm run validate:prod
```

Der Build ersetzt im erzeugten Manifest Taskpane- und Icon-URLs, `validDomains` und Entwicklerangaben. Die vorhandene Add-in-GUID bleibt erhalten. Die Manifestversion muss für jedes veröffentlichte Update erhöht werden; das aktuelle Manifest steht auf `1.0.1`. Für Node-Builds und die SWA-Function-Runtime ist Node.js 22 festgelegt.

Es existiert derzeit kein Git-Release-Tag. Einen Tag wie `v1.0.1` erst auf dem finalen Release-Commit anlegen, nicht auf einem noch uncommitteten Arbeitsstand.

Danach `dist/` und die API über die gewählte Azure-Pipeline veröffentlichen, HTTPS und `/api/projects` prüfen und die hochgeladene Manifestdatei validieren. Keine Zugangsdaten in die statischen `dist/`-Dateien aufnehmen.

## M365-Rollout

Zuerst eine Pilotgruppe in Outlook Web und den unterstützten Desktop-Clients testen. Anschließend kann ein M365-Administrator die Produktionsmanifestdatei über **Admin Center → Settings → Integrated apps** einer Sicherheitsgruppe zuweisen. Rolloutstatus und Verfügbarkeit in den Zielclients prüfen; die Synchronisationsdauer ist nicht garantiert.
