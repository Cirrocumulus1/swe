# Produktionsbereitstellung

## Sicherheitsmodell

Die 85 Projektnamen sind für alle 25 Mitarbeitenden bestimmt. Das Add-in fordert deshalb kein Entra-SSO für die Liste an. Der Client ruft ausschließlich `/api/projects` auf; die Azure Function enthält nur den fest verdrahteten, lesenden Teamspace-Abruf und paginiert serverseitig. Der Teamspace-Benutzername und das Passwort bleiben in Function-App-Einstellungen beziehungsweise einem unterstützten Key-Vault-Setup.

Der Endpunkt ist anonym erreichbar und liefert die Projektliste an jeden Aufrufer, der seine URL kennt. Das ist beabsichtigt, weil die Liste ohnehin für alle Mitarbeitenden sichtbar sein soll. Er ist kein beliebiger Teamspace-Proxy und bietet nur `GET /api/projects`; Teamspace-Zugangsdaten oder Schreiboperationen werden nicht an den Client weitergegeben. Trotzdem muss der Teamspace-Dienstbenutzer auf die nötigen Leserechte beschränkt sein. Falls die Projektliste nicht öffentlich aus dem Internet abrufbar sein soll, braucht der Endpunkt eine Zugangsbeschränkung auf Netzwerk-/Hostingebene.

Die Function cached die Teamspace-Antwort 15 Minuten im Arbeitsspeicher. Der Browser cached sie vier Stunden in `localStorage`; manuelles Neuladen umgeht den Browsercache, nicht den serverseitigen 15-Minuten-Cache.

## Sofortmaßnahme: Teamspace-Zugang

Die zuvor im Client enthaltenen Zugangsdaten als offengelegt behandeln: im Teamspace rotieren oder sperren, Berechtigungen prüfen und Audit-Logs kontrollieren. Neue, lesebeschränkte Zugangsdaten ausschließlich als Azure Function App Settings beziehungsweise in einem unterstützten Key-Vault-Setup hinterlegen. Nicht in Manifest, Clientcode, Buildvariablen oder Git einchecken. `api/local.settings.json` ist ignoriert und darf nicht committet werden.

## Azure-Konfiguration

Die Function benötigt diese App Settings:

- `TEAMSPACE_API_BASE_URL`
- `TEAMSPACE_USERNAME`
- `TEAMSPACE_PASSWORD`

Azure Static Web Apps kann die statischen Dateien aus `dist/` und die Functions aus `api/` gemeinsam bereitstellen. Alternativ eine dedizierte Azure Function App hinter derselben Add-in-Domain/API-Route verwenden. `/api` soll same-origin bleiben, damit keine CORS-Freigabe nötig ist. Server-/Front-Door-Limits können zusätzlich vor übermäßigem Abruf schützen; der Endpoint selbst nimmt keine URL- oder Filterparameter vom Aufrufer entgegen.

Für lokale Entwicklung Azure Functions Core Tools v4 installieren, `api/local.settings.example.json` als `api/local.settings.json` kopieren und dort lokal neue Zugangsdaten eintragen. Danach `npm run start:api` und in einem zweiten Terminal `npm run dev-server` starten. Der Dev-Server leitet `/api` an `localhost:7071` weiter.

## Produktionsmanifest und Build

`production-build-variables.example` enthält nicht geheime Buildwerte. Beispieldomain und Organisationsangaben durch echte HTTPS-Werte ersetzen. Die ältere `.env.production.example` ist ein nicht mehr verwendeter SSO-Entwurf; Webpack liest sie nicht.

In PowerShell:

```powershell
$env:ADDIN_BASE_URL = "https://<echte-add-in-domain>/"
$env:DEVELOPER_NAME = "<Organisation>"
$env:DEVELOPER_WEBSITE_URL = "https://<Organisationswebsite>/"
$env:PRIVACY_URL = "https://<Datenschutzseite>/"
$env:TERMS_OF_USE_URL = "https://<Nutzungsbedingungen>/"
npm run validate
npm run build
npm run validate:prod
```

Der Build ersetzt im erzeugten Manifest Taskpane- und Icon-URLs, `validDomains` und Entwicklerangaben. Die vorhandene Add-in-GUID bleibt erhalten. Die Manifestversion muss für jedes veröffentlichte Update erhöht werden; derzeit steht sie auf `1.0.0`.

Danach `dist/` und die API über die gewählte Azure-Pipeline veröffentlichen, HTTPS und `/api/projects` prüfen und die hochgeladene Manifestdatei validieren. Keine Zugangsdaten in die statischen `dist/`-Dateien aufnehmen.

## M365-Rollout

Zuerst eine Pilotgruppe in Outlook Web und den unterstützten Desktop-Clients testen. Anschließend kann ein M365-Administrator die Produktionsmanifestdatei über **Admin Center → Settings → Integrated apps** einer Sicherheitsgruppe zuweisen. Rolloutstatus und Verfügbarkeit in den Zielclients prüfen; die Synchronisationsdauer ist nicht garantiert.
