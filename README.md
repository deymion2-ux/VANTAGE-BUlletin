# VANTAGE Intranet V3.8.9 - Security Update (Entwicklerpaket)

**Stand:** 05.11.2026 · **Basis:** VANTAGE Bulletin V3.8.8 · **Status:** Technischer Überarbeitungsvorschlag; vor produktiver Freigabe Integrations- und Sicherheitstests durchführen.

## Dateien

- `index.html` - Anwendung mit CSP und ohne Inline-JavaScript/Eventhandler.
- `assets/app.js` - Anwendungslogik, MSAL v2-kompatibel, sichere DOM-Verarbeitung.
- `assets/styles.css` - statisch erzeugtes Tailwind CSS v4.1.10 + ursprüngliche Layoutregeln; kein Tailwind-CDN.
- `assets/icons.js` - lokal ausgelieferte Symbolgrafiken (Font Awesome Free Solid, siehe Lizenz).
- `nginx-security-headers.conf` - **Muster** für HTTP-Sicherheitsheader inkl. `frame-ancestors 'none'`.
- `LICENSE_FontAwesome.txt` - Lizenzhinweise für die lokal eingebundenen Symbole.

## Sicherheitskorrekturen

1. **XSS:** Graph-Daten (Kalenderbetreff, Benutzername, E-Mail-Adresse, Fehlermeldungen) werden ausschließlich als Text via `textContent` ausgegeben; keine Einbettung in HTML-Strings oder Inline-JavaScript.
2. **Graph-Minimalprinzip:** Nur delegierte Scopes `User.Read`, `Calendars.ReadBasic`, `User.ReadBasic.All`. Kein Zugriff auf Mails oder SharePoint per Graph, kein ungenutztes `Presence.Read.All`.
3. **CSP:** Kein `unsafe-inline`, kein `unsafe-eval` in `script-src`; entfernte Inline-Eventhandler; CSP-Hinweis zusätzlich als HTTP-Header mit `frame-ancestors`.
4. **Statische Styles/Icons:** Kein Tailwind-Play-CDN und kein fremdes Icon-CDN; kompiliertes CSS und Icon-Skript werden lokal ausgeliefert.
5. **Abmeldung:** Lokale Mitarbeiterdaten werden entfernt, Suchfeld geleert, ausstehende Graph-Fetches abgebrochen, Authentifizierungszustand zurückgesetzt.
6. **Authentifizierung:** Eindeutige aktive Benutzerkonten, kontrollierte Interaktions-Fallbacks, sicherere Fehlermeldungen, Outlook-URL-Validierung per `URL`/Allowlist.
7. **Funktionswahrheit:** Keine erfundenen grünen Online-Punkte; Mitarbeiter-Verzeichnis statt Teams-Live-Präsenz; nicht implementierter Chatversand deaktiviert.
8. **Benutzerverzeichnis:** Paginierung bis max. 5 Seiten à 100 Einträge, nur Basisfelder (`id`, `displayName`, `mail`), keine Abfrage von `jobTitle`.
9. **Kalender:** `/me/calendarView` zeigt den persönlichen Kalender; UTC-Rückgabe wird korrekt als UTC-Zeit verarbeitet und lokal dargestellt.

## Voraussetzungen vor dem Test

1. **Alle Dateien in derselben Ordnerstruktur** auf einem HTTPS-Webserver bereitstellen. `file://` / direkter Doppelklick auf `index.html` ist für die Microsoft-365-Anmeldung nicht geeignet.
2. In **Entra ID → App-Registrierungen** für die bestehende Client-ID die korrekte, exakt passende **SPA Redirect-URI** registrieren. Die Anwendung nutzt `window.location.origin + window.location.pathname`, z. B. `https://intranet.example.org/index.html`. Die verwendete Seite muss zu dieser URI passen.
3. Die **API-Berechtigungen tatsächlich in Entra ID anpassen**. Das Entfernen alter Scopes aus dem Code widerruft *nicht* bereits gewährte `Mail.Read`, `Sites.Read.All`, `Presence.Read.All` oder andere Einwilligungen.
4. `nginx-security-headers.conf` an den tatsächlichen Webserver anpassen, aktivieren und die CSP-Fehler in den Browser-DevTools prüfen. Insbesondere `frame-ancestors` funktioniert nicht in einem `<meta>`-Tag.
5. Den realen Login, Consent, Tokenrefresh, Logout (auch Popup-Blocker und Safari), Graph-Abfragen und Kalender-Zeitzonen im echten Tenant prüfen.

## Noch offener Punkt: MSAL-Abhängigkeit

**Wichtig:** Um die bestehende Microsoft-365-Anmeldung nicht ungetestet zu brechen, nutzt diese Version noch das von Microsoft gehostete **MSAL Browser 2.38.0** (`https://alcdn.msauth.net`). Diese externe Abhängigkeit und die alte Hauptversion sind **nicht** vollständig behoben. Für eine produktive Freigabe sollte der Entwickler eine unterstützte MSAL-Version per npm beziehen, lokal bündeln und die Initialisierung (`await initialize()` ab v3), Redirect-Behandlung und Login/Logout-Abläufe vollständig neu testen. Danach kann die externe CDN-Domain aus der CSP entfernt werden. **Der aktuelle Stand ist deshalb noch keine abschließende Produktions-Sicherheitsfreigabe.**

## Angepasste beziehungsweise nicht implementierte Funktionen

- Outlook-Mailübersicht und SharePoint-Graph-Mapping waren bereits in V3.8.8 nicht mehr vorhanden; sie wurden **nicht** wiederhergestellt.
- Es wird **kein** Teams-Live-Präsenzstatus abgefragt. Die neutralen Markierungen im Verzeichnis sind ausdrücklich keine Präsenzanzeige. Für echte Präsenz müsste ein separater, begründeter Scope `Presence.Read.All` plus Graph Presence API ergänzt werden.
- Der Chat-Button ist deaktiviert, weil `sendDetailChatMessage()` in der Vorlage nicht implementiert ist.
- Externe Links bleiben als von den Entwicklern vorgegebene Konstanten erhalten; SharePoint-Berechtigungen für die Zielseite richten sich weiterhin nach den Benutzerrechten.

## Prüfliste (vor produktiver Freigabe)

- [ ] XSS-Test: präparierter Outlook-Terminbetreff führt zu keiner Skriptausführung.
- [ ] XSS-Test: präparierter M365-Anzeigename/E-Mail-Adresse wird nur als Text dargestellt.
- [ ] Keine JS-Fehler, keine unerwarteten CSP-Verletzungen, keine Inline-Eventhandler.
- [ ] HTTP-CSP mit `frame-ancestors` aktiv; weitere HTTP-Header kontrolliert.
- [ ] Login/Logout/Refresh/Redirect/Safari und Mehrkonten-Situation im Tenant getestet.
- [ ] `Calendars.ReadBasic` und `User.ReadBasic.All` reichen im realen Tenant aus.
- [ ] Nicht mehr benötigte *früher gewährte* API-Scopes in Entra ID widerrufen.
- [ ] Verzeichnisdaten und Kalender werden nach Logout nicht wieder eingeblendet; parallele Fetches abgebrochen.
- [ ] Gemeinsamen Kalender fachlich geklärt; `/me/calendarView` entspricht dem persönlichen Kalender.
- [ ] MSAL modernisiert, lokal gebündelt und vollständig getestet.
- [ ] Externe Linkziele, HTTPS, CORS/CSP und Deployment-Berechtigungen im Betrieb geprüft.

## Für Entwickler: Schnelltest

Nach dem Entpacken kann der lokale Mock-Test mit `node tests/security-smoke.cjs` durchgeführt werden. Er testet u. a. Textausgabe präparierter Graph-Inhalte und Outlook-Link-Validierung, ersetzt aber keinen Browser-/Tenant-Integrationstest.

## Grenzen der bereitgestellten Prüfung

Die Auslieferung wurde statisch geprüft (`node --check`, HTML-Inspektion) und mit einem kleinen DOM-/API-Mock-Smoke-Test (`node tests/security-smoke.cjs`) getestet. Es wurde **kein echter Zugriff** auf den Entra-Mandanten, seine tatsächlich erteilten Scopes oder Kalenderdaten durchgeführt. Ein vollwertiger Chromium-Integrationstest war in der Erstellungsumgebung technisch nicht ausführbar. Die Test-Checkliste ist daher verbindlich vom Entwicklungsteam auszuführen, bevor eine Freigabe erfolgen kann.

## Verwendete Referenzen

- Graph-Berechtigungen: https://learn.microsoft.com/en-us/graph/permissions-reference
- Calendar View: https://learn.microsoft.com/en-us/graph/api/user-list-calendarview
- MSAL Browser initialisieren: https://learn.microsoft.com/en-us/entra/msal/javascript/browser/initialization
- MSAL CDN: https://learn.microsoft.com/en-us/entra/msal/javascript/browser/cdn-usage
- Tailwind Production: https://v3.tailwindcss.com/docs/optimizing-for-production
