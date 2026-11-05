const msalConfig = {
            auth: {
                clientId: "cee7a128-4c3e-4a07-b0c8-89b97475a71c",
                authority: "https://login.microsoftonline.com/7dfd912b-740c-430b-ae2b-10a0948f1378",
                redirectUri: window.location.origin + window.location.pathname
            },
            cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false }
        };

        const loginRequest = { scopes: ["User.Read", "Calendars.ReadBasic", "User.ReadBasic.All"] };
        const msalInstance = window.msal?.PublicClientApplication ? new msal.PublicClientApplication(msalConfig) : null;
        let currentUserAccount = null;
        let fetchedUsersList = [];
        let graphRequestController = null;
        let sessionEpoch = 0;

        function isSafariOrWebKit() {
            const ua = navigator.userAgent.toLowerCase();
            return ua.includes('safari') && !ua.includes('chrome') && !ua.includes('android');
        }

        function getActiveAccount() {
            const accounts = msalInstance?.getAllAccounts() || [];
            if (accounts.length === 0) return null;
            return msalInstance.getActiveAccount() || (accounts.length === 1 ? accounts[0] : null);
        }

        async function loginM365() {
            try {
                if (isSafariOrWebKit()) {
                    await msalInstance.loginRedirect(loginRequest);
                    return;
                }
                const response = await msalInstance.loginPopup(loginRequest);
                currentUserAccount = response.account || getActiveAccount();
                if (currentUserAccount) {
                    msalInstance?.setActiveAccount(currentUserAccount);
                    updateUserUI(currentUserAccount);
                    await loadGraphData();
                }
            } catch (err) {
                const code = String(err?.errorCode || '');
                // A denied consent, cancellation or authorization failure must not trigger a second login.
                if (['popup_window_error', 'empty_window_error', 'popup_window_timeout'].includes(code)) {
                    await msalInstance.loginRedirect(loginRequest);
                    return;
                }
                console.warn('M365-Anmeldung fehlgeschlagen:', code);
                showLoginMessage('Microsoft-365-Anmeldung fehlgeschlagen oder abgebrochen.');
            }
        }

        function showLoginMessage(message) {
            const el = document.getElementById('sync-status');
            if (el) el.textContent = message;
        }

        function resetPrivateViews() {
            sessionEpoch += 1;
            if (graphRequestController) graphRequestController.abort();
            graphRequestController = null;
            fetchedUsersList = [];
            currentUserAccount = null;
            msalInstance?.setActiveAccount(null);
            const search = document.getElementById('employee-search');
            if (search) search.value = '';
            document.getElementById('user-profile-badge').classList.add('hidden');
            document.getElementById('login-btn').classList.remove('hidden');
            document.getElementById('sync-status').textContent = 'Graph API: Abgemeldet';
            for (const id of ['today-calendar-events', 'upcoming-calendar-events']) {
                const node = document.getElementById(id);
                node.replaceChildren();
                const div = document.createElement('div');
                div.className = 'p-2.5 bg-[var(--vantage-panel)] rounded-sm border border-[var(--vantage-border)] text-slate-500 italic';
                div.textContent = 'Bitte bei M365 anmelden, um Termine zu laden.';
                node.appendChild(div);
            }
            const presenceContainer = document.getElementById('sidebar-presence-list');
            presenceContainer.textContent = 'Bitte bei M365 anmelden, um das Mitarbeiterverzeichnis zu laden.';
        }

        async function logoutM365() {
            const account = currentUserAccount || getActiveAccount();
            resetPrivateViews();
            try {
                if (account) await msalInstance.logoutPopup({ account, mainWindowRedirectUri: window.location.href });
            } catch (err) {
                console.warn('SSO-Logout nicht vollständig:', err?.errorCode || 'Unbekannter Fehler');
                // Clear local authentication material even if the Microsoft logout popup was blocked.
                try { await msalInstance.clearCache(); } catch { /* older MSAL versions differ */ }
            }
        }

        function updateUserUI(account) {
            document.getElementById('login-btn').classList.add('hidden');
            const badge = document.getElementById('user-profile-badge');
            badge.classList.remove('hidden');
            document.getElementById('user-display-name').textContent = account.name || account.username;
        }

        async function getAccessToken() {
            const account = getActiveAccount();
            if (!account) return null;
            try {
                const result = await msalInstance.acquireTokenSilent({ scopes: loginRequest.scopes, account });
                return result.accessToken;
            } catch (err) {
                const code = String(err?.errorCode || '');
                // Only start an interactive flow if Microsoft explicitly requests interaction.
                const interactionRequired = err instanceof msal.InteractionRequiredAuthError ||
                    ['interaction_required', 'consent_required', 'login_required'].includes(code);
                if (!interactionRequired) throw err;
                if (isSafariOrWebKit()) {
                    await msalInstance.acquireTokenRedirect({ scopes: loginRequest.scopes, account });
                    return null;
                }
                const result = await msalInstance.acquireTokenPopup({ scopes: loginRequest.scopes, account });
                return result.accessToken;
            }
        }

        async function loadGraphData() {
            try {
                const token = await getAccessToken();
                if (!token || !getActiveAccount()) return;
                if (graphRequestController) graphRequestController.abort();
                graphRequestController = new AbortController();
                const epoch = ++sessionEpoch;
                const signal = graphRequestController.signal;
                await Promise.all([
                    fetchOutlookCalendars(token, signal, epoch),
                    fetchDirectoryUsers(token, signal, epoch)
                ]);
            } catch (err) {
                if (err?.name !== 'AbortError') {
                    console.error('Graph-Daten konnten nicht geladen werden:', err?.message || err);
                    showLoginMessage('Graph API: Anmeldung oder Datenabruf fehlgeschlagen');
                }
            }
        }

        function isCurrentRequest(epoch, signal) {
            return epoch === sessionEpoch && !signal?.aborted && Boolean(currentUserAccount);
        }

        async function fetchDirectoryUsers(token, signal, epoch) {
            const presenceContainer = document.getElementById('sidebar-presence-list');
            try {
                // User.ReadBasic.All: use only basic profile fields; no fabricated online status.
                let next = 'https://graph.microsoft.com/v1.0/users?$top=100&$select=id,displayName,mail';
                const users = [];
                for (let page = 0; page < 5 && next; page++) {
                    const url = new URL(next);
                    if (url.protocol !== 'https:' || url.hostname !== 'graph.microsoft.com' || !url.pathname.startsWith('/v1.0/users')) {
                        throw new Error('Unerwartete Graph-Paginierungs-URL');
                    }
                    const res = await fetch(url.toString(), { headers: { Authorization: `Bearer ${token}` }, signal });
                    if (!res.ok) throw new Error(`Benutzerverzeichnis: HTTP ${res.status}`);
                    const data = await res.json();
                    if (!isCurrentRequest(epoch, signal)) return;
                    if (Array.isArray(data.value)) users.push(...data.value);
                    next = typeof data['@odata.nextLink'] === 'string' ? data['@odata.nextLink'] : null;
                }
                if (!isCurrentRequest(epoch, signal)) return;
                fetchedUsersList = users;
                renderDirectoryUsers(users);
            } catch (err) {
                if (err?.name === 'AbortError' || !isCurrentRequest(epoch, signal)) return;
                console.warn('Mitarbeiterverzeichnis:', err?.message || err);
                presenceContainer.textContent = 'Mitarbeiterverzeichnis konnte nicht geladen werden.';
            }
        }

        function renderDirectoryUsers(users) {
            const container = document.getElementById('sidebar-presence-list');
            container.replaceChildren();
            if (!users.length) {
                container.textContent = 'Keine Mitarbeiter gefunden.';
                return;
            }
            const fragment = document.createDocumentFragment();
            for (const user of users) {
                const item = document.createElement('div');
                item.className = 'flex items-center justify-between p-2 rounded-sm bg-[var(--vantage-bg)] border border-[var(--vantage-border)] text-xs font-normal not-italic';
                const wrapper = document.createElement('div');
                wrapper.className = 'truncate';
                const name = document.createElement('div');
                name.className = 'font-semibold text-[var(--vantage-text-main)] truncate';
                name.textContent = user.displayName || 'Unbekannt';
                const mail = document.createElement('div');
                mail.className = 'text-[10px] text-[var(--vantage-text-muted)] truncate';
                mail.textContent = user.mail || 'M365 Benutzer';
                wrapper.append(name, mail);
                // Not a presence indicator: this endpoint provides no Teams availability.
                const indicator = document.createElement('span');
                indicator.className = 'w-2 h-2 rounded-full bg-slate-400 shrink-0';
                indicator.title = 'Kein Live-Präsenzstatus verfügbar';
                item.append(wrapper, indicator);
                fragment.appendChild(item);
            }
            container.appendChild(fragment);
        }

        function filterPresenceList() {
            if (!currentUserAccount) return;
            const query = document.getElementById('employee-search').value.toLocaleLowerCase();
            renderDirectoryUsers(fetchedUsersList.filter(u =>
                String(u.displayName || '').toLocaleLowerCase().includes(query)
            ));
        }

        function safeOutlookUrl(rawUrl) {
            try {
                const u = new URL(rawUrl);
                if (u.protocol === 'https:' &&
                    ['outlook.office.com', 'outlook.office365.com', 'outlook.office.de'].includes(u.hostname)) {
                    return u.href;
                }
            } catch { /* invalid or missing URL */ }
            return 'https://outlook.office.com/calendar/';
        }

        function calendarMessage(container, message, error = false) {
            container.replaceChildren();
            const element = document.createElement('div');
            element.className = error
                ? 'p-2.5 bg-rose-50 rounded-sm border border-rose-200 text-xs text-rose-600'
                : 'p-2.5 bg-[var(--vantage-panel)] rounded-sm border border-[var(--vantage-border)] text-xs text-slate-500';
            element.textContent = message;
            container.appendChild(element);
        }

        async function fetchOutlookCalendars(token, signal, epoch) {
            const todayContainer = document.getElementById('today-calendar-events');
            const upcomingContainer = document.getElementById('upcoming-calendar-events');
            const syncStatus = document.getElementById('sync-status');
            try {
                const now = new Date();
                const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
                const endOfThreeDays = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 3);
                const graphUrl = new URL('https://graph.microsoft.com/v1.0/me/calendarView');
                graphUrl.searchParams.set('startDateTime', startOfDay.toISOString());
                graphUrl.searchParams.set('endDateTime', endOfThreeDays.toISOString());
                graphUrl.searchParams.set('$top', '50');
                graphUrl.searchParams.set('$select', 'start,subject,webLink');
                graphUrl.searchParams.set('$orderby', 'start/dateTime');
                // Request UTC consistently; Graph returns dateTime without a timezone suffix.
                const res = await fetch(graphUrl.toString(), {
                    headers: { Authorization: `Bearer ${token}`, Prefer: 'outlook.timezone="UTC"' }, signal
                });
                if (!res.ok) throw new Error(`Kalender: HTTP ${res.status}`);
                const data = await res.json();
                if (!isCurrentRequest(epoch, signal)) return;
                syncStatus.textContent = 'Graph API Kalender: Aktiv';
                syncStatus.className = 'text-[10px] uppercase tracking-[1px] text-emerald-600 font-semibold bg-[var(--vantage-bg)] px-2 py-0.5 rounded-sm border border-[var(--vantage-border)]';
                todayContainer.replaceChildren();
                upcomingContainer.replaceChildren();
                let todayCount = 0, upcomingCount = 0;
                for (const evt of Array.isArray(data.value) ? data.value : []) {
                    const rawStart = String(evt.start?.dateTime || '');
                    const start = new Date(/(?:Z|[+-]\d\d:\d\d)$/i.test(rawStart) ? rawStart : `${rawStart}Z`);
                    if (!Number.isFinite(start.getTime())) continue;
                    const isToday = start.toDateString() === now.toDateString();
                    const item = document.createElement('div');
                    item.className = 'flex items-center justify-between p-2.5 bg-[var(--vantage-panel)] rounded-sm border border-[var(--vantage-border)] hover:border-[var(--vantage-gold)] transition cursor-pointer';
                    item.addEventListener('click', () => window.open(safeOutlookUrl(evt.webLink), '_blank', 'noopener,noreferrer'));
                    const line = document.createElement('span');
                    line.className = 'font-medium text-[var(--vantage-text-main)] truncate';
                    const hour = start.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                    const day = isToday ? '' : start.toLocaleDateString([], { day: '2-digit', month: '2-digit' }) + ', ';
                    line.textContent = `${day}${hour} - ${evt.subject || '(Ohne Titel)'}`;
                    item.appendChild(line);
                    (isToday ? todayContainer : upcomingContainer).appendChild(item);
                    if (isToday) todayCount++; else upcomingCount++;
                }
                if (!todayCount) calendarMessage(todayContainer, 'Keine Termine für heute.');
                if (!upcomingCount) calendarMessage(upcomingContainer, 'Keine Termine in der Vorschau.');
            } catch (err) {
                if (err?.name === 'AbortError' || !isCurrentRequest(epoch, signal)) return;
                syncStatus.textContent = 'Graph API Kalender: Fehler';
                calendarMessage(todayContainer, `Kalenderdaten konnten nicht geladen werden (${err?.message || 'Unbekannter Fehler'}).`, true);
                calendarMessage(upcomingContainer, 'Kalenderdaten konnten nicht geladen werden.', true);
            }
        }

        // SPRACHEN & FACHBEREICHE ÜBERSETZUNGSDATENBANK
        let currentLanguage = "de";
        const translations = {
            de: {
                teamsStatus: "M365 Verzeichnis",
                searchPlaceholder: "Mitarbeiter suchen...",
                login: "M365 Anmelden",
                bulletinBoard: "BULLETIN BOARD",
                bulletinSub: "Vantage News & Outlook Kalender",
                todayCalendar: "Heute im Kalender",
                upcomingCalendar: "Vorschau nächste Tage",
                vantageCalendar: "Persönlicher Kalender",
                quickAccess: "Schnellzugriffe",
                qaCases: "Leerkoffer",
                qaWiki: "Kamera Wiki",
                qaOptics: "Optik Service",
                qaSupport: "Support",
                qaDocs: "Dokumente",
                deptHeader: "Fachbereiche & System-Channels",
                btnTestbetreuung: "Testbetreuung (SharePoint)",
                btnWhoWithWhom: "WHO WITH WHOM Suche",
                btnBack: "Schritt zurück",
                btnHome: "Zum Dashboard",
                subCategoriesHeader: "Unterkategorien & Tools",
                detailSub: "Live Teams-Kommentare & Diskussionsverlauf",
                btnOpenTeams: "In Teams öffnen",
                chatHeader: "Diskussion (noch nicht angebunden)",
                chatInputPlaceholder: "Antwort schreiben...",
                btnSend: "Senden",
                footerText: "Vantage Intranet PWA V3.8.9 • MS365 & Teams Integration",
                
                deptTechRentalTitle: "Technical Departments Rental",
                deptTechRentalDesc: "Kamera, Mechanik, NLN & Werkstatt-Protokolle.",
                deptBookingTitle: "Booking Department & Client Contacts",
                deptBookingDesc: "Buchungen, Kundenkontakte & Verfügbarkeiten.",
                deptProductsTitle: "VANTAGE & HAWK Products",
                deptProductsDesc: "Eigenentwicklungen, HAWK Lenses & Datenblätter.",
                deptAccountingTitle: "Buchhaltung",
                deptAccountingDesc: "Rechnungen, Spesen & Finanzprozesse.",
                deptPreparationTitle: "Vorbereitung / Job Handling",
                deptPreparationDesc: "Job-Vorbereitung, Checklisten & Check-Outs.",
                deptItTitle: "IT / Software",
                deptItDesc: "Infrastruktur, Tools, Netzwerke & Support.",
                deptMarketingTitle: "Marketing",
                deptMarketingDesc: "Social Media, Events, Presse & CI-Materialien.",

                subFirmware: "Firmware Updates",
                subSoftwareUpdates: "Software Updates",
                subGroundglass: "Mattscheiben-Rechner S35",
                subCaseLists: "Kofferlisten (Leerkoffer)",
                subArriCatalog: "ARRI Ersatzteil-Katalog",
                subWorkshopProtocols: "Werkstatt Protokolle",
                subWhoWithWhom: "WHO WITH WHOM Suche",
                subPhotoStation: "Vantage PhotoStation",

                ggTitle: "Mattscheiben-Rechner S35",
                ggSub: "Berechnung von Zielformaten & Schnittabständen für Framing-Schablonen",
                ggCloseBtn: "Tool schließen",
                ggTypeLabel: "Mattscheibe Typ",
                ggAspectLabel: "Ziel-Seitenverhältnis (z. B. 1.78, 1.85, 2.39)",
                ggMaxW: "Max. Breite",
                ggMaxH: "Max. Höhe",
                ggTargetW: "Zielbreite",
                ggTargetH: "Zielhöhe",
                ggCenterDist: "Mitte nach oben/unten",
                ggEdgeDist: "Abstand Ober-/Unterkante",
                ggPrintBtn: "Schablone mit Grafik drucken",
                ggPreviewHint: "Visuelle Vorschau (Rot = Zielformat)",

                snakeTitle: "Vantage Snake 🕹️",
                snakeSub: "Benutze die Pfeiltasten zum Steuern!",
                snakeScore: "Score",
                snakeHighscore: "Highscore",
                snakeNewBtn: "Neues Spiel",
                snakeCloseBtn: "Schließen"
            },
            en: {
                teamsStatus: "M365 Directory",
                searchPlaceholder: "Search employee...",
                login: "Sign in to M365",
                bulletinBoard: "BULLETIN BOARD",
                bulletinSub: "Vantage News & Outlook Calendar",
                todayCalendar: "Today's Calendar",
                upcomingCalendar: "Upcoming Days",
                vantageCalendar: "Personal Calendar",
                quickAccess: "Quick Access",
                qaCases: "Empty Cases",
                qaWiki: "Camera Wiki",
                qaOptics: "Optics Service",
                qaSupport: "Support",
                qaDocs: "Documents",
                deptHeader: "Departments & System Channels",
                btnTestbetreuung: "Test Support (SharePoint)",
                btnWhoWithWhom: "WHO WITH WHOM Search",
                btnBack: "Step Back",
                btnHome: "To Dashboard",
                subCategoriesHeader: "Subcategories & Tools",
                detailSub: "Live Teams comments & discussion thread",
                btnOpenTeams: "Open in Teams",
                chatHeader: "Discussion (not connected)",
                chatInputPlaceholder: "Write a reply...",
                btnSend: "Send",
                footerText: "Vantage Intranet PWA V3.8.9 • MS365 & Teams Integration",

                deptTechRentalTitle: "Technical Departments Rental",
                deptTechRentalDesc: "Camera, mechanics, NLN & workshop protocols.",
                deptBookingTitle: "Booking Department & Client Contacts",
                deptBookingDesc: "Bookings, client contacts & availabilities.",
                deptProductsTitle: "VANTAGE & HAWK Products",
                deptProductsDesc: "Proprietary developments, HAWK Lenses & datasheets.",
                deptAccountingTitle: "Accounting",
                deptAccountingDesc: "Invoices, expenses & financial workflows.",
                deptPreparationTitle: "Preparation / Job Handling",
                deptPreparationDesc: "Job preparation, checklists & check-outs.",
                deptItTitle: "IT / Software",
                deptItDesc: "Infrastructure, tools, networks & support.",
                deptMarketingTitle: "Marketing",
                deptMarketingDesc: "Social media, events, PR & CI materials.",

                subFirmware: "Firmware Updates",
                subSoftwareUpdates: "Software Updates",
                subGroundglass: "Groundglass Calculator S35",
                subCaseLists: "Case Lists",
                subArriCatalog: "ARRI Spare Parts Catalog",
                subWorkshopProtocols: "Workshop Protocols",
                subWhoWithWhom: "WHO WITH WHOM Search",
                subPhotoStation: "Vantage PhotoStation",

                ggTitle: "Groundglass Calculator S35",
                ggSub: "Calculation of target formats & cutting distances for framing templates",
                ggCloseBtn: "Close Tool",
                ggTypeLabel: "Groundglass Type",
                ggAspectLabel: "Target Aspect Ratio (e.g. 1.78, 1.85, 2.39)",
                ggMaxW: "Max Width",
                ggMaxH: "Max Height",
                ggTargetW: "Target Width",
                ggTargetH: "Target Height",
                ggCenterDist: "Center to Top/Bottom",
                ggEdgeDist: "Top/Bottom Edge Distance",
                ggPrintBtn: "Print Template with Graphic",
                ggPreviewHint: "Visual Preview (Red = Target Format)",

                snakeTitle: "Vantage Snake 🕹️",
                snakeSub: "Use arrow keys to steer!",
                snakeScore: "Score",
                snakeHighscore: "Highscore",
                snakeNewBtn: "New Game",
                snakeCloseBtn: "Close"
            },
            fr: {
                teamsStatus: "Annuaire M365",
                searchPlaceholder: "Chercher employé...",
                login: "Connexion M365",
                bulletinBoard: "TABLEAU D’INFORMATION",
                bulletinSub: "Actualités Vantage & Calendrier Outlook",
                todayCalendar: "Agenda du jour",
                upcomingCalendar: "Prochains jours",
                vantageCalendar: "Calendrier personnel",
                quickAccess: "Accès rapides",
                qaCases: "Valises vides",
                qaWiki: "Wiki Caméra",
                qaOptics: "Service Optique",
                qaSupport: "Support",
                qaDocs: "Documents",
                deptHeader: "Services & Canaux Système",
                btnTestbetreuung: "Support de test (SharePoint)",
                btnWhoWithWhom: "Recherche WHO WITH WHOM",
                btnBack: "Retour",
                btnHome: "Au Tableau de bord",
                subCategoriesHeader: "Sous-catégories & Outils",
                detailSub: "Commentaires Teams en direct",
                btnOpenTeams: "Ouvrir dans Teams",
                chatHeader: "Fil de discussion",
                chatInputPlaceholder: "Écrire une réponse...",
                btnSend: "Envoyer",
                footerText: "Vantage Intranet PWA V3.8.9 • MS365 & Teams Integration",

                deptTechRentalTitle: "Technical Departments Rental",
                deptTechRentalDesc: "Caméra, mécanique, NLN & protocoles d'atelier.",
                deptBookingTitle: "Booking Department & Client Contacts",
                deptBookingDesc: "Réservations, contacts clients & disponibilités.",
                deptProductsTitle: "VANTAGE & HAWK Products",
                deptProductsDesc: "Développements, HAWK Lenses & fiches techniques.",
                deptAccountingTitle: "Accounting",
                deptAccountingDesc: "Factures, notes de frais & processus financiers.",
                deptPreparationTitle: "Preparation / Job Handling",
                deptPreparationDesc: "Préparation des jobs, checklists & check-outs.",
                deptItTitle: "IT / Software",
                deptItDesc: "Infrastructure, outils, réseaux & support.",
                deptMarketingTitle: "Marketing",
                deptMarketingDesc: "Réseaux sociaux, événements & supports CI.",

                subFirmware: "Mises à jour Firmware",
                subSoftwareUpdates: "Mises à jour Logiciel",
                subGroundglass: "Calculateur de dépoli S35",
                subCaseLists: "Listes de valises",
                subArriCatalog: "Catalogue de pièces ARRI",
                subWorkshopProtocols: "Protocoles d'atelier",
                subWhoWithWhom: "Recherche WHO WITH WHOM",
                subPhotoStation: "Vantage PhotoStation",

                ggTitle: "Calculateur de dépoli S35",
                ggSub: "Calcul des formats cibles & des distances de coupe pour les gabarits",
                ggCloseBtn: "Fermer l'outil",
                ggTypeLabel: "Type de dépoli",
                ggAspectLabel: "Format d'image cible (ex. 1.78, 1.85, 2.39)",
                ggMaxW: "Largeur max.",
                ggMaxH: "Hauteur max.",
                ggTargetW: "Largeur cible",
                ggTargetH: "Hauteur cible",
                ggCenterDist: "Centre vers haut/bas",
                ggEdgeDist: "Distance bord sup./inf.",
                ggPrintBtn: "Imprimer le gabarit avec graphique",
                ggPreviewHint: "Aperçu visuel (Rouge = Format cible)",

                snakeTitle: "Vantage Snake 🕹️",
                snakeSub: "Utilisez les flèches pour diriger !",
                snakeScore: "Score",
                snakeHighscore: "Meilleur score",
                snakeNewBtn: "Nouvelle partie",
                snakeCloseBtn: "Fermer"
            },
            cs: {
                teamsStatus: "Adresář M365",
                searchPlaceholder: "Hledat zaměstnance...",
                login: "Přihlásit M365",
                bulletinBoard: "NÁSTĚNKA",
                bulletinSub: "Novinky Vantage & Outlook Kalendář",
                todayCalendar: "Dnešní kalendář",
                upcomingCalendar: "Nadcházející dny",
                vantageCalendar: "Osobní kalendář",
                quickAccess: "Rychlý přístup",
                qaCases: "Prázdné kufry",
                qaWiki: "Kamera Wiki",
                qaOptics: "Servis optiky",
                qaSupport: "Podpora",
                qaDocs: "Dokumenty",
                deptHeader: "Oddělení & Systémové kanály",
                btnTestbetreuung: "Správa testů (SharePoint)",
                btnWhoWithWhom: "Vyhledávání WHO WITH WHOM",
                btnBack: "Zpět",
                btnHome: "Na nástěnku",
                subCategoriesHeader: "Podkategorie & Nástroje",
                detailSub: "Živé komentáře Teams",
                btnOpenTeams: "Otevřít v Teams",
                chatHeader: "Diskuze & Novinky",
                chatInputPlaceholder: "Napsat odpověď...",
                btnSend: "Odeslat",
                footerText: "Vantage Intranet PWA V3.8.9 • MS365 & Teams Integration",

                deptTechRentalTitle: "Technical Departments Rental",
                deptTechRentalDesc: "Kamera, mechanika, NLN & dílenské protokoly.",
                deptBookingTitle: "Booking Department & Client Contacts",
                deptBookingDesc: "Rezervace, kontakty na klienty & dostupnost.",
                deptProductsTitle: "VANTAGE & HAWK Products",
                deptProductsDesc: "Vlastní vývoj, HAWK Lenses & datové listy.",
                deptAccountingTitle: "Accounting",
                deptAccountingDesc: "Faktury, výdaje & finanční procesy.",
                deptPreparationTitle: "Preparation / Job Handling",
                deptPreparationDesc: "Příprava zakázek, kontrolní seznamy & check-outy.",
                deptItTitle: "IT / Software",
                deptItDesc: "Infrastruktur, nástroje, sítě & podpora.",
                deptMarketingTitle: "Marketing",
                deptMarketingDesc: "Sociální sítě, akce, PR & CI materiály.",

                subFirmware: "Aktualizace Firmware",
                subSoftwareUpdates: "Aktualizace Softwaru",
                subGroundglass: "Kalkulátor matnic S35",
                subCaseLists: "Seznamy kufrů",
                subArriCatalog: "Katalog náhradních dílů ARRI",
                subWorkshopProtocols: "Dílenské protokoly",
                subWhoWithWhom: "Vyhledávání WHO WITH WHOM",
                subPhotoStation: "Vantage PhotoStation",

                ggTitle: "Kalkulátor matnic S35",
                ggSub: "Výpočet cílových formátů & vzdáleností ořezu pro šablony rámečků",
                ggCloseBtn: "Zavřít nástroj",
                ggTypeLabel: "Typ matnice",
                ggAspectLabel: "Cílový poměr stran (např. 1.78, 1.85, 2.39)",
                ggMaxW: "Max. šířka",
                ggMaxH: "Max. výška",
                ggTargetW: "Cílová šířka",
                ggTargetH: "Cílová výška",
                ggCenterDist: "Střed nahoru/dolů",
                ggEdgeDist: "Vzdálenost horní/dolní okraj",
                ggPrintBtn: "Tisknout šablonu s grafikou",
                ggPreviewHint: "Vizuální náhled (Červená = Cílový formát)",

                snakeTitle: "Vantage Snake 🕹️",
                snakeSub: "Použijte šipky pro ovládání!",
                snakeScore: "Score",
                snakeHighscore: "Highscore",
                snakeNewBtn: "Neues Spiel",
                snakeCloseBtn: "Schließen"
            }
        };

        const departmentsConfig = [
            {
                id: "techRental", titleKey: "deptTechRentalTitle", descKey: "deptTechRentalDesc", icon: "camera",
                hasSharePointLink: true, sharePointUrl: "https://vantageweiden.sharepoint.com/sites/Testbetreuung2/Lists/Testbetreuung/AllItems.aspx",
                subcategories: [
                    { nameKey: "subFirmware", icon: "download-cloud" },
                    { nameKey: "subGroundglass", icon: "layout-template", isTool: true, toolType: "groundglass" },
                    { nameKey: "subCaseLists", icon: "box" },
                    { nameKey: "subArriCatalog", icon: "external-link", externalUrl: "https://arri.academy/SPC" },
                    { nameKey: "subWorkshopProtocols", icon: "clipboard-list" },
                    { nameKey: "subWhoWithWhom", icon: "external-link", externalUrl: "WhoWithWhom-Tool.html" }
                ]
            },
            {
                id: "booking", titleKey: "deptBookingTitle", descKey: "deptBookingDesc", icon: "calendar-check",
                hasSharePointLink: true, sharePointUrl: "https://vantageweiden.sharepoint.com/sites/Testbetreuung2/Lists/Testbetreuung/AllItems.aspx",
                subcategories: [
                    { nameKey: "subWhoWithWhom", icon: "external-link", externalUrl: "WhoWithWhom-Tool.html" }
                ]
            },
            { id: "products", titleKey: "deptProductsTitle", descKey: "deptProductsDesc", icon: "cpu", subcategories: [] },
            { id: "accounting", titleKey: "deptAccountingTitle", descKey: "deptAccountingDesc", icon: "receipt", subcategories: [] },
            { id: "preparation", titleKey: "deptPreparationTitle", descKey: "deptPreparationDesc", icon: "check-square", subcategories: [] },
            { 
                id: "it", titleKey: "deptItTitle", descKey: "deptItDesc", icon: "server",
                subcategories: [{ nameKey: "subSoftwareUpdates", icon: "shield" }] 
            },
            { 
                id: "marketing", titleKey: "deptMarketingTitle", descKey: "deptMarketingDesc", icon: "megaphone",
                subcategories: [{ nameKey: "subPhotoStation", icon: "image", externalUrl: "https://photostation.vantagefilm.com" }] 
            }
        ];

        function renderDepartments() {
            const grid = document.getElementById('department-grid');
            if (!grid) return;
            grid.innerHTML = "";
            const t = translations[currentLanguage];

            departmentsConfig.forEach(dept => {
                const title = t[dept.titleKey] || dept.titleKey;
                const desc = t[dept.descKey] || dept.descKey;

                const card = document.createElement('div');
                card.className = "bg-[var(--vantage-panel)] border border-[var(--vantage-border)] hover:border-[var(--vantage-gold)] rounded-sm p-5 transition-all duration-200 hover:shadow-md flex flex-col justify-between cursor-pointer group";
                card.onclick = () => openDepartment(dept.titleKey);
                
                card.innerHTML = `
                    <div>
                        <div class="flex items-center justify-between mb-3">
                            <div class="w-10 h-10 rounded-sm bg-[var(--vantage-bg)] border border-[var(--vantage-border)] text-[var(--vantage-gold)] flex items-center justify-center group-hover:scale-105 transition">
                                <i data-lucide="${dept.icon}" class="w-5 h-5"></i>
                            </div>
                        </div>
                        <h3 class="font-semibold text-[var(--vantage-text-main)] text-base mb-1 uppercase tracking-[0.5px] group-hover:text-[var(--vantage-gold-dark)] transition">${title}</h3>
                        <p class="text-xs text-[var(--vantage-text-muted)] mb-4 leading-relaxed">${desc}</p>
                    </div>
                `;
                grid.appendChild(card);
            });
            lucide.createIcons();
            updateTranslations();
        }

        function updateTranslations() {
            const lang = currentLanguage;
            document.querySelectorAll("[data-i18n]").forEach(el => {
                const key = el.dataset.i18n;
                if (translations[lang]?.[key]) el.textContent = translations[lang][key];
            });
            document.querySelectorAll("[data-i18n-placeholder]").forEach(el => {
                const key = el.dataset.i18nPlaceholder;
                if (translations[lang]?.[key]) el.placeholder = translations[lang][key];
            });
            document.documentElement.lang = lang;
        }

        function changeLanguage(lang) {
            currentLanguage = lang;
            localStorage.setItem("vantage-language", lang);
            renderDepartments();
            if (window._currentActiveDeptKey) {
                openDepartment(window._currentActiveDeptKey, true);
            }
            if (!document.getElementById('tool-view-container').classList.contains('hidden')) {
                loadGroundglassTool();
            }
        }

        function initializeLanguage() {
            const langSel = document.getElementById("language-select");
            const saved = localStorage.getItem("vantage-language");
            if (saved && translations[saved]) {
                currentLanguage = saved;
                langSel.value = saved;
            }
        }

        function toggleMobileSidebar() {
            const sidebar = document.getElementById('sidebar');
            const backdrop = document.getElementById('mobile-sidebar-backdrop');
            sidebar.classList.toggle('-translate-x-full');
            backdrop.classList.toggle('hidden');
        }

        window._currentActiveDeptKey = null;

        function openDepartment(departmentTitleKey, keepToolState = false) {
            window._currentActiveDeptKey = departmentTitleKey;
            document.getElementById('dashboard-view').classList.add('hidden');
            document.getElementById('detail-view').classList.remove('hidden');
            
            const t = translations[currentLanguage];
            const translatedTitle = t[departmentTitleKey] || departmentTitleKey;

            document.getElementById('detail-title').textContent = translatedTitle;
            document.getElementById('breadcrumb-path').textContent = translatedTitle;

            const actionContainer = document.getElementById('detail-action-buttons');
            actionContainer.innerHTML = "";

            const foundDept = departmentsConfig.find(d => d.titleKey === departmentTitleKey);
            if (foundDept && foundDept.hasSharePointLink) {
                actionContainer.innerHTML = `
                    <a href="${foundDept.sharePointUrl}" target="_blank" class="bg-[var(--vantage-gold)] hover:bg-[var(--vantage-gold-dark)] text-white px-3.5 py-2.5 rounded-sm text-xs font-semibold uppercase tracking-wider transition flex items-center gap-1.5 shadow-sm">
                        <i data-lucide="external-link" class="w-4 h-4"></i> <span data-i18n="btnTestbetreuung">${t.btnTestbetreuung}</span>
                    </a>
                `;
            }

            if (!keepToolState) closeEmbeddedTool();

            const subContainer = document.getElementById('subcategories-container');
            subContainer.innerHTML = "";
            const subcategories = foundDept ? foundDept.subcategories : [];

            subcategories.forEach(sub => {
                const subName = t[sub.nameKey] || sub.nameKey;
                const btn = document.createElement('div');
                btn.className = "flex items-center justify-between p-3 bg-[var(--vantage-bg)] hover:bg-[var(--vantage-panel)] border border-[var(--vantage-border)] hover:border-[var(--vantage-gold)] rounded-sm cursor-pointer transition group";
                
                if (sub.externalUrl) {
                    btn.onclick = () => window.open(sub.externalUrl, '_blank', 'noopener,noreferrer');
                } else if (sub.isTool && sub.toolType === "groundglass") {
                    btn.onclick = () => loadGroundglassTool();
                }

                btn.innerHTML = `
                    <span class="flex items-center gap-2.5 text-xs font-semibold text-[var(--vantage-text-main)] group-hover:text-[var(--vantage-gold-dark)] uppercase tracking-wide">
                        <i data-lucide="${sub.icon}" class="w-4 h-4 text-[var(--vantage-gold)]"></i> ${subName}
                    </span>
                    <i data-lucide="chevron-right" class="w-4 h-4 text-[var(--vantage-text-muted)] group-hover:text-[var(--vantage-gold)]"></i>
                `;
                subContainer.appendChild(btn);
            });
            lucide.createIcons();
        }

        function goHome() {
            window._currentActiveDeptKey = null;
            document.getElementById('detail-view').classList.add('hidden');
            document.getElementById('dashboard-view').classList.remove('hidden');
            closeEmbeddedTool();
        }

        function goBack() { goHome(); }

        function closeEmbeddedTool() {
            document.getElementById('tool-view-container').classList.add('hidden');
            document.getElementById('tool-view-container').innerHTML = "";
            document.getElementById('chat-stream-card').classList.remove('hidden');
        }

        // =========================================================================
        // MATTSCHEIBEN-RECHNER S35
        // =========================================================================
        function loadGroundglassTool() {
            document.getElementById('chat-stream-card').classList.add('hidden');
            const toolContainer = document.getElementById('tool-view-container');
            toolContainer.classList.remove('hidden');
            const t = translations[currentLanguage];

            toolContainer.innerHTML = `
                <div class="bg-[var(--vantage-dark-panel)] text-slate-100 rounded-sm p-6 border border-[var(--vantage-border)] shadow-lg space-y-6 no-print">
                    <div class="flex items-center justify-between border-b border-slate-700 pb-4">
                        <div>
                            <h3 class="text-base font-semibold text-[var(--vantage-gold)] uppercase tracking-wider">${t.ggTitle}</h3>
                            <p class="text-xs text-slate-400">${t.ggSub}</p>
                        </div>
                        <button data-action="close-tool" class="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-sm text-xs font-semibold uppercase tracking-wider transition">${t.ggCloseBtn}</button>
                    </div>

                    <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <div class="space-y-4">
                            <div>
                                <label class="block text-xs font-semibold text-slate-400 mb-1 uppercase tracking-wider">${t.ggTypeLabel}</label>
                                <select id="gg-mattscheibe" data-change-action="calculate-gg" class="w-full p-2.5 bg-slate-950 border border-slate-700 rounded-sm text-xs text-white focus:border-[var(--vantage-gold)] outline-none"></select>
                            </div>
                            <div>
                                <label class="block text-xs font-semibold text-slate-400 mb-1 uppercase tracking-wider">${t.ggAspectLabel}</label>
                                <input type="number" id="gg-aspect" step="0.01" value="" placeholder="z. B. 1.78" data-input-action="calculate-gg" class="w-full p-2.5 bg-slate-950 border border-slate-700 rounded-sm text-xs text-white focus:border-[var(--vantage-gold)] outline-none">
                            </div>

                            <div class="grid grid-cols-2 gap-2 pt-2">
                                <div class="bg-slate-950 p-3 rounded-sm border border-slate-800"><div class="text-[10px] text-slate-400 uppercase tracking-wider">Max. Breite</div><div id="gg-res-breite" class="text-sm font-bold text-white mt-1">-</div></div>
                                <div class="bg-slate-950 p-3 rounded-sm border border-slate-800"><div class="text-[10px] text-slate-400 uppercase tracking-wider">Max. Höhe</div><div id="gg-res-hoehe" class="text-sm font-bold text-white mt-1">-</div></div>
                                <div class="bg-slate-950 p-3 rounded-sm border border-slate-800"><div class="text-[10px] text-slate-400 uppercase tracking-wider">Zielbreite</div><div id="gg-res-zielbreite" class="text-sm font-bold text-white mt-1">-</div></div>
                                <div class="bg-slate-950 p-3 rounded-sm border border-slate-800"><div class="text-[10px] text-slate-400 uppercase tracking-wider">Zielhöhe</div><div id="gg-res-zielhoehe" class="text-sm font-bold text-white mt-1">-</div></div>
                                <div class="bg-slate-950 p-3 rounded-sm border border-slate-800"><div class="text-[10px] text-slate-400 uppercase tracking-wider">Mitte Abstand</div><div id="gg-res-mitte" class="text-sm font-bold text-white mt-1">-</div></div>
                                <div class="bg-slate-950 p-3 rounded-sm border border-slate-800"><div class="text-[10px] text-slate-400 uppercase tracking-wider">Rand Abstand</div><div id="gg-res-kante" class="text-sm font-bold text-white mt-1">-</div></div>
                            </div>

                            <div class="pt-2 flex justify-end">
                                <button data-action="print" class="bg-[var(--vantage-gold)] hover:bg-[var(--vantage-gold-dark)] text-white font-bold px-4 py-2 rounded-sm text-xs uppercase tracking-wider transition flex items-center gap-2 cursor-pointer">
                                    <i data-lucide="printer" class="w-4 h-4"></i> ${t.ggPrintBtn}
                                </button>
                            </div>
                        </div>

                        <div class="flex flex-col items-center justify-center bg-slate-950 p-4 rounded-sm border border-slate-800">
                            <canvas id="gg-previewCanvas" width="400" height="300" class="bg-black rounded-sm border border-slate-800 max-w-full h-auto"></canvas>
                            <span class="text-[10px] text-slate-500 mt-2 uppercase tracking-wide">${t.ggPreviewHint}</span>
                        </div>
                    </div>
                </div>

                <!-- DRUCK ANSICHT -->
                <div class="print-only print-template">
                    <h1>Druck-Schablone & Grafische Vorschau</h1>
                    <table class="print-table">
                        <thead><tr><th>Parameter</th><th>Spezifikation / Wert</th></tr></thead>
                        <tbody>
                            <tr><td>Gewählte Mattscheibe</td><td id="print-mattscheibe">-</td></tr>
                            <tr><td>Max. Breite × Höhe</td><td><span id="print-breite">-</span> mm &times; <span id="print-hoehe">-</span> mm</td></tr>
                            <tr><td>Ziel-Seitenverhältnis</td><td id="print-aspect">-</td></tr>
                            <tr><td>Zielhöhe (Schnittfenster)</td><td><strong id="print-zielhoehe">-</strong> mm</td></tr>
                        </tbody>
                    </table>
                    <div class="print-preview-section">
                        <canvas id="gg-printCanvas" width="280" height="200" class="print-preview-canvas"></canvas>
                    </div>
                </div>
            `;
            lucide.createIcons();
            initGGSelect();
        }

        const mattscheibenData = [
            { name: "-- Bitte auswählen --", breite: 0, hoehe: 0 },
            { name: "ANSI S35 Silent 1.33", breite: 24.0, hoehe: 18.00 },
            { name: "ANSI S35 TV 1.33 safe", breite: 23.0, hoehe: 17.25 },
            { name: "ANSI S35 TV 1.78 transmitted", breite: 24.0, hoehe: 13.50 },
            { name: "ANSI S35 1.85", breite: 24.0, hoehe: 13.00 },
            { name: "ANSI S35 2.35 centric", breite: 24.0, hoehe: 10.20 }
        ];

        function initGGSelect() {
            const select = document.getElementById("gg-mattscheibe");
            if (!select) return;
            select.innerHTML = "";
            mattscheibenData.forEach((item, index) => {
                const opt = document.createElement("option");
                opt.value = index;
                opt.textContent = item.name;
                select.appendChild(opt);
            });
            select.selectedIndex = 0;
        }

        function calculateGG() {
            const select = document.getElementById("gg-mattscheibe");
            if (!select || select.value == "0") return;
            const selectedItem = mattscheibenData[select.value];
            const targetAR = parseFloat(document.getElementById("gg-aspect").value) || 0;
            if (!selectedItem || targetAR <= 0) return;

            const breite = selectedItem.breite;
            const hoehe = selectedItem.hoehe;
            const zielbreite = breite;
            const zielhoehe = zielbreite / targetAR;
            const abstandMitte = zielhoehe / 2;
            const abstandKante = (hoehe - zielhoehe) / 2;

            document.getElementById("gg-res-breite").textContent = breite.toFixed(2) + " mm";
            document.getElementById("gg-res-hoehe").textContent = hoehe.toFixed(2) + " mm";
            document.getElementById("gg-res-zielbreite").textContent = zielbreite.toFixed(2) + " mm";
            document.getElementById("gg-res-zielhoehe").textContent = zielhoehe.toFixed(2) + " mm";
            document.getElementById("gg-res-mitte").textContent = abstandMitte.toFixed(2) + " mm";
            document.getElementById("gg-res-kante").textContent = abstandKante.toFixed(2) + " mm";

            document.getElementById("print-mattscheibe").textContent = selectedItem.name;
            document.getElementById("print-breite").textContent = breite.toFixed(2);
            document.getElementById("print-hoehe").textContent = hoehe.toFixed(2);
            document.getElementById("print-aspect").textContent = targetAR.toFixed(2);
            document.getElementById("print-zielhoehe").textContent = zielhoehe.toFixed(2);

            drawGGCanvas("gg-previewCanvas", breite, hoehe, zielbreite, zielhoehe, 400, 300, false);
            drawGGCanvas("gg-printCanvas", breite, hoehe, zielbreite, zielhoehe, 280, 200, true);
        }

        function drawGGCanvas(canvasId, sensorW, sensorH, targetW, targetH, width, height, isPrint) {
            const canvas = document.getElementById(canvasId);
            if (!canvas) return;
            const ctx = canvas.getContext("2d");
            ctx.clearRect(0, 0, width, height);
            if (isPrint) { ctx.fillStyle = "#ffffff"; ctx.fillRect(0, 0, width, height); }

            const margin = 25;
            const scale = Math.min((width - margin * 2) / sensorW, (height - margin * 2) / sensorH);
            const drawSensorW = sensorW * scale, drawSensorH = sensorH * scale;
            const startX = (width - drawSensorW) / 2, startY = (height - drawSensorH) / 2;

            ctx.strokeStyle = isPrint ? "#000000" : "#94a3b8";
            ctx.lineWidth = 2;
            ctx.strokeRect(startX, startY, drawSensorW, drawSensorH);

            if (targetW > 0 && targetH > 0) {
                const drawTargetW = targetW * scale, drawTargetH = targetH * scale;
                ctx.strokeStyle = "#dc2626";
                ctx.lineWidth = 2;
                ctx.strokeRect(startX + (drawSensorW - drawTargetW) / 2, startY + (drawSensorH - drawTargetH) / 2, drawTargetW, drawTargetH);
            }
        }

        // =========================================================================
        // EASTER EGG LOGIC (Vantage Snake)
        // =========================================================================
        let konamiCode = ['ArrowUp', 'ArrowUp', 'ArrowDown', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'ArrowLeft', 'ArrowRight', 'b', 'a'];
        let konamiIndex = 0;
        let logoClickCount = 0;
        let logoClickTimer = null;

        document.addEventListener('keydown', (e) => {
            if (e.key === konamiCode[konamiIndex]) {
                konamiIndex++;
                if (konamiIndex === konamiCode.length) { openSnakeGame(); konamiIndex = 0; }
            } else { konamiIndex = 0; }
        });

        function triggerEasterEggClick() {
            logoClickCount++;
            clearTimeout(logoClickTimer);
            if (logoClickCount === 3) { openSnakeGame(); logoClickCount = 0; }
            else { logoClickTimer = setTimeout(() => { logoClickCount = 0; }, 800); }
        }

        let snakeCanvas, snakeCtx;
        let snake = [], food = {};
        let snakeDir = 'RIGHT', nextDir = 'RIGHT';
        let snakeScore = 0, snakeHighscore = 0;
        let snakeGameLoop = null;
        const gridSize = 15, tileCount = 20;

        function openSnakeGame() {
            document.getElementById('easter-egg-modal').classList.remove('hidden');
            snakeCanvas = document.getElementById('snakeCanvas');
            snakeCtx = snakeCanvas.getContext('2d');
            startSnakeGame();
            document.addEventListener('keydown', handleSnakeControls);
        }

        function closeSnakeGame() {
            document.getElementById('easter-egg-modal').classList.add('hidden');
            clearInterval(snakeGameLoop);
            document.removeEventListener('keydown', handleSnakeControls);
        }

        function startSnakeGame() {
            clearInterval(snakeGameLoop);
            snake = [{ x: 10, y: 10 }, { x: 9, y: 10 }, { x: 8, y: 10 }];
            snakeDir = 'RIGHT'; nextDir = 'RIGHT';
            snakeScore = 0;
            document.getElementById('snake-score').textContent = snakeScore;
            spawnFood();
            snakeGameLoop = setInterval(updateSnake, 140);
        }

        function spawnFood() {
            food = { x: Math.floor(Math.random() * tileCount), y: Math.floor(Math.random() * tileCount) };
        }

        function handleSnakeControls(e) {
            if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) e.preventDefault();
            if (e.key === 'ArrowUp' && snakeDir !== 'DOWN') nextDir = 'UP';
            if (e.key === 'ArrowDown' && snakeDir !== 'UP') nextDir = 'DOWN';
            if (e.key === 'ArrowLeft' && snakeDir !== 'RIGHT') nextDir = 'LEFT';
            if (e.key === 'ArrowRight' && snakeDir !== 'LEFT') nextDir = 'RIGHT';
        }

        function updateSnake() {
            snakeDir = nextDir;
            let head = { ...snake[0] };
            if (snakeDir === 'UP') head.y--;
            if (snakeDir === 'DOWN') head.y++;
            if (snakeDir === 'LEFT') head.x--;
            if (snakeDir === 'RIGHT') head.x++;

            if (head.x < 0 || head.x >= tileCount || head.y < 0 || head.y >= tileCount || snake.some(s => s.x === head.x && s.y === head.y)) {
                clearInterval(snakeGameLoop);
                if (snakeScore > snakeHighscore) {
                    snakeHighscore = snakeScore;
                    document.getElementById('snake-highscore').textContent = snakeHighscore;
                }
                return;
            }

            snake.unshift(head);
            if (head.x === food.x && head.y === food.y) {
                snakeScore += 10;
                document.getElementById('snake-score').textContent = snakeScore;
                spawnFood();
            } else { snake.pop(); }

            snakeCtx.fillStyle = '#000000';
            snakeCtx.fillRect(0, 0, snakeCanvas.width, snakeCanvas.height);
            snakeCtx.fillStyle = '#B8923B';
            snakeCtx.fillRect(food.x * gridSize, food.y * gridSize, gridSize - 1, gridSize - 1);
            snake.forEach((seg, i) => {
                snakeCtx.fillStyle = i === 0 ? '#10b981' : '#ffffff';
                snakeCtx.fillRect(seg.x * gridSize, seg.y * gridSize, gridSize - 1, gridSize - 1);
            });
        }


        // CSP-safe delegated actions instead of inline event attributes.
        document.addEventListener('click', (event) => {
            const element = event.target.closest('[data-action]');
            if (!element || element.disabled) return;
            switch (element.dataset.action) {
                case 'toggle-sidebar': toggleMobileSidebar(); break;
                case 'easter-egg': triggerEasterEggClick(); break;
                case 'go-back': goBack(); break;
                case 'go-home': goHome(); break;
                case 'login': loginM365(); break;
                case 'logout': logoutM365(); break;
                case 'close-snake': closeSnakeGame(); break;
                case 'start-snake': startSnakeGame(); break;
                case 'close-tool': closeEmbeddedTool(); break;
                case 'print': window.print(); break;
            }
        });
        document.addEventListener('change', (event) => {
            if (event.target.dataset?.changeAction === 'change-language') changeLanguage(event.target.value);
            if (event.target.dataset?.changeAction === 'calculate-gg') calculateGG();
        });
        document.addEventListener('input', (event) => {
            if (event.target.dataset?.inputAction === 'filter-presence') filterPresenceList();
            if (event.target.dataset?.inputAction === 'calculate-gg') calculateGG();
        });

        window.addEventListener("DOMContentLoaded", async () => {
            initializeLanguage();
            renderDepartments();
            lucide.createIcons();

            try {
                if (!msalInstance) {
                    showLoginMessage('Authentifizierungsbibliothek nicht geladen. Netzwerk oder CSP prüfen.');
                    return;
                }
                const response = await msalInstance.handleRedirectPromise();
                currentUserAccount = response ? response.account : getActiveAccount();
                if (currentUserAccount) {
                    msalInstance?.setActiveAccount(currentUserAccount);
                    updateUserUI(currentUserAccount);
                    await loadGraphData();
                }
            } catch (err) { console.warn("MSAL Init Error:", err); }
        });
