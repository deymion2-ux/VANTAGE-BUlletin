const msalConfig = {
    auth: {
        clientId: "cee7a128-4c3e-4a07-b0c8-89b97475a71c",
        authority: "https://login.microsoftonline.com/7dfd912b-740c-430b-ae2b-10a0948f1378",
        redirectUri: new URL("auth-redirect.html", window.location.href).href
    },
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false }
};

const loginRequest = { scopes: ["User.Read", "Calendars.Read"] };
const msalInstance = new msal.PublicClientApplication(msalConfig);
let currentUserAccount = null;

function getActiveAccount() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) return null;
    return msalInstance.getActiveAccount() || accounts[0];
}

function setCalendarMessage(container, text, isError = false) {
    if (!container) return;
    const message = document.createElement("div");
    message.className = isError
        ? "p-2.5 bg-rose-50 rounded-sm border border-rose-200 text-xs text-rose-600"
        : "p-2.5 bg-[var(--vantage-panel)] rounded-sm border border-[var(--vantage-border)] text-xs text-slate-500";
    message.textContent = text;
    container.replaceChildren(message);
}

function setCalendarSyncStatus(text, isError = false) {
    const syncStatus = document.getElementById("sync-status");
    if (!syncStatus) return;
    syncStatus.textContent = text;
    syncStatus.className = isError
        ? "text-[10px] uppercase tracking-[1px] text-rose-600 font-semibold bg-[var(--vantage-bg)] px-2 py-0.5 rounded-sm border border-[var(--vantage-border)]"
        : "text-[10px] uppercase tracking-[1px] text-emerald-600 font-semibold bg-[var(--vantage-bg)] px-2 py-0.5 rounded-sm border border-[var(--vantage-border)]";
}

async function loginM365() {
    try {
        const loginResult = await msalInstance.loginPopup(loginRequest);
        currentUserAccount = loginResult?.account || null;
        if (!currentUserAccount) throw new Error("Microsoft hat kein Benutzerkonto zurückgegeben.");
        msalInstance.setActiveAccount(currentUserAccount);
        updateUserUI(currentUserAccount);
        await loadGraphData();
    } catch (err) {
        console.error("Login Error:", err);
        setCalendarSyncStatus("M365 Anmeldung fehlgeschlagen", true);
        setCalendarMessage(document.getElementById("today-calendar-events"), "Anmeldung fehlgeschlagen. Bitte erneut anmelden.", true);
        setCalendarMessage(document.getElementById("upcoming-calendar-events"), "Anmeldung fehlgeschlagen. Bitte erneut anmelden.", true);
    }
}

async function logoutM365() {
    const account = currentUserAccount || getActiveAccount();
    try {
        if (account) await msalInstance.logoutPopup({ account });
    } catch (err) {
        console.error("Logout Error:", err);
        msalInstance.clearCache();
    }
    currentUserAccount = null;
    msalInstance.setActiveAccount(null);
    document.getElementById("user-profile-badge")?.classList.add("hidden");
    document.getElementById("login-btn")?.classList.remove("hidden");
    const status = document.getElementById("sync-status");
    if (status) status.textContent = "SharePoint Sync: Abgemeldet";
    setCalendarMessage(document.getElementById("today-calendar-events"), "Bitte bei M365 anmelden, um Termine zu laden.");
    setCalendarMessage(document.getElementById("upcoming-calendar-events"), "Bitte bei M365 anmelden, um Termine zu laden.");
}

function updateUserUI(account) {
    document.getElementById("login-btn")?.classList.add("hidden");
    document.getElementById("user-profile-badge")?.classList.remove("hidden");
    const name = document.getElementById("user-display-name");
    if (name) name.textContent = account?.name || account?.username || "M365";
}

async function getAccessToken() {
    const activeAcc = currentUserAccount || getActiveAccount();
    if (!activeAcc) return null;
    try {
        const res = await msalInstance.acquireTokenSilent({ scopes: loginRequest.scopes, account: activeAcc });
        return res.accessToken;
    } catch (err) {
        if (err instanceof msal.InteractionRequiredAuthError) {
            const res = await msalInstance.acquireTokenPopup({ scopes: loginRequest.scopes, account: activeAcc });
            if (res.account) {
                currentUserAccount = res.account;
                msalInstance.setActiveAccount(res.account);
            }
            return res.accessToken;
        }
        throw err;
    }
}

async function loadGraphData() {
    try {
        const token = await getAccessToken();
        if (token) await fetchOutlookCalendars(token);
    } catch (err) {
        console.error("M365 data load error:", err);
        setCalendarSyncStatus("Graph API Sync: Fehler", true);
        setCalendarMessage(document.getElementById("today-calendar-events"), "Kalender konnte nicht geladen werden. Bitte erneut anmelden.", true);
        setCalendarMessage(document.getElementById("upcoming-calendar-events"), "Kalender konnte nicht geladen werden. Bitte erneut anmelden.", true);
    }
}

function createCalendarEventElement(evt, isToday, startDate) {
    const item = document.createElement("div");
    item.className = "flex items-center justify-between p-2.5 bg-[var(--vantage-panel)] rounded-sm border border-[var(--vantage-border)] hover:border-[var(--vantage-gold)] transition cursor-pointer";
    const time = startDate.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    const date = startDate.toLocaleDateString([], { day: "2-digit", month: "2-digit" });
    const left = document.createElement("div");
    left.className = "flex items-center gap-2 truncate mr-2";
    const dot = document.createElement("span");
    dot.className = "w-2 h-2 rounded-full bg-[var(--vantage-gold)] shrink-0";
    const title = document.createElement("span");
    title.className = "font-medium text-[var(--vantage-text-main)] truncate";
    title.textContent = `${isToday ? "" : `${date}, `}${time} - ${evt.subject || "(Kein Betreff)"}`;
    left.append(dot, title);
    item.appendChild(left);
    const location = evt.location?.displayName;
    if (location) {
        const badge = document.createElement("span");
        badge.className = "text-[10px] bg-[var(--vantage-bg)] text-[var(--vantage-text-muted)] px-2 py-0.5 rounded-sm shrink-0 border border-[var(--vantage-border)] truncate max-w-[120px]";
        badge.textContent = location;
        badge.title = location;
        item.appendChild(badge);
    }
    item.addEventListener("click", () => {
        let target = "https://outlook.office.com/calendar/";
        try {
            const url = new URL(evt.webLink);
            const allowedHosts = ["outlook.office.com", "outlook.office365.com", "outlook.office.de"];
            if (url.protocol === "https:" && allowedHosts.includes(url.hostname) && !url.username && !url.password && url.port === "") target = url.href;
        } catch { /* Invalid link: use the trusted Outlook calendar fallback. */ }
        window.open(target, "_blank", "noopener,noreferrer");
    });
    return item;
}

async function fetchOutlookCalendars(token) {
    const todayContainer = document.getElementById("today-calendar-events");
    const upcomingContainer = document.getElementById("upcoming-calendar-events");
    try {
        const now = new Date();
        const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate()).toISOString();
        const endOfThreeDays = new Date(now.getTime() + 72 * 60 * 60 * 1000).toISOString();
        const url = new URL("https://graph.microsoft.com/v1.0/me/calendarView");
        url.search = new URLSearchParams({ startDateTime: startOfDay, endDateTime: endOfThreeDays, "$top": "10", "$orderby": "start/dateTime" });
        const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
        if (!res.ok) throw new Error(`Graph API Fehler: ${res.status}`);
        const data = await res.json();
        if (!data || !Array.isArray(data.value)) throw new Error("Ungültige Kalenderantwort.");
        setCalendarSyncStatus("Graph API Sync: Aktiv");
        todayContainer?.replaceChildren();
        upcomingContainer?.replaceChildren();
        const todayStr = now.toDateString();
        let todayCount = 0, upcomingCount = 0;
        data.value.forEach(evt => {
            const startDate = new Date(evt?.start?.dateTime);
            if (!Number.isFinite(startDate.getTime())) return;
            const isToday = startDate.toDateString() === todayStr;
            const item = createCalendarEventElement(evt, isToday, startDate);
            if (isToday) { todayCount++; todayContainer?.appendChild(item); }
            else { upcomingCount++; upcomingContainer?.appendChild(item); }
        });
        if (!todayCount) setCalendarMessage(todayContainer, "Keine Termine für heute.");
        if (!upcomingCount) setCalendarMessage(upcomingContainer, "Keine Termine in der Vorschau.");
    } catch (err) {
        console.error("Calendar load error:", err);
        setCalendarSyncStatus("Graph API Sync: Fehler", true);
        // Do not expose untrusted API or exception text in the page.
        setCalendarMessage(todayContainer, "Kalender konnte nicht geladen werden. Bitte später erneut versuchen.", true);
        setCalendarMessage(upcomingContainer, "Kalender konnte nicht geladen werden. Bitte später erneut versuchen.", true);
    }
}

// SPRACHEN & FACHBEREICHE ÜBERSETZUNGSDATENBANK
let currentLanguage = "de";
const translations = {
    de: {
        teamsStatus: "Teams Status",
        searchPlaceholder: "Mitarbeiter suchen...",
        login: "M365 Anmelden",
        bulletinBoard: "BULLETIN BOARD",
        bulletinSub: "Vantage News & Outlook Kalender",
        todayCalendar: "Heute im Kalender",
        upcomingCalendar: "Vorschau nächste Tage",
        vantageCalendar: "Vantage Kalender",
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
        chatHeader: "Diskussionsverlauf & Neuigkeiten",
        chatInputPlaceholder: "Antwort schreiben...",
        btnSend: "Senden",
        footerText: "Vantage Intranet PWA V3.9.1 • MS365 & Teams Integration",
        
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
        teamsStatus: "Teams Presence",
        searchPlaceholder: "Search employee...",
        login: "Sign in to M365",
        bulletinBoard: "BULLETIN BOARD",
        bulletinSub: "Vantage News & Outlook Calendar",
        todayCalendar: "Today's Calendar",
        upcomingCalendar: "Upcoming Days",
        vantageCalendar: "Vantage Calendar",
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
        chatHeader: "Discussion Feed & Updates",
        chatInputPlaceholder: "Write a reply...",
        btnSend: "Send",
        footerText: "Vantage Intranet PWA V3.9.1 • MS365 & Teams Integration",

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
        teamsStatus: "Statut Teams",
        searchPlaceholder: "Chercher employé...",
        login: "Connexion M365",
        bulletinBoard: "TABLEAU D’INFORMATION",
        bulletinSub: "Actualités Vantage & Calendrier Outlook",
        todayCalendar: "Agenda du jour",
        upcomingCalendar: "Prochains jours",
        vantageCalendar: "Calendrier Vantage",
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
        footerText: "Vantage Intranet PWA V3.9.1 • MS365 & Teams Integration",

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
        teamsStatus: "Stav Teams",
        searchPlaceholder: "Hledat zaměstnance...",
        login: "Přihlásit M365",
        bulletinBoard: "NÁSTĚNKA",
        bulletinSub: "Novinky Vantage & Outlook Kalendář",
        todayCalendar: "Dnešní kalendář",
        upcomingCalendar: "Nadcházející dny",
        vantageCalendar: "Vantage Kalendář",
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
        footerText: "Vantage Intranet PWA V3.9.1 • MS365 & Teams Integration",

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
        snakeScore: "Skóre",
        snakeHighscore: "Nejlepší skóre",
        snakeNewBtn: "Nová hra",
        snakeCloseBtn: "Zavřít"
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
    sessionStorage.setItem("vantage-language", lang);
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
    const saved = sessionStorage.getItem("vantage-language");
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
                <button onclick="closeEmbeddedTool()" class="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-sm text-xs font-semibold uppercase tracking-wider transition">${t.ggCloseBtn}</button>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div class="space-y-4">
                    <div>
                        <label class="block text-xs font-semibold text-slate-400 mb-1 uppercase tracking-wider">${t.ggTypeLabel}</label>
                        <select id="gg-mattscheibe" onchange="calculateGG()" class="w-full p-2.5 bg-slate-950 border border-slate-700 rounded-sm text-xs text-white focus:border-[var(--vantage-gold)] outline-none"></select>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-slate-400 mb-1 uppercase tracking-wider">${t.ggAspectLabel}</label>
                        <input type="number" id="gg-aspect" step="0.01" value="" placeholder="z. B. 1.78" oninput="calculateGG()" class="w-full p-2.5 bg-slate-950 border border-slate-700 rounded-sm text-xs text-white focus:border-[var(--vantage-gold)] outline-none">
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
                        <button onclick="window.print()" class="bg-[var(--vantage-gold)] hover:bg-[var(--vantage-gold-dark)] text-white font-bold px-4 py-2 rounded-sm text-xs uppercase tracking-wider transition flex items-center gap-2 cursor-pointer">
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
            <div style="margin-top: 1cm;">
                <canvas id="gg-printCanvas" width="280" height="200" style="background:#fff; border:1px solid #000;"></canvas>
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
    const sensorAR = breite / hoehe;
    let zielbreite;
    let zielhoehe;
    if (targetAR >= sensorAR) {
        zielbreite = breite;
        zielhoehe = breite / targetAR;
    } else {
        zielhoehe = hoehe;
        zielbreite = hoehe * targetAR;
    }
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
    const printZielbreite = document.getElementById("print-zielbreite");
    if (printZielbreite) printZielbreite.textContent = zielbreite.toFixed(2);

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

window.addEventListener("DOMContentLoaded", async () => {
    initializeLanguage();
    renderDepartments();
    lucide.createIcons();

    try {
        const response = await msalInstance.handleRedirectPromise();
        currentUserAccount = response?.account || getActiveAccount();
        if (currentUserAccount) {
            msalInstance.setActiveAccount(currentUserAccount);
            updateUserUI(currentUserAccount);
            await loadGraphData();
        }
    } catch (err) {
        console.warn("MSAL Init Error:", err);
        setCalendarSyncStatus("M365 Initialisierung fehlgeschlagen", true);
    }
});

const INACTIVITY_TIMEOUT = 60 * 60 * 1000;
let inactivityTimer;
function resetInactivityTimer(){ 
    clearTimeout(inactivityTimer); 
    inactivityTimer = setTimeout(() => { 
        if(typeof logoutM365 === 'function'){ logoutM365(); } 
    }, INACTIVITY_TIMEOUT);
}
['mousemove','mousedown','keypress','scroll','touchstart'].forEach(e => document.addEventListener(e, resetInactivityTimer));
resetInactivityTimer();