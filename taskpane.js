/**
 * ════════════════════════════════════════════════════════════════
 * ScheduleHound — Outlook Add-in (Serverless)
 * Court Deadline Extractor & Calendar Creator
 *
 * This version calls Gemini directly from the client.
 * The API key is stored here — acceptable for internal/small-team use.
 * ════════════════════════════════════════════════════════════════
 */

// ┌──────────────────────────────────────────────────┐
// │  CONFIGURATION — EDIT THESE VALUES               │
// └──────────────────────────────────────────────────┘
const CONFIG = {
  // Your Gemini API key (get one at https://aistudio.google.com/apikey)
  GEMINI_API_KEY: "AIzaSyAn9-0Q8kgrq4Bv5rnsqkk1EaO_9Kf0v8c",

  // Gemini model to use
  GEMINI_MODEL: "gemini-2.5-pro",

  // Max file size in MB
  MAX_FILE_SIZE_MB: 20,

  // Default event color (Outlook category — set to null to skip)
  EVENT_CATEGORY: "Court Deadline",

  // Default reminders for timed events [minutes before]
  TIMED_EVENT_REMINDERS: [1440, 60],
};

// ┌──────────────────────────────────────────────────┐
// │  MODULE: ScheduleHound (SH namespace)            │
// └──────────────────────────────────────────────────┘
const SH = (() => {
  "use strict";

  // ── State ────────────────────────────────────────
  let fileBase64 = null;
  let currentFileName = null;
  let caseInfo = null;
  let deadlines = [];
  let rawTitles = [];
  let cardStates = [];
  let approvedCount = 0;
  let skippedCount = 0;
  let bulkProcessing = false;
  let allExpanded = false;

  // Calendar / Auth state
  let graphToken = null;       // Microsoft Graph access token
  let calendarMode = "graph";  // "graph" | "ics" (fallback)
  let availableCalendars = []; // fetched from Graph API

  // ══════════════════════════════════════════════════
  // ABBREVIATION MAPS (customize as needed)
  // ══════════════════════════════════════════════════
  const DEFENDANT_ABBREVIATIONS = [
    { match: "dr horton", abbr: "DRH" },
    { match: "d r horton", abbr: "DRH" },
    { match: "drhorton", abbr: "DRH" },
    // Add more: { match: "state farm", abbr: "State Farm" },
  ];

  const ORDER_ABBREVIATIONS = [
    { match: "case management plan", abbr: "CMP" },
    { match: "case management order", abbr: "CMO" },
    { match: "amended case management", abbr: "Amended CMO" },
    { match: "scheduling order", abbr: "SO" },
  ];

  const REMINDER_OPTIONS = [
    { label: "None", value: "none" },
    { label: "At time of event", value: "[0]" },
    { label: "15 min before", value: "[15]" },
    { label: "30 min before", value: "[30]" },
    { label: "1 hour before", value: "[60]" },
    { label: "1 day before", value: "[1440]" },
    { label: "1 day + 1 hr before", value: "[1440,60]" },
    { label: "2 days before", value: "[2880]" },
    { label: "1 week before", value: "[10080]" },
    { label: "1 week + 1 day", value: "[10080,1440]" },
  ];

  // ══════════════════════════════════════════════════
  // HELPERS
  // ══════════════════════════════════════════════════
  const $ = (id) => document.getElementById(id);
  const normalize = (s) => (s || "").toLowerCase().replace(/[.,]/g, "").replace(/\s+/g, " ").trim();
  const capitalize = (s) => s ? s.charAt(0).toUpperCase() + s.slice(1) : "";
  const escapeHtml = (s) => { const d = document.createElement("div"); d.textContent = s || ""; return d.innerHTML; };
  const escapeAttr = (s) => (s || "").replace(/"/g, "&quot;");

  function abbreviateDefendant(name) {
    if (!name) return "";
    const norm = normalize(name);
    for (const e of DEFENDANT_ABBREVIATIONS) { if (norm.includes(e.match)) return e.abbr; }
    const t = name.trim();
    if (/,\s*(inc|llc|corp|ltd)/i.test(t)) return t.replace(/,?\s*(inc\.?|llc\.?|corp\.?|ltd\.?)$/i, "").trim();
    if (!/\b(inc|llc|corp|ltd|company|group|city|county|state)\b/i.test(t)) {
      const w = t.split(/\s+/);
      if (w.length > 1) return w[w.length - 1];
    }
    return t;
  }

  function abbreviateOrderTitle(title) {
    if (!title) return "Order";
    const norm = normalize(title);
    for (const e of ORDER_ABBREVIATIONS) { if (norm.includes(e.match)) return e.abbr; }
    if (norm.includes("trial")) return "Trial Order";
    return "Order";
  }

  function formatOrderDateShort(d) {
    if (!d) return "";
    try { const p = d.split("-"); return `${parseInt(p[1], 10)}.${String(parseInt(p[2], 10)).padStart(2, "0")}.${p[0].slice(-2)}`; }
    catch { return d; }
  }

  function formatDateDisplay(d) {
    try { return new Date(d + "T00:00:00").toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric", year: "numeric" }); }
    catch { return d; }
  }

  function formatDateISO(d) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }

  function hasTime(dl) { return dl.time && dl.time !== "null" && dl.time !== ""; }
  function getDefaultReminder(dl) { return hasTime(dl) ? "[1440,60]" : "none"; }
  function getReminderLabel(dl) { return hasTime(dl) ? "🔔 1 day + 1 hr before" : "🔕 No reminder"; }

  // ══════════════════════════════════════════════════
  // TITLE FORMATTING
  // ══════════════════════════════════════════════════
  function formatTitle(rawTitle) {
    if (!caseInfo) return rawTitle;
    const fmt = $("titleFormat").value;
    if (fmt === "custom") return formatCustomTitle(rawTitle);

    const pl = caseInfo.plaintiffName || "";
    const def = abbreviateDefendant(caseInfo.defendantName || "");
    let prefix = "", suffix = "";

    switch (fmt) {
      case "plaintiff": prefix = pl; break;
      case "pvd": prefix = (pl && def) ? `${pl}/${def}` : (pl || def); break;
      case "pvd_order": {
        prefix = (pl && def) ? `${pl}/${def}` : (pl || def);
        const oa = abbreviateOrderTitle(caseInfo.orderTitle);
        const od = formatOrderDateShort(caseInfo.orderDate);
        if (oa && od) suffix = ` Per ${oa} ${od}`;
        else if (oa) suffix = ` Per ${oa}`;
        break;
      }
      case "casenum": prefix = caseInfo.caseNumber || ""; break;
    }
    let r = prefix ? `${prefix} – ${rawTitle}` : rawTitle;
    if (suffix) r += suffix;
    return r;
  }

  function formatCustomTitle(rawTitle) {
    const pattern = $("customFormatInput").value || "{event}";
    return pattern
      .replace(/\{plaintiff\}/gi, caseInfo?.plaintiffName || "")
      .replace(/\{defendant\}/gi, abbreviateDefendant(caseInfo?.defendantName || ""))
      .replace(/\{event\}/gi, rawTitle)
      .replace(/\{casenum\}/gi, caseInfo?.caseNumber || "")
      .replace(/\{order\}/gi, abbreviateOrderTitle(caseInfo?.orderTitle || ""))
      .replace(/\{orderdate\}/gi, formatOrderDateShort(caseInfo?.orderDate || ""));
  }

  function onTitleFormatChange() {
    const fmt = $("titleFormat").value;
    $("customFormatWrap").classList.toggle("visible", fmt === "custom");
    updateTitlePreview();
    for (let i = 0; i < rawTitles.length; i++) {
      const f = formatTitle(rawTitles[i]);
      const el = document.querySelector(`#card-${i} .dl-summary-title`);
      if (el) el.textContent = f;
      const ef = $(`edit-title-${i}`);
      if (ef) ef.value = f;
    }
  }

  function updateTitlePreview() {
    const fmt = $("titleFormat").value;
    const pl = caseInfo?.plaintiffName || "Smith";
    const def = caseInfo ? abbreviateDefendant(caseInfo.defendantName) || "Jones" : "DRH";
    const ord = caseInfo ? abbreviateOrderTitle(caseInfo.orderTitle) : "CMO";
    const dt = caseInfo?.orderDate ? formatOrderDateShort(caseInfo.orderDate) : "2.21.25";
    const cn = caseInfo?.caseNumber || "CV2401234";
    let ex = "";
    switch (fmt) {
      case "plaintiff": ex = `${pl} – Pretrial Conference`; break;
      case "pvd": ex = `${pl}/${def} – Pretrial Conference`; break;
      case "pvd_order": ex = `${pl}/${def} – Pretrial Conference Per ${ord} ${dt}`; break;
      case "casenum": ex = `${cn} – Pretrial Conference`; break;
      case "custom": {
        const p = $("customFormatInput").value || "{event}";
        ex = p.replace(/\{plaintiff\}/gi, pl).replace(/\{defendant\}/gi, def)
          .replace(/\{event\}/gi, "Pretrial Conference").replace(/\{casenum\}/gi, cn)
          .replace(/\{order\}/gi, ord).replace(/\{orderdate\}/gi, dt);
        break;
      }
    }
    $("titlePreview").textContent = ex;
  }

  function insertToken(tok) {
    const inp = $("customFormatInput");
    const pos = inp.selectionStart;
    inp.value = inp.value.slice(0, pos) + tok + inp.value.slice(pos);
    inp.focus();
    inp.selectionStart = inp.selectionEnd = pos + tok.length;
    updateTitlePreview();
  }

  // ══════════════════════════════════════════════════
  // OFFICE.JS INITIALIZATION
  // ══════════════════════════════════════════════════
  function initOffice() {
    Office.onReady((info) => {
      console.log("Office.js ready. Host:", info.host, "Platform:", info.platform);

      // Set reference date
      $("refDate").value = new Date().toISOString().split("T")[0];
      updateTitlePreview();

      // Try to get Graph token for calendar access
      authenticateGraph();
    });
  }

  // ══════════════════════════════════════════════════
  // MICROSOFT GRAPH AUTHENTICATION
  // ══════════════════════════════════════════════════
  async function authenticateGraph() {
    try {
      // Try SSO first (silent, no popup)
      const tokenResponse = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });

      // The SSO token is a bootstrap token — we need to exchange it
      // for a Graph token via the on-behalf-of flow, which requires a server.
      // For serverless: we use MSAL.js popup flow instead.
      // But first, try using the token directly with Graph (works in some configs)
      graphToken = tokenResponse;
      calendarMode = "graph";
      $("authBanner").classList.remove("visible");
      $("icsFallback").classList.remove("visible");
      loadCalendars();
    } catch (ssoError) {
      console.log("SSO failed, trying MSAL popup:", ssoError);
      try {
        await authenticateWithMSAL();
      } catch (msalError) {
        console.log("MSAL also failed, falling back to ICS:", msalError);
        setICSFallbackMode();
      }
    }
  }

  async function authenticateWithMSAL() {
    // Dynamically load MSAL if not already loaded
    if (typeof msal === "undefined") {
      await loadScript("https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js");
    }

    const msalConfig = {
      auth: {
        // Replace with your Azure AD app registration client ID
        // Register at: https://portal.azure.com → App registrations
        clientId: "YOUR_AZURE_CLIENT_ID_HERE",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin + "/taskpane.html",
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
      },
    };

    const msalInstance = new msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();

    const loginRequest = {
      scopes: ["Calendars.ReadWrite", "User.Read"],
    };

    try {
      // Try silent token acquisition first
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        const silentResult = await msalInstance.acquireTokenSilent({
          ...loginRequest,
          account: accounts[0],
        });
        graphToken = silentResult.accessToken;
      } else {
        // Interactive popup
        const popupResult = await msalInstance.acquireTokenPopup(loginRequest);
        graphToken = popupResult.accessToken;
      }
      calendarMode = "graph";
      $("authBanner").classList.remove("visible");
      $("icsFallback").classList.remove("visible");
      loadCalendars();
    } catch (e) {
      throw e;
    }
  }

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = src;
      s.onload = resolve;
      s.onerror = reject;
      document.head.appendChild(s);
    });
  }

  function setICSFallbackMode() {
    calendarMode = "ics";
    graphToken = null;
    $("authBanner").classList.remove("visible");
    $("icsFallback").classList.add("visible");

    // Show a simple "default calendar" option
    const cl = $("calendarList");
    cl.innerHTML = `
      <label class="cal-item">
        <input type="checkbox" checked value="ics-default"> Default Calendar (via .ics download)
        <span class="cal-tag">ICS</span>
      </label>`;
  }

  // ══════════════════════════════════════════════════
  // CALENDAR LOADING (Graph API)
  // ══════════════════════════════════════════════════
  async function loadCalendars() {
    if (calendarMode !== "graph" || !graphToken) {
      setICSFallbackMode();
      return;
    }

    try {
      const response = await fetch("https://graph.microsoft.com/v1.0/me/calendars", {
        headers: { Authorization: `Bearer ${graphToken}` },
      });

      if (!response.ok) {
        console.log("Calendar fetch failed:", response.status);
        setICSFallbackMode();
        return;
      }

      const data = await response.json();
      availableCalendars = data.value || [];
      renderCalendarList();
    } catch (e) {
      console.log("Calendar load error:", e);
      setICSFallbackMode();
    }
  }

  function renderCalendarList() {
    const cl = $("calendarList");
    cl.innerHTML = "";

    if (availableCalendars.length === 0) {
      cl.innerHTML = `<label class="cal-item"><input type="checkbox" checked value="primary"> Calendar</label>`;
      return;
    }

    // Group: owned vs shared
    const owned = availableCalendars.filter(c => c.isOwner !== false);
    const shared = availableCalendars.filter(c => c.isOwner === false);

    if (owned.length) {
      const gl = document.createElement("div");
      gl.className = "cal-group-label";
      gl.textContent = "My Calendars";
      cl.appendChild(gl);
      owned.forEach(cal => cl.appendChild(makeCalItem(cal, cal.isDefaultCalendar)));
    }
    if (shared.length) {
      const gl = document.createElement("div");
      gl.className = "cal-group-label";
      gl.textContent = "Shared Calendars";
      cl.appendChild(gl);
      shared.forEach(cal => cl.appendChild(makeCalItem(cal, false)));
    }
  }

  function makeCalItem(cal, isDefault) {
    const label = document.createElement("label");
    label.className = "cal-item";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = cal.id;
    if (isDefault) cb.checked = true;
    label.appendChild(cb);
    const span = document.createElement("span");
    span.textContent = cal.name;
    label.appendChild(span);
    if (isDefault) {
      const tag = document.createElement("span");
      tag.className = "cal-tag";
      tag.textContent = "Default";
      label.appendChild(tag);
    }
    return label;
  }

  function getSelectedCalendarIds() {
    return Array.from(document.querySelectorAll("#calendarList input:checked")).map(cb => cb.value);
  }

  // ══════════════════════════════════════════════════
  // FILE HANDLING
  // ══════════════════════════════════════════════════
  function setupFileHandlers() {
    const zone = $("uploadZone");
    const input = $("fileInput");

    zone.addEventListener("dragover", (e) => { e.preventDefault(); zone.classList.add("dragover"); });
    zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
    zone.addEventListener("drop", (e) => {
      e.preventDefault();
      zone.classList.remove("dragover");
      if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });
    input.addEventListener("change", (e) => {
      if (e.target.files.length) handleFile(e.target.files[0]);
    });
  }

  function handleFile(file) {
    if (file.type !== "application/pdf") { showError("Please upload a PDF file."); return; }
    if (file.size > CONFIG.MAX_FILE_SIZE_MB * 1024 * 1024) {
      showError(`File too large (max ${CONFIG.MAX_FILE_SIZE_MB} MB).`);
      return;
    }
    currentFileName = file.name;
    $("fileName").textContent = file.name;
    $("fileTag").style.display = "inline-flex";
    $("extractBtn").disabled = false;
    hideError();
    const r = new FileReader();
    r.onload = (e) => { fileBase64 = e.target.result.split(",")[1]; };
    r.readAsDataURL(file);
  }

  function clearFile() {
    fileBase64 = null;
    currentFileName = null;
    $("fileTag").style.display = "none";
    $("extractBtn").disabled = true;
    $("fileInput").value = "";
  }

  // ══════════════════════════════════════════════════
  // GEMINI AI — DEADLINE EXTRACTION
  // ══════════════════════════════════════════════════
  async function startExtraction() {
    if (!fileBase64) return;

    if (CONFIG.GEMINI_API_KEY === "YOUR_GEMINI_API_KEY_HERE") {
      showError("Please set your Gemini API key in taskpane.js → CONFIG.GEMINI_API_KEY");
      return;
    }

    // Reset state
    deadlines = []; rawTitles = []; cardStates = [];
    approvedCount = 0; skippedCount = 0;
    $("extractBtn").disabled = true;
    showStatus("Analyzing document with Gemini AI… This may take 15–30 seconds.");
    hideError();
    ["caseBanner", "deadlinesSection", "summary"].forEach(id => $(id).classList.remove("visible"));
    $("deadlinesContainer").innerHTML = "";

    const today = $("refDate").value || new Date().toISOString().split("T")[0];

    const prompt = `You are a legal document analyst. Analyze this court document (PDF) and extract ALL court-ordered deadlines, filing dates, hearing dates, response deadlines, and any other time-sensitive obligations.

IMPORTANT — Identify the parties and the source order:
- "plaintiff_name": The FIRST-NAMED plaintiff's LAST NAME only (e.g., "DeAgostino", "Smith"). For criminal cases, use the defendant's last name.
- "defendant_name": The FIRST-NAMED defendant's full name EXACTLY as it appears in the caption (e.g., "D.R. Horton, Inc.", "Smith", "City of Springfield"). Include business suffixes like Inc., LLC, Corp. if present.
- "order_title": The title or type of this court order/document EXACTLY as it appears (e.g., "Case Management Order", "Scheduling Order", "Case Management Plan", "Order Setting Trial"). Use the actual title from the document.
- "order_date": The date this order was entered/filed/signed, in YYYY-MM-DD format. Look for "ENTERED:", "DATED:", "SO ORDERED", "Filed:", or the judge's signature date. If not found, use null.

For each deadline, provide:
1. "title": A SHORT description of the event WITHOUT any party names (e.g., "Pretrial Conference", "File Motion to Dismiss", "Discovery Cutoff"). Do NOT include plaintiff or defendant names — they will be added automatically.
2. "date": YYYY-MM-DD format. Today is ${today}. Resolve relative dates. Use "UNKNOWN" only if truly indeterminate.
3. "time": HH:MM 24-hour if specified, otherwise null
4. "description": 2-3 sentences max. Key court language, who must act, consequences. No newlines inside the string.
5. "category": One of: filing, hearing, discovery, motion, trial, conference, compliance, other
6. "urgency": One of: critical, important, informational
7. "parties_responsible": Who must act
8. "virtual_info": Virtual/remote attendance details (Zoom, Teams, dial-in, meeting ID, passcode, etc.) as a single string. If none, use null.
9. "trial_end_date": For multi-day trials only. Set "date" to first day, "trial_end_date" to last day (YYYY-MM-DD). Otherwise null.

CRITICAL: Keep descriptions SHORT (under 200 chars). No newlines inside JSON strings. Output ONLY valid compact JSON, no markdown, no backticks. Ensure JSON is complete and properly closed.

Format: {"case_name":"string","case_number":"string","court":"string","document_type":"string","plaintiff_name":"LAST NAME","defendant_name":"FULL NAME AS IN CAPTION","order_title":"string","order_date":"YYYY-MM-DD or null","deadlines":[{"title":"string","date":"YYYY-MM-DD","time":"HH:MM or null","description":"string","category":"string","urgency":"string","parties_responsible":"string","virtual_info":"string or null","trial_end_date":"YYYY-MM-DD or null"}]}

If no deadlines found, return with empty deadlines array.`;

    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
      const payload = {
        contents: [{
          parts: [
            { inlineData: { mimeType: "application/pdf", data: fileBase64 } },
            { text: prompt },
          ]
        }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 16384 },
      };

      const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      const responseBody = await response.json();

      if (!response.ok) {
        hideStatus();
        showError(`Gemini API error (${response.status}): ${responseBody?.error?.message || "Unknown error"}`);
        $("extractBtn").disabled = false;
        return;
      }

      let text = responseBody?.candidates?.[0]?.content?.parts?.[0]?.text;
      if (!text) {
        hideStatus();
        showError("Gemini returned an empty response. Try a different document.");
        $("extractBtn").disabled = false;
        return;
      }

      // Clean markdown fences
      text = text.replace(/```json\s*/gi, "").replace(/```\s*/gi, "").trim();

      let parsed;
      try {
        parsed = JSON.parse(text);
      } catch (parseError) {
        parsed = repairAndParseJSON(text);
        if (!parsed) {
          hideStatus();
          showError(`Failed to parse AI response. Try a shorter document. Error: ${parseError.message}`);
          $("extractBtn").disabled = false;
          return;
        }
      }

      let extractedDeadlines = parsed.deadlines || [];
      extractedDeadlines = expandTrialPeriods(extractedDeadlines);

      hideStatus();
      handleExtractionResult({
        success: true,
        caseInfo: {
          caseName: parsed.case_name || "Unknown Case",
          caseNumber: parsed.case_number || "N/A",
          court: parsed.court || "Unknown Court",
          documentType: parsed.document_type || "Court Document",
          plaintiffName: parsed.plaintiff_name || "",
          defendantName: parsed.defendant_name || "",
          orderTitle: parsed.order_title || parsed.document_type || "Court Order",
          orderDate: parsed.order_date || null,
          sourceFile: currentFileName,
        },
        deadlines: extractedDeadlines,
      });

    } catch (e) {
      hideStatus();
      showError(`Network error: ${e.message}. Check your internet connection.`);
      $("extractBtn").disabled = false;
    }
  }

  // ── JSON repair (same logic as Apps Script version) ──
  function repairAndParseJSON(text) {
    try {
      const ds = text.indexOf('"deadlines"');
      if (ds === -1) return null;
      const as = text.indexOf('[', ds);
      if (as === -1) return null;
      let l = text.lastIndexOf('}');
      while (l > as) {
        try { const r = JSON.parse(text.substring(0, l + 1) + ']}'); if (r.deadlines) return r; } catch (e) { }
        try { const r = JSON.parse(text.substring(0, l + 1) + '}]}'); if (r.deadlines) return r; } catch (e) { }
        l = text.lastIndexOf('}', l - 1);
      }
    } catch (e) { }
    for (const c of [']}', '"]}',' "}]}']) {
      try { const r = JSON.parse(text + c); if (r.deadlines) return r; } catch (e) { }
    }
    return null;
  }

  // ── Trial period expansion ──
  function expandTrialPeriods(dls) {
    const expanded = [];
    for (const dl of dls) {
      expanded.push(dl);
      const endDate = dl.trial_end_date;
      if (!endDate || endDate === "null" || endDate === dl.date) continue;
      const start = new Date(dl.date + "T00:00:00"), end = new Date(endDate + "T00:00:00");
      if (isNaN(start.getTime()) || isNaN(end.getTime()) || end <= start) continue;
      const bdays = getBusinessDays(start, end);
      for (let i = 0; i < bdays.length; i++) {
        expanded.push({
          title: `TRIAL PERIOD (Day ${i + 1} of ${bdays.length})`,
          date: formatDateISO(bdays[i]),
          time: null,
          description: `Trial day ${i + 1} of ${bdays.length}. ${dl.description || ""}`.trim(),
          category: "trial",
          urgency: "critical",
          parties_responsible: dl.parties_responsible || "All Parties",
          virtual_info: dl.virtual_info || null,
          trial_end_date: null,
          _isTrialDay: true,
        });
      }
    }
    return expanded;
  }

  function getBusinessDays(s, e) {
    const d = [];
    const c = new Date(s);
    while (c <= e) {
      if (c.getDay() !== 0 && c.getDay() !== 6) d.push(new Date(c));
      c.setDate(c.getDate() + 1);
    }
    return d;
  }

  // ══════════════════════════════════════════════════
  // EXTRACTION RESULT HANDLER
  // ══════════════════════════════════════════════════
  function handleExtractionResult(result) {
    $("extractBtn").disabled = false;
    if (!result.success) { showError(result.error); return; }

    caseInfo = result.caseInfo;
    deadlines = result.deadlines;
    rawTitles = deadlines.map(dl => dl.title);
    cardStates = deadlines.map(() => "pending");
    approvedCount = 0;
    skippedCount = 0;

    // Populate case banner
    $("caseName").textContent = caseInfo.caseName;
    $("courtName").textContent = caseInfo.court;
    $("caseNumber").textContent = caseInfo.caseNumber;
    $("plaintiffName").textContent = caseInfo.plaintiffName || "N/A";
    $("defendantName").textContent = caseInfo.defendantName || "N/A";
    let os = caseInfo.orderTitle || "Court Order";
    if (caseInfo.orderDate) os += ` — Entered ${formatOrderDateShort(caseInfo.orderDate)}`;
    $("orderSource").textContent = os;
    $("caseBanner").classList.add("visible");

    if (!deadlines.length) { showError("No deadlines found in this document."); return; }

    $("deadlinesSubtitle").textContent =
      `${deadlines.length} deadline${deadlines.length !== 1 ? "s" : ""} found. Review, select, and approve.`;
    $("deadlinesSection").classList.add("visible");
    const container = $("deadlinesContainer");
    deadlines.forEach((dl, i) => container.appendChild(createDeadlineCard(dl, i)));
    updateSelectInfo();
    updateTitlePreview();
  }

  // ══════════════════════════════════════════════════
  // CARD BUILDER
  // ══════════════════════════════════════════════════
  function createDeadlineCard(dl, i) {
    const card = document.createElement("div");
    card.className = "dl-card";
    card.id = `card-${i}`;
    const ft = formatTitle(rawTitles[i]);
    const dateDisp = dl.date === "UNKNOWN" ? "⚠️ Unknown" : formatDateDisplay(dl.date);
    const timeDisp = hasTime(dl) ? ` at ${dl.time}` : "";
    const defReminder = getDefaultReminder(dl);
    const remLabel = getReminderLabel(dl);
    const isTrialDay = dl._isTrialDay === true;
    const hasVirtual = dl.virtual_info && dl.virtual_info !== "null" && (dl.virtual_info + "").trim();
    const urgency = dl.urgency || "important";

    let remOpts = "";
    REMINDER_OPTIONS.forEach(o => {
      remOpts += `<option value='${o.value}' ${o.value === defReminder ? "selected" : ""}>${o.label}</option>`;
    });

    card.innerHTML = `
      <div class="dl-collapsed" onclick="SH.toggleExpand(event, ${i})">
        <div class="cb-wrap" onclick="event.stopPropagation()">
          <input type="checkbox" id="check-${i}" checked onchange="SH.updateSelectInfo()">
        </div>
        <div class="dl-summary">
          <div class="dl-summary-title">${escapeHtml(ft)}</div>
          <div class="dl-summary-meta">
            <span>📅 ${dateDisp}${timeDisp}</span>
            <span>📂 ${capitalize(dl.category || "other")}</span>
          </div>
        </div>
        ${isTrialDay ? '<span class="dl-badge trial">Trial Day</span>' : ''}
        <span class="dl-badge ${urgency}">${urgency}</span>
        <span class="chevron">▾</span>
      </div>
      <div class="dl-body">
        <dl class="dl-detail-grid">
          <dt>Date</dt><dd>${dateDisp}${timeDisp}</dd>
          <dt>Category</dt><dd>${capitalize(dl.category || "other")}</dd>
          <dt>Responsible</dt><dd>${escapeHtml(dl.parties_responsible || "N/A")}</dd>
          <dt>Urgency</dt><dd><span class="dl-badge ${urgency}" style="font-size:10px;">${urgency}</span></dd>
        </dl>
        <div class="dl-description">${escapeHtml(dl.description || "No description.")}</div>
        ${hasVirtual ? `<div class="dl-virtual">💻 <strong>Virtual:</strong> ${escapeHtml(dl.virtual_info)}</div>` : ""}
        <div class="dl-notification-display" id="notif-${i}">${remLabel}</div>
        <div class="edit-section" id="edit-${i}">
          <div class="field-group"><label>Event Title</label><input type="text" id="edit-title-${i}" value="${escapeAttr(ft)}"></div>
          <div class="field-row">
            <div class="field-group"><label>Date</label><input type="date" id="edit-date-${i}" value="${dl.date !== 'UNKNOWN' ? dl.date : ''}"></div>
            <div class="field-group"><label>Time</label><input type="time" id="edit-time-${i}" value="${hasTime(dl) ? dl.time : ''}"></div>
          </div>
          <div class="field-group"><label>Reminder</label><select id="edit-reminder-${i}" onchange="SH.onReminderChange(${i})">${remOpts}</select></div>
          ${hasVirtual ? `<div class="field-group"><label>Virtual Info</label><textarea id="edit-virtual-${i}">${escapeHtml(dl.virtual_info)}</textarea></div>` : ""}
          <div class="field-group"><label>Description</label><textarea id="edit-desc-${i}">${escapeHtml(dl.description || "")}</textarea></div>
        </div>
      </div>
      <div class="dl-actions" id="actions-${i}">
        <button class="btn btn-ghost btn-xs" onclick="SH.toggleEdit(${i})">✏️ Edit</button>
        <span class="spacer"></span>
        <button class="btn btn-danger-ghost btn-xs" onclick="SH.skipDeadline(${i})">Skip</button>
        <button class="btn btn-success btn-xs" onclick="SH.approveSingle(${i})">✅ Approve</button>
      </div>
      <div class="dl-status" id="status-${i}"></div>`;
    return card;
  }

  // ══════════════════════════════════════════════════
  // EXPAND / COLLAPSE
  // ══════════════════════════════════════════════════
  function toggleExpand(event, i) {
    $(`card-${i}`).classList.toggle("expanded");
  }

  function toggleExpandAll() {
    allExpanded = !allExpanded;
    document.querySelectorAll(".dl-card").forEach(c => {
      c.classList.toggle("expanded", allExpanded);
    });
    $("expandIcon").textContent = allExpanded ? "▲" : "▼";
    $("expandLabel").textContent = allExpanded ? "Collapse All" : "Expand All";
  }

  // ══════════════════════════════════════════════════
  // SELECTION
  // ══════════════════════════════════════════════════
  function toggleSelectAll() {
    const cbs = document.querySelectorAll("[id^='check-']");
    const allChecked = Array.from(cbs)
      .filter(cb => cardStates[parseInt(cb.id.split("-")[1])] === "pending")
      .every(cb => cb.checked);
    cbs.forEach(cb => {
      const idx = parseInt(cb.id.split("-")[1]);
      if (cardStates[idx] === "pending") cb.checked = !allChecked;
    });
    updateSelectInfo();
  }

  function updateSelectInfo() {
    const c = getSelectedPendingIndices().length;
    $("selectInfo").textContent = `${c} selected`;
    $("approveSelectedBtn").disabled = c === 0;
  }

  function getSelectedPendingIndices() {
    const r = [];
    for (let i = 0; i < deadlines.length; i++) {
      const cb = $(`check-${i}`);
      if (cb && cb.checked && cardStates[i] === "pending") r.push(i);
    }
    return r;
  }

  // ══════════════════════════════════════════════════
  // EDIT
  // ══════════════════════════════════════════════════
  function toggleEdit(i) {
    const card = $(`card-${i}`);
    if (!card.classList.contains("expanded")) card.classList.add("expanded");
    $(`edit-${i}`).classList.toggle("visible");
  }

  function onReminderChange(i) {
    const s = $(`edit-reminder-${i}`);
    const d = $(`notif-${i}`);
    d.textContent = s.value === "none" ? "🔕 No reminder" : "🔔 " + s.options[s.selectedIndex].text;
  }

  function collectUserEdits(i) {
    const rv = $(`edit-reminder-${i}`).value;
    let rm;
    if (rv === "none") rm = [];
    else { try { rm = JSON.parse(rv); } catch (e) { rm = null; } }

    const e = {
      title: $(`edit-title-${i}`).value,
      date: $(`edit-date-${i}`).value,
      time: $(`edit-time-${i}`).value || null,
      description: $(`edit-desc-${i}`).value,
      reminderMinutes: rm,
    };
    const vf = $(`edit-virtual-${i}`);
    if (vf) e.virtualInfo = vf.value;
    return e;
  }

  // ══════════════════════════════════════════════════
  // APPROVE / SKIP
  // ══════════════════════════════════════════════════
  function approveSingle(i) {
    const calIds = getSelectedCalendarIds();
    if (!calIds.length) { alert("Please select at least one calendar."); return; }
    const edits = collectUserEdits(i);
    if (!edits.date) { alert("Please set a valid date."); toggleEdit(i); return; }
    processDeadline(i, edits, calIds);
  }

  async function approveSelected() {
    const calIds = getSelectedCalendarIds();
    if (!calIds.length) { alert("Please select at least one calendar."); return; }
    const indices = getSelectedPendingIndices();
    if (!indices.length) return;
    bulkProcessing = true;
    $("approveSelectedBtn").disabled = true;
    showStatus(`Approving 0 of ${indices.length}…`);

    for (let n = 0; n < indices.length; n++) {
      const idx = indices[n];
      const edits = collectUserEdits(idx);
      if (!edits.date) {
        setCardError(idx, "Skipped — no date");
        cardStates[idx] = "skipped";
        skippedCount++;
        continue;
      }
      $("statusText").textContent = `Approving ${n + 1} of ${indices.length}…`;
      await processDeadline(idx, edits, calIds);
    }
    bulkProcessing = false;
    hideStatus();
    updateSelectInfo();
    checkAllDone();
  }

  function skipDeadline(i) {
    if (cardStates[i] !== "pending") return;
    const card = $(`card-${i}`);
    const actions = $(`actions-${i}`);
    const status = $(`status-${i}`);
    const cb = $(`check-${i}`);
    card.classList.add("skipped");
    card.classList.remove("expanded");
    cardStates[i] = "skipped";
    skippedCount++;
    actions.style.display = "none";
    $(`edit-${i}`).classList.remove("visible");
    status.className = "dl-status status-skipped";
    status.innerHTML = "⏭️ Skipped";
    if (cb) cb.disabled = true;
    if (!bulkProcessing) { updateSelectInfo(); checkAllDone(); }
  }

  function skipSelected() {
    getSelectedPendingIndices().forEach(i => skipDeadline(i));
  }

  // ══════════════════════════════════════════════════
  // CALENDAR EVENT CREATION
  // ══════════════════════════════════════════════════
  async function processDeadline(i, edits, calIds) {
    const dl = deadlines[i];
    const actions = $(`actions-${i}`);
    const status = $(`status-${i}`);
    const cb = $(`check-${i}`);

    cardStates[i] = "processing";
    actions.style.display = "none";
    $(`edit-${i}`).classList.remove("visible");
    status.className = "dl-status status-processing";
    status.innerHTML = `<div class="spinner" style="width:14px;height:14px;border-width:2px;"></div> Creating…`;

    const title = edits.title;
    const description = buildEventDescription(dl, edits);
    const isTimed = edits.time && edits.time !== "";

    let errors = [];
    let created = 0;

    if (calendarMode === "graph") {
      // Use Microsoft Graph API
      for (const calId of calIds) {
        try {
          await createGraphEvent(calId, title, edits, description, isTimed);
          created++;
        } catch (e) {
          errors.push(e.message);
        }
      }
    } else {
      // ICS fallback — download a .ics file
      try {
        downloadICS(title, edits, description, isTimed);
        created = 1;
      } catch (e) {
        errors.push(e.message);
      }
    }

    // Update card
    if (errors.length === 0) {
      $(`card-${i}`).classList.add("approved");
      $(`card-${i}`).classList.remove("expanded");
      cardStates[i] = "approved";
      approvedCount++;
      const label = calendarMode === "graph"
        ? `✅ Created on ${created} calendar${created > 1 ? "s" : ""}`
        : "✅ Downloaded .ics file";
      status.className = "dl-status status-approved";
      status.innerHTML = label;
      if (cb) cb.disabled = true;
    } else if (created > 0) {
      $(`card-${i}`).classList.add("approved");
      cardStates[i] = "approved";
      approvedCount++;
      status.className = "dl-status status-approved";
      status.innerHTML = `⚠️ Created on ${created} of ${calIds.length} (${errors.length} failed)`;
      if (cb) cb.disabled = true;
    } else {
      cardStates[i] = "pending";
      status.className = "dl-status status-error";
      status.innerHTML = `❌ ${escapeHtml(errors[0])}`;
      actions.style.display = "flex";
    }

    if (!bulkProcessing) { updateSelectInfo(); checkAllDone(); }
  }

  // ── Graph API event creation ──
  async function createGraphEvent(calendarId, title, edits, description, isTimed) {
    const eventBody = {
      subject: title,
      body: { contentType: "Text", content: description },
      categories: CONFIG.EVENT_CATEGORY ? [CONFIG.EVENT_CATEGORY] : [],
    };

    if (isTimed) {
      const startDT = `${edits.date}T${edits.time}:00`;
      const endDT = new Date(new Date(startDT).getTime() + 3600000).toISOString();
      eventBody.start = { dateTime: startDT, timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone };
      eventBody.end = { dateTime: endDT.replace("Z", ""), timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone };
      eventBody.isAllDay = false;
    } else {
      eventBody.start = { dateTime: `${edits.date}T00:00:00`, timeZone: "UTC" };
      eventBody.end = { dateTime: `${edits.date}T00:00:00`, timeZone: "UTC" };
      eventBody.isAllDay = true;
    }

    // Reminders
    if (edits.reminderMinutes && edits.reminderMinutes.length > 0) {
      eventBody.isReminderOn = true;
      eventBody.reminderMinutesBeforeStart = edits.reminderMinutes[0];
    } else {
      eventBody.isReminderOn = false;
    }

    const url = calendarId === "primary"
      ? "https://graph.microsoft.com/v1.0/me/events"
      : `https://graph.microsoft.com/v1.0/me/calendars/${calendarId}/events`;

    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${graphToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(eventBody),
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err?.error?.message || `Graph API error ${response.status}`);
    }
  }

  // ── ICS File Download (fallback) ──
  function downloadICS(title, edits, description, isTimed) {
    const uid = `sh-${Date.now()}-${Math.random().toString(36).slice(2, 8)}@schedulehound`;
    const now = new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d{3}/, "");
    const descEscaped = description.replace(/\n/g, "\\n").replace(/,/g, "\\,").replace(/;/g, "\\;");

    let ics = `BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//ScheduleHound//EN\r\nBEGIN:VEVENT\r\nUID:${uid}\r\nDTSTAMP:${now}\r\n`;

    if (isTimed) {
      const dtStart = edits.date.replace(/-/g, "") + "T" + edits.time.replace(/:/g, "") + "00";
      const endTime = new Date(new Date(`${edits.date}T${edits.time}:00`).getTime() + 3600000);
      const dtEnd = formatDateISO(endTime).replace(/-/g, "") + "T" + String(endTime.getHours()).padStart(2, "0") + String(endTime.getMinutes()).padStart(2, "0") + "00";
      ics += `DTSTART:${dtStart}\r\nDTEND:${dtEnd}\r\n`;
    } else {
      const dtStart = edits.date.replace(/-/g, "");
      const nextDay = new Date(new Date(edits.date + "T00:00:00").getTime() + 86400000);
      const dtEnd = formatDateISO(nextDay).replace(/-/g, "");
      ics += `DTSTART;VALUE=DATE:${dtStart}\r\nDTEND;VALUE=DATE:${dtEnd}\r\n`;
    }

    if (edits.reminderMinutes && edits.reminderMinutes.length > 0) {
      const mins = edits.reminderMinutes[0];
      ics += `BEGIN:VALARM\r\nTRIGGER:-PT${mins}M\r\nACTION:DISPLAY\r\nDESCRIPTION:${title}\r\nEND:VALARM\r\n`;
    }

    ics += `SUMMARY:${title}\r\nDESCRIPTION:${descEscaped}\r\nEND:VEVENT\r\nEND:VCALENDAR`;

    const blob = new Blob([ics], { type: "text/calendar;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${title.replace(/[^a-zA-Z0-9 -]/g, "").slice(0, 60)}.ics`;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  // ── Event description builder ──
  function buildEventDescription(dl, edits) {
    const sep = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━";
    let d = `📋 CASE INFORMATION\n${sep}\n`;
    d += `Case:   ${caseInfo.caseName}\nNumber: ${caseInfo.caseNumber}\nCourt:  ${caseInfo.court}\n`;
    const ot = caseInfo.orderTitle || caseInfo.documentType || "Court Order";
    d += caseInfo.orderDate ? `Source: Per ${ot} Entered ${formatOrderDateShort(caseInfo.orderDate)}\n` : `Source: Per ${ot}\n`;
    d += `File:   ${caseInfo.sourceFile}\n\n`;

    const vi = edits.virtualInfo || dl.virtual_info;
    if (vi && vi !== "null" && vi.trim()) {
      d += `💻 VIRTUAL ATTENDANCE\n${sep}\n${vi}\n\n`;
    }

    d += `⚖️ DEADLINE DETAILS\n${sep}\nCategory:    ${capitalize(dl.category)}\nUrgency:     ${getUrgencyText(dl.urgency)}\nResponsible: ${dl.parties_responsible}\n\n`;
    d += `📄 DESCRIPTION & COURT LANGUAGE\n${sep}\n${edits.description || dl.description}\n\n`;
    d += `${sep}\n⚠️ DISCLAIMER: This deadline was extracted automatically by AI. Always verify against the original document. Not legal advice.\n`;
    d += `\nExtracted on: ${new Date().toISOString().replace("T", " ").slice(0, 19)}\n`;
    return d;
  }

  function getUrgencyText(u) {
    switch (u) {
      case "critical": return "🔴 CRITICAL";
      case "important": return "🟠 IMPORTANT";
      case "informational": return "🟢 INFORMATIONAL";
      default: return "⚪ Unknown";
    }
  }

  function setCardError(i, msg) {
    const s = $(`status-${i}`);
    const a = $(`actions-${i}`);
    a.style.display = "none";
    s.className = "dl-status status-skipped";
    s.style.display = "flex";
    s.innerHTML = `⏭️ ${msg}`;
    const c = $(`check-${i}`);
    if (c) c.disabled = true;
  }

  function checkAllDone() {
    if (cardStates.filter(s => s === "pending").length === 0) {
      const sm = $("summary");
      const modeNote = calendarMode === "graph" ? "Check your Outlook Calendar!" : "Open the downloaded .ics files to import.";
      $("summaryText").textContent = `${approvedCount} event${approvedCount !== 1 ? "s" : ""} created, ${skippedCount} skipped. ${modeNote}`;
      sm.classList.add("visible");
      sm.scrollIntoView({ behavior: "smooth", block: "center" });
    }
  }

  // ══════════════════════════════════════════════════
  // UI HELPERS
  // ══════════════════════════════════════════════════
  function showStatus(msg) { $("statusText").textContent = msg; $("statusBar").classList.add("visible"); }
  function hideStatus() { $("statusBar").classList.remove("visible"); }
  function showError(msg) { $("errorBanner").textContent = msg; $("errorBanner").classList.add("visible"); }
  function hideError() { $("errorBanner").classList.remove("visible"); }

  // ══════════════════════════════════════════════════
  // INITIALIZATION
  // ══════════════════════════════════════════════════
  function init() {
    setupFileHandlers();

    // Check if we're running inside Office
    if (typeof Office !== "undefined" && Office.onReady) {
      initOffice();
    } else {
      // Running standalone (e.g., for testing outside Outlook)
      console.log("Office.js not available — running in standalone mode");
      $("refDate").value = new Date().toISOString().split("T")[0];
      updateTitlePreview();
      setICSFallbackMode();
    }
  }

  // Start
  init();

  // ══════════════════════════════════════════════════
  // PUBLIC API (exposed as SH.methodName)
  // ══════════════════════════════════════════════════
  return {
    startExtraction,
    clearFile,
    onTitleFormatChange,
    updateTitlePreview,
    insertToken,
    toggleExpand,
    toggleExpandAll,
    toggleSelectAll,
    updateSelectInfo,
    approveSelected,
    approveSingle,
    skipDeadline,
    skipSelected,
    toggleEdit,
    onReminderChange,
    authenticateGraph,
  };
})();
