# ScheduleHound — Outlook Add-in (Serverless Version)

Extract court deadlines from PDF documents using AI and add them directly to your Outlook calendar.

## Architecture

This is the **serverless** version — no backend server required. The add-in runs entirely in the browser:

- **Gemini AI** is called directly from the client-side JavaScript
- **Microsoft Graph API** creates calendar events using the user's own OAuth token
- **ICS fallback** — if Graph API isn't available (POP/IMAP accounts), events are downloaded as `.ics` files that can be imported into any calendar

```
┌─────────────────────────────────────────────────┐
│  Outlook (Desktop / Web / Mobile)               │
│  ┌───────────────────────────────────────────┐  │
│  │  ScheduleHound Task Pane                  │  │
│  │  ┌─────────┐    ┌──────────────────────┐  │  │
│  │  │ PDF     │───>│ Gemini API (direct)  │  │  │
│  │  │ Upload  │    │ Returns deadlines    │  │  │
│  │  └─────────┘    └──────────────────────┘  │  │
│  │       │                                    │  │
│  │       v                                    │  │
│  │  ┌──────────────────────────────────────┐  │  │
│  │  │ Review & Approve Deadlines           │  │  │
│  │  │ (expand/collapse, edit, bulk approve) │  │  │
│  │  └──────────────────────────────────────┘  │  │
│  │       │                                    │  │
│  │       v                                    │  │
│  │  ┌─────────────┐  OR  ┌───────────────┐  │  │
│  │  │ Graph API   │      │ .ics download  │  │  │
│  │  │ (Exchange/  │      │ (POP/IMAP     │  │  │
│  │  │  M365)      │      │  fallback)     │  │  │
│  │  └─────────────┘      └───────────────┘  │  │
│  └───────────────────────────────────────────┘  │
└─────────────────────────────────────────────────┘
```

## Setup Instructions

### 1. Get a Gemini API Key

1. Go to [Google AI Studio](https://aistudio.google.com/apikey)
2. Create an API key
3. Open `taskpane.js` and replace `YOUR_GEMINI_API_KEY_HERE` with your key:
   ```js
   GEMINI_API_KEY: "AIzaSy...",
   ```

### 2. (Optional) Register an Azure AD App — for Graph API calendar access

If you want ScheduleHound to create events directly in Outlook calendars (not just download .ics files), you need an Azure AD app registration:

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click **New registration**
   - Name: `ScheduleHound`
   - Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
   - Redirect URI: **Single-page application (SPA)** → `https://localhost:3000/taskpane.html`
3. After registration, copy the **Application (client) ID**
4. Go to **API permissions** → Add:
   - `Calendars.ReadWrite`
   - `User.Read`
5. Open `taskpane.js` and replace `YOUR_AZURE_CLIENT_ID_HERE` with your client ID

**Note:** If you skip this step, ScheduleHound will automatically fall back to .ics file downloads, which work with any email account.

### 3. Host the Files over HTTPS

Office Add-ins require HTTPS. Options:

#### Option A: Local Development (easiest to start)
```bash
# Install a simple HTTPS server
npm install -g http-server

# Generate a self-signed cert (one time)
openssl req -nodes -new -x509 -keyout key.pem -out cert.pem -days 365

# Serve the files
http-server -S -C cert.pem -K key.pem -p 3000
```

#### Option B: GitHub Pages (free, permanent)
1. Push the files to a GitHub repo
2. Enable GitHub Pages in repo settings
3. Update all URLs in `manifest.xml` from `https://localhost:3000/` to your GitHub Pages URL

#### Option C: Any Static Host
Vercel, Netlify, Cloudflare Pages, Azure Static Web Apps — any of these work. Just upload the files and update the URLs in `manifest.xml`.

### 4. Update the Manifest

Open `manifest.xml` and replace all instances of `https://localhost:3000/` with your actual hosting URL.

### 5. Sideload the Add-in

#### Outlook on the Web (outlook.office.com)
1. Go to **Settings** (gear icon) → **Manage integrations** or **Get Add-ins**
2. Click **My add-ins** → **Add a custom add-in** → **Add from file**
3. Upload `manifest.xml`

#### Outlook Desktop (Windows)
1. Go to **File** → **Manage Add-ins** (opens browser)
2. Click **My add-ins** → **Add a custom add-in** → **Add from file**
3. Upload `manifest.xml`

#### Microsoft 365 Admin Center (for org-wide deployment)
1. Go to [admin.microsoft.com](https://admin.microsoft.com)
2. Settings → **Integrated apps** → **Upload custom apps**
3. Upload `manifest.xml`
4. Assign to users/groups

## Usage

1. Open any email or calendar item in Outlook
2. Click the **ScheduleHound** button in the ribbon (or find it in the add-in panel)
3. Upload a PDF court document
4. Configure settings (reference date, title format, calendar)
5. Click **Extract Deadlines**
6. Review the extracted deadlines — expand to see full details, edit as needed
7. **Approve Selected** to create all events, or approve one-by-one

## Title Format Options

| Format | Example |
|--------|---------|
| Plaintiff – Event | `Smith – Pretrial Conference` |
| Plaintiff/Defendant – Event | `Smith/DRH – Pretrial Conference` |
| Plaintiff/Def – Event Per Order Date | `Smith/DRH – Pretrial Conference Per CMO 2.21.25` |
| Case Number – Event | `CV2401234 – Pretrial Conference` |
| Custom | Any pattern using `{plaintiff}`, `{defendant}`, `{event}`, `{casenum}`, `{order}`, `{orderdate}` |

## Customization

### Defendant Abbreviations

In `taskpane.js`, find `DEFENDANT_ABBREVIATIONS` and add entries:
```js
const DEFENDANT_ABBREVIATIONS = [
  { match: "dr horton", abbr: "DRH" },
  { match: "state farm", abbr: "State Farm" },
  // Add your own...
];
```

### Order Title Abbreviations

Find `ORDER_ABBREVIATIONS`:
```js
const ORDER_ABBREVIATIONS = [
  { match: "case management order", abbr: "CMO" },
  { match: "scheduling order", abbr: "SO" },
  // Add your own...
];
```

## Security Notes

- The Gemini API key is stored in client-side JavaScript. Anyone using the add-in can view it via browser dev tools. This is acceptable for internal/small-team use.
- For wider distribution, use the **server-proxied version** instead (see the `schedulehound-server/` sibling project).
- Set API quota limits in [Google Cloud Console](https://console.cloud.google.com/) to prevent abuse.
- Microsoft Graph tokens are obtained via the user's own OAuth flow — no credentials are stored.

## Files

| File | Purpose |
|------|---------|
| `manifest.xml` | Office Add-in manifest — tells Outlook where to find the add-in |
| `taskpane.html` | Main UI shell |
| `taskpane.css` | All styles |
| `taskpane.js` | All logic: Gemini calls, Graph API, ICS fallback, UI |
| `assets/` | Add-in icons (replace with your own) |

## Troubleshooting

- **"Gemini API key not configured"** → Edit `taskpane.js` CONFIG section
- **Calendar shows "ICS fallback"** → Either Azure AD app isn't registered, or you're using a POP/IMAP account. ICS downloads still work fine.
- **Add-in doesn't appear** → Check that manifest.xml URLs match your hosting URL. Ensure HTTPS.
- **CORS error calling Gemini** → Gemini's API allows browser requests by default. If issues persist, check for browser extensions blocking requests.
