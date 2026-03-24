# GHL Contact Capture — Outlook Add-in

Capture email sender details from Outlook and add them as contacts in GoHighLevel (LeadConnector) with one click.

Built using Microsoft's **Office.js API** — reads sender name, email, and subject directly from Outlook's official API. No DOM scraping, works on Outlook Web, Desktop (Windows/Mac), and eventually Mobile.

---

## How it works

1. Open any email in Outlook
2. Click **"Add to GHL"** in the ribbon (top toolbar)
3. The sidebar opens with the sender's name + email pre-filled
4. Add optional phone, company, tags, and a note
5. Click **Add to GHL** → contact created in GoHighLevel ✓

---

## Setup — GitHub Pages Hosting

### Step 1: Create the repo

1. Go to [github.com/new](https://github.com/new)
2. Name it exactly: `ghl-outlook-addin`
3. Set to **Public**
4. Click **Create repository**

### Step 2: Push these files

```bash
cd ghl-outlook-addin
git init
git add .
git commit -m "Initial add-in"
git branch -M main
git remote add origin https://github.com/karan12chawla/ghl-outlook-addin.git
git push -u origin main
```

### Step 3: Enable GitHub Pages

1. Go to your repo → **Settings** → **Pages**
2. Source: **Deploy from a branch**
3. Branch: **main** / **/ (root)**
4. Click **Save**
5. Wait ~2 minutes → your add-in is live at:
   `https://karan12chawla.github.io/ghl-outlook-addin/`

### Step 4: Verify it's live

Open this URL in your browser — you should see the taskpane HTML:
`https://karan12chawla.github.io/ghl-outlook-addin/taskpane.html`

---

## Install — Sideload for Testing

### Outlook Web (outlook.office.com or outlook.live.com)

1. Open Outlook Web in Chrome
2. Open any email
3. Click the **three dots (…)** in the email toolbar → **Get Add-ins**
4. Click **My add-ins** tab → **Add a custom add-in** → **Add from URL**
5. Paste: `https://karan12chawla.github.io/ghl-outlook-addin/manifest.xml`
6. Click **Install** → **OK** on the warning
7. Close the dialog — refresh Outlook
8. Open any email → you'll see **"Add to GHL"** in the ribbon ✓

### Outlook Desktop (Windows)

1. Open Outlook Desktop
2. Go to **File** → **Manage Add-ins** (opens Outlook Web)
3. Follow the same steps as Outlook Web above

### Organisation-wide deployment (via M365 Admin)

1. Go to [admin.microsoft.com](https://admin.microsoft.com)
2. **Settings** → **Integrated apps** → **Upload custom apps**
3. Upload `manifest.xml`
4. Assign to users/groups
5. Add-in appears automatically in their Outlook within 24h

---

## Configure the Add-in

On first use, fill in the **Settings** section at the bottom of the taskpane:

| Field | Where to find it |
|---|---|
| **GHL API Key** | GHL → Settings → Integrations → API Keys → Create Key |
| **Location ID** | GHL → Settings → Business Info (or from the URL when logged in) |

Settings are saved in the browser's localStorage — you only need to enter them once per browser/device.

---

## AppSource Submission (when ready)

1. Sign up at [partner.microsoft.com](https://partner.microsoft.com) (free)
2. Go to **Marketplace offers** → **New offer** → **Office add-in**
3. Upload `manifest.xml` + screenshots + privacy policy URL
4. Microsoft reviews in 3–5 business days
5. Published — users find it via **Get Add-ins** in any Outlook

---

## File structure

```
ghl-outlook-addin/
├── manifest.xml      ← registered with Microsoft — points to GitHub Pages URLs
├── taskpane.html     ← sidebar UI
├── taskpane.js       ← Office.js contact reading + GHL API calls
├── taskpane.css      ← styles
├── commands.html     ← required stub for ribbon function file
├── README.md
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-64.png
    └── icon-80.png
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| "Add to GHL" doesn't appear in ribbon | Reload Outlook, wait 1 min after installing |
| Taskpane shows loading forever | Check browser console — likely a CORS or manifest URL issue |
| "API key missing" warning | Fill in Settings section and click Save |
| Contact created but no note | Notes use the v2 API — confirm your key has `contacts.write` and `notes.write` scope |
| Sideloading fails | Ensure GitHub Pages is enabled and manifest.xml is accessible at the URL |
| Manifest validation error | Validate at [aka.ms/officeaddinvalidator](https://aka.ms/officeaddinvalidator) |

---

## Technical notes

- Uses **Office.js Mailbox API 1.3** — supported in all modern Outlook versions
- Reads email via `Office.context.mailbox.item.from.getAsync()` — official Microsoft API, never breaks
- GHL API: tries **v2 LeadConnector** first, falls back to **v1** automatically
- No backend required — all API calls go directly from the taskpane to GHL
- Settings stored in `localStorage` — scoped to `karan12chawla.github.io`
