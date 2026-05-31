# NEXUS Full Build — Wiring Guide

## What's in this zip
```
nexus-pwa/          → Frontend (deploy to Vercel)
  index.html        → Full app with Automation Control Panel
  manifest.json     → PWA config
  sw.js             → Offline service worker
  icons/            → App icons

nexus-backend/
  Code.gs           → Full Apps Script (API + Automations + Demo Seed + Triggers)
```

---

## STEP 1 — Apps Script Setup (10 min)

1. Open your RescueTap Google Sheet
2. **Extensions → Apps Script** → delete everything → paste `Code.gs`
3. Click **Save**
4. **Deploy → New Deployment**
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy** → Authorize → **Copy the Exec URL**

---

## STEP 2 — Seed Demo Data (30 seconds)

Two ways:

**A. From the app** (once URL is set):
- Open the app → click **⬇ SEED DEMO** in the topbar
- Confirm the dialog → all 7 tabs auto-fill with realistic RescueTap data

**B. From Apps Script directly**:
- In the Apps Script editor, select function `seedDemoData` → click **Run**

---

## STEP 3 — Paste URL into Frontend

Open `nexus-pwa/index.html`, find line ~215:
```js
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxdMI8Ztnjkza8lc883cIfS_e6uypTwHw1NnqNvoEkD-IiTX2Zh61CGX0y7SKpfhBs6/exec';
```
Replace with your Exec URL from Step 1.

---

## STEP 4 — Install Time-Based Triggers (once)

In Apps Script editor:
- Select function `installTriggers` → click **Run**
- This sets up:
  - `dailyOverdueCheck` → fires every day at 8 AM
  - `dailyEODNudge` → fires every day at 4:30 PM

---

## STEP 5 — Deploy Frontend to Vercel

1. Push `nexus-pwa/` folder to a GitHub repo
2. Go to vercel.com → New Project → Import repo
3. Framework: **Other** (static) — no config needed
4. Deploy → your URL is live over HTTPS
5. Open on Android/Chrome → install banner appears → tap Install

---

## STEP 6 — Wire n8n (optional, for email/WhatsApp automation)

1. Deploy n8n on Railway.app (free tier)
2. Create 3 workflows:
   - **Assignment webhook** → email the assigned person
   - **Overdue webhook** → batch email reminders
   - **EOD nudge webhook** → WhatsApp/email missing reporters
3. Paste webhook URLs into `Code.gs` at the top:
```js
const N8N = {
  ASSIGNMENT : 'https://your-n8n.railway.app/webhook/assignment',
  OVERDUE    : 'https://your-n8n.railway.app/webhook/overdue',
  EOD_NUDGE  : 'https://your-n8n.railway.app/webhook/eod-nudge',
};
```

---

## Automation Control Panel

- Navigate to **Automations** in the sidebar
- Enter your name (must be Ayo, Shina, Darlington, Chizaram, or ABION)
- Toggle any automation on/off — syncs to the `8. Automation Controls` tab in the sheet instantly
- Activity log shows last 10 toggle actions with timestamp + who did it

To add more leadership names, edit the `LEADERS` array in `Code.gs`:
```js
const LEADERS = ['abion', 'ayo', 'shina', 'darlington', 'chizaram'];
```
