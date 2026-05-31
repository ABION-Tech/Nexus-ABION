# NEXUS PWA — RescueTap Operations Console

## Files
```
index.html      → Main app (all HTML + CSS + JS)
manifest.json   → PWA metadata (name, icons, shortcuts)
sw.js           → Service worker (offline cache + API fallback)
icons/          → App icons (192px + 512px)
```

## Setup (3 steps)

### 1. Paste Your Apps Script URL
Open `index.html`, find line ~210:
```js
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxdMI8Ztnjkza8lc883cIfS_e6uypTwHw1NnqNvoEkD-IiTX2Zh61CGX0y7SKpfhBs6/exec';
```
Replace with your deployed Apps Script exec URL.

### 2. Deploy to Vercel
1. Push all files to a GitHub repo (or drag-drop to vercel.com)
2. Vercel auto-detects static site — no config needed
3. Your URL: `https://nexus-[yourname].vercel.app`

> ⚠️ **Must be served over HTTPS for PWA install to work.** Vercel gives you this for free.

### 3. Install as App
- **Android/Chrome**: Banner appears automatically — tap "Install"
- **iOS/Safari**: Share → "Add to Home Screen"
- **Desktop/Chrome**: Click the ⊕ icon in the address bar

## Sidebar Bug Fix
The old bug (`sidebar-closed` applied on desktop) is fully resolved.  
The sidebar is **always visible on desktop** (≥700px) as a flex column.  
On mobile (<700px) it slides in/out via the hamburger menu.

## Data Connection
The dashboard reads from your Google Sheet via Apps Script.  
Make sure your `doGet` / `doPost` functions return:
```json
{
  "tasks": [...],
  "partners": [...],
  "kpi": [...],
  "cos": [...],
  "orgs": [...],
  "marketing": [...]
}
```
Auto-refreshes every 2 minutes. Manual sync via the `↻ SYNC` button.
