# IBI Staff Attendance — v2.10

v2.10 (27 Oct 2026): edge to edge on every screen size: no centred 440 px column on laptops; full-bleed cards with 10 px inner padding; wide screens show two columns of cards.

v2.9 (27 Oct 2026): on phones the logo sits below the top-right buttons, so the Share button never covers it.

v2.8 (27 Oct 2026): new "Share this app" button in the header (ibi-share.js).

Geofenced staff attendance PWA for India Business International.
Live: https://staff.indiabusinessinternational.online/

## What's new in v2.0
- **Server-side geofence** — punches are refused by the backend unless the phone is inside the
  warehouse zone with a good GPS fix (the old build recorded any/no coordinates). Staff allow-list too.
- **PWA** — installable (header button + in-menu "Install as App"), offline shell (`sw.js`), 1-day
  install re-prompt, iOS add-to-home-screen instructions.
- **Modern UI** — version badge (top-left), settings menu, edge-to-edge mobile, entry date/time shown as
  `28 May 2026, Thursday, 01:38:00 PM`, dark/light toggle. Headings #00c5ff on black, Roboto throughout.
- **Open Graph + Twitter** social preview (`og-banner.png`).
- **ERP integration** — exposes `getStaffSalaries`; the IBI ERP System pulls it as de-duplicated Salary expenses.

## Files
- `index.html` — the app (GitHub Pages serves this).
- `manifest.json`, `sw.js`, `icon-192.png`, `icon-512.png`, `icon-512-mask.png`, `og-banner.png` — PWA + social assets.
- `Code.gs` — Google Apps Script backend (versioned backup; the live copy lives in the bound Apps Script project).
- `logo.png`, `CNAME`.

## Backend deploy (Apps Script)
Paste `Code.gs` over the bound script → **Deploy → Manage deployments → Edit → New version → Deploy**
(the `/exec` URL stays the same). Keep `AUTH_TOKEN`, `WH_LAT/WH_LON`, `STAFF_NAMES` in sync with `index.html`.
