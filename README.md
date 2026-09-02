# University of Ghana Sports Directorate — Multi-Activity Registration System

Five activities, one Google Sheets backend, four static front-ends.

## Activities & member codes

| Activity | Code prefix | Example | Categories | Plans |
|---|---|---|---|---|
| Gym Membership | `G` | `G1234567` | UG Student, UG Staff, Non-UG Student, Public | Walk-in, Monthly, Semesterly (UG Student only), Quarterly, Half-yearly, Yearly |
| Leisure Tennis | `T` | `T1234567` | UG Student, UG Staff, Non-UG Student, Public | Walk-in, Monthly, Semesterly (UG Student only), Quarterly, Half-yearly, Yearly |
| Leisure Swimming | `S` | `S1234567` | UG Student, UG Staff, Non-UG Student, Public | Walk-in, Monthly, Semesterly (UG Student only), Quarterly, Half-yearly, Yearly |
| Tennis Lessons | `TL` | `TL1234567` | UG Student, UG Staff, UG Staff Relation (Under 17 / 17 & Above), Public Child (Under 17), Public Adult (17 & Above), Family Package (Max 4) | Walk-in, Monthly (30-day window) |
| Swimming Lessons | `SL` | `SL1234567` | Same 7 categories as Tennis Lessons | Walk-in, Monthly (6-week / 42-day window, capped at 12 sessions — whichever limit hits first ends the package) |

A code is always `<prefix>` + 7 digits, auto-generated for every category except UG Student / UG Staff, who keep entering their own student/staff ID unprefixed — exactly as before.

## Files

- **`Code.gs`** — the Google Apps Script backend (JSON API). Deploy this as a Web App; every front-end below talks to the same `/exec` URL.
- **`registration-app.html`** — the public-facing app: pick a program, then Register / Renew / Sign In / Sign Out. Share this one link with members.
- **`front-desk-dashboard.html`** — the main front desk. Approves/rejects everything, for all 5 activities (switch between them with the activity pills at the top), plus registration tables, visit logs, and Excel export.
- **`tennis-front-desk.html`** — satellite front desk for Leisure Tennis + Tennis Lessons. Shows full registrant detail and the visit log; can only approve/reject **walk-ins** (enforced by the backend, not just hidden in the UI) — new registrations and renewals still need the main front desk.
- **`swimming-front-desk.html`** — same, for Leisure Swimming + Swimming Lessons.

## Setup

1. Create a new Google Sheet (sheets.new).
2. Extensions → Apps Script. Delete any starter code, paste in `Code.gs`, save.
3. From the function dropdown, select `setup`, click Run. Authorize when asked. This creates 20 tabs (Pending / Registrations / Visits / Alerts × 5 activities), each with the right headers.
4. Deploy → New deployment → gear icon → Web app.
   - Execute as: **Me**
   - Who has access: **Anyone**
   
   Deploy, authorize (Drive access is needed for photos/exports), copy the URL ending in `/exec`.
5. Paste that URL into `SCRIPT_URL` near the top of **each** of the four HTML files (`registration-app.html`, `front-desk-dashboard.html`, `tennis-front-desk.html`, `swimming-front-desk.html`) — they all share the same backend.
6. Host the four HTML files wherever you like (a static host, or just open them locally) and distribute the links: `registration-app.html` to members, `front-desk-dashboard.html` to the main desk, `tennis-front-desk.html` to the tennis court desk, `swimming-front-desk.html` to the pool desk.
7. Any time `Code.gs` is edited again: Deploy → Manage deployments → pencil icon → Version: New version → Deploy, or the live URL won't see the change.

Default staff PIN on every front-desk app is `1234` — change the `STAFF_PIN` constant near the top of each file's `<script>` before going live.

See the large header comment at the top of `Code.gs` for details on photo/signature storage, the date-grouped Registrations sheet, walk-ins, renewals, and the Excel export.
