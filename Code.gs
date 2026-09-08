/**
 * MULTI-ACTIVITY SPORTS REGISTRATION — GOOGLE SHEETS BACKEND
 * (Gym, Leisure Tennis, Tennis Lessons, Leisure Swimming, Swimming Lessons)
 * ------------------------------------------------------------------
 * One Apps Script project is the JSON backend for FOUR static HTML
 * front-ends that live alongside this file in the repo:
 *   - registration-app.html      (public: Register / Renew / Sign In / Sign Out,
 *                                  covers all 5 activities via an activity picker)
 *   - front-desk-dashboard.html  (main front desk: full approvals for every
 *                                  activity, registration tables, visit logs)
 *   - tennis-front-desk.html     (Leisure Tennis + Tennis Lessons only —
 *                                  read-only registrant data + a Walk-in entry
 *                                  form; cannot approve/reject anything, and
 *                                  has no access to Pending or the Visit Log
 *                                  at all — everything pending, walk-ins
 *                                  included, waits for the main front desk)
 *   - swimming-front-desk.html   (same, for Leisure Swimming + Swimming Lessons)
 *
 * SHEET LAYOUT:
 * - Pending is ONE shared sheet across all 5 activities, distinguished
 *   by an "activity" column (gym / leisureTennis / etc. — see
 *   PENDING_HEADERS). A registration sits here from submission until a
 *   front desk approves or rejects it.
 * - Registrations and Visits are back to ONE SHEET PER ACTIVITY (e.g.
 *   "Registrations - Gym", "Visits - Gym" — see ACTIVITIES below), the
 *   same as the very first version of this project. Once a Pending row
 *   is approved, it's copied into that activity's own Registrations
 *   sheet — see approvePendingRow().
 * - Only Pending is shared; Registrations/Visits are isolated per
 *   activity simply by being different sheets, so there's nothing to
 *   filter — reading "Registrations - Gym" can only ever return Gym
 *   rows. No sheet gets a basic Sheets filter (the little dropdown
 *   arrows on the header row) automatically anymore — it didn't play
 *   well with the merged date-header banner rows on Registrations
 *   sheets (see "DATE-GROUPED REGISTRATIONS SHEETS" below).
 *
 * SETUP (fresh sheet):
 * 1. Create a new Google Sheet (sheets.new).
 * 2. Extensions -> Apps Script. Delete any starter code, paste this
 *    whole file in, save (disk icon / Ctrl+S / Cmd+S).
 * 3. From the function dropdown (next to Run/Debug), select "setup",
 *    click Run. Authorize when asked (Advanced -> "Go to (project)
 *    (unsafe)" -> Allow). This creates the shared Pending sheet plus
 *    each activity's own Registrations/Visits sheets (11 sheets total),
 *    each with the right headers. (The Drive folders — a "Pending
 *    Registration Photos" folder for not-yet-approved photos/
 *    signatures, and a "Registration Photos" folder for approved ones —
 *    are both created automatically the first time they're needed.)
 * 4. Deploy -> New deployment -> gear icon -> Web app.
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Deploy, authorize if asked (this version also asks for Drive
 *    access, since photos/signatures are saved there), copy the
 *    URL ending in /exec.
 * 5. Paste that URL into SCRIPT_URL in EACH of the four HTML files
 *    listed above — they all talk to the same backend.
 * 6. Any time you edit this file again: Deploy -> Manage deployments
 *    -> pencil icon -> Version: New version -> Deploy, or the live
 *    URL won't see your change.
 *
 * UPGRADING FROM THE ORIGINAL 15-SHEET LAYOUT (one Pending/
 * Registrations/Visits sheet PER activity, e.g. "Pending - Gym"):
 * 1. Paste this file in, save, redeploy (step 6 above).
 * 2. Run migratePendingToSharedSheet() ONCE from the function dropdown.
 *    It copies every row out of the old per-activity Pending sheets
 *    into the new shared Pending sheet, tagged with the right activity.
 *    Additive and safe to re-run (skips anything already present);
 *    never touches or deletes the old sheets. The old per-activity
 *    Registrations/Visits sheets need no migration at all — that's
 *    still exactly the shape this version reads.
 * 3. Once you've checked the shared Pending sheet looks right, you can
 *    run deleteLegacyPendingSheets() to remove the old per-activity
 *    Pending tabs for good (that one's a real, permanent delete — see
 *    its comment).
 *
 * UPGRADING FROM THE FULLY-MERGED LAYOUT (a short-lived version of
 * this project where Pending, Registrations AND Visits were all one
 * shared sheet each, just called "Pending"/"Registrations"/"Visits"):
 * 1. Paste this file in, save, redeploy. Pending needs no migration —
 *    it's already shaped exactly like this version expects it.
 * 2. Run splitSharedRegistrationsAndVisits() ONCE from the function
 *    dropdown. It reads the old shared "Registrations"/"Visits" sheets
 *    and copies each row into its own activity's Registrations/Visits
 *    sheet. Additive and safe to re-run; never touches or deletes the
 *    old shared sheets.
 * 3. Once you've checked the per-activity sheets look right, run
 *    deleteSharedRegistrationsAndVisitsSheets() to remove the old
 *    shared "Registrations"/"Visits" tabs for good (permanent delete).
 *
 * HOW ACTIVITIES WORK:
 * - ACTIVITIES (below) is the single source of truth for the 5
 *   activities: member-code prefix, which categories are offered,
 *   which categories must supply their own ID vs get an
 *   auto-generated code, which duration/plan options exist (with
 *   day-length and, for Swimming Lessons, a session cap), and which
 *   sheets belong to it.
 * - EVERY request (doGet view= and doPost action=) carries an
 *   "activity" key (gym / leisureTennis / leisureSwimming /
 *   tennisLessons / swimmingLessons) that says which activity's
 *   sheets to read/write. There is no cross-activity data — a member
 *   registered for the gym and for tennis lessons is two entirely
 *   separate rows in two entirely separate sheets (possibly with the
 *   same raw ID, if they used their student/staff ID both times).
 *
 * HOW IDENTITY / MEMBER CODES WORK:
 * - idNo is the one field that identifies a member within an
 *   activity: it's what UG Student / UG Staff type in themselves,
 *   and what every other category is auto-assigned the moment they
 *   submit — a random unique code shaped "<prefix><7 digits>", e.g.
 *   G1234567 (Gym), T1234567 (Leisure Tennis), S1234567 (Leisure
 *   Swimming), TL1234567 (Tennis Lessons), SL1234567 (Swimming
 *   Lessons). idNo is also the row's key for approve/reject (paired
 *   with activity on the shared Pending sheet, since a manually-typed
 *   ID can appear on two different activities' rows there), and it's
 *   the "code" a member later types into Sign In / Sign Out.
 * - Because an auto-generated idNo is created at submission (not at
 *   approval), it can serve as the pending row's key right away — but
 *   the app never shows it to the member until their registration is
 *   actually approved, via the Sign In tab's phone-number lookup.
 * - A manually-entered ID (UG Student / UG Staff) is never prefixed —
 *   it's stored exactly as typed. The SAME ID number can legitimately
 *   appear on rows in two different activities' sheets (one person,
 *   two memberships) — that's why every Pending lookup filters by
 *   activity as well as idNo, never idNo alone (a per-activity
 *   Registrations/Visits sheet needs no such filter — the sheet itself
 *   is already scoped to one activity).
 *
 * HOW SWIMMING LESSONS' SESSION CAP WORKS:
 * - Swimming Lessons has exactly ONE plan (no Walk-in): the
 *   "12-Session Package", which is BOTH a 6-week (42-day) window AND
 *   capped at 12 sign-outs, whichever is hit first — see the
 *   "sessionCap" property on that duration in ACTIVITIES, and
 *   isExpired()/getExpiryDate() below. A session only counts as used
 *   once the member SIGNS OUT (not at sign-in — see "checkout" below),
 *   incrementing that member's "sessionsUsed" column in Registrations;
 *   a renewal or fresh approval resets it back to blank.
 *
 * HOW A FAMILY PACKAGE REGISTRATION WORKS:
 * - "Family Package (Max 5)" (FAMILY_CATEGORY) is NOT one row for the
 *   whole family. The person filling out the form becomes one full
 *   row (dob/gender/medical/photo/etc. all captured normally, exactly
 *   like any other registration); every additional family member they
 *   list (up to 4 more, so 5 people total) becomes its own lightweight
 *   row — full name, gender, their relationship to the primary
 *   registrant, and optionally that person's OWN medical conditions
 *   (stored in the same hasMedicalCondition/medicalConditionDetails
 *   columns the primary registrant uses), plus the SAME shared
 *   phone/email/address/emergency-contact — but no separate
 *   dob/photo — each with its own auto-generated member code. All rows
 *   in one family also share one "familyGroupId" (a random ID stamped
 *   at submission), which is what lets doApprove()/doReject() act on
 *   the whole family in one request instead of the front desk having
 *   to approve or reject each member one at a time — see
 *   findRowIndicesByFamilyGroup() below. See the "submit" handler.
 * - Because every row in the family shares one phone number, the Sign
 *   In tab's "lookup" action (phone-number code retrieval) naturally
 *   returns every family member's code at once when the head enters
 *   that shared number — see "lookup" below.
 *
 * DATE-GROUPED REGISTRATIONS SHEETS:
 * - Every approval calls insertRegistrationIntoDateGroup(), which drops the
 *   newly-approved row straight into today's date block on its activity's
 *   Registrations sheet (creating that block at the top if this is the
 *   first approval of the day). This is a cheap, targeted insert — a
 *   single-column note scan plus one insertRowAfter()/insertRowsBefore()
 *   — so grouping is live the moment a registration is approved, including
 *   several back to back for a Family Package.
 * - A renewal is approved through this exact same insert path, as a
 *   brand new row in today's block — the member's prior Registrations
 *   row is left exactly as it was, not edited or removed. So a
 *   member's renewal history simply accumulates as separate rows over
 *   time, each one dated to when it was approved.
 * - regroupAllRegistrations() additionally rebuilds every activity's
 *   Registrations sheet from scratch (sorted newest-date-first, with a
 *   bold, shaded, merged banner row above each date's block) once a night
 *   as part of runNightlyMaintenance() (see installNightlyMaintenanceTrigger()
 *   below). It's a self-healing backstop, not the primary mechanism —
 *   it exists to catch any remaining drift (e.g. a stale banner count
 *   after a row is deleted from a block by something other than the
 *   insert path above) rather than to do the day-to-day grouping work.
 *   A full rebuild is deliberately NEVER run synchronously inside an
 *   approval request — an earlier version of this project did that, and
 *   rewriting/reformatting the whole sheet on every single approval is
 *   what caused approvals to time out with "check the connection".
 *   You can still also run the full rebuild by hand
 *   (Run > regroupAllRegistrations) any time you don't want to wait for
 *   the nightly run.
 *
 * (Photo upload, e-signature, the Excel export, walk-ins, and
 * renew/update-details all work exactly as in the original
 * single-activity version — see the inline comments near each
 * function below.)
 */


// ------------------------------------------------------------------
// Activity registry — the one place that defines the 5 activities
// ------------------------------------------------------------------

// Exact category label used everywhere a family-package row needs to be
// identified (submit(), and the multi-name family registration below).
const FAMILY_CATEGORY = "Family Package (Max 5)";

// Shared category set for Leisure Tennis, Tennis Lessons, Leisure
// Swimming and Swimming Lessons (per the printed rate cards) — Gym
// keeps its own plain 4-category set.
const LESSON_STYLE_CATEGORIES = [
  "UG Student", "UG Staff",
  "UG Staff Relation (Under 17)", "UG Staff Relation (17 & Above)",
  "Public Child (Under 17)", "Public Adult (17 & Above)",
  FAMILY_CATEGORY
];

// A "UG Staff Relation" registrant must name the UG staff member they're
// related to, that staff member's own ID number, and their relationship
// to them — restricted to Spouse or Child, nothing else is eligible.
// See the "submit" handler below.
const UG_STAFF_RELATION_CATEGORIES = ["UG Staff Relation (Under 17)", "UG Staff Relation (17 & Above)"];
const STAFF_RELATIONSHIP_OPTIONS = ["Spouse", "Child"];

const ACTIVITIES = {
  gym: {
    key: "gym",
    label: "Gym Membership",
    prefix: "G",
    // Old per-activity Pending sheet name — only ever read by
    // migratePendingToSharedSheet()/deleteLegacyPendingSheets() below,
    // never by normal request handling anymore (Pending is shared now).
    legacyPendingSheet: "Pending - Gym",
    // Registrations/Visits are back to one sheet per activity — these
    // ARE the live sheet names, read/written on every request.
    registrationsSheet: "Registrations - Gym",
    visitsSheet: "Visits - Gym",
    categories: ["UG Student", "UG Staff", "Non-UG Student", "Public"],
    idRequiredCategories: ["UG Student", "UG Staff"],
    deptRequiredCategories: ["UG Student", "UG Staff"],
    durations: {
      "Walk-in": { days: 1 },
      "Monthly": { days: 30 },
      "Semesterly": { days: 120, onlyFor: ["UG Student"] },
      "Quarterly": { days: 90, hideFor: ["UG Student"] },
      "Half-yearly": { days: 182 },
      "Yearly": { days: 365 },
      // legacy value from before Semesterly/Quarterly were split apart —
      // kept so older approved rows still calculate a correct expiry.
      "Semesterly/Quarterly": { days: 120, legacy: true }
    }
  },
  // Leisure Tennis, Tennis Lessons, Leisure Swimming and Swimming Lessons
  // all share this same 7-category set (per the printed rate cards) —
  // only Gym keeps the plain 4-category set above.
  leisureTennis: {
    key: "leisureTennis",
    label: "Leisure Tennis",
    prefix: "T",
    legacyPendingSheet: "Pending - Leisure Tennis",
    registrationsSheet: "Registrations - Leisure Tennis",
    visitsSheet: "Visits - Leisure Tennis",
    categories: LESSON_STYLE_CATEGORIES,
    idRequiredCategories: ["UG Student", "UG Staff"],
    deptRequiredCategories: ["UG Student", "UG Staff"],
    // Leisure Tennis is Walk-in/Monthly only — no Semesterly/Quarterly/
    // Half-yearly/Yearly (unlike Gym and Leisure Swimming).
    durations: {
      "Walk-in": { days: 1 },
      "Monthly": { days: 30 }
    }
  },
  leisureSwimming: {
    key: "leisureSwimming",
    label: "Leisure Swimming",
    prefix: "S",
    legacyPendingSheet: "Pending - Leisure Swimming",
    registrationsSheet: "Registrations - Leisure Swimming",
    visitsSheet: "Visits - Leisure Swimming",
    categories: LESSON_STYLE_CATEGORIES,
    idRequiredCategories: ["UG Student", "UG Staff"],
    deptRequiredCategories: ["UG Student", "UG Staff"],
    durations: {
      "Walk-in": { days: 1 },
      "Monthly": { days: 30 },
      "Semesterly": { days: 120, onlyFor: ["UG Student"] },
      "Quarterly": { days: 90, hideFor: ["UG Student"] },
      "Half-yearly": { days: 182 },
      "Yearly": { days: 365 }
    }
  },
  tennisLessons: {
    key: "tennisLessons",
    label: "Tennis Lessons",
    prefix: "TL",
    legacyPendingSheet: "Pending - Tennis Lessons",
    registrationsSheet: "Registrations - Tennis Lessons",
    visitsSheet: "Visits - Tennis Lessons",
    categories: LESSON_STYLE_CATEGORIES,
    idRequiredCategories: ["UG Student", "UG Staff"],
    deptRequiredCategories: ["UG Student", "UG Staff"],
    durations: {
      "Walk-in": { days: 1 },
      "Monthly": { days: 30 }
    }
  },
  swimmingLessons: {
    key: "swimmingLessons",
    label: "Swimming Lessons",
    prefix: "SL",
    legacyPendingSheet: "Pending - Swimming Lessons",
    registrationsSheet: "Registrations - Swimming Lessons",
    visitsSheet: "Visits - Swimming Lessons",
    categories: LESSON_STYLE_CATEGORIES,
    idRequiredCategories: ["UG Student", "UG Staff"],
    deptRequiredCategories: ["UG Student", "UG Staff"],
    // Only ONE plan exists for Swimming Lessons — no Walk-in. 6 weeks =
    // 42 days, AND capped at 12 sign-ins — whichever comes first ends
    // the package. See isExpired() below.
    durations: {
      "12-Session Package": { days: 42, sessionCap: 12 }
    },
    planDisclaimer: "Swimming lessons consist of 12 sessions held over 6 weeks. All 12 sessions must be completed within that 6-week window — sessions do not carry over beyond it."
  }
};

function getActivity(key) {
  return Object.prototype.hasOwnProperty.call(ACTIVITIES, key) ? ACTIVITIES[key] : null;
}


// All server-generated dates/times are formatted with these, in the
// spreadsheet's own timezone (File -> Settings -> Time zone in the
// Sheet — set that to your facility's actual timezone once). Keeping
// one format used everywhere means the "date" column never
// accidentally carries a time, and the "time"/"timeIn"/"timeOut"
// columns never accidentally carry a date.
const DATE_FORMAT = "M/d/yyyy";
const TIME_FORMAT = "h:mm a";

function formatNowDate() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMAT);
}
function formatNowTime() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), TIME_FORMAT);
}
function formatDateMDY(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), DATE_FORMAT);
}

// Approved members' photos/signatures live here — this is the folder
// whose files get the public "Anyone with link" sharing that lets the
// front desk's <img> thumbnails render.
const PHOTOS_FOLDER_NAME = "Registration Photos";
// A submission's photo/signature land here FIRST, while still pending —
// same public sharing (the front desk's approval card shows the photo
// so staff can check it against the person before deciding, and that
// screen isn't a Google-authenticated page, so the file has to stay
// link-viewable even before approval), but kept in a separate folder so
// pending and already-approved members' files aren't mixed together.
// On approval the file is moved into PHOTOS_FOLDER_NAME (see doApprove);
// on rejection it's trashed (see doReject); a Walk-in's photo is trashed
// on approval too, since a Walk-in becomes a Visits row with no photo
// column — see doApprove's Walk-in branch.
const PENDING_PHOTOS_FOLDER_NAME = "Pending Registration Photos";

// A "Clear List" on the front desk's Registration Table doesn't touch
// the Google Sheet at all — it just records a per-activity cutoff
// timestamp here, and the Registration Table view (and therefore the
// Excel export, which is built from that same view) hides any
// registration approved at or before it. Sign-in/out/verify/lookup
// all read the Registrations sheet directly, not through this filter,
// so members are completely unaffected by a clear.
const VIEW_CLEARED_AT_PREFIX = "VIEW_CLEARED_AT_";

// The one shared Pending sheet name — every activity's pending rows
// live here, distinguished by the "activity" column (see
// PENDING_HEADERS). Registrations/Visits are per-activity — see
// ACTIVITIES[key].registrationsSheet/.visitsSheet instead.
const PENDING_SHEET_NAME = "Pending";

// Marker written as a NOTE (not the cell value) on the idNo cell of a
// synthetic banner row, used to group a Registrations sheet into dated
// blocks — see regroupRegistrationsByDate() (full rebuild) and
// insertRegistrationIntoDateGroup() (incremental insert on approval)
// below. The actual note value is this marker PLUS the block's raw
// date string (e.g. "§DATE_HEADER§9/8/2026"), so a fresh approval can
// find "does today's block already exist" without re-parsing every
// banner's display label. sheetToObjects() filters any row whose note
// STARTS WITH this marker out of every normal read, so a banner row
// never shows up as if it were a real member.
const DATE_HEADER_MARKER = "§DATE_HEADER§";


// Columns in the shared Pending sheet. "activity" comes first so it's
// the first thing you see scanning a row, and so every lookup below
// can filter on it before ever comparing idNo.
//
// "sessionsUsed" is only ever populated for a duration with a
// sessionCap (currently just Swimming Lessons' package) — it sits
// blank/unused otherwise.
//
// "relatedStaffName"/"relatedStaffIdNo"/"staffRelationship" are only
// ever populated for a UG_STAFF_RELATION_CATEGORIES class ("UG Staff
// Relation (Under 17)"/"(17 & Above)") — see the "submit" handler
// below — and sit blank/unused otherwise.
//
// A Family Package registration is NOT one row for the whole family —
// see the "submit" handler below: the person filling the form becomes
// one full row (dob/gender/medical/photo/etc. all captured normally),
// and each additional family member they list by name becomes its own
// lightweight row (name + gender + shared phone/email/address/emergency
// contact only — no separate dob/medical/photo, except each member can
// still state their own medical conditions). All rows in one family
// also share one "familyGroupId" — see findRowIndicesByFamilyGroup().
// "familyRelationship" ("relationship to you") is likewise only ever
// populated for an additional family member's row.
//
// NOTE: PENDING_HEADERS is positional — every sheet already has its
// physical columns laid out in this exact order, and the header row
// itself is only (re)written for a brand-new sheet (see
// getOrCreateSheet). A new field must always be appended at the END of
// this array, never inserted in the middle, or every column after it
// will silently misalign with already-written rows. REGISTRATIONS_HEADERS
// below is derived from this array, so the same rule applies there too.
const PENDING_HEADERS = [
  "activity",
  "idNo", "name", "dob", "gender", "nationality", "hasMedicalCondition",
  "medicalConditionDetails", "address",
  "email", "phone", "department", "class",
  "relatedStaffName", "relatedStaffIdNo", "staffRelationship",
  "duration", "sessionsUsed",
  "date", "time", "emergencyName", "emergencyPhone", "emergencyRelationship",
  "photoUrl", "signatureUrl", "isRenewal",
  "familyRelationship", "familyGroupId"
];

// Registrations sheets are per-activity, so they don't need an
// "activity" column — the sheet itself already says which activity it
// belongs to. Otherwise identical to PENDING_HEADERS, so a Pending
// row's fields map onto a Registrations row one-for-one on approval
// (see approvePendingRow()).
const REGISTRATIONS_HEADERS = PENDING_HEADERS.filter(h => h !== "activity");

// Visits sheets are per-activity too, for the same reason.
const VISIT_HEADERS = ["visitId", "idNo", "name", "class", "date", "timeIn", "timeOut", "phone"];

// Expired/used-up-membership sign-in attempts, so every front desk for
// that activity (main and satellite alike) can be alerted even when
// they aren't the one watching that sign-in. These are deliberately
// NOT a sheet — they're a short-lived script property instead: a
// small JSON array, keyed per activity, that a staff member's
// "Dismiss" click removes an alert from. See
// getAlerts()/addAlert()/dismissAlert() below.
const ALERTS_PROPERTY_PREFIX = "ALERTS_";
// A front desk only ever needs to see recent, still-unacknowledged
// alerts — this bounds how many are kept per activity, well under
// PropertiesService's 9KB-per-value limit.
const MAX_STORED_ALERTS = 50;

function getAlerts(activity) {
  const raw = PropertiesService.getScriptProperties().getProperty(ALERTS_PROPERTY_PREFIX + activity.key);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

function addAlert(activity, alert) {
  const alerts = getAlerts(activity);
  alerts.push(alert);
  const trimmed = alerts.slice(-MAX_STORED_ALERTS);
  PropertiesService.getScriptProperties().setProperty(ALERTS_PROPERTY_PREFIX + activity.key, JSON.stringify(trimmed));
}

function dismissAlert(activity, alertId) {
  const alerts = getAlerts(activity).filter(a => a.alertId !== alertId);
  PropertiesService.getScriptProperties().setProperty(ALERTS_PROPERTY_PREFIX + activity.key, JSON.stringify(alerts));
}


function getDurationConfig(activity, duration) {
  return activity.durations[duration];
}

function getExpiryDate(activity, dateStr, duration) {
  const regDate = parseDateSafe(dateStr);
  if (!regDate) return null;
  const cfg = getDurationConfig(activity, duration);
  if (!cfg) return null;
  const expiry = new Date(regDate);
  expiry.setHours(0, 0, 0, 0);
  expiry.setDate(expiry.getDate() + cfg.days);
  return expiry;
}

// True if the date window has passed, OR (for a duration with a
// sessionCap, e.g. Swimming Lessons' Monthly) the member has used up
// their sessions — whichever comes first.
function isExpired(activity, dateStr, duration, sessionsUsed) {
  const cfg = getDurationConfig(activity, duration);
  if (!cfg) return false;
  const expiry = getExpiryDate(activity, dateStr, duration);
  if (expiry) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    if (expiry < today) return true;
  }
  if (cfg.sessionCap) {
    const used = Number(sessionsUsed) || 0;
    if (used >= cfg.sessionCap) return true;
  }
  return false;
}

function durationAllowedForCategory(cfg, category) {
  if (!cfg) return false;
  if (cfg.onlyFor && cfg.onlyFor.indexOf(category) === -1) return false;
  if (cfg.hideFor && cfg.hideFor.indexOf(category) !== -1) return false;
  return true;
}


function setup() {
  getOrCreateSheet(PENDING_SHEET_NAME, PENDING_HEADERS);
  Object.keys(ACTIVITIES).forEach(key => {
    const activity = ACTIVITIES[key];
    getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
    getOrCreateSheet(activity.visitsSheet, VISIT_HEADERS);
  });
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    // Only needed once, right when the sheet is created — this used to
    // run on every single call (getOrCreateSheet fires on every
    // request, often several times per request), and setNumberFormat
    // across the sheet's whole row capacity is expensive enough that
    // doing it every time was making every sync slow. The apostrophe-
    // prefix trick (sheetSafeText/forceLiteralText, used on every
    // write) is what actually keeps phone/ID/date/time values literal
    // — this format pass is just an extra safety net for a brand-new
    // sheet, not something that needs re-checking on every read.
    ensureTextFormatForPhoneColumns(sheet, headers);
  }
  // No automatic Sheets filter (the dropdown arrows on the header row)
  // — it doesn't play well with the merged date-header banner rows
  // (blank/weird entries in the dropdowns, filtering rows out from
  // under a banner), so it's left off. Add one by hand from the Data
  // menu if you want it for a particular sheet.
  return sheet;
}

// Phone numbers are stored like "+233 24 123 4567", and some ID numbers
// have leading zeros. A cell value that starts with "+" (or "-" or "=")
// gets read by Sheets as the start of a formula, which fails to parse
// and leaves the cell showing #ERROR! instead of the number — and a
// leading zero on a plain number gets silently dropped. "date"/"time"/
// "timeIn"/"timeOut" have the same underlying problem for a different
// reason: Sheets recognizes those strings as dates/times and silently
// converts the cell. Forcing all of these columns to Plain Text format
// stops it.
function ensureTextFormatForPhoneColumns(sheet, headers) {
  const cols = ["idNo", "phone", "emergencyPhone", "date", "time", "timeIn", "timeOut"]
    .map(h => headers.indexOf(h) + 1)
    .filter(i => i > 0);
  if (cols.length === 0) return;
  const numRows = Math.max(sheet.getMaxRows() - 1, 1);
  cols.forEach(col => {
    sheet.getRange(2, col, numRows, 1).setNumberFormat("@");
  });
}

// Belt-and-braces fix for the same "+233 24 123 4567" problem the
// Plain Text formatting above targets. A leading apostrophe is the
// standard, bulletproof way to force Sheets to store a value as
// literal text no matter what it starts with; Sheets strips that
// apostrophe automatically whenever the value is read back (via
// getValue/getValues), so nothing downstream ever sees it. Use this
// on every value going into idNo, phone, or emergencyPhone.
function sheetSafeText(v) {
  const s = (v === null || v === undefined) ? "" : String(v);
  return /^[+\-=]/.test(s) ? "'" + s : s;
}

// Unconditional version of the trick above, used for date/time strings
// like "9/1/2026" or "10:30 AM" — the problem there is that Sheets
// recognizes the PATTERN as a date or time and silently converts the
// cell to a real date/time value instead of storing the text.
function forceLiteralText(v) {
  const s = (v === null || v === undefined) ? "" : String(v);
  return s === "" ? s : "'" + s;
}

// Self-healing read: if a cell still holds a real Date object (either
// a legacy row from before this fix, or a manual edit in the sheet),
// format it back out as a plain string instead of letting a raw Date
// leak into the JSON response. headerName picks the right format.
function cellToDisplayValue(v, headerName) {
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    const fmt = (headerName === "time" || headerName === "timeIn" || headerName === "timeOut")
      ? TIME_FORMAT : DATE_FORMAT;
    return Utilities.formatDate(v, Session.getScriptTimeZone(), fmt);
  }
  return v;
}


function sheetToObjects(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values.shift();
  const idIdx = headers.indexOf("idNo");
  // The §DATE_HEADER§ banner-row marker only ever lives as a note on
  // the idNo column — fetching notes for every column via
  // range.getNotes() is much more expensive than this single column,
  // for something most sheets never have at all.
  const idNotes = (idIdx === -1 || values.length === 0)
    ? []
    : sheet.getRange(2, idIdx + 1, values.length, 1).getNotes();
  return values
    .map((row, i) => ({ row, note: idIdx === -1 ? "" : (idNotes[i] ? idNotes[i][0] : "") }))
    .filter(({ row }) => row.join("") !== "")
    .filter(({ note }) => note.indexOf(DATE_HEADER_MARKER) !== 0)
    .map(({ row }) => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = cellToDisplayValue(row[i], h));
      return obj;
    });
}


// General-purpose row-by-idNo finder.
// - On the shared Pending sheet, pass PENDING_HEADERS and an
//   activityKey — idNo alone isn't a safe key there, since a
//   manually-typed UG Student/Staff ID can legitimately appear on rows
//   for two different activities (one person, two memberships).
// - On a per-activity Registrations sheet, pass REGISTRATIONS_HEADERS
//   and omit activityKey — the sheet is already scoped to one
//   activity, so there's nothing to filter (REGISTRATIONS_HEADERS has
//   no "activity" column to check against anyway).
function findRowIndexByIdNo(sheet, idNo, headers, activityKey) {
  const idColIndex = headers.indexOf("idNo") + 1; // 1-based
  const activityColIndex = headers.indexOf("activity") + 1; // 0 if headers has no "activity" column
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, idColIndex, lastRow - 1, 1).getValues();
  const activities = activityColIndex > 0
    ? sheet.getRange(2, activityColIndex, lastRow - 1, 1).getValues()
    : null;
  const targetId = String(idNo).trim();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() !== targetId) continue;
    if (activities && String(activities[i][0]).trim() !== activityKey) continue;
    return i + 2;
  }
  return -1;
}

// Fetches exactly one Registrations row by idNo, shaped the same way
// sheetToObjects() would produce it, without reading or mapping the
// rest of the sheet. checkin/checkout/verify only ever need one
// member's row — reading (and note-scanning) the whole sheet just to
// Array.find() one match gets slower every time the sheet grows,
// which, now that a renewal appends a new row instead of updating in
// place, happens faster than it used to. Returns null if idNo isn't
// found.
function getRegistrationRowByIdNo(sheet, headers, idNo) {
  const rowIdx = findRowIndexByIdNo(sheet, idNo, headers);
  if (rowIdx === -1) return null;
  const row = sheet.getRange(rowIdx, 1, 1, headers.length).getValues()[0];
  const obj = {};
  headers.forEach((h, i) => obj[h] = cellToDisplayValue(row[i], h));
  return obj;
}

// Fetches every Registrations row whose idNo is in `codes` — used by
// "checkApproved" to resolve a handful of freshly-submitted codes (at
// most 5, a Family Package's cap) without reading the whole sheet.
// Reads just the idNo column once (cheap) to find which rows match,
// then reads only those matched rows in full.
function getRegistrationRowsByIdNos(sheet, headers, codes) {
  if (codes.length === 0) return [];
  const idColIndex = headers.indexOf("idNo") + 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const ids = sheet.getRange(2, idColIndex, lastRow - 1, 1).getValues();
  const wanted = new Set(codes);
  const matchedRows = [];
  for (let i = 0; i < ids.length; i++) {
    if (wanted.has(String(ids[i][0]).trim())) matchedRows.push(i + 2);
  }
  return matchedRows.map(rowNum => {
    const row = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
    const obj = {};
    headers.forEach((h, i) => obj[h] = cellToDisplayValue(row[i], h));
    return obj;
  });
}

// Fetches every Registrations row with a given phone number — used by
// "lookup" across all 5 activities' sheets. Same shape of optimization
// as getRegistrationRowsByIdNos() above: one cheap single-column read
// to find matches, then full reads of just those rows.
function getRegistrationRowsByPhone(sheet, headers, phone) {
  const phoneColIndex = headers.indexOf("phone") + 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !phone) return [];
  const phones = sheet.getRange(2, phoneColIndex, lastRow - 1, 1).getValues();
  const matchedRows = [];
  for (let i = 0; i < phones.length; i++) {
    if (String(phones[i][0]).trim() === phone) matchedRows.push(i + 2);
  }
  return matchedRows.map(rowNum => {
    const row = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
    const obj = {};
    headers.forEach((h, i) => obj[h] = cellToDisplayValue(row[i], h));
    return obj;
  });
}

// Counts Pending rows (any activity) with a given phone number — used
// by "lookup" to report how many of a phone number's submissions are
// still awaiting approval. A single phone-column read instead of
// sheetToObjects() reading/mapping the whole (shared, so potentially
// busiest) Pending sheet just to count matches.
function countPendingByPhone(pending, phone) {
  const phoneColIndex = PENDING_HEADERS.indexOf("phone") + 1;
  const lastRow = pending.getLastRow();
  if (lastRow < 2) return 0;
  const phones = pending.getRange(2, phoneColIndex, lastRow - 1, 1).getValues();
  let count = 0;
  for (let i = 0; i < phones.length; i++) {
    if (String(phones[i][0]).trim() === phone) count++;
  }
  return count;
}

// Every Pending row sharing one Family Package submission's
// familyGroupId (within one activity), as 1-based sheet row numbers —
// used by doApprove()/doReject() to act on a whole family at once
// instead of one member at a time. Returned in DESCENDING order so a
// caller can delete/process each row without an earlier deleteRow()
// shifting a later index still waiting to be handled. Pending-only —
// Registrations/Visits sheets don't need this (see doApprove()).
function findRowIndicesByFamilyGroup(sheet, groupId, activityKey) {
  const activityColIndex = PENDING_HEADERS.indexOf("activity") + 1;
  const groupColIndex = PENDING_HEADERS.indexOf("familyGroupId") + 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const activities = sheet.getRange(2, activityColIndex, lastRow - 1, 1).getValues();
  const groups = sheet.getRange(2, groupColIndex, lastRow - 1, 1).getValues();
  const rows = [];
  for (let i = 0; i < activities.length; i++) {
    if (String(activities[i][0]).trim() === activityKey && String(groups[i][0]).trim() === groupId) {
      rows.push(i + 2);
    }
  }
  return rows.sort((a, b) => b - a);
}


// Loads every idNo currently used by this activity — Pending (filtered
// to this activity, since Pending is shared across all 5) plus this
// activity's own Registrations sheet — into one in-memory Set, with a
// single column read from each sheet. Used by the "submit" handler so
// validating/generating up to 5 idNos (a Family Package's primary
// registrant plus up to 4 extra members) never re-reads either sheet:
// the old idNoExists()/generateUniqueIdNo() re-read both idNo columns
// on every single check, which meant up to 10 full-column reads for
// one family submission, and got slower as the sheets grew.
function loadUsedIdNoSet(activity) {
  const set = new Set();
  const pending = getOrCreateSheet(PENDING_SHEET_NAME, PENDING_HEADERS);
  const pLastRow = pending.getLastRow();
  if (pLastRow >= 2) {
    const idColIndex = PENDING_HEADERS.indexOf("idNo") + 1;
    const actColIndex = PENDING_HEADERS.indexOf("activity") + 1;
    const ids = pending.getRange(2, idColIndex, pLastRow - 1, 1).getValues();
    const acts = pending.getRange(2, actColIndex, pLastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(acts[i][0]).trim() === activity.key) set.add(String(ids[i][0]).trim());
    }
  }
  const registrations = getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
  const rLastRow = registrations.getLastRow();
  if (rLastRow >= 2) {
    const idColIndex = REGISTRATIONS_HEADERS.indexOf("idNo") + 1;
    const ids = registrations.getRange(2, idColIndex, rLastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) set.add(String(ids[i][0]).trim());
  }
  return set;
}

// Draws a random "<prefix><7 digits>" code and keeps re-rolling until
// it finds one not already in usedIdNos (see loadUsedIdNoSet() above)
// — an in-memory check instead of re-reading a sheet on every attempt.
// Capped at MAX_ATTEMPTS so this can never spin forever. Returns null
// if it still comes up empty — the caller must handle that rather than
// assume a code back. Does NOT add the code to usedIdNos itself — the
// caller must do that once it's actually accepted, so a second call
// (e.g. for the next family member) doesn't hand out the same code
// twice.
function generateUniqueIdNoFromSet(activity, usedIdNos) {
  const MAX_ATTEMPTS = 200;
  for (let i = 0; i < MAX_ATTEMPTS; i++) {
    const digits = String(Math.floor(1000000 + Math.random() * 9000000)); // 7 digits
    const code = activity.prefix + digits;
    if (!usedIdNos.has(code)) return code;
  }
  return null;
}


// ------------------------------------------------------------------
// Photo storage
// ------------------------------------------------------------------

// DriveApp.getFoldersByName() is a Drive-wide search — one of the
// slower calls in this whole project — and getPhotosFolder()/
// getPendingPhotosFolder() used to run it on every single submit,
// approve, and addPhoto. The folder's ID never changes once created,
// so it's cached in Script Properties after the first lookup; every
// call after that is a plain getFolderById(), which is fast.
function getOrCreateFolderCached(name) {
  const props = PropertiesService.getScriptProperties();
  const key = "FOLDER_ID_" + name;
  const cachedId = props.getProperty(key);
  if (cachedId) {
    try { return DriveApp.getFolderById(cachedId); } catch (err) { /* folder was deleted/moved by hand — fall through and re-resolve it */ }
  }
  const folders = DriveApp.getFoldersByName(name);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
  props.setProperty(key, folder.getId());
  return folder;
}

function getPhotosFolder() {
  return getOrCreateFolderCached(PHOTOS_FOLDER_NAME);
}

function getPendingPhotosFolder() {
  return getOrCreateFolderCached(PENDING_PHOTOS_FOLDER_NAME);
}

// Strips characters Drive/Windows/macOS dislike in filenames and
// collapses whitespace, so an applicant's name can be dropped straight
// into a filename. Falls back to "Unnamed" if nothing usable is left.
function sanitizeForFilename(name) {
  const cleaned = String(name || "")
    .replace(/[\/\\:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned || "Unnamed";
}

// Decodes a base64 (optionally data-URL-prefixed) image and saves it to
// Drive, returning a viewable URL. Returns "" (never throws) on failure,
// so a photo problem never blocks a whole registration submission.
// filenameBase is the full filename (minus extension) to save under —
// callers build this from the applicant's name + idNo (which is unique
// per-activity, and globally unique for auto-generated codes since
// every activity has its own letter prefix). folder is required — pass
// getPendingPhotosFolder() for a brand-new submission (see the "submit"
// handler) or getPhotosFolder() when adding a photo directly to an
// already-approved member (see the "addPhoto" handler).
function savePhotoAndGetUrl(filenameBase, base64Data, mimeType, folder) {
  if (!base64Data) return "";
  try {
    const cleaned = base64Data.indexOf(",") !== -1 ? base64Data.split(",")[1] : base64Data;
    const type = mimeType || "image/jpeg";
    const ext = type.indexOf("png") !== -1 ? "png" : "jpg";
    const bytes = Utilities.base64Decode(cleaned);
    const blob = Utilities.newBlob(bytes, type, `${filenameBase}.${ext}`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    return "";
  }
}

// Trashes a photo/signature file saved by savePhotoAndGetUrl, given the
// Drive URL stored in the sheet — used when a pending registration is
// rejected, so no personal info (photo, signature) is left sitting in
// Drive for someone who was never approved. Trashed rather than
// permanently deleted, so it's still recoverable from Drive's trash if
// rejected by mistake. Never throws — a stray file left behind is not
// worth failing the reject over, and an empty/unparseable url is a
// no-op.
function deleteDriveFileIfAny(url) {
  const fileId = parseDriveFileId(url);
  if (!fileId) return;
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) { /* already gone, or never a real file — nothing to clean up */ }
}

// Moves a photo/signature file (saved into the Pending folder at
// submission time) into the given folder — used on approval to move it
// out of PENDING_PHOTOS_FOLDER_NAME into PHOTOS_FOLDER_NAME. Safe to
// call on a file that's already in the destination folder (a no-op) —
// e.g. a renewal's pending row reuses the member's existing, already-
// approved photo rather than a freshly uploaded one.
function moveDriveFileIfAny(url, folder) {
  const fileId = parseDriveFileId(url);
  if (!fileId) return;
  try {
    const file = DriveApp.getFileById(fileId);
    const parents = file.getParents();
    while (parents.hasNext()) parents.next().removeFile(file);
    folder.addFile(file);
  } catch (err) { /* already gone, or never a real file — nothing to move */ }
}

// Parses either a bare Drive file ID or a full Drive URL
// (".../file/d/<ID>/view", "...?id=<ID>", etc.) into just the ID — used
// by deleteDriveFileIfAny()/moveDriveFileIfAny() above to resolve a
// photoUrl/signatureUrl cell back to the file it points at.
function parseDriveFileId(input) {
  const s = String(input || "").trim();
  if (!s) return "";
  const m = s.match(/\/d\/([-\w]{10,})/) || s.match(/[?&]id=([-\w]{10,})/);
  if (m) return m[1];
  return /^[-\w]{10,}$/.test(s) ? s : "";
}


// ------------------------------------------------------------------
// Date helpers
// ------------------------------------------------------------------

function parseDateSafe(v) {
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function dateLabelFor(dateStr) {
  const d = parseDateSafe(dateStr);
  if (!d) return String(dateStr || "Unknown date");
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "EEEE, MMM d, yyyy");
}


// ------------------------------------------------------------------
// Shared approve/reject logic
// ------------------------------------------------------------------

// Approves exactly one Pending row. A Walk-in never becomes a
// Registrations row — it's a one-off visit, so approving it writes a
// Visits row directly (checked in right now, no code needed later) and
// removes it from Pending. A renewal request is treated exactly like a
// brand new registration: it's appended as its own row (new duration,
// expiry restarted from right now, sessionsUsed reset to blank), into
// today's date block, same as any other fresh approval — the member's
// prior Registrations row is left exactly as it was, not edited or
// removed. Returns the idNo that was approved. See doApprove() below
// for how a Family Package's several rows are grouped and each run
// through this one at a time. "registrations" is that activity's own
// Registrations sheet, already resolved by the caller.
function approvePendingRow(activity, pending, registrations, idx) {
  const rowValues = pending.getRange(idx, 1, 1, PENDING_HEADERS.length).getValues()[0];
  const idNo = String(rowValues[PENDING_HEADERS.indexOf("idNo")]).replace(/^'/, "").trim();

  if (String(rowValues[PENDING_HEADERS.indexOf("duration")]).trim() === "Walk-in") {
    const visits = getOrCreateSheet(activity.visitsSheet, VISIT_HEADERS);
    const now = new Date();
    visits.appendRow(VISIT_HEADERS.map(h => {
      if (h === "visitId") return Utilities.getUuid();
      if (h === "idNo") return sheetSafeText(rowValues[PENDING_HEADERS.indexOf("idNo")]);
      if (h === "name") return rowValues[PENDING_HEADERS.indexOf("name")];
      if (h === "class") return rowValues[PENDING_HEADERS.indexOf("class")];
      // forceLiteralText, same as everywhere else a "date"/time-shaped
      // string is written — otherwise Sheets can silently store it as a
      // real Date, and "checkout" below compares this column against a
      // plain string, which would then never match.
      if (h === "date") return forceLiteralText(formatDateMDY(now));
      if (h === "timeIn") return forceLiteralText(now.toLocaleTimeString());
      if (h === "phone") return sheetSafeText(rowValues[PENDING_HEADERS.indexOf("phone")]);
      return ""; // timeOut
    }));
    // A Walk-in visit has no photo column — the photo/signature
    // captured at submission (in the Pending folder) would just sit
    // there unreferenced forever, so trash them now rather than moving
    // them anywhere.
    deleteDriveFileIfAny(rowValues[PENDING_HEADERS.indexOf("photoUrl")]);
    deleteDriveFileIfAny(rowValues[PENDING_HEADERS.indexOf("signatureUrl")]);
    pending.deleteRow(idx);
    return idNo;
  }

  // The photo/signature were uploaded into the Pending folder at
  // submission time (see the "submit" handler) — now that this row is
  // becoming a real member, move them into the approved folder. Safe to
  // call even for a renewal, whose rowValues.photoUrl/signatureUrl are
  // just copied from the member's existing (already-approved) row —
  // moving a file into the folder it's already in is a no-op.
  const approvedPhotosFolder = getPhotosFolder();
  moveDriveFileIfAny(rowValues[PENDING_HEADERS.indexOf("photoUrl")], approvedPhotosFolder);
  moveDriveFileIfAny(rowValues[PENDING_HEADERS.indexOf("signatureUrl")], approvedPhotosFolder);

  // getValues() strips any leading apostrophe on the way out, so a
  // value like "+233 24 123 4567" comes back plain again — re-guard it
  // before this appendRow()/setValue() re-triggers the same formula
  // parsing.
  ["idNo", "phone", "emergencyPhone"].forEach(h => {
    const i = PENDING_HEADERS.indexOf(h);
    rowValues[i] = sheetSafeText(rowValues[i]);
  });

  // Stamp "date"/"time" with the actual moment of approval — that's
  // what the expiry countdown is based on.
  const approvedNow = new Date();
  rowValues[PENDING_HEADERS.indexOf("date")] = forceLiteralText(formatDateMDY(approvedNow));
  rowValues[PENDING_HEADERS.indexOf("time")] = forceLiteralText(approvedNow.toLocaleTimeString());
  // A freshly (re)approved package always starts with 0 sessions used.
  rowValues[PENDING_HEADERS.indexOf("sessionsUsed")] = "";

  const isRenewalIdx = PENDING_HEADERS.indexOf("isRenewal");
  rowValues[isRenewalIdx] = ""; // flag is spent once applied — never carried into Registrations

  // A renewal is a brand new registration record, not an edit of the
  // old one — the member's prior Registrations row is left exactly as
  // it is, and this renewed row (new duration/date/time, sessionsUsed
  // reset — stamped above) is appended as its own row, same as any
  // other fresh approval. So there is no renewal-specific branch here:
  // every approval, renewal or not, takes this same path.

  // rowValues is Pending-shaped (has "activity" as its first field);
  // the per-activity Registrations sheet has no such column, so map by
  // field NAME rather than position.
  const regRowValues = REGISTRATIONS_HEADERS.map(h => rowValues[PENDING_HEADERS.indexOf(h)]);
  // Insert straight into today's date block (creating it if this is the
  // first approval of the day) instead of a plain append, so the sheet
  // stays grouped by date immediately — not just after the nightly
  // regroupAllRegistrations() backstop. Cheap: touches only the one block.
  insertRegistrationIntoDateGroup(registrations, REGISTRATIONS_HEADERS, regRowValues, formatDateMDY(approvedNow));
  pending.deleteRow(idx);
  return idNo;
}

// A Family Package submission is several Pending rows sharing one
// familyGroupId (see the PENDING_HEADERS comment above) — approving any
// one of them approves the whole family in one go, rather than making
// the front desk approve each member individually. A renewal is never
// grouped even though its cloned row carries the member's old
// familyGroupId forward: only one member is ever renewing at a time, so
// sweeping in the rest of the family (who aren't renewing anything)
// would be wrong. Rows are processed highest-row-number first so an
// earlier deleteRow() never shifts a not-yet-processed index.
function doApprove(activity, idNo) {
  const pending = getOrCreateSheet(PENDING_SHEET_NAME, PENDING_HEADERS);
  const idx = findRowIndexByIdNo(pending, idNo, PENDING_HEADERS, activity.key);
  if (idx === -1) return ok({ message: "Already handled" });

  const isRenewal = String(pending.getRange(idx, PENDING_HEADERS.indexOf("isRenewal") + 1).getValue()).trim().toUpperCase() === "TRUE";
  const groupId = isRenewal ? "" : String(pending.getRange(idx, PENDING_HEADERS.indexOf("familyGroupId") + 1).getValue()).trim();
  const rowIndices = groupId ? findRowIndicesByFamilyGroup(pending, groupId, activity.key) : [idx];

  const registrations = getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
  const approvedIdNos = rowIndices.map(rowIdx => approvePendingRow(activity, pending, registrations, rowIdx));

  return ok({ idNo: idNo, approvedIdNos: approvedIdNos });
}

// Rejecting a pending registration leaves nothing behind — the photo
// and signature saved to Drive at submission time are trashed along
// with the row, not just orphaned in the Photos folder forever. A
// Family Package's rows are grouped and rejected together too, same as
// doApprove() above (and for the same reason, a renewal is never
// grouped).
function doReject(activity, idNo) {
  const pending = getOrCreateSheet(PENDING_SHEET_NAME, PENDING_HEADERS);
  const idx = findRowIndexByIdNo(pending, idNo, PENDING_HEADERS, activity.key);
  if (idx === -1) return ok({ message: "Already handled" });

  const isRenewal = String(pending.getRange(idx, PENDING_HEADERS.indexOf("isRenewal") + 1).getValue()).trim().toUpperCase() === "TRUE";
  const groupId = isRenewal ? "" : String(pending.getRange(idx, PENDING_HEADERS.indexOf("familyGroupId") + 1).getValue()).trim();
  const rowIndices = groupId ? findRowIndicesByFamilyGroup(pending, groupId, activity.key) : [idx];

  const rejectedIdNos = rowIndices.map(rowIdx => {
    const rowValues = pending.getRange(rowIdx, 1, 1, PENDING_HEADERS.length).getValues()[0];
    const rejectedIdNo = String(rowValues[PENDING_HEADERS.indexOf("idNo")]).replace(/^'/, "").trim();
    deleteDriveFileIfAny(rowValues[PENDING_HEADERS.indexOf("photoUrl")]);
    deleteDriveFileIfAny(rowValues[PENDING_HEADERS.indexOf("signatureUrl")]);
    pending.deleteRow(rowIdx);
    return rejectedIdNo;
  });

  return ok({ rejectedIdNos: rejectedIdNos });
}


// ------------------------------------------------------------------
// HTTP handlers
// ------------------------------------------------------------------

// Registration rows are stamped with "date"/"time" at the moment
// they're approved (see doApprove) — that pair is what a Clear List
// cutoff compares against. Unparseable values are treated as "new"
// (kept visible) rather than silently hidden.
function registrationTimestampMs(row) {
  const raw = `${row.date || ""} ${row.time || ""}`.trim();
  if (!raw) return Infinity;
  const ms = new Date(raw).getTime();
  return isNaN(ms) ? Infinity : ms;
}

// One activity's Registrations rows, with the Clear List cutoff (if
// any) already applied. Shared by the plain "registrations" view, the
// main "dashboard" view, and the satellite "registrantsDashboard" view
// so the filtering logic lives in one place.
function getVisibleRegistrations(activity) {
  const sheet = getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
  let rows = sheetToObjects(sheet);
  const clearedAt = PropertiesService.getScriptProperties().getProperty(VIEW_CLEARED_AT_PREFIX + activity.key);
  if (clearedAt) {
    const cutoffMs = new Date(clearedAt).getTime();
    rows = rows.filter(r => registrationTimestampMs(r) > cutoffMs);
  }
  return { rows: rows, clearedAt: clearedAt || null };
}

function doGet(e) {
  try {
    // No specific activity needed — one sheet read total, instead of
    // the front desk making a separate request per activity just to
    // populate the pending-count badges. (Pending is still the one
    // shared sheet, so this is still cheap.)
    if (e.parameter.view === 'allPendingCounts') {
      const sheet = getOrCreateSheet(PENDING_SHEET_NAME, PENDING_HEADERS);
      const counts = {};
      Object.keys(ACTIVITIES).forEach(key => { counts[key] = 0; });
      sheetToObjects(sheet).forEach(r => { if (counts[r.activity] !== undefined) counts[r.activity]++; });
      return ok({ counts: counts });
    }

    const activity = getActivity(e.parameter.activity);
    if (!activity) return errorMsg("Unknown or missing activity.");

    const view = e.parameter.view;
    if (view === 'alerts') {
      return ok({ rows: getAlerts(activity) });
    }
    if (view === 'visits') {
      const sheet = getOrCreateSheet(activity.visitsSheet, VISIT_HEADERS);
      return ok({ rows: sheetToObjects(sheet) });
    }
    if (view === 'pending') {
      const sheet = getOrCreateSheet(PENDING_SHEET_NAME, PENDING_HEADERS);
      return ok({ rows: sheetToObjects(sheet).filter(r => r.activity === activity.key) });
    }
    if (view === 'dashboard') {
      // Everything the MAIN front desk's auto-refresh cycle needs for
      // one activity, in a single execution instead of 3-4 separate
      // ones (pending/registrations/visits/alerts each cost their own
      // request overhead — spreadsheet open, auth — on top of the
      // actual read).
      const pendingSheet = getOrCreateSheet(PENDING_SHEET_NAME, PENDING_HEADERS);
      const visitsSheet = getOrCreateSheet(activity.visitsSheet, VISIT_HEADERS);
      const registrations = getVisibleRegistrations(activity);
      return ok({
        pending: sheetToObjects(pendingSheet).filter(r => r.activity === activity.key),
        registrations: registrations.rows,
        clearedAt: registrations.clearedAt,
        visits: sheetToObjects(visitsSheet),
        alerts: getAlerts(activity)
      });
    }
    if (view === 'registrantsDashboard') {
      // The satellite (tennis/swimming) front desks only have access to
      // approved registrants and expired-membership alerts — no
      // Pending, no Visit Log (both are handled at the main front
      // desk) — so this is a leaner view than "dashboard" above, and
      // still just one request per refresh cycle.
      const registrations = getVisibleRegistrations(activity);
      return ok({
        registrations: registrations.rows,
        clearedAt: registrations.clearedAt,
        alerts: getAlerts(activity)
      });
    }
    const registrations = getVisibleRegistrations(activity);
    return ok({ rows: registrations.rows, clearedAt: registrations.clearedAt });
  } catch (err) {
    return errorOut(err);
  }
}


function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || "submit";
    const activity = getActivity(data.activity);
    if (!activity) return errorMsg("Unknown or missing activity.");


    if (action === "submit") {
      if (activity.categories.indexOf(data.class) === -1) {
        return errorMsg("Invalid category for this activity.");
      }
      const durCfg = getDurationConfig(activity, data.duration);
      if (!durCfg || !durationAllowedForCategory(durCfg, data.class)) {
        return errorMsg("That plan isn't available for your category.");
      }

      // A UG Staff Relation must name the UG staff member they're
      // related to, that staff member's own ID number, and their
      // relationship — restricted to Spouse or Child, nothing else is
      // eligible for this category.
      const relatedStaffName = String(data.relatedStaffName || "").trim();
      const relatedStaffIdNo = String(data.relatedStaffIdNo || "").trim();
      const staffRelationship = String(data.staffRelationship || "").trim();
      if (UG_STAFF_RELATION_CATEGORIES.indexOf(data.class) !== -1) {
        if (!relatedStaffName || !relatedStaffIdNo) {
          return errorMsg("Please provide the full name and ID number of the UG staff member you're related to.");
        }
        if (STAFF_RELATIONSHIP_OPTIONS.indexOf(staffRelationship) === -1) {
          return errorMsg("Only a Spouse or Child of a UG staff member is eligible to register under this category.");
        }
      }

      // One read of every idNo already used in this activity (Pending +
      // Registrations), instead of re-reading both sheets for every
      // idNo checked/generated below — a Family Package checks/
      // generates up to 5 of these in one submission. usedIdNos is
      // updated in place as each one is accepted, so two family
      // members can never end up handed the same generated code.
      const usedIdNos = loadUsedIdNoSet(activity);

      // ID number is optional for categories NOT in idRequiredCategories
      // — an auto-generated "<prefix><7 digits>" code is used if they
      // don't have or didn't provide one. idRequiredCategories (UG
      // Student / UG Staff) must supply their own.
      const idRequired = activity.idRequiredCategories.indexOf(data.class) !== -1;
      let idNo;
      if (!idRequired) {
        idNo = String(data.idNo || "").trim();
        if (!idNo) {
          idNo = generateUniqueIdNoFromSet(activity, usedIdNos);
          if (!idNo) {
            return errorMsg("Couldn't generate a member code right now — the code pool may be full. Please ask the front desk to register you with a manual ID number instead.");
          }
        } else if (usedIdNos.has(idNo)) {
          return errorMsg("This ID number is already registered or pending approval.");
        }
      } else {
        idNo = String(data.idNo || "").trim();
        if (!idNo) return errorMsg("An ID number is required for this category.");
        if (usedIdNos.has(idNo)) return errorMsg("This ID number is already registered or pending approval.");
      }
      usedIdNos.add(idNo);

      // Name-first filenames (idNo tucked in parentheses for uniqueness)
      // so photos/signatures can be found by applicant name in Drive.
      // Lands in the Pending folder for now — doApprove() moves it into
      // the approved folder once (if) this registration is approved,
      // and doReject()/a Walk-in approval trashes it otherwise.
      const applicantFileName = sanitizeForFilename(data.name);
      const pendingPhotosFolder = getPendingPhotosFolder();
      const photoUrl = savePhotoAndGetUrl(`${applicantFileName} (${idNo})`, data.photoBase64, data.photoMimeType, pendingPhotosFolder);
      const signatureUrl = savePhotoAndGetUrl(`${applicantFileName} (${idNo}) - Signature`, data.signatureBase64, data.signatureMimeType, pendingPhotosFolder);

      // A Family Package registration isn't one row for the whole
      // family — see the PENDING_HEADERS comment above. The person
      // filling the form (idNo/photo/etc. above) is one full row;
      // every additional family member they list gets their own
      // lightweight row below — name, and optionally their own medical
      // conditions (stored in the same
      // hasMedicalCondition/medicalConditionDetails columns the
      // primary registrant uses) — generated and validated up front so
      // the whole submission fails cleanly (nothing written) rather
      // than partially, if the code pool or the 5-person cap is hit.
      // Shared by every row in this family (the primary registrant and
      // each additional member below) so doApprove()/doReject() can
      // find and act on the whole family at once — see
      // findRowIndicesByFamilyGroup(). Blank for a non-family
      // registration.
      const familyGroupId = data.class === FAMILY_CATEGORY ? Utilities.getUuid() : "";

      let extraFamilyMembers = [];
      if (data.class === FAMILY_CATEGORY) {
        const rawMembers = Array.isArray(data.familyMembers) ? data.familyMembers : [];
        const members = rawMembers
          .map(m => ({
            name: String((m && m.name) || "").trim(),
            gender: String((m && m.gender) || "").trim(),
            relationship: String((m && m.relationship) || "").trim(),
            medicalConditions: String((m && m.medicalConditions) || "").trim()
          }))
          .filter(m => m.name !== "");
        if (members.length > 4) {
          return errorMsg("A family package covers at most 5 people, including you — please list at most 4 additional family members.");
        }
        for (const m of members) {
          const extraIdNo = generateUniqueIdNoFromSet(activity, usedIdNos);
          if (!extraIdNo) {
            return errorMsg("Couldn't generate member codes for the whole family right now — the code pool may be full. Please ask the front desk to register the family manually instead.");
          }
          usedIdNos.add(extraIdNo);
          extraFamilyMembers.push({
            name: m.name,
            idNo: extraIdNo,
            gender: m.gender,
            familyRelationship: m.relationship,
            hasMedicalCondition: m.medicalConditions ? "Yes" : "No",
            medicalConditionDetails: m.medicalConditions
          });
        }
      }

      const sheet = getOrCreateSheet(PENDING_SHEET_NAME, PENDING_HEADERS);

      const primaryRow = PENDING_HEADERS.map(h => {
        if (h === "activity") return activity.key;
        if (h === "idNo") return sheetSafeText(idNo);
        if (h === "photoUrl") return photoUrl;
        if (h === "signatureUrl") return signatureUrl;
        if (h === "sessionsUsed") return "";
        if (h === "phone" || h === "emergencyPhone" || h === "relatedStaffIdNo") return sheetSafeText(data[h] || "");
        // Force "date"/"time" to literal text too — otherwise Sheets
        // silently converts them to real date/time values.
        if (h === "date" || h === "time") return forceLiteralText(data[h] || "");
        if (h === "familyGroupId") return familyGroupId;
        return data[h] || "";
      });

      const familyMemberRows = extraFamilyMembers.map(member =>
        PENDING_HEADERS.map(h => {
          if (h === "activity") return activity.key;
          if (h === "idNo") return sheetSafeText(member.idNo);
          if (h === "name") return member.name;
          if (h === "gender") return member.gender;
          if (h === "familyRelationship") return member.familyRelationship;
          if (h === "class") return data.class;
          if (h === "duration") return data.duration;
          if (h === "hasMedicalCondition") return member.hasMedicalCondition;
          if (h === "medicalConditionDetails") return member.medicalConditionDetails;
          if (h === "phone" || h === "emergencyPhone") return sheetSafeText(data[h] || "");
          if (h === "email" || h === "address" || h === "emergencyName" || h === "emergencyRelationship") return data[h] || "";
          if (h === "date" || h === "time") return forceLiteralText(data[h] || "");
          if (h === "familyGroupId") return familyGroupId;
          // dob/nationality/department/photo/signature/sessionsUsed/
          // isRenewal are all left blank for an additional family
          // member — only their name, gender, relationship to the
          // primary registrant, optional medical conditions, and the
          // shared contact details are collected.
          return "";
        })
      );

      // One write for the primary registrant plus every extra family
      // member (up to 5 rows total), instead of a separate appendRow()
      // per row — each appendRow() is its own round trip to the Sheets
      // service, and a Family Package could mean up to 5 of them back
      // to back.
      const allNewRows = [primaryRow].concat(familyMemberRows);
      sheet.getRange(sheet.getLastRow() + 1, 1, allNewRows.length, PENDING_HEADERS.length).setValues(allNewRows);

      const response = { idNo: idNo };
      if (data.class === FAMILY_CATEGORY) {
        response.familyMembers = [{ name: data.name, idNo: idNo }].concat(extraFamilyMembers);
      }
      return ok(response);
    }


    if (action === "approve") {
      return doApprove(activity, data.idNo);
    }


    if (action === "reject") {
      return doReject(activity, data.idNo);
    }


    if (action === "checkin") {
      const registrations = getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
      const code = String(data.code || "").trim();
      const match = getRegistrationRowByIdNo(registrations, REGISTRATIONS_HEADERS, code);
      if (!match) return errorMsg("Code not recognized");

      const now = new Date();

      // Expired/used-up membership: don't log a visit — alert the
      // front desk instead so a staff member can sort it out with the
      // member in person, rather than letting an expired code silently
      // work.
      if (isExpired(activity, match.date, match.duration, match.sessionsUsed)) {
        const expiry = getExpiryDate(activity, match.date, match.duration);
        const expiredOnLabel = expiry ? Utilities.formatDate(expiry, Session.getScriptTimeZone(), DATE_FORMAT) : "";
        addAlert(activity, {
          alertId: Utilities.getUuid(),
          idNo: match.idNo,
          name: match.name,
          class: match.class,
          duration: match.duration,
          expiredOn: expiredOnLabel,
          date: formatDateMDY(now),
          time: now.toLocaleTimeString()
        });
        const cfg = getDurationConfig(activity, match.duration);
        const usedUp = cfg && cfg.sessionCap && (Number(match.sessionsUsed) || 0) >= cfg.sessionCap;
        return errorMsg(usedUp
          ? `You've used all ${cfg.sessionCap} sessions on this package. Please see the front desk to renew.`
          : ("Your membership expired" + (expiredOnLabel ? ` on ${expiredOnLabel}` : "") + ". Please see the front desk to renew."));
      }

      const visits = getOrCreateSheet(activity.visitsSheet, VISIT_HEADERS);
      visits.appendRow(VISIT_HEADERS.map(h => {
        if (h === "visitId") return Utilities.getUuid();
        if (h === "idNo") return sheetSafeText(match.idNo);
        if (h === "name") return match.name;
        if (h === "class") return match.class;
        // forceLiteralText, same as everywhere else a "date"/time-shaped
        // string is written — otherwise Sheets can silently store it as
        // a real Date, and "checkout" below compares this column
        // against a plain string, which would then never match.
        if (h === "date") return forceLiteralText(formatDateMDY(now));
        if (h === "timeIn") return forceLiteralText(now.toLocaleTimeString());
        return ""; // timeOut, phone stay blank at check-in
      }));
      return ok({ member: match });
    }


    if (action === "checkout") {
      const registrations = getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
      const code = String(data.code || "").trim();
      const match = getRegistrationRowByIdNo(registrations, REGISTRATIONS_HEADERS, code);
      if (!match) return errorMsg("Code not recognized");

      const visits = getOrCreateSheet(activity.visitsSheet, VISIT_HEADERS);
      const lastRow = visits.getLastRow();
      if (lastRow < 2) return errorMsg("No sign-in found for this code. Please sign in first.");

      const idColIndex = VISIT_HEADERS.indexOf("idNo");
      const timeOutColIndex = VISIT_HEADERS.indexOf("timeOut");
      const values = visits.getRange(2, 1, lastRow - 1, VISIT_HEADERS.length).getValues();

      // Finds this member's most recent STILL-OPEN visit (no timeOut
      // yet), whatever day it was signed in on — not just one recorded
      // as "today". Requiring an exact same-day match here used to mean
      // a sign-in made late at night, or any drift between the sheet's
      // time zone and the venue's, could leave a member unable to sign
      // out at all even though their visit was genuinely still open;
      // scanning from the last row down already finds the newest one
      // first, and the nightly auto sign-out (see autoSignOutAt9pm)
      // closes anything left open at day's end anyway, so there's
      // nothing an exact-date check was actually protecting against.
      let targetRow = -1;
      for (let i = values.length - 1; i >= 0; i--) {
        const row = values[i];
        if (String(row[idColIndex]).trim() === match.idNo && !row[timeOutColIndex]) {
          targetRow = i + 2; // sheet row number
          break;
        }
      }
      if (targetRow === -1) return errorMsg("No open sign-in found for this code. Please sign in first.");

      visits.getRange(targetRow, timeOutColIndex + 1).setValue(new Date().toLocaleTimeString());

      // Duration has a session cap (Swimming Lessons' package) — a
      // session only counts as "used" once the member actually signs
      // out, not when they sign in (so a session in progress doesn't
      // get counted early, and a forgotten sign-in with no sign-out
      // doesn't burn a session at all).
      const cfg = getDurationConfig(activity, match.duration);
      if (cfg && cfg.sessionCap) {
        const regIdx = findRowIndexByIdNo(registrations, match.idNo, REGISTRATIONS_HEADERS);
        if (regIdx !== -1) {
          const newUsed = (Number(match.sessionsUsed) || 0) + 1;
          registrations.getRange(regIdx, REGISTRATIONS_HEADERS.indexOf("sessionsUsed") + 1).setValue(newUsed);
          match.sessionsUsed = String(newUsed);
        }
      }
      return ok({ member: match });
    }


    if (action === "checkoutByPhone") {
      // For members without a usable code — Walk-ins above all, since
      // they're never in Registrations and never shown a code — this
      // finds their open Visits row by the phone number they signed in
      // with instead.
      const phone = String(data.phone || "").trim();
      if (!phone) return errorMsg("Enter the phone number you signed in with.");

      const visits = getOrCreateSheet(activity.visitsSheet, VISIT_HEADERS);
      const lastRow = visits.getLastRow();
      if (lastRow < 2) return errorMsg("No sign-in found for this phone number. Please sign in first.");

      const phoneColIndex = VISIT_HEADERS.indexOf("phone");
      const timeOutColIndex = VISIT_HEADERS.indexOf("timeOut");
      const values = visits.getRange(2, 1, lastRow - 1, VISIT_HEADERS.length).getValues();

      // Most recent STILL-OPEN visit for this phone number, whatever day
      // it was signed in on — see the matching comment in "checkout"
      // above for why an exact same-day match isn't required.
      let targetRow = -1;
      for (let i = values.length - 1; i >= 0; i--) {
        const row = values[i];
        if (String(row[phoneColIndex]).trim() === phone && !row[timeOutColIndex]) {
          targetRow = i + 2; // sheet row number
          break;
        }
      }
      if (targetRow === -1) {
        return errorMsg("No open sign-in found for this phone number. Please sign in first, or ask the front desk.");
      }

      visits.getRange(targetRow, timeOutColIndex + 1).setValue(new Date().toLocaleTimeString());
      const rowValues = visits.getRange(targetRow, 1, 1, VISIT_HEADERS.length).getValues()[0];
      const visit = {};
      VISIT_HEADERS.forEach((h, i) => visit[h] = rowValues[i]);
      return ok({ member: visit });
    }


    if (action === "addPhoto") {
      const idNo = String(data.idNo || "").trim();
      if (!idNo) return errorMsg("Enter your code or ID number.");
      if (!data.photoBase64) return errorMsg("No photo received.");

      const pending = getOrCreateSheet(PENDING_SHEET_NAME, PENDING_HEADERS);
      let idx = findRowIndexByIdNo(pending, idNo, PENDING_HEADERS, activity.key);
      let targetSheet = null, targetHeaders = null;
      if (idx !== -1) {
        targetSheet = pending;
        targetHeaders = PENDING_HEADERS;
      } else {
        const registrations = getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
        idx = findRowIndexByIdNo(registrations, idNo, REGISTRATIONS_HEADERS);
        if (idx !== -1) { targetSheet = registrations; targetHeaders = REGISTRATIONS_HEADERS; }
      }
      if (!targetSheet) return errorMsg("Code not recognized");

      const photoColIndex = targetHeaders.indexOf("photoUrl") + 1;
      const nameColIndex = targetHeaders.indexOf("name") + 1;
      const applicantName = targetSheet.getRange(idx, nameColIndex).getValue();
      const applicantFileName = sanitizeForFilename(applicantName);
      // Pending vs. already-approved decides which folder — same split
      // as the "submit" handler.
      const folder = targetSheet === pending ? getPendingPhotosFolder() : getPhotosFolder();
      const photoUrl = savePhotoAndGetUrl(`${applicantFileName} (${idNo})`, data.photoBase64, data.photoMimeType, folder);
      if (!photoUrl) return errorMsg("Couldn't save the photo — please try again.");

      targetSheet.getRange(idx, photoColIndex).setValue(photoUrl);
      return ok({});
    }


    if (action === "lookup") {
      // Returns EVERY approved row sharing this phone number, across
      // EVERY activity, not just the one the request happened to be
      // made from — so a phone number connected to more than one
      // subscription (e.g. Gym AND Tennis Lessons) gets every code for
      // every one of them back in a single lookup. Registrations lives
      // in its own sheet per activity, so this loops all 5 and tags
      // each match with which activity it came from (a per-activity
      // sheet has no "activity" column of its own to read that off
      // of), so the caller can label which subscription each code
      // belongs to. A Family Package's shared contact number retrieves
      // every family member's code the same way (each family member is
      // its own row — see the PENDING_HEADERS comment above).
      const phone = String(data.phone || "").trim();

      const pending = getOrCreateSheet(PENDING_SHEET_NAME, PENDING_HEADERS);
      const pendingCount = countPendingByPhone(pending, phone);

      const approvedMembers = [];
      Object.keys(ACTIVITIES).forEach(key => {
        const act = ACTIVITIES[key];
        const registrations = getOrCreateSheet(act.registrationsSheet, REGISTRATIONS_HEADERS);
        getRegistrationRowsByPhone(registrations, REGISTRATIONS_HEADERS, phone)
          .forEach(r => { r.activity = act.key; approvedMembers.push(r); });
      });

      if (approvedMembers.length === 0 && pendingCount === 0) return ok({ found: false });
      return ok({ found: true, approvedMembers: approvedMembers, pendingCount: pendingCount });
    }


    if (action === "checkApproved") {
      // Lets the registration app poll right after a fresh submission
      // and pop the code up automatically the moment the front desk
      // approves it, instead of making the registrant come back later
      // and look themselves up by phone. Takes the idNo(s) handed back
      // by "submit" (one per family member for a Family Package) and
      // reports back whichever of those are now approved — the rest
      // may still be pending, so the caller keeps polling for those.
      const codes = Array.isArray(data.idNos)
        ? data.idNos.map(c => String(c || "").trim()).filter(Boolean)
        : [];
      if (codes.length === 0) return errorMsg("No codes to check.");
      const registrations = getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
      const approvedMembers = getRegistrationRowsByIdNos(registrations, REGISTRATIONS_HEADERS, codes);
      return ok({ approvedMembers: approvedMembers });
    }


    if (action === "verify") {
      // Looks a member up by code for the Renew tab's gate — deliberately
      // does NOT log a Visits row (unlike "checkin"). Approved members only.
      const registrations = getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
      const code = String(data.code || "").trim();
      if (!code) return errorMsg("Enter your code or ID number.");
      const match = getRegistrationRowByIdNo(registrations, REGISTRATIONS_HEADERS, code);
      if (!match) return errorMsg("Code not recognized");
      return ok({ member: match });
    }


    if (action === "requestRenewal") {
      const code = String(data.code || "").trim();
      const duration = String(data.duration || "").trim();
      if (!code) return errorMsg("Enter your code or ID number.");
      if (!duration) return errorMsg("Please choose a plan.");

      const registrations = getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
      const idx = findRowIndexByIdNo(registrations, code, REGISTRATIONS_HEADERS);
      if (idx === -1) return errorMsg("Code not recognized");

      const currentClass = registrations.getRange(idx, REGISTRATIONS_HEADERS.indexOf("class") + 1).getValue();
      const durCfg = getDurationConfig(activity, duration);
      if (!durCfg || !durationAllowedForCategory(durCfg, currentClass)) {
        return errorMsg("That plan isn't available for your category.");
      }

      const pending = getOrCreateSheet(PENDING_SHEET_NAME, PENDING_HEADERS);
      if (findRowIndexByIdNo(pending, code, PENDING_HEADERS, activity.key) !== -1) {
        return errorMsg("You already have a request awaiting approval at the front desk.");
      }

      const rowValues = registrations.getRange(idx, 1, 1, REGISTRATIONS_HEADERS.length).getValues()[0];
      const now = new Date();
      // Mapped by field NAME, not position — the Registrations row has
      // no "activity" column, so PENDING_HEADERS (which does) can't be
      // filled in positionally the way it could when both sheets used
      // to share one layout.
      const pendingRow = PENDING_HEADERS.map(h => {
        if (h === "activity") return activity.key;
        if (h === "duration") return duration;
        if (h === "date") return forceLiteralText(formatDateMDY(now));
        if (h === "time") return forceLiteralText(now.toLocaleTimeString());
        if (h === "isRenewal") return "TRUE";
        const regFieldIdx = REGISTRATIONS_HEADERS.indexOf(h);
        const raw = regFieldIdx === -1 ? "" : rowValues[regFieldIdx];
        if (h === "idNo" || h === "phone" || h === "emergencyPhone") return sheetSafeText(raw);
        return raw;
      });
      pending.appendRow(pendingRow);
      return ok({});
    }


    if (action === "updateDetails") {
      const code = String(data.code || "").trim();
      if (!code) return errorMsg("Enter your code or ID number.");

      const registrations = getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
      const idx = findRowIndexByIdNo(registrations, code, REGISTRATIONS_HEADERS);
      if (idx === -1) return errorMsg("Code not recognized");

      // Only these fields are editable from the Renew tab's form —
      // idNo, name, class, duration, dates, and photo stay untouched.
      const editableFields = ["phone", "email", "address", "emergencyName", "emergencyPhone", "emergencyRelationship"];
      editableFields.forEach(h => {
        if (data[h] === undefined) return;
        const val = (h === "phone" || h === "emergencyPhone") ? sheetSafeText(data[h]) : data[h];
        registrations.getRange(idx, REGISTRATIONS_HEADERS.indexOf(h) + 1).setValue(val);
      });

      const rowValues = registrations.getRange(idx, 1, 1, REGISTRATIONS_HEADERS.length).getValues()[0];
      const member = {};
      REGISTRATIONS_HEADERS.forEach((h, i) => member[h] = rowValues[i]);
      return ok({ member: member });
    }


    if (action === "acknowledgeAlert") {
      const alertId = String(data.alertId || "").trim();
      if (!alertId) return errorMsg("Missing alert id.");
      dismissAlert(activity, alertId);
      return ok({});
    }


    // "Clear List" only records a cutoff timestamp — see
    // registrationTimestampMs()/doGet above. Nothing is deleted from
    // the Registrations sheet, so sign-in/out/verify/lookup and the
    // sheet itself are completely unaffected; only the front desk's
    // Registration Table view (and Excel exports built from it) hide
    // anything approved at or before the cutoff.
    if (action === "clearRegistrationsView") {
      const clearedAt = new Date().toISOString();
      PropertiesService.getScriptProperties().setProperty(VIEW_CLEARED_AT_PREFIX + activity.key, clearedAt);
      return ok({ clearedAt: clearedAt });
    }

    if (action === "restoreRegistrationsView") {
      PropertiesService.getScriptProperties().deleteProperty(VIEW_CLEARED_AT_PREFIX + activity.key);
      return ok({});
    }


    return errorMsg("Unknown action");


  } catch (err) {
    return errorOut(err);
  }
}


function ok(extra) {
  return ContentService
    .createTextOutput(JSON.stringify(Object.assign({ status: "ok" }, extra)))
    .setMimeType(ContentService.MimeType.JSON);
}


function errorMsg(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "error", message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}


function errorOut(err) {
  return errorMsg(String(err));
}


// ------------------------------------------------------------------
// Migrating Pending from the old per-activity sheets
// ------------------------------------------------------------------

// Run this ONCE (Run > migratePendingToSharedSheet) after deploying
// this version of Code.gs, if you're coming from the very first,
// fully-per-activity layout (one Pending/Registrations/Visits sheet
// PER activity). Copies every row out of the old per-activity Pending
// sheets ("Pending - Gym" etc.) into the new shared Pending sheet,
// tagged with which activity each row belongs to. Additive and safe to
// re-run — it only appends a row if one for that idNo isn't already
// present for that activity, so running it twice never duplicates
// anything. Never touches or deletes the old per-activity Pending
// sheets — see deleteLegacyPendingSheets() below once you're ready to
// remove them for good. The old per-activity Registrations/Visits
// sheets need no migration at all — this version reads those directly.
function migratePendingToSharedSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let migrated = 0;

  const sharedPending = getOrCreateSheet(PENDING_SHEET_NAME, PENDING_HEADERS);
  const existingKeys = new Set(sheetToObjects(sharedPending).map(r => r.activity + "|" + String(r.idNo).trim()));

  Object.keys(ACTIVITIES).forEach(key => {
    const activity = ACTIVITIES[key];
    const legacyPending = ss.getSheetByName(activity.legacyPendingSheet);
    if (!legacyPending) return;
    sheetToObjects(legacyPending).forEach(row => {
      const dedupeKey = activity.key + "|" + String(row.idNo).trim();
      if (existingKeys.has(dedupeKey)) return;
      existingKeys.add(dedupeKey);
      sharedPending.appendRow(PENDING_HEADERS.map(h => {
        if (h === "activity") return activity.key;
        if (h === "idNo" || h === "phone" || h === "emergencyPhone" || h === "relatedStaffIdNo") return sheetSafeText(row[h] || "");
        if (h === "date" || h === "time") return forceLiteralText(row[h] || "");
        return row[h] || "";
      }));
      migrated++;
    });
  });

  Logger.log(`Migrated ${migrated} pending row(s) into the shared Pending sheet.`);
}

// The old per-activity Pending sheets (e.g. "Pending - Gym") are unused
// dead weight once migratePendingToSharedSheet() has copied everything
// into the shared Pending sheet. This PERMANENTLY DELETES all 5 of
// them. Run by hand only (Run > deleteLegacyPendingSheets) once you've
// checked the shared Pending sheet looks right — Google Sheets' own
// version history (File > Version history) can still recover a deleted
// sheet for a while after, but this script can't undo it.
function deleteLegacyPendingSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let deleted = 0;
  Object.keys(ACTIVITIES).forEach(key => {
    const sheet = ss.getSheetByName(ACTIVITIES[key].legacyPendingSheet);
    if (sheet) { ss.deleteSheet(sheet); deleted++; }
  });
  Logger.log(`Deleted ${deleted} legacy per-activity Pending sheet(s).`);
}


// ------------------------------------------------------------------
// Splitting Registrations/Visits back into per-activity sheets, for
// anyone coming from the short-lived fully-merged layout (Pending,
// Registrations AND Visits were all one shared sheet each)
// ------------------------------------------------------------------

const SHARED_REGISTRATIONS_SHEET_NAME = "Registrations";
const SHARED_VISITS_SHEET_NAME = "Visits";

// Run this ONCE (Run > splitSharedRegistrationsAndVisits) if your sheet
// still has the old shared "Registrations"/"Visits" tabs from the
// fully-merged version of this project. Reads every row out of them
// and copies it into that row's own activity's Registrations/Visits
// sheet. Additive and safe to re-run — it only appends a row if one for
// that idNo (or, for Visits, that visitId) isn't already present in the
// destination sheet, so running it twice never duplicates anything.
// Never touches or deletes the old shared sheets — see
// deleteSharedRegistrationsAndVisitsSheets() below once you're ready to
// remove them for good. Pending needs no split — it's already the one
// shared sheet this version expects.
function splitSharedRegistrationsAndVisits() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let movedRegistrations = 0, movedVisits = 0;

  const sharedRegistrations = ss.getSheetByName(SHARED_REGISTRATIONS_SHEET_NAME);
  if (sharedRegistrations) {
    const byActivity = {};
    sheetToObjects(sharedRegistrations).forEach(row => {
      const key = row.activity;
      if (!byActivity[key]) byActivity[key] = [];
      byActivity[key].push(row);
    });
    Object.keys(byActivity).forEach(key => {
      const activity = ACTIVITIES[key];
      if (!activity) return; // unknown/stale activity value — skip rather than guess
      const sheet = getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
      const existingIds = new Set(sheetToObjects(sheet).map(r => String(r.idNo).trim()));
      byActivity[key].forEach(row => {
        const dedupeKey = String(row.idNo).trim();
        if (existingIds.has(dedupeKey)) return;
        existingIds.add(dedupeKey);
        sheet.appendRow(REGISTRATIONS_HEADERS.map(h => {
          if (h === "idNo" || h === "phone" || h === "emergencyPhone" || h === "relatedStaffIdNo") return sheetSafeText(row[h] || "");
          if (h === "date" || h === "time") return forceLiteralText(row[h] || "");
          return row[h] || "";
        }));
        movedRegistrations++;
      });
    });
  }

  const sharedVisits = ss.getSheetByName(SHARED_VISITS_SHEET_NAME);
  if (sharedVisits) {
    const byActivity = {};
    sheetToObjects(sharedVisits).forEach(row => {
      const key = row.activity;
      if (!byActivity[key]) byActivity[key] = [];
      byActivity[key].push(row);
    });
    Object.keys(byActivity).forEach(key => {
      const activity = ACTIVITIES[key];
      if (!activity) return;
      const sheet = getOrCreateSheet(activity.visitsSheet, VISIT_HEADERS);
      const existingVisitIds = new Set(sheetToObjects(sheet).map(r => String(r.visitId).trim()));
      byActivity[key].forEach(row => {
        const visitId = String(row.visitId || "").trim();
        if (!visitId || existingVisitIds.has(visitId)) return;
        existingVisitIds.add(visitId);
        sheet.appendRow(VISIT_HEADERS.map(h => {
          if (h === "idNo" || h === "phone") return sheetSafeText(row[h] || "");
          if (h === "date" || h === "timeIn" || h === "timeOut") return forceLiteralText(row[h] || "");
          return row[h] || "";
        }));
        movedVisits++;
      });
    });
  }

  Logger.log(`Split ${movedRegistrations} registration(s) and ${movedVisits} visit(s) out into per-activity sheets.`);
}

// The old shared "Registrations"/"Visits" sheets are unused dead weight
// once splitSharedRegistrationsAndVisits() has copied everything into
// the per-activity sheets. This PERMANENTLY DELETES both of them. Run
// by hand only (Run > deleteSharedRegistrationsAndVisitsSheets) once
// you've checked the per-activity sheets look right.
function deleteSharedRegistrationsAndVisitsSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let deleted = 0;
  [SHARED_REGISTRATIONS_SHEET_NAME, SHARED_VISITS_SHEET_NAME].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) { ss.deleteSheet(sheet); deleted++; }
  });
  Logger.log(`Deleted ${deleted} shared sheet(s).`);
}

// Reorders and color-codes every managed tab — purely cosmetic, doesn't
// touch any data. Run by hand (Run > organizeSheets) any time. Pending
// first, then each activity's own Registrations/Visits pair grouped
// together.
function organizeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let position = 1;

  const pendingSheet = ss.getSheetByName(PENDING_SHEET_NAME);
  if (pendingSheet) {
    ss.setActiveSheet(pendingSheet);
    ss.moveActiveSheet(position++);
    pendingSheet.setTabColor("#EDA100");
  }

  Object.keys(ACTIVITIES).forEach(key => {
    const activity = ACTIVITIES[key];
    const reg = ss.getSheetByName(activity.registrationsSheet);
    if (reg) {
      ss.setActiveSheet(reg);
      ss.moveActiveSheet(position++);
      reg.setTabColor("#15369E");
    }
    const vis = ss.getSheetByName(activity.visitsSheet);
    if (vis) {
      ss.setActiveSheet(vis);
      ss.moveActiveSheet(position++);
      vis.setTabColor("#1BAF7A");
    }
  });

  Logger.log("Sheet tabs reordered and color-coded.");
}

// Alerts moved off sheets entirely a while back (see
// getAlerts()/addAlert() above) — any "Alerts - X" tabs left over from
// before that change are unused dead weight too. This PERMANENTLY
// DELETES them. Run by hand only (Run > deleteUnusedAlertSheets) once
// you're sure you don't need their history.
function deleteUnusedAlertSheets() {
  const NAMES = [
    "Alerts - Gym",
    "Alerts - Leisure Tennis",
    "Alerts - Leisure Swimming",
    "Alerts - Tennis Lessons",
    "Alerts - Swimming Lessons"
  ];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let deleted = 0;
  NAMES.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) { ss.deleteSheet(sheet); deleted++; }
  });
  Logger.log(`Deleted ${deleted} unused Alerts sheet(s).`);
}


// ------------------------------------------------------------------
// Date-grouped Registrations sheets (see the top-of-file doc comment)
// ------------------------------------------------------------------

// Rebuilds one activity's Registrations sheet so rows are sorted
// newest-date-first and each date's block has a bold, shaded, merged
// header row above it. Safe to call any time; it re-derives everything
// from the real member rows (ignoring any existing header rows) so
// it's idempotent.
//
// NOT called from doApprove() — rewriting and reformatting the ENTIRE
// sheet on every single approval got slow as a sheet grew, and made
// approving a Family Package's several members back to back time out
// with "check the connection". Instead it's run automatically once a
// night by runNightlyMaintenance() (see regroupAllRegistrations()
// below and installNightlyMaintenanceTrigger()) — you can also run
// regroupAllRegistrations() by hand any time you don't want to wait
// for the nightly run.
function regroupRegistrationsByDate(activity) {
  const sheet = getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
  const lastCol = REGISTRATIONS_HEADERS.length;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const idColIndex = REGISTRATIONS_HEADERS.indexOf("idNo");
  const dateColIndex = REGISTRATIONS_HEADERS.indexOf("date");

  const allValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const allNotes = sheet.getRange(2, 1, lastRow - 1, lastCol).getNotes();
  const memberRows = allValues.filter((row, i) =>
    row.join("") !== "" && String(allNotes[i][idColIndex] || "").indexOf(DATE_HEADER_MARKER) !== 0
  );

  // Self-healing: normalize the date cell to plain text before grouping,
  // in case a row still holds a real Date object.
  memberRows.forEach(row => {
    row[dateColIndex] = cellToDisplayValue(row[dateColIndex], "date");
  });

  // Clear everything below the header row — content and formatting —
  // before rewriting, and unmerge any previous date-header rows.
  const clearRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, lastCol);
  try { clearRange.breakApart(); } catch (e) { /* nothing merged yet */ }
  clearRange.clearContent();
  clearRange.clearNote();
  clearRange.setBackground(null).setFontWeight("normal").setFontColor(null);

  if (memberRows.length === 0) return;

  const groups = new Map();
  memberRows.forEach(row => {
    const key = row[dateColIndex] || "Unknown date";
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  });

  const dateKeys = [...groups.keys()].sort((a, b) => {
    const da = parseDateSafe(a), db = parseDateSafe(b);
    if (!da || !db) return 0;
    return db - da; // newest date first
  });

  // Same guard as elsewhere: getValues() already stripped any leading
  // apostrophe from idNo/phone/emergencyPhone/date/time, so re-apply it
  // before writing these rows back with setValues() below.
  const phoneGuardCols = ["idNo", "phone", "emergencyPhone"].map(h => REGISTRATIONS_HEADERS.indexOf(h));
  const dateTimeGuardCols = ["date", "time"].map(h => REGISTRATIONS_HEADERS.indexOf(h));
  memberRows.forEach(row => {
    phoneGuardCols.forEach(i => { row[i] = sheetSafeText(row[i]); });
    dateTimeGuardCols.forEach(i => { row[i] = forceLiteralText(row[i]); });
  });

  const outRows = [];
  const headerRows = []; // { offset, key } — 0-based offset into outRows, plus that block's date key

  dateKeys.forEach(key => {
    const rowsForDate = groups.get(key);
    const headerRow = new Array(lastCol).fill("");
    headerRow[idColIndex] =
      `${dateLabelFor(key)}  —  ${rowsForDate.length} registration${rowsForDate.length === 1 ? "" : "s"}`;
    headerRows.push({ offset: outRows.length, key: key });
    outRows.push(headerRow);
    rowsForDate.forEach(r => outRows.push(r));
  });

  sheet.getRange(2, 1, outRows.length, lastCol).setValues(outRows);

  headerRows.forEach(({ offset, key }) => {
    const rowNum = offset + 2; // +2: row 1 is the column header, outRows is 0-based
    const range = sheet.getRange(rowNum, 1, 1, lastCol);
    range.merge();
    range.setFontWeight("bold");
    range.setBackground("#DCE4F0");
    range.setFontColor("#0F1D3B");
    range.setHorizontalAlignment("left");
    sheet.getRange(rowNum, idColIndex + 1, 1, 1).setNote(DATE_HEADER_MARKER + key);
  });
}

// Tidies every activity's Registrations sheet into dated, banner-grouped
// blocks. Runs automatically once a night via runNightlyMaintenance()
// (see installNightlyMaintenanceTrigger()) — takes a while on a large
// sheet, which is exactly why it runs there and not after every single
// approval. You can also run it by hand (Apps Script editor's function
// dropdown -> Run) any time you don't want to wait for the nightly run.
function regroupAllRegistrations() {
  Object.keys(ACTIVITIES).forEach(key => regroupRegistrationsByDate(ACTIVITIES[key]));
}

// Drops a freshly-approved row straight into its Registrations sheet's
// TODAY block — creating that block at the very top if this is the
// first approval of the day — instead of just appending it to the
// bottom, so date grouping is live the moment a registration is
// approved rather than waiting for the nightly regroupAllRegistrations().
// Cheap by construction: it only reads the idNo column's notes (to
// find/verify blocks) and rewrites the one block being touched, never
// the whole sheet — unlike a full regroupRegistrationsByDate() rebuild,
// this is safe to run on every single approval, including several back
// to back for a Family Package.
//
// Only ever called for a FRESH (non-renewal) approval, and dateKey is
// always today's date (every fresh approval is stamped with "now" —
// see approvePendingRow()), which is why this never needs to search
// for a sorted insertion point among older blocks: a brand new block
// is always the newest one, so it always goes at the very top. A
// renewal updates its member's row in place instead of moving it, so
// it can end up sitting inside a stale (non-today) block until the
// next nightly regroupAllRegistrations() run sorts it back out —
// that's an accepted, self-healing gap, not a bug.
function insertRegistrationIntoDateGroup(sheet, headers, regRowValues, dateKey) {
  const lastCol = headers.length;
  const idColIndex = headers.indexOf("idNo");
  const lastRow = sheet.getLastRow();
  const bannerNote = DATE_HEADER_MARKER + dateKey;

  if (lastRow < 2) {
    insertNewDateBlock(sheet, lastCol, idColIndex, regRowValues, dateKey, 2);
    return;
  }

  // Single-column note read (cheap) to find today's banner, if any,
  // and — by scanning forward until the next banner or the end of the
  // sheet — the last row currently in that block.
  const idNotes = sheet.getRange(2, idColIndex + 1, lastRow - 1, 1).getNotes();
  let bannerRow = -1;
  let blockEndRow = -1;
  for (let i = 0; i < idNotes.length; i++) {
    if (idNotes[i][0] !== bannerNote) continue;
    bannerRow = i + 2;
    blockEndRow = bannerRow;
    for (let j = i + 1; j < idNotes.length; j++) {
      if (String(idNotes[j][0] || "").indexOf(DATE_HEADER_MARKER) === 0) break; // next block starts here
      blockEndRow = j + 2;
    }
    break;
  }

  if (bannerRow === -1) {
    // No block for today yet — today is always the newest date this
    // can be, so the new block goes right after the column header.
    insertNewDateBlock(sheet, lastCol, idColIndex, regRowValues, dateKey, 2);
    return;
  }

  // Today's block already exists — add this member at the end of it
  // (inheriting that row's plain formatting, not the banner's bold
  // one), then bump the banner's count label.
  sheet.insertRowAfter(blockEndRow);
  const newRow = blockEndRow + 1;
  sheet.getRange(newRow, 1, 1, lastCol).setValues([regRowValues]);
  const newCount = (blockEndRow - bannerRow) + 1; // members already in the block, plus this one
  sheet.getRange(bannerRow, idColIndex + 1).setValue(
    `${dateLabelFor(dateKey)}  —  ${newCount} registration${newCount === 1 ? "" : "s"}`
  );
}

// Inserts a brand-new two-row (banner + one member) date block at
// atRow, formatted the same way regroupRegistrationsByDate() formats
// one. Used by insertRegistrationIntoDateGroup() above.
function insertNewDateBlock(sheet, lastCol, idColIndex, regRowValues, dateKey, atRow) {
  sheet.insertRowsBefore(atRow, 2);

  const bannerRange = sheet.getRange(atRow, 1, 1, lastCol);
  bannerRange.merge();
  bannerRange.setFontWeight("bold");
  bannerRange.setBackground("#DCE4F0");
  bannerRange.setFontColor("#0F1D3B");
  bannerRange.setHorizontalAlignment("left");
  sheet.getRange(atRow, idColIndex + 1).setValue(`${dateLabelFor(dateKey)}  —  1 registration`);
  sheet.getRange(atRow, idColIndex + 1).setNote(DATE_HEADER_MARKER + dateKey);

  sheet.getRange(atRow + 1, 1, 1, lastCol).setValues([regRowValues]);
}


// ------------------------------------------------------------------
// ONE-TIME REPAIR utilities
// ------------------------------------------------------------------

// Run once from the function dropdown (Run > repairPhoneNumbers), then
// redeploy. Switches idNo/phone/emergencyPhone columns to Plain Text
// so "#ERROR!" can't happen again, and recovers what it can from cells
// currently showing that error.
function repairPhoneNumbers() {
  const sheetsToRepair = [{ name: PENDING_SHEET_NAME, headers: PENDING_HEADERS }];
  Object.keys(ACTIVITIES).forEach(key => {
    sheetsToRepair.push({ name: ACTIVITIES[key].registrationsSheet, headers: REGISTRATIONS_HEADERS });
  });

  sheetsToRepair.forEach(({ name, headers }) => {
    const sheet = getOrCreateSheet(name, headers);
    ensureTextFormatForPhoneColumns(sheet, headers); // re-applied here on purpose — this is the manual repair path
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    ["idNo", "phone", "emergencyPhone"].forEach(h => {
      const col = headers.indexOf(h) + 1;
      if (col < 1) return;
      const range = sheet.getRange(2, col, lastRow - 1, 1);
      const formulas = range.getFormulas();
      const values = range.getValues();
      let changed = false;

      const fixed = values.map((row, i) => {
        const formula = formulas[i][0];
        if (formula && formula.toString().indexOf("=") === 0) {
          changed = true;
          return [formula.toString().slice(1)]; // drop the leading "="
        }
        return [row[0]];
      });

      if (changed) range.setValues(fixed);
    });
  });

  Logger.log("Phone/ID number formatting repaired. If any cells still show #ERROR!, " +
    "the original text couldn't be recovered automatically — retype those by hand in the sheet.");
}

// Run once from the function dropdown (Run > repairDateTimeColumns),
// then redeploy. Switches date/time columns to Plain Text and rewrites
// any cell that's still a real Date object as clean formatted text.
function repairDateTimeColumns() {
  const sheetsToRepair = [{ name: PENDING_SHEET_NAME, headers: PENDING_HEADERS }];
  Object.keys(ACTIVITIES).forEach(key => {
    const activity = ACTIVITIES[key];
    sheetsToRepair.push({ name: activity.registrationsSheet, headers: REGISTRATIONS_HEADERS });
    sheetsToRepair.push({ name: activity.visitsSheet, headers: VISIT_HEADERS });
  });

  sheetsToRepair.forEach(({ name, headers }) => {
    const sheet = getOrCreateSheet(name, headers);
    ensureTextFormatForPhoneColumns(sheet, headers); // re-applied here on purpose — this is the manual repair path
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    ["date", "time", "timeIn", "timeOut"].forEach(h => {
      const col = headers.indexOf(h) + 1;
      if (col < 1) return;
      const range = sheet.getRange(2, col, lastRow - 1, 1);
      const values = range.getValues();
      let changed = false;

      const fixed = values.map(row => {
        const displayVal = cellToDisplayValue(row[0], h);
        if (displayVal === row[0]) return [row[0]]; // wasn't a Date object, leave as-is
        changed = true;
        return [forceLiteralText(displayVal)];
      });

      if (changed) range.setValues(fixed);
    });
  });

  Logger.log("Date/time formatting repaired. Any date or time cells that had been " +
    "auto-converted by Sheets are now plain text, formatted as " + DATE_FORMAT + " / " + TIME_FORMAT + ".");
}


// ------------------------------------------------------------------
// Automatic 9pm sign-out
// ------------------------------------------------------------------

// Closes out every still-open visit (no timeOut yet), across every
// activity's own Visits sheet, as if that member had signed out at
// closing time — for anyone who used the facility but forgot to sign
// out themselves. Meant to run automatically once a day via a
// time-driven trigger — see installNightlyMaintenanceTrigger() below,
// which sets that up. Safe to run by hand too (Run > autoSignOutAt9pm)
// if you ever need to close everything out early.
function autoSignOutAt9pm() {
  const CLOSING_TIME_LABEL = "9:00 PM";
  let totalClosed = 0;

  Object.keys(ACTIVITIES).forEach(key => {
    const activity = ACTIVITIES[key];
    const visits = getOrCreateSheet(activity.visitsSheet, VISIT_HEADERS);
    const lastRow = visits.getLastRow();
    if (lastRow < 2) return;

    const idColIndex = VISIT_HEADERS.indexOf("idNo");
    const timeOutColIndex = VISIT_HEADERS.indexOf("timeOut");
    const values = visits.getRange(2, 1, lastRow - 1, VISIT_HEADERS.length).getValues();

    const registrations = getOrCreateSheet(activity.registrationsSheet, REGISTRATIONS_HEADERS);
    const regByIdNo = {};
    sheetToObjects(registrations).forEach(r => { regByIdNo[String(r.idNo).trim()] = r; });

    for (let i = 0; i < values.length; i++) {
      if (values[i][timeOutColIndex]) continue; // already signed out
      const rowNum = i + 2;
      visits.getRange(rowNum, timeOutColIndex + 1).setValue(CLOSING_TIME_LABEL);
      totalClosed++;

      // Same session-cap bookkeeping a normal checkout does (see the
      // "checkout" action above) — they used the facility today even
      // though they didn't sign out themselves.
      const idNo = String(values[i][idColIndex]).trim();
      const match = regByIdNo[idNo];
      const cfg = match && getDurationConfig(activity, match.duration);
      if (cfg && cfg.sessionCap) {
        const regIdx = findRowIndexByIdNo(registrations, idNo, REGISTRATIONS_HEADERS);
        if (regIdx !== -1) {
          const newUsed = (Number(match.sessionsUsed) || 0) + 1;
          registrations.getRange(regIdx, REGISTRATIONS_HEADERS.indexOf("sessionsUsed") + 1).setValue(newUsed);
        }
      }
    }
  });

  if (totalClosed > 0) {
    Logger.log(`Auto sign-out: closed ${totalClosed} open visit(s) across all activities.`);
  }
}

function runNightlyMaintenance() {
  autoSignOutAt9pm();
  regroupAllRegistrations();
}

// Run this ONCE from the function dropdown (Run > installNightlyMaintenanceTrigger),
// then approve the permissions prompt. Schedules runNightlyMaintenance()
// (the 9pm auto sign-out, then the date-grouped Registrations
// tidy-up — kept as its own wrapper in case more nightly jobs get
// added later) to run automatically every day at 9pm, in this
// project's time zone (Project Settings (gear icon) -> Time zone —
// set that first if it isn't already the venue's local time zone).
// Safe to re-run: it removes any existing trigger for this function
// (and the older autoSignOutAt9pm/compiled-sheet triggers, if you'd
// set either of those up before) first, so you'll never end up with
// duplicates firing the same night.
function installNightlyMaintenanceTrigger() {
  ["autoSignOutAt9pm", "runNightlyMaintenance"].forEach(fn => {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t);
    });
  });
  ScriptApp.newTrigger("runNightlyMaintenance")
    .timeBased()
    .everyDays(1)
    .atHour(21)
    .create();
  Logger.log("Installed: runNightlyMaintenance will now run automatically every day at 9pm.");
}
