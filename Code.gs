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
 *                                  read-only registrant data + visit log +
 *                                  sign-in/out kiosk; can only approve/reject
 *                                  WALK-INS, enforced server-side below)
 *   - swimming-front-desk.html   (same, for Leisure Swimming + Swimming Lessons)
 *
 * SETUP (fresh sheet):
 * 1. Create a new Google Sheet (sheets.new).
 * 2. Extensions -> Apps Script. Delete any starter code, paste this
 *    whole file in, save (disk icon / Ctrl+S / Cmd+S).
 * 3. From the function dropdown (next to Run/Debug), select "setup",
 *    click Run. Authorize when asked (Advanced -> "Go to (project)
 *    (unsafe)" -> Allow). This creates a Pending/Registrations/Visits/
 *    Alerts tab for EACH of the 5 activities (20 tabs total), each
 *    with the right headers. (The photo folder and export folder in
 *    Drive are created automatically the first time they're needed.)
 * 4. Deploy -> New deployment -> gear icon -> Web app.
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Deploy, authorize if asked (this version also asks for Drive
 *    access, since photos and exports are saved there), copy the
 *    URL ending in /exec.
 * 5. Paste that URL into SCRIPT_URL in EACH of the four HTML files
 *    listed above — they all talk to the same backend.
 * 6. Any time you edit this file again: Deploy -> Manage deployments
 *    -> pencil icon -> Version: New version -> Deploy, or the live
 *    URL won't see your change.
 *
 * HOW ACTIVITIES WORK:
 * - ACTIVITIES (below) is the single source of truth for the 5
 *   activities: which sheets they use, which member-code prefix,
 *   which categories are offered, which categories must supply their
 *   own ID vs get an auto-generated code, and which duration/plan
 *   options exist (with day-length and, for Swimming Lessons, a
 *   session cap).
 * - EVERY request (doGet view= and doPost action=) carries an
 *   "activity" key (gym / leisureTennis / leisureSwimming /
 *   tennisLessons / swimmingLessons) that selects which set of
 *   sheets to read/write. There is no cross-activity data — a member
 *   registered for the gym and for tennis lessons is two entirely
 *   separate rows (possibly with the same raw ID, if they used their
 *   student/staff ID both times) in two different sheets.
 *
 * HOW IDENTITY / MEMBER CODES WORK:
 * - idNo is the one field that identifies a member within an
 *   activity: it's what UG Student / UG Staff type in themselves,
 *   and what every other category is auto-assigned the moment they
 *   submit — a random unique code shaped "<prefix><7 digits>", e.g.
 *   G1234567 (Gym), T1234567 (Leisure Tennis), S1234567 (Leisure
 *   Swimming), TL1234567 (Tennis Lessons), SL1234567 (Swimming
 *   Lessons). idNo is also the row's unique key for approve/reject,
 *   and it's the "code" a member later types into Sign In / Sign Out.
 * - Because an auto-generated idNo is created at submission (not at
 *   approval), it can serve as the pending row's key right away — but
 *   the app never shows it to the member until their registration is
 *   actually approved, via the Sign In tab's phone-number lookup.
 * - A manually-entered ID (UG Student / UG Staff) is never prefixed —
 *   it's stored exactly as typed, same as before this multi-activity
 *   version.
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
 *   row — full name, and optionally that person's OWN medical
 *   conditions (stored in the same hasMedicalCondition/
 *   medicalConditionDetails columns the primary registrant uses), plus
 *   the SAME shared phone/email/address/emergency-contact — but no
 *   separate dob/gender/photo — each with its own auto-generated
 *   member code. See the "submit" handler below.
 * - Because every row in the family shares one phone number, the Sign
 *   In tab's "lookup" action (phone-number code retrieval) naturally
 *   returns every family member's code at once when the head enters
 *   that shared number — see "lookup" below.
 *
 * (Photo upload, e-signature, date-grouped Registrations sheet, the
 * Excel export, walk-ins, and renew/update-details all work exactly
 * as in the original single-activity version — see the inline
 * comments near each function below — just scoped per-activity now.)
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
    pendingSheet: "Pending - Gym",
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
    pendingSheet: "Pending - Leisure Tennis",
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
    pendingSheet: "Pending - Leisure Swimming",
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
    pendingSheet: "Pending - Tennis Lessons",
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
    pendingSheet: "Pending - Swimming Lessons",
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

const PHOTOS_FOLDER_NAME = "Registration Photos";
const EXPORT_FOLDER_NAME = "Registration Exports";

// A "Clear List" on the front desk's Registration Table doesn't touch
// the Google Sheet at all — it just records a per-activity cutoff
// timestamp here, and the Registration Table view (and therefore the
// Excel export, which is built from that same view) hides any
// registration approved at or before it. Sign-in/out/verify/lookup
// all read the Registrations sheet directly, not through this filter,
// so members are completely unaffected by a clear.
const VIEW_CLEARED_AT_PREFIX = "VIEW_CLEARED_AT_";

// Which Drive file a given activity's Excel export is linked to, if
// any — see saveExportFile()/parseDriveFileId() and the
// setExportFileId/exportLink/exportFile doPost/doGet actions below.
const EXPORT_FILE_ID_PREFIX = "EXPORT_FILE_ID_";

// Marker written as a NOTE (not the cell value) on the idNo cell of a
// synthetic "date header" row in a Registrations sheet, so the app
// can tell it apart from a real member row. It's a note rather than a
// value because that row's cells get merged into one wide banner, and
// a merge keeps the top-left cell's VALUE as the visible text — that
// cell needs to show the human-readable date label, not this marker.
const DATE_HEADER_MARKER = "§DATE_HEADER§";


// Columns shared by every activity's Pending/Registrations sheets.
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
// share the same phone number, which is how "lookup" (the Sign In tab's
// phone-number code retrieval) returns every family member's code at
// once — there's no separate "family group" column needed for that.
// "familyRelationship" ("relationship to you") is likewise only ever
// populated for an additional family member's row.
//
// NOTE: HEADERS is positional — every existing Pending/Registrations
// sheet already has its physical columns laid out in this exact order,
// and the header row itself is only (re)written for a brand-new sheet
// (see getOrCreateSheet). A new field must always be appended at the
// END of this array, never inserted in the middle, or every column
// after it will silently misalign with already-written sheets.
const HEADERS = [
  "idNo", "name", "dob", "gender", "nationality", "hasMedicalCondition",
  "medicalConditionDetails", "address",
  "email", "phone", "department", "class",
  "relatedStaffName", "relatedStaffIdNo", "staffRelationship",
  "duration", "sessionsUsed",
  "date", "time", "emergencyName", "emergencyPhone", "emergencyRelationship",
  "photoUrl", "signatureUrl", "isRenewal",
  "familyRelationship"
];
const VISIT_HEADERS = ["visitId", "idNo", "name", "class", "date", "timeIn", "timeOut", "phone"];

// Expired/used-up-membership sign-in attempts, so every front desk for
// that activity (main and satellite alike) can be alerted even when
// they aren't the one watching that sign-in. These are deliberately
// NOT a sheet — there's no per-activity "Alerts" tab — they're a
// short-lived script property instead: a small JSON array, keyed per
// activity, that a staff member's "Dismiss" click removes an alert
// from. See getAlerts()/addAlert()/dismissAlert() below.
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
  Object.keys(ACTIVITIES).forEach(key => {
    const activity = ACTIVITIES[key];
    getOrCreateSheet(activity.pendingSheet, HEADERS);
    getOrCreateSheet(activity.registrationsSheet, HEADERS);
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
  }
  ensureTextFormatForPhoneColumns(sheet, headers);
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
// stops it, for every sheet this script touches (called on every
// request, not just at setup).
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
  const notes = range.getNotes();
  const headers = values.shift();
  notes.shift();
  const idIdx = headers.indexOf("idNo");
  return values
    .map((row, i) => ({ row, note: idIdx === -1 ? "" : notes[i][idIdx] }))
    .filter(({ row }) => row.join("") !== "")
    .filter(({ note }) => note !== DATE_HEADER_MARKER)
    .map(({ row }) => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = cellToDisplayValue(row[i], h));
      return obj;
    });
}


function findRowIndexByIdNo(sheet, idNo) {
  const idColIndex = HEADERS.indexOf("idNo") + 1; // 1-based
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, idColIndex, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(idNo).trim()) return i + 2;
  }
  return -1;
}


function idNoExists(activity, idNo) {
  const pending = getOrCreateSheet(activity.pendingSheet, HEADERS);
  const registrations = getOrCreateSheet(activity.registrationsSheet, HEADERS);
  return findRowIndexByIdNo(pending, idNo) !== -1 || findRowIndexByIdNo(registrations, idNo) !== -1;
}


// Draws a random "<prefix><7 digits>" code and keeps re-rolling until
// it finds one that isn't already used by a Pending or Registrations
// row IN THIS ACTIVITY (codes from different activities can never
// collide anyway, since each activity has its own prefix). Capped at
// MAX_ATTEMPTS so this can never spin forever. Returns null if it
// still comes up empty — the caller must handle that rather than
// assume a code back.
function generateUniqueIdNo(activity) {
  const MAX_ATTEMPTS = 200;
  for (let i = 0; i < MAX_ATTEMPTS; i++) {
    const digits = String(Math.floor(1000000 + Math.random() * 9000000)); // 7 digits
    const code = activity.prefix + digits;
    if (!idNoExists(activity, code)) return code;
  }
  return null;
}


// ------------------------------------------------------------------
// Photo storage
// ------------------------------------------------------------------

function getPhotosFolder() {
  const folders = DriveApp.getFoldersByName(PHOTOS_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTOS_FOLDER_NAME);
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
// every activity has its own letter prefix).
function savePhotoAndGetUrl(filenameBase, base64Data, mimeType) {
  if (!base64Data) return "";
  try {
    const cleaned = base64Data.indexOf(",") !== -1 ? base64Data.split(",")[1] : base64Data;
    const type = mimeType || "image/jpeg";
    const ext = type.indexOf("png") !== -1 ? "png" : "jpg";
    const bytes = Utilities.base64Decode(cleaned);
    const blob = Utilities.newBlob(bytes, type, `${filenameBase}.${ext}`);
    const folder = getPhotosFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    return "";
  }
}


// ------------------------------------------------------------------
// Date-grouped Registrations sheet
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

// Rebuilds one activity's Registrations sheet so rows are sorted
// newest-date-first and each date's block has a bold, shaded, merged
// header row above it. Safe to call any time; it re-derives everything
// from the real member rows (ignoring any existing header rows) so
// it's idempotent.
function regroupRegistrationsByDate(activity) {
  const sheet = getOrCreateSheet(activity.registrationsSheet, HEADERS);
  const lastCol = HEADERS.length;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const idColIndex = HEADERS.indexOf("idNo");
  const dateColIndex = HEADERS.indexOf("date");

  const allValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const allNotes = sheet.getRange(2, 1, lastRow - 1, lastCol).getNotes();
  const memberRows = allValues.filter((row, i) =>
    row.join("") !== "" && allNotes[i][idColIndex] !== DATE_HEADER_MARKER
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
  const phoneGuardCols = ["idNo", "phone", "emergencyPhone"].map(h => HEADERS.indexOf(h));
  const dateTimeGuardCols = ["date", "time"].map(h => HEADERS.indexOf(h));
  memberRows.forEach(row => {
    phoneGuardCols.forEach(i => { row[i] = sheetSafeText(row[i]); });
    dateTimeGuardCols.forEach(i => { row[i] = forceLiteralText(row[i]); });
  });

  const outRows = [];
  const headerRowOffsets = []; // 0-based offsets into outRows

  dateKeys.forEach(key => {
    const rowsForDate = groups.get(key);
    const headerRow = new Array(lastCol).fill("");
    headerRow[idColIndex] =
      `${dateLabelFor(key)}  —  ${rowsForDate.length} registration${rowsForDate.length === 1 ? "" : "s"}`;
    headerRowOffsets.push(outRows.length);
    outRows.push(headerRow);
    rowsForDate.forEach(r => outRows.push(r));
  });

  sheet.getRange(2, 1, outRows.length, lastCol).setValues(outRows);

  headerRowOffsets.forEach(offset => {
    const rowNum = offset + 2; // +2: row 1 is the column header, outRows is 0-based
    const range = sheet.getRange(rowNum, 1, 1, lastCol);
    range.merge();
    range.setFontWeight("bold");
    range.setBackground("#DCE4F0");
    range.setFontColor("#0F1D3B");
    range.setHorizontalAlignment("left");
    sheet.getRange(rowNum, idColIndex + 1, 1, 1).setNote(DATE_HEADER_MARKER);
  });
}


// ------------------------------------------------------------------
// Excel export -> single Drive file per activity, same link every time
// ------------------------------------------------------------------
//
// Keeps re-using the exact same Drive file (same file ID, same share
// link) across every export of a given activity, by overwriting its
// content in place instead of deleting and recreating it. That
// requires the "Drive API" advanced service to be turned on in this
// Apps Script project (Services (+) -> Drive API -> Add, then
// redeploy). Each activity's file ID is remembered in Script
// Properties under its own key the first time it's exported.
//
// If the Drive API service isn't enabled yet, this automatically
// falls back to the old "trash old copy, create a new one" approach
// (same name/folder, but a new link each time) so exporting still
// works — you just won't get a permanent link until you enable it.

function getExportFolder() {
  const folders = DriveApp.getFoldersByName(EXPORT_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(EXPORT_FOLDER_NAME);
}

function driveApiIsEnabled() {
  // "Drive" is only defined once the Drive API advanced service has
  // been added via Services (+) in the Apps Script editor.
  return typeof Drive !== "undefined" && !!Drive.Files;
}

function saveExportFile(activity, base64Data) {
  const cleaned = base64Data.indexOf(",") !== -1 ? base64Data.split(",")[1] : base64Data;
  const bytes = Utilities.base64Decode(cleaned);
  const mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const fileName = `${activity.label} Registrations.xlsx`;
  const blob = Utilities.newBlob(bytes, mimeType, fileName);

  const props = PropertiesService.getScriptProperties();
  const propKey = EXPORT_FILE_ID_PREFIX + activity.key;
  const storedFileId = props.getProperty(propKey);

  if (storedFileId) {
    if (!driveApiIsEnabled()) {
      // A file is linked, but without the Drive API advanced service
      // there's no way to overwrite an arbitrary file's binary content
      // in place — say so plainly instead of silently exporting
      // somewhere else, which would look like the link isn't working.
      return {
        url: null,
        usedLinkedFile: false,
        warning: "A file is linked, but the Drive API advanced service isn't enabled in Apps Script (Services → Drive API), so it can't be updated in place. Nothing was written to it this time."
      };
    }
    try {
      Drive.Files.update({}, storedFileId, blob);
      return { url: "https://drive.google.com/file/d/" + storedFileId + "/view", usedLinkedFile: true };
    } catch (err) {
      // The stored file was probably deleted/moved outside the app —
      // fall through and create a fresh one below, but say so, since
      // the export otherwise silently lands somewhere the user didn't
      // expect.
      const folder = getExportFolder();
      const existing = folder.getFilesByName(fileName);
      while (existing.hasNext()) existing.next().setTrashed(true);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      props.deleteProperty(propKey);
      return {
        url: file.getUrl(),
        usedLinkedFile: false,
        warning: "The linked file couldn't be reached (it may have been deleted or moved) — exported to a new file instead, and the link was cleared. Please re-link."
      };
    }
  }

  const folder = getExportFolder();
  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  if (driveApiIsEnabled()) {
    props.setProperty(propKey, file.getId());
  }
  return { url: file.getUrl(), usedLinkedFile: false };
}

// Parses either a bare Drive file ID or a full Drive URL
// (".../file/d/<ID>/view", "...?id=<ID>", etc.) into just the ID.
function parseDriveFileId(input) {
  const s = String(input || "").trim();
  if (!s) return "";
  const m = s.match(/\/d\/([-\w]{10,})/) || s.match(/[?&]id=([-\w]{10,})/);
  if (m) return m[1];
  return /^[-\w]{10,}$/.test(s) ? s : "";
}


// ------------------------------------------------------------------
// Shared approve/reject logic (used by both the full main-app actions
// and the walk-in-only actions the satellite front desks are allowed
// to call)
// ------------------------------------------------------------------

// Approves a Pending row. A Walk-in never becomes a Registrations row —
// it's a one-off visit, so approving it writes a Visits row directly
// (checked in right now, no code needed later) and removes it from
// Pending. A renewal request overwrites the member's EXISTING
// Registrations row (new duration, expiry restarted from right now,
// sessionsUsed reset to blank) instead of appending a duplicate row.
function doApprove(activity, idNo) {
  const pending = getOrCreateSheet(activity.pendingSheet, HEADERS);
  const idx = findRowIndexByIdNo(pending, idNo);
  if (idx === -1) return ok({ message: "Already handled" });
  const rowValues = pending.getRange(idx, 1, 1, HEADERS.length).getValues()[0];

  if (String(rowValues[HEADERS.indexOf("duration")]).trim() === "Walk-in") {
    const visits = getOrCreateSheet(activity.visitsSheet, VISIT_HEADERS);
    const now = new Date();
    visits.appendRow([
      Utilities.getUuid(),
      sheetSafeText(rowValues[HEADERS.indexOf("idNo")]),
      rowValues[HEADERS.indexOf("name")],
      rowValues[HEADERS.indexOf("class")],
      formatDateMDY(now), now.toLocaleTimeString(), "",
      sheetSafeText(rowValues[HEADERS.indexOf("phone")])
    ]);
    pending.deleteRow(idx);
    return ok({ idNo: idNo });
  }

  // getValues() strips any leading apostrophe on the way out, so a
  // value like "+233 24 123 4567" comes back plain again — re-guard it
  // before this appendRow()/setValue() re-triggers the same formula
  // parsing.
  ["idNo", "phone", "emergencyPhone"].forEach(h => {
    const i = HEADERS.indexOf(h);
    rowValues[i] = sheetSafeText(rowValues[i]);
  });

  // Stamp "date"/"time" with the actual moment of approval — that's
  // what the expiry countdown is based on, and it's also the date the
  // header-grouping below groups rows under.
  const approvedNow = new Date();
  rowValues[HEADERS.indexOf("date")] = forceLiteralText(formatDateMDY(approvedNow));
  rowValues[HEADERS.indexOf("time")] = forceLiteralText(approvedNow.toLocaleTimeString());
  // A freshly (re)approved package always starts with 0 sessions used.
  rowValues[HEADERS.indexOf("sessionsUsed")] = "";

  const isRenewalIdx = HEADERS.indexOf("isRenewal");
  const isRenewal = String(rowValues[isRenewalIdx]).trim().toUpperCase() === "TRUE";
  rowValues[isRenewalIdx] = ""; // flag is spent once applied — never carried into Registrations

  const registrations = getOrCreateSheet(activity.registrationsSheet, HEADERS);

  if (isRenewal) {
    const idNoVal = String(rowValues[HEADERS.indexOf("idNo")]).replace(/^'/, "").trim();
    const regIdx = findRowIndexByIdNo(registrations, idNoVal);
    if (regIdx !== -1) {
      registrations.getRange(regIdx, HEADERS.indexOf("duration") + 1).setValue(rowValues[HEADERS.indexOf("duration")]);
      registrations.getRange(regIdx, HEADERS.indexOf("date") + 1).setValue(rowValues[HEADERS.indexOf("date")]);
      registrations.getRange(regIdx, HEADERS.indexOf("time") + 1).setValue(rowValues[HEADERS.indexOf("time")]);
      registrations.getRange(regIdx, HEADERS.indexOf("sessionsUsed") + 1).setValue("");
      pending.deleteRow(idx);
      regroupRegistrationsByDate(activity);
      return ok({ idNo: idNo });
    }
    // Member's row is gone somehow (e.g. deleted by hand) — fall
    // through and append the clone as a fresh row instead of silently
    // dropping the request.
  }

  registrations.appendRow(rowValues);
  pending.deleteRow(idx);
  regroupRegistrationsByDate(activity);
  return ok({ idNo: idNo });
}

function doReject(activity, idNo) {
  const pending = getOrCreateSheet(activity.pendingSheet, HEADERS);
  const idx = findRowIndexByIdNo(pending, idNo);
  if (idx === -1) return ok({ message: "Already handled" });
  pending.deleteRow(idx);
  return ok({});
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

function doGet(e) {
  try {
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
    if (view === 'exportLink') {
      const fileId = PropertiesService.getScriptProperties().getProperty(EXPORT_FILE_ID_PREFIX + activity.key);
      if (!fileId) return ok({ linked: false });
      try {
        const f = DriveApp.getFileById(fileId);
        return ok({ linked: true, fileId: fileId, fileName: f.getName(), fileUrl: f.getUrl(), driveApiEnabled: driveApiIsEnabled() });
      } catch (err) {
        return ok({ linked: true, fileId: fileId, fileName: null, fileUrl: null, driveApiEnabled: driveApiIsEnabled() });
      }
    }
    if (view === 'exportFile') {
      // Lets the export button read back whatever's currently in the
      // linked file (if any) so it can merge in only genuinely new
      // rows instead of duplicating members already recorded there.
      const fileId = PropertiesService.getScriptProperties().getProperty(EXPORT_FILE_ID_PREFIX + activity.key);
      if (!fileId) return ok({ exists: false });
      try {
        const blob = DriveApp.getFileById(fileId).getBlob();
        return ok({ exists: true, base64: Utilities.base64Encode(blob.getBytes()) });
      } catch (err) {
        return ok({ exists: false });
      }
    }
    if (view === 'pending') {
      const sheet = getOrCreateSheet(activity.pendingSheet, HEADERS);
      return ok({ rows: sheetToObjects(sheet) });
    }
    const sheet = getOrCreateSheet(activity.registrationsSheet, HEADERS);
    let rows = sheetToObjects(sheet);
    const clearedAt = PropertiesService.getScriptProperties().getProperty(VIEW_CLEARED_AT_PREFIX + activity.key);
    if (clearedAt) {
      const cutoffMs = new Date(clearedAt).getTime();
      rows = rows.filter(r => registrationTimestampMs(r) > cutoffMs);
    }
    return ok({ rows: rows, clearedAt: clearedAt || null });
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

      // ID number is optional for categories NOT in idRequiredCategories
      // — an auto-generated "<prefix><7 digits>" code is used if they
      // don't have or didn't provide one. idRequiredCategories (UG
      // Student / UG Staff) must supply their own.
      const idRequired = activity.idRequiredCategories.indexOf(data.class) !== -1;
      let idNo;
      if (!idRequired) {
        idNo = String(data.idNo || "").trim();
        if (!idNo) {
          idNo = generateUniqueIdNo(activity);
          if (!idNo) {
            return errorMsg("Couldn't generate a member code right now — the code pool may be full. Please ask the front desk to register you with a manual ID number instead.");
          }
        } else if (idNoExists(activity, idNo)) {
          return errorMsg("This ID number is already registered or pending approval.");
        }
      } else {
        idNo = String(data.idNo || "").trim();
        if (!idNo) return errorMsg("An ID number is required for this category.");
        if (idNoExists(activity, idNo)) return errorMsg("This ID number is already registered or pending approval.");
      }

      // Name-first filenames (idNo tucked in parentheses for uniqueness)
      // so photos/signatures can be found by applicant name in Drive.
      const applicantFileName = sanitizeForFilename(data.name);
      const photoUrl = savePhotoAndGetUrl(`${applicantFileName} (${idNo})`, data.photoBase64, data.photoMimeType);
      const signatureUrl = savePhotoAndGetUrl(`${applicantFileName} (${idNo}) - Signature`, data.signatureBase64, data.signatureMimeType);

      // A Family Package registration isn't one row for the whole
      // family — see the HEADERS comment above. The person filling the
      // form (idNo/photo/etc. above) is one full row; every additional
      // family member they list gets their own lightweight row below
      // — name, and optionally their own medical conditions (stored in
      // the same hasMedicalCondition/medicalConditionDetails columns
      // the primary registrant uses) — generated and validated up
      // front so the whole submission fails cleanly (nothing written)
      // rather than partially, if the code pool or the 5-person cap is
      // hit.
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
          const extraIdNo = generateUniqueIdNo(activity);
          if (!extraIdNo) {
            return errorMsg("Couldn't generate member codes for the whole family right now — the code pool may be full. Please ask the front desk to register the family manually instead.");
          }
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

      const sheet = getOrCreateSheet(activity.pendingSheet, HEADERS);
      sheet.appendRow(HEADERS.map(h => {
        if (h === "idNo") return sheetSafeText(idNo);
        if (h === "photoUrl") return photoUrl;
        if (h === "signatureUrl") return signatureUrl;
        if (h === "sessionsUsed") return "";
        if (h === "phone" || h === "emergencyPhone" || h === "relatedStaffIdNo") return sheetSafeText(data[h] || "");
        // Force "date"/"time" to literal text too — otherwise Sheets
        // silently converts them to real date/time values.
        if (h === "date" || h === "time") return forceLiteralText(data[h] || "");
        return data[h] || "";
      }));

      extraFamilyMembers.forEach(member => {
        sheet.appendRow(HEADERS.map(h => {
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
          // dob/nationality/department/photo/signature/sessionsUsed/
          // isRenewal are all left blank for an additional family
          // member — only their name, gender, relationship to the
          // primary registrant, optional medical conditions, and the
          // shared contact details are collected.
          return "";
        }));
      });

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


    // ---- Restricted actions for the satellite (tennis/swimming) front
    // desks: approvals there are ONLY allowed on a Walk-in row. This is
    // enforced here server-side (not just by hiding the button in the
    // satellite UI) since both apps call the same backend URL. ----
    if (action === "approveWalkin" || action === "rejectWalkin") {
      const pending = getOrCreateSheet(activity.pendingSheet, HEADERS);
      const idx = findRowIndexByIdNo(pending, data.idNo);
      if (idx === -1) return ok({ message: "Already handled" });
      const duration = String(pending.getRange(idx, HEADERS.indexOf("duration") + 1).getValue()).trim();
      if (duration !== "Walk-in") {
        return errorMsg("This front desk can only approve or reject walk-ins — new registrations and renewals need the main front desk.");
      }
      return action === "approveWalkin" ? doApprove(activity, data.idNo) : doReject(activity, data.idNo);
    }


    if (action === "checkin") {
      const registrations = getOrCreateSheet(activity.registrationsSheet, HEADERS);
      const rows = sheetToObjects(registrations);
      const code = String(data.code || "").trim();
      const match = rows.find(r => String(r.idNo).trim() === code);
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
      visits.appendRow([
        Utilities.getUuid(), match.idNo, match.name, match.class,
        formatDateMDY(now), now.toLocaleTimeString(), "",
      ]);
      return ok({ member: match });
    }


    if (action === "checkout") {
      const registrations = getOrCreateSheet(activity.registrationsSheet, HEADERS);
      const rows = sheetToObjects(registrations);
      const code = String(data.code || "").trim();
      const match = rows.find(r => String(r.idNo).trim() === code);
      if (!match) return errorMsg("Code not recognized");

      const visits = getOrCreateSheet(activity.visitsSheet, VISIT_HEADERS);
      const lastRow = visits.getLastRow();
      if (lastRow < 2) return errorMsg("No sign-in found for this code today. Please sign in first.");

      const idColIndex = VISIT_HEADERS.indexOf("idNo");
      const timeOutColIndex = VISIT_HEADERS.indexOf("timeOut");
      const dateColIndex = VISIT_HEADERS.indexOf("date");
      const values = visits.getRange(2, 1, lastRow - 1, VISIT_HEADERS.length).getValues();

      const todayStr = formatDateMDY(new Date());
      let targetRow = -1;
      for (let i = values.length - 1; i >= 0; i--) {
        const row = values[i];
        if (String(row[idColIndex]).trim() === match.idNo && !row[timeOutColIndex] && row[dateColIndex] === todayStr) {
          targetRow = i + 2; // sheet row number
          break;
        }
      }
      if (targetRow === -1) return errorMsg("No open sign-in found for this code today. Please sign in first.");

      visits.getRange(targetRow, timeOutColIndex + 1).setValue(new Date().toLocaleTimeString());

      // Duration has a session cap (Swimming Lessons' package) — a
      // session only counts as "used" once the member actually signs
      // out, not when they sign in (so a session in progress doesn't
      // get counted early, and a forgotten sign-in with no sign-out
      // doesn't burn a session at all).
      const cfg = getDurationConfig(activity, match.duration);
      if (cfg && cfg.sessionCap) {
        const regIdx = findRowIndexByIdNo(registrations, match.idNo);
        if (regIdx !== -1) {
          const newUsed = (Number(match.sessionsUsed) || 0) + 1;
          registrations.getRange(regIdx, HEADERS.indexOf("sessionsUsed") + 1).setValue(newUsed);
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
      if (lastRow < 2) return errorMsg("No sign-in found for this phone number today. Please sign in first.");

      const phoneColIndex = VISIT_HEADERS.indexOf("phone");
      const timeOutColIndex = VISIT_HEADERS.indexOf("timeOut");
      const dateColIndex = VISIT_HEADERS.indexOf("date");
      const values = visits.getRange(2, 1, lastRow - 1, VISIT_HEADERS.length).getValues();

      const todayStr = formatDateMDY(new Date());
      let targetRow = -1;
      for (let i = values.length - 1; i >= 0; i--) {
        const row = values[i];
        if (String(row[phoneColIndex]).trim() === phone && !row[timeOutColIndex] && row[dateColIndex] === todayStr) {
          targetRow = i + 2; // sheet row number
          break;
        }
      }
      if (targetRow === -1) {
        return errorMsg("No open sign-in found for this phone number today. Please sign in first, or ask the front desk.");
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

      const photoColIndex = HEADERS.indexOf("photoUrl") + 1;
      const nameColIndex = HEADERS.indexOf("name") + 1;

      const pending = getOrCreateSheet(activity.pendingSheet, HEADERS);
      let idx = findRowIndexByIdNo(pending, idNo);
      let targetSheet = null;
      if (idx !== -1) {
        targetSheet = pending;
      } else {
        const registrations = getOrCreateSheet(activity.registrationsSheet, HEADERS);
        idx = findRowIndexByIdNo(registrations, idNo);
        if (idx !== -1) targetSheet = registrations;
      }
      if (!targetSheet) return errorMsg("Code not recognized");

      const applicantName = targetSheet.getRange(idx, nameColIndex).getValue();
      const applicantFileName = sanitizeForFilename(applicantName);
      const photoUrl = savePhotoAndGetUrl(`${applicantFileName} (${idNo})`, data.photoBase64, data.photoMimeType);
      if (!photoUrl) return errorMsg("Couldn't save the photo — please try again.");

      targetSheet.getRange(idx, photoColIndex).setValue(photoUrl);
      return ok({});
    }


    if (action === "lookup") {
      // Returns EVERY approved row sharing this phone number, not just
      // one — so a Family Package's shared contact number retrieves
      // every family member's code at once (each family member is its
      // own row — see the HEADERS comment above), and anyone else who
      // happens to share a phone number with another registrant sees
      // all of theirs too.
      const phone = String(data.phone || "").trim();

      const pending = getOrCreateSheet(activity.pendingSheet, HEADERS);
      const pendingCount = sheetToObjects(pending).filter(r => String(r.phone).trim() === phone).length;

      const registrations = getOrCreateSheet(activity.registrationsSheet, HEADERS);
      const approvedMembers = sheetToObjects(registrations).filter(r => String(r.phone).trim() === phone);

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
      const registrations = getOrCreateSheet(activity.registrationsSheet, HEADERS);
      const approvedMembers = sheetToObjects(registrations).filter(r => codes.indexOf(String(r.idNo).trim()) !== -1);
      return ok({ approvedMembers: approvedMembers });
    }


    if (action === "verify") {
      // Looks a member up by code for the Renew tab's gate — deliberately
      // does NOT log a Visits row (unlike "checkin"). Approved members only.
      const registrations = getOrCreateSheet(activity.registrationsSheet, HEADERS);
      const code = String(data.code || "").trim();
      if (!code) return errorMsg("Enter your code or ID number.");
      const match = sheetToObjects(registrations).find(r => String(r.idNo).trim() === code);
      if (!match) return errorMsg("Code not recognized");
      return ok({ member: match });
    }


    if (action === "requestRenewal") {
      const code = String(data.code || "").trim();
      const duration = String(data.duration || "").trim();
      if (!code) return errorMsg("Enter your code or ID number.");
      if (!duration) return errorMsg("Please choose a plan.");

      const registrations = getOrCreateSheet(activity.registrationsSheet, HEADERS);
      const idx = findRowIndexByIdNo(registrations, code);
      if (idx === -1) return errorMsg("Code not recognized");

      const currentClass = registrations.getRange(idx, HEADERS.indexOf("class") + 1).getValue();
      const durCfg = getDurationConfig(activity, duration);
      if (!durCfg || !durationAllowedForCategory(durCfg, currentClass)) {
        return errorMsg("That plan isn't available for your category.");
      }

      const pending = getOrCreateSheet(activity.pendingSheet, HEADERS);
      if (findRowIndexByIdNo(pending, code) !== -1) {
        return errorMsg("You already have a request awaiting approval at the front desk.");
      }

      const rowValues = registrations.getRange(idx, 1, 1, HEADERS.length).getValues()[0];
      const now = new Date();
      const pendingRow = HEADERS.map((h, i) => {
        if (h === "duration") return duration;
        if (h === "date") return forceLiteralText(formatDateMDY(now));
        if (h === "time") return forceLiteralText(now.toLocaleTimeString());
        if (h === "isRenewal") return "TRUE";
        if (h === "idNo" || h === "phone" || h === "emergencyPhone") return sheetSafeText(rowValues[i]);
        return rowValues[i];
      });
      pending.appendRow(pendingRow);
      return ok({});
    }


    if (action === "updateDetails") {
      const code = String(data.code || "").trim();
      if (!code) return errorMsg("Enter your code or ID number.");

      const registrations = getOrCreateSheet(activity.registrationsSheet, HEADERS);
      const idx = findRowIndexByIdNo(registrations, code);
      if (idx === -1) return errorMsg("Code not recognized");

      // Only these fields are editable from the Renew tab's form —
      // idNo, name, class, duration, dates, and photo stay untouched.
      const editableFields = ["phone", "email", "address", "emergencyName", "emergencyPhone", "emergencyRelationship"];
      editableFields.forEach(h => {
        if (data[h] === undefined) return;
        const val = (h === "phone" || h === "emergencyPhone") ? sheetSafeText(data[h]) : data[h];
        registrations.getRange(idx, HEADERS.indexOf(h) + 1).setValue(val);
      });

      const rowValues = registrations.getRange(idx, 1, 1, HEADERS.length).getValues()[0];
      const member = {};
      HEADERS.forEach((h, i) => member[h] = rowValues[i]);
      return ok({ member: member });
    }


    if (action === "saveExport") {
      if (!data.fileBase64) return errorMsg("No file data received.");
      const result = saveExportFile(activity, data.fileBase64);
      return ok(result);
    }

    if (action === "setExportFileId") {
      const fileId = parseDriveFileId(data.fileId);
      if (!fileId) return errorMsg("Couldn't find a valid Drive file ID in that link.");
      let fileName;
      try {
        fileName = DriveApp.getFileById(fileId).getName();
      } catch (err) {
        return errorMsg("Couldn't open that file — check the link and make sure it's shared with this app's Google account.");
      }
      PropertiesService.getScriptProperties().setProperty(EXPORT_FILE_ID_PREFIX + activity.key, fileId);
      return ok({ fileId: fileId, fileName: fileName, driveApiEnabled: driveApiIsEnabled() });
    }

    if (action === "clearExportFileId") {
      PropertiesService.getScriptProperties().deleteProperty(EXPORT_FILE_ID_PREFIX + activity.key);
      return ok({});
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
// ONE-TIME REPAIR utilities (same fixes as the original single-activity
// version, just looped across every activity's sheets)
// ------------------------------------------------------------------

// Run once from the function dropdown (Run > repairPhoneNumbers), then
// redeploy. Switches idNo/phone/emergencyPhone columns to Plain Text
// so "#ERROR!" can't happen again, and recovers what it can from cells
// currently showing that error.
function repairPhoneNumbers() {
  Object.keys(ACTIVITIES).forEach(key => {
    const activity = ACTIVITIES[key];
    [activity.pendingSheet, activity.registrationsSheet].forEach(name => {
      const sheet = getOrCreateSheet(name, HEADERS); // also (re)applies text formatting
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;

      ["idNo", "phone", "emergencyPhone"].forEach(h => {
        const col = HEADERS.indexOf(h) + 1;
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
  });

  Logger.log("Phone/ID number formatting repaired across all activities. If any cells still show #ERROR!, " +
    "the original text couldn't be recovered automatically — retype those by hand in the sheet.");
}

// Run once from the function dropdown (Run > repairDateTimeColumns),
// then redeploy. Switches date/time columns to Plain Text and rewrites
// any cell that's still a real Date object as clean formatted text.
function repairDateTimeColumns() {
  const sheetsToRepair = [];
  Object.keys(ACTIVITIES).forEach(key => {
    const activity = ACTIVITIES[key];
    sheetsToRepair.push({ name: activity.pendingSheet, headers: HEADERS });
    sheetsToRepair.push({ name: activity.registrationsSheet, headers: HEADERS });
    sheetsToRepair.push({ name: activity.visitsSheet, headers: VISIT_HEADERS });
  });

  sheetsToRepair.forEach(({ name, headers }) => {
    const sheet = getOrCreateSheet(name, headers); // also (re)applies text formatting
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

  Logger.log("Date/time formatting repaired across all activities. Any date or time cells that had been " +
    "auto-converted by Sheets are now plain text, formatted as " + DATE_FORMAT + " / " + TIME_FORMAT + ".");
}
