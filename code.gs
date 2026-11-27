/**
 * Code.gs  –  Google Apps Script backend for Time Slot Scheduler
 * Social-media only (WhatsApp removed).
 *
 * SCRIPT PROPERTIES:
 *   BUFFER_TOKEN_SOCIAL / BUFFER_TOKEN_VIDEO  — Buffer tokens
 *   BUFFER_ORG_SOCIAL / BUFFER_ORG_VIDEO      — (optional) Buffer organisation IDs
 *   BUFFER_DEFAULT_CHANNELS                   — "channelId:orgType,..." (auto-publish fallback)
 *   TZ_OFFSET_HOURS                           — e.g. 5.5 for Sri Lanka
 *   CLOUDINARY_CLOUD_NAME / _API_KEY / _API_SECRET
 *   TOKEN_SECRET                              — signs login tokens
 *   OWNER_PHONE / OWNER_NAME                  — owner login phone and display name
 *   OWNER_PASSWORD                            — type the password here, then run setupOwnerPassword() once:
 *                                               it stores OWNER_PASSWORD_HASH / OWNER_PASSWORD_SALT and deletes the plain text
 *   OWNER_EMAIL                               — (optional) owner can also log in by email
 *   CLOUDINARY_UPLOAD_FOLDER                  — (optional) default "time-slot-scheduler"
 *   CLOUDINARY_ALLOWED_FORMATS / CLOUDINARY_SIGNED_PRESET — (optional) upload limits / signed preset
 *   INSTAGRAM_STORIES_ENABLED                 — "true" turns Instagram Stories on (missing = off; Facebook Stories unaffected)
 *   YOUTUBE_PHOTO_POSTS                       — "true" lets image posts be tried on YouTube (missing = off)
 *   Run clearCaches() after changing sheet columns, projects or role lists by hand.
 *   (managed from the admin page, do not edit by hand)
 *   ADMIN_ENTRIES, CLUB_ROLES, PROJECT_ROLES, SYSTEM_PAUSED
 *
 * ADMIN RULES: an admin is a person listed in the admin whitelist WITH a club role.
 * Removing the entry removes both. The owner is exempt.
 * Run authorizeMail() once from the editor to allow sending Gmail.
 */

const SHEET_NAME         = "Database";
const PROJECTS_SHEET     = "Projects";
const BUFFER_GRAPHQL_URL = "https://api.buffer.com/graphql";
const BOD_PROJECT        = "BOD Post";          // club-level posts
const SLOT_GRACE_MIN     = 15;                  // a slot closes 15 min after its start
const LOCK_MS            = 10 * 60 * 1000;      // edit lock lasts 10 min
const MAINTENANCE_MSG    = "The system is under maintenance. Please try again later.";
const SITE_URL           = "https://leosofuop.github.io/Share/User/Submit.html";   // link added to the post emails

const DEFAULT_PROJECT_ROLES = ["Director", "Co-Director", "Chair", "Co-Chair", "Secretary",
                               "Designer", "Caption Writer", "OC Member"];
const DEFAULT_CLUB_ROLES    = ["President", "Vice President", "Treasurer", "Chief Media Officer",
                               "Chief PR Officer", "PR Panel"];

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function getBufferToken(orgType) {
  const key = orgType === "social" ? "BUFFER_TOKEN_SOCIAL" : "BUFFER_TOKEN_VIDEO";
  return PropertiesService.getScriptProperties().getProperty(key);
}

function getBufferOrgId(orgType) {
  const props = PropertiesService.getScriptProperties();
  return String(props.getProperty(orgType === "social" ? "BUFFER_ORG_SOCIAL" : "BUFFER_ORG_VIDEO") || "");
}

function getTzOffsetHours() {
  const v = PropertiesService.getScriptProperties().getProperty("TZ_OFFSET_HOURS");
  return v ? parseFloat(v) : 5.5; // default: Sri Lanka UTC+5:30
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildColMap(headers) {
  const map = {};
  headers.forEach((h, i) => { if (h) map[h.toString().toLowerCase()] = i; });
  return map;
}

function toSheetRow(rowIndex) { return Number(rowIndex) + 1; }

const DB_HEADERS = [
  "title","content","author","timeslot","date","media","status",
  "cloudinary_public_id","shared","buffer_post_ids","platforms","scheduled",
  "approved_by","rejected_by","auto_social_targets",
  "youtube_title","youtube_category_id",
  "submitted_by","submitted_at","editing_by","editing_until",
  "album_media","album_public_ids","reminder_sent"
];

// The posts sheet and its column map. The map is cached for 10 minutes, so most requests skip the header setup.
// If you add or move a column by hand in the Google Sheet, run clearCaches() once (or wait 10 minutes).
let _dbMemo_ = null;
function getDb_() {
  if (_dbMemo_) return _dbMemo_;
  const cache = CacheService.getScriptCache();
  const key = "dbcols_" + DB_HEADERS.length;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  let col = null;
  if (sheet) { try { const raw = cache.get(key); if (raw) col = JSON.parse(raw); } catch (e) { col = null; } }
  if (col && !DB_HEADERS.every(h => col[h] !== undefined)) col = null;
  if (!col) {
    ensureColumns();                                  // creates the tab and any missing column (first time only)
    sheet = ensureDatabaseSheet();
    col = buildColMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
    try { cache.put(key, JSON.stringify(col), 600); } catch (e) {}
  }
  return (_dbMemo_ = { sheet: sheet, col: col });
}

// Run from the editor after changing the sheet columns, projects or role lists by hand
function clearCaches() {
  CacheService.getScriptCache().removeAll(["dbcols_" + DB_HEADERS.length, "projects_v1", "roles_v1"]);
  console.log("Caches cleared.");
}

// Slot checks read ONLY the date, time slot and status columns (not the whole sheet).
// "row" is the same number the Admin page uses for a post.
function readSlotRows_(sheet, col) {
  const n = sheet.getLastRow() - 1;
  if (n < 1) return [];
  const dates = sheet.getRange(2, col["date"]     + 1, n, 1).getValues();
  const slots = sheet.getRange(2, col["timeslot"] + 1, n, 1).getDisplayValues();
  const stats = sheet.getRange(2, col["status"]   + 1, n, 1).getValues();
  const out = [];
  for (let i = 0; i < n; i++) out.push({ row: i + 1, date: normDateServer(dates[i][0]), slot: normSlotServer(slots[i][0]), status: String(stats[i][0] || "") });
  return out;
}
function slotTaken_(sheet, col, date, slot, ownRow) {
  return readSlotRows_(sheet, col).some(r => r.row !== ownRow && r.slot === slot && r.date === date && r.status !== "rejected");
}

// Creates the "Database" (posts) tab with its header row if it is missing
function ensureDatabaseSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, DB_HEADERS.length).setValues([DB_HEADERS]);
    sheet.getRange("D:E").setNumberFormat("@");   // keep time slot / date as plain text
  }
  return sheet;
}

function ensureColumns() {
  const sheet = ensureDatabaseSheet();
  const headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const existing = new Set(headers.map(h => h.toString().toLowerCase()));
  DB_HEADERS.forEach(c => {
    if (!existing.has(c)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(c);
      existing.add(c);
    }
  });
}

function ensureProjectsSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(PROJECTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PROJECTS_SHEET);
    sheet.getRange(1, 1, 1, 3).setValues([["name", "status", "color"]]);
  }
  return sheet;
}

let _usersMemo_ = null, _usersMemoOn_ = false;
function usersData_() {
  if (_usersMemoOn_ && _usersMemo_) return _usersMemo_;
  const d = ensureUsersSheet().getDataRange().getValues();
  if (_usersMemoOn_) _usersMemo_ = d;
  return d;
}

function ensureUsersSheet() {
  _usersMemo_ = null;                       // anything that opens the sheet for writing clears the saved copy
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Users");
  if (!sheet) sheet = ss.insertSheet("Users");
  const want = ["phone", "name", "role", "projects", "email", "salt", "hash", "createdAt", "club_role", "pending_profile"];
  const have = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), want.length)).getValues()[0];
  let changed = false;
  want.forEach((h, i) => { if (!have[i]) { have[i] = h; changed = true; } });
  if (changed) sheet.getRange(1, 1, 1, want.length).setValues([have.slice(0, want.length)]);
  return sheet;
}

function ensureActivityLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("ActivityLog");
  if (!sheet) {
    sheet = ss.insertSheet("ActivityLog");
    sheet.getRange(1, 1, 1, 6).setValues([["timestamp", "phone", "name", "role", "action", "detail"]]);
  }
  return sheet;
}

// Normalise a date / time cell value so comparisons work however Sheets stored it
function normDateServer(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return (v || "").toString().trim();
}
function normSlotServer(v) {
  const to12 = (h, m) => {
    const ap = h >= 12 ? "PM" : "AM";
    const h12 = h % 12 === 0 ? 12 : h % 12;
    return h12 + ":" + String(m).padStart(2, "0") + " " + ap;
  };
  if (v instanceof Date) return to12(v.getUTCHours(), v.getUTCMinutes());
  const s = (v || "").toString().trim();
  let m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/i);
  if (m) return parseInt(m[1], 10) + ":" + m[2] + " " + m[3].toUpperCase();
  m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (m) return to12(parseInt(m[1], 10), parseInt(m[2], 10));
  return s;
}

/**
 * Parse "8:00 AM" / "12:00 PM" on dateStr "YYYY-MM-DD"
 * Returns a UTC Date accounting for local timezone offset.
 */
function slotToDate(dateStr, slotStr, tzOffsetHours) {
  const m = slotStr.trim().match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (!m) return null;
  let hour   = parseInt(m[1]);
  const min  = parseInt(m[2]);
  const ampm = m[3].toUpperCase();
  if (ampm === "PM" && hour !== 12) hour += 12;
  if (ampm === "AM" && hour === 12) hour  = 0;
  const [y, mo, d]    = dateStr.split("-").map(Number);
  const tzOffsetMs    = tzOffsetHours * 60 * 60 * 1000;
  const localMidnight = Date.UTC(y, mo - 1, d) - tzOffsetMs;
  return new Date(localMidnight + (hour * 60 + min) * 60 * 1000);
}

function dateToYMD(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

// ---------------------------------------------------------------------------
// System pause (owner switch)
// ---------------------------------------------------------------------------
function isSystemPaused() {
  return PropertiesService.getScriptProperties().getProperty("SYSTEM_PAUSED") === "true";
}

function handleSetSystemPaused(data, session) {
  if (!session || session.role !== "owner") return json({ success: false, error: "Unauthorized." });
  PropertiesService.getScriptProperties().setProperty("SYSTEM_PAUSED", data.paused ? "true" : "false");
  logActivity(session, "setSystemPaused", "paused=" + !!data.paused);
  return json({ success: true, paused: !!data.paused });
}

// ---------------------------------------------------------------------------
// Role lists (editable by the owner)
// ---------------------------------------------------------------------------
function readList_(key, fallback) {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(key);
    if (raw) {
      const arr = JSON.parse(raw);
      if (Array.isArray(arr) && arr.length) return arr.map(String);
    }
  } catch (e) {}
  return fallback.slice();
}

function getRoleLists() {
  const cache = CacheService.getScriptCache();
  const raw = cache.get("roles_v1");
  if (raw) { try { return JSON.parse(raw); } catch (e) {} }
  const v = { projectRoles: readList_("PROJECT_ROLES", DEFAULT_PROJECT_ROLES),
              clubRoles:    readList_("CLUB_ROLES",    DEFAULT_CLUB_ROLES) };
  try { cache.put("roles_v1", JSON.stringify(v), 600); } catch (e) {}
  return v;
}

function handleSaveRoles(data, session) {
  if (!session || session.role !== "owner") return json({ success: false, error: "Unauthorized." });
  const clean = (arr) => Array.from(new Set((Array.isArray(arr) ? arr : [])
    .map(s => String(s).trim()).filter(Boolean))).slice(0, 40);
  const props = PropertiesService.getScriptProperties();
  if (data.projectRoles) props.setProperty("PROJECT_ROLES", JSON.stringify(clean(data.projectRoles)));
  if (data.clubRoles)    props.setProperty("CLUB_ROLES",    JSON.stringify(clean(data.clubRoles)));
  CacheService.getScriptCache().remove("roles_v1");
  logActivity(session, "saveRoles", "");
  return json(Object.assign({ success: true }, getRoleLists()));
}

// ---------------------------------------------------------------------------
// Auth — Phone / role / admin-entry helpers
// ---------------------------------------------------------------------------

/** Strip non-digits; local 07xxxxxxxx / 7xxxxxxxx → 947xxxxxxxx (Sri Lanka) */
function normalizePhone(raw) {
  let digits = String(raw || "").replace(/\D/g, "");
  if (digits.startsWith("0") && digits.length === 10) {
    digits = "94" + digits.slice(1);
  } else if (digits.length === 9 && digits.startsWith("7")) {
    digits = "94" + digits;
  }
  return digits;
}

// A whitelist id is an email (lower-case) or a normalised phone number
function normId_(v) {
  v = String(v || "").trim();
  if (!v) return "";
  return v.indexOf("@") >= 0 ? v.toLowerCase() : normalizePhone(v);
}

let _adminEntriesMemo_ = null;   // lives only for one request, so it is never stale across requests
function getAdminEntries_() {
  if (_adminEntriesMemo_) return _adminEntriesMemo_;
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty("ADMIN_ENTRIES");
  if (raw) {
    try { const a = JSON.parse(raw); if (Array.isArray(a)) return (_adminEntriesMemo_ = a); } catch (e) {}
  }
  // one-time migration from the old comma-separated ADMIN_WHITELIST
  const old = String(props.getProperty("ADMIN_WHITELIST") || "")
    .split(",").map(s => normId_(s)).filter(Boolean);
  return (_adminEntriesMemo_ = old.map(id => ({ id: id, role: "Admin" })));
}

function saveAdminEntries_(arr) {
  _adminEntriesMemo_ = null;
  PropertiesService.getScriptProperties().setProperty("ADMIN_ENTRIES", JSON.stringify(arr));
}

/** Owner = OWNER_EMAIL. (The owner phone login is handled separately in handleLoginPassword.) */
function isOwnerEmail_(email) {
  const owner = String(PropertiesService.getScriptProperties().getProperty("OWNER_EMAIL") || "").trim().toLowerCase();
  return !!owner && String(email || "").trim().toLowerCase() === owner;
}

/**
 * Access of a registered person, from the admin whitelist:
 *   listed → role "admin" + their club role;   not listed → role "user", no club role.
 */
function resolveAccess(email, phone) {
  if (isOwnerEmail_(email)) return { role: "owner", clubRole: "Owner", canShare: true };
  const e = normId_(email), p = normalizePhone(phone);
  const hit = getAdminEntries_().find(x => (e && x.id === e) || (p && x.id === p));
  // canShare: Admin (Approve + Share). Older entries without the flag keep sharing until the owner changes them.
  return hit
    ? { role: "admin", clubRole: hit.role || "Admin", canShare: hit.canShare !== false }
    : { role: "user", clubRole: "", canShare: false };
}

/** Only Admin (Approve + Share) and the Owner may publish / share posts. Checked on the server for every request. */
function canShare_(session) {
  return !!session && (session.role === "owner" || (session.role === "admin" && session.canShare === true));
}
/** Admin (Approve), Admin (Approve + Share) and the Owner may approve, reject, edit, move and delete posts. */
function canApprove_(session) {
  return !!session && (session.role === "owner" || session.role === "admin");
}

// ---------------------------------------------------------------------------
// Auth — User helpers
// ---------------------------------------------------------------------------

function userFromRow(row, i) {
  let projects = [];
  try { projects = JSON.parse(row[3] || "[]"); } catch(e) {}
  return {
    phone:    normalizePhone(String(row[0])),
    name:     row[1] || "",
    role:     row[2] || "user",
    projects: projects,
    email:    String(row[4] || "").trim().toLowerCase(),
    salt:     String(row[5] || ""),
    hash:     String(row[6] || ""),
    clubRole: String(row[8] || ""),
    isNew:    false,
    row:      i + 1
  };
}

function findUser(phone) {
  const data = usersData_();
  for (let i = 1; i < data.length; i++) {
    if (normalizePhone(String(data[i][0])) === phone) return userFromRow(data[i], i);
  }
  return null;
}

function findUserByEmail(email) {
  email = String(email || "").trim().toLowerCase();
  const data = usersData_();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4] || "").trim().toLowerCase() === email) return userFromRow(data[i], i);
  }
  return null;
}

function logActivity(session, action, detail) {
  try {
    ensureActivityLogSheet().appendRow([
      new Date().toISOString(),
      session ? (session.phone || "") : "",
      session ? (session.name  || "") : "",
      session ? (session.role  || "") : "",
      action  || "",
      detail  || ""
    ]);
  } catch(e) {
    console.error("logActivity error:", e);
  }
}

// ---------------------------------------------------------------------------
// Auth — Token issue / verify
// ---------------------------------------------------------------------------

// Run ONCE from the Apps Script editor if TOKEN_SECRET is missing. It creates a random secret.
// Everyone will have to log in again after this runs.
function setupTokenSecret() {
  const props = PropertiesService.getScriptProperties();
  const cur = props.getProperty("TOKEN_SECRET");
  if (cur && cur.length >= 16) { console.log("TOKEN_SECRET is already set (" + cur.length + " characters). Nothing changed."); return; }
  const secret = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "");
  props.setProperty("TOKEN_SECRET", secret);
  console.log("TOKEN_SECRET created. You can now log in.");
}

function getTokenSecret_() {
  const s = PropertiesService.getScriptProperties().getProperty("TOKEN_SECRET");
  if (!s || s.length < 16) throw new Error("TOKEN_SECRET is not set.");
  return s;
}
// Owner password: stored as a salted hash (OWNER_PASSWORD_HASH + OWNER_PASSWORD_SALT).
// Until setupOwnerPassword() has been run once, a plain OWNER_PASSWORD is still accepted so nobody is locked out.
function ownerPasswordOk_(pw) {
  const P = PropertiesService.getScriptProperties();
  const hash = P.getProperty("OWNER_PASSWORD_HASH"), salt = P.getProperty("OWNER_PASSWORD_SALT");
  if (hash && salt) return safeEqual_(hashPassword(String(pw), salt), hash);
  const plain = P.getProperty("OWNER_PASSWORD") || "";
  return !!plain && safeEqual_(String(pw), plain);
}

// Run ONCE from the Apps Script editor: turns the plain OWNER_PASSWORD into a hash and deletes the plain text.
// To change the owner password later: type the new one into OWNER_PASSWORD, run this again.
function setupOwnerPassword() {
  const P = PropertiesService.getScriptProperties();
  const plain = P.getProperty("OWNER_PASSWORD");
  if (!plain) { console.log("OWNER_PASSWORD is not set, nothing to do. (A hash may already be stored.)"); return; }
  if (plain.length < 8) { console.log("OWNER_PASSWORD is too short. Use at least 8 characters, then run this again."); return; }
  const salt = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "");
  P.setProperty("OWNER_PASSWORD_SALT", salt);
  P.setProperty("OWNER_PASSWORD_HASH", hashPassword(plain, salt));
  P.deleteProperty("OWNER_PASSWORD");
  console.log("Owner password is now stored as a hash. The plain text was deleted.");
}

function getOwnerPhone_() {
  return normalizePhone(PropertiesService.getScriptProperties().getProperty("OWNER_PHONE") || "");
}

function toHex_(bytes) { return bytes.map(b => ("0" + (b & 0xff).toString(16)).slice(-2)).join(""); }
// Compares two strings without stopping at the first difference
function safeEqual_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let d = 0;
  for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
// HMAC-SHA256 of the token data with TOKEN_SECRET
function signToken_(payload, secret) { return toHex_(Utilities.computeHmacSha256Signature(payload, secret)); }

function issueToken(phone, role, name) {
  const secret  = getTokenSecret_();
  const expires = Date.now() + 8 * 60 * 60 * 1000; // 8 hours
  const safeName = Utilities.base64Encode(name || "");
  const payload = `${phone}|${role}|${expires}|${safeName}`;
  const sig = signToken_(payload, secret);
  return Utilities.base64Encode(payload + "|" + sig);
}

function verifyToken(token) {
  if (!token) return null;
  try {
    const decoded  = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    const lastPipe = decoded.lastIndexOf("|");
    const sig      = decoded.substring(lastPipe + 1);
    const payload  = decoded.substring(0, lastPipe);
    const secret   = getTokenSecret_();
    if (!safeEqual_(sig, signToken_(payload, secret))) return null;
    const parts   = payload.split("|");
    if (parts.length < 3) return null;
    const phone   = parts[0];
    const role    = parts[1];
    const expires = Number(parts[2]);
    const name    = parts[3] ? Utilities.newBlob(Utilities.base64Decode(parts[3])).getDataAsString() : "";
    if (Date.now() > expires) return null;
    return { phone, role, name };
  } catch(e) {
    return null;
  }
}

// ---------------------------------------------------------------------------
// GET handler
// ---------------------------------------------------------------------------
function handleGetProfiles(session) {
  if (!canShare_(session)) return json({ success: false, error: "You don't have permission to share posts." });
  try {
    const fetchChannels = (orgType) => {
      const t = getBufferToken(orgType);
      if (!t) return [];
      const orgId = getBufferOrgId(orgType);
      const r = UrlFetchApp.fetch(BUFFER_GRAPHQL_URL, {
        method: "post", contentType: "application/json",
        headers: { Authorization: "Bearer " + t },
        payload: JSON.stringify({ query:
          `{ channels(input: { organizationId: "${orgId}" }) { id name service } }` }),
        muteHttpExceptions: true
      });
      const body = JSON.parse(r.getContentText());
      if (body.errors) { console.error("getProfiles " + orgType, JSON.stringify(body.errors)); return []; }
      return (body.data?.channels) || [];
    };

    const social = fetchChannels("social").map(ch => Object.assign({}, ch, { orgType: "social" }));
    const video  = fetchChannels("video").map(ch => Object.assign({}, ch, { orgType: "video" }));
    const seen   = {};
    const all    = social.concat(video).filter(ch => {
      if (seen[ch.id]) return false;
      seen[ch.id] = true;
      return true;
    });

    return json(all.map(ch => ({
      id:          ch.id,
      serviceName: (ch.service || "unknown").toLowerCase(),
      username:    ch.name || ch.id,
      orgType:     ch.orgType,
      connected:   true
    })));
  } catch (err) {
    console.error("getProfiles error:", err);
    return json({ success: false, error: "Could not load the channels." });
  }
}

function getAllPosts_() {
  const sheet = getDb_().sheet;
  const range   = sheet.getDataRange();
  const data    = range.getDisplayValues();
  const raw     = range.getValues();
  const tz      = Session.getScriptTimeZone();
  const headers = data[0];
  const col     = buildColMap(headers);
  const posts   = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[col["title"]] && !r[col["date"]]) continue;
    let dateVal = raw[i][col["date"]];
    dateVal = (dateVal instanceof Date)
      ? Utilities.formatDate(dateVal, tz, "yyyy-MM-dd")
      : (r[col["date"]] || "");
    const slotVal = normSlotServer(r[col["timeslot"]]);

    posts.push({
      row: i, title: r[col["title"]] || "", project: r[col["title"]] || "",
      content: r[col["content"]] || "",
      author: r[col["author"]] || "", timeslot: slotVal,
      date: dateVal, media: r[col["media"]] || "", album: parseAlbum_(r[col["album_media"]]),
      status: r[col["status"]] || "pending", shared: r[col["shared"]] === "TRUE",
      bufferPostIds: r[col["buffer_post_ids"]] || "", platforms: r[col["platforms"]] || "",
      scheduled: (r[col["scheduled"]] || "false").toString().toLowerCase(),
      approvedBy: r[col["approved_by"]] || "",
      rejectedBy: r[col["rejected_by"]] || "",
      submittedAt: r[col["submitted_at"]] || ""
    });
  }
  return posts;
}

function doGet(e) {
  const action = e?.parameter?.action;

  if (action === "getProjects") {
    try {
      return json(getProjectsList_());
    } catch (err) {
      console.error("getProjects error:", err);
      return json({ error: "Something went wrong. Please try again." });
    }
  }

  if (action === "getRoles")        return json(getRoleLists());
  if (action === "getSystemPaused") return json({ success: true, paused: isSystemPaused() });

  return json({ error: "Not available." });
}

function handleGetPosts(session) {
  if (!session || session.role === "user") return json({ success: false, error: "Unauthorized." });
  return json({ success: true, posts: getAllPosts_() });
}
function handleGetTakenSlots(session) {
  if (!session) return json({ success: false, error: "Please log in again." });
  const db = getDb_();
  const slots = readSlotRows_(db.sheet, db.col).filter(r => r.status !== "rejected" && r.date)
    .map(r => ({ date: r.date, timeslot: r.slot }));
  return json({ success: true, slots: slots });
}
function handleGetContact(session) {
  if (!session) return json({ success: false });
  return json({ success: true, whatsapp:
    String(PropertiesService.getScriptProperties().getProperty("CONTACT_WHATSAPP") || "").replace(/\D/g, "") });
}

// ---------------------------------------------------------------------------
// POST handler
// ---------------------------------------------------------------------------
const DB_ACTIONS_ = { publishToBuffer: 1, update: 1, edit: 1, changeSlot: 1, deleteRow: 1, setScheduled: 1, saveAutoDestinations: 1,
  getMyPosts: 1, editMyPost: 1, addAlbumPhotos: 1, lockPost: 1, unlockPost: 1, reorderAlbum: 1, submit: 1 };
const USERS_READONLY_ACTIONS_ = { submit: true, getMyPosts: true, editMyPost: true, getPosts: true, getTakenSlots: true };

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // ── No session required ─────────────────────────────────────────────────
    if (data.action === "emailSendOtp")    return handleEmailSendOtp(data);
    if (data.action === "emailVerifyOtp")  return handleEmailVerifyOtp(data);
    if (data.action === "registerAccount") return handleRegisterAccount(data);
    if (data.action === "loginPassword")   return handleLoginPassword(data);
    if (data.action === "forgotSendOtp")   return handleForgotSendOtp(data);
    if (data.action === "forgotReset")     return handleForgotReset(data);
    if (data.action === "getSystemPaused") return json({ success: true, paused: isSystemPaused() });

    // ── Everything else needs a valid session token ─────────────────────────
    const session = verifyToken(data.token);

    // During these actions the Users sheet is only read, so one read can be reused inside the request
    _usersMemoOn_ = data.action === undefined || data.action === null || USERS_READONLY_ACTIONS_[String(data.action)] === true;

    // An admin who was removed from the whitelist becomes a normal user immediately
    if (session) {
      if (session.role === "admin") {
        const u = findUser(session.phone);
        const acc = u ? resolveAccess(u.email, u.phone) : { role: "user", canShare: false };
        session.role = acc.role;
        session.canShare = acc.canShare;
      } else {
        session.canShare = session.role === "owner";
      }
    }

    // Allowed even while paused, so the maintenance screen can show the contact link
    if (data.action === "getContact") return handleGetContact(session);

    // While the system is paused only admins and the owner can act
    if (isSystemPaused() && (!session || session.role === "user")) {
      return json({ success: false, paused: true, error: MAINTENANCE_MSG });
    }

    // Only the actions that work on posts open the posts sheet
    let sheet = null, col = null;
    if (DB_ACTIONS_[String(data.action)] || data.action === undefined || data.action === null) {
      const db = getDb_(); sheet = db.sheet; col = db.col;
    }

    switch (data.action) {
      case "getPosts":      return handleGetPosts(session);
      case "getTakenSlots": return handleGetTakenSlots(session);
      case "getProfiles":   return handleGetProfiles(session);
      case "publishToBuffer": return handlePublishToBuffer(sheet, col, data, session);
      case "update":          return handleUpdateStatus(sheet, col, data, session);
      case "edit":            return handleEditPost(sheet, col, data, session);
      case "changeSlot":      return handleChangeSlot(sheet, col, data, session);
      case "deleteRow":       return handleDeleteRow(sheet, col, data, session);
      case "setScheduled":    return handleSetScheduled(sheet, col, data, session);
      case "addProject":      return handleAddProject(data, session);
      case "removeProject":   return handleRemoveProject(data, session);
      case "updateProject":   return handleUpdateProject(data, session);
      case "getWhitelist":    return handleGetWhitelist(session);
      case "setWhitelist":    return handleSetWhitelist(data, session);
      case "getAdminEntries": return handleGetAdminEntries(session);
      case "setAdminEntry":   return handleSetAdminEntry(data, session);
      case "removeAdminEntry":return handleRemoveAdminEntry(data, session);
      case "saveRoles":       return handleSaveRoles(data, session);
      case "getSheetUrls":    return handleGetSheetUrls(session);
      case "setSystemPaused": return handleSetSystemPaused(data, session);
      case "saveAutoDestinations": return handleSaveAutoDestinations(sheet, col, data, session);
      case "saveAccountProjects":  return handleSaveAccountProjects(data, session);
      case "changePassword":  return handleChangePassword(data, session);
      case "getMyPosts":      return handleGetMyPosts(sheet, col, session);
      case "quickScan":       return handleScan_(session, false);
      case "fullScan":        return handleScan_(session, true);
      case "diagDeleteAsset": return handleDiagDeleteAsset(data, session);
      case "editMyPost":      return handleEditMyPost(sheet, col, data, session);
      case "addAlbumPhotos":  return handleAddAlbumPhotos(sheet, col, data, session);
      case "getMyProfile":    return handleGetMyProfile(session);
      case "profileEmailOtp": return handleProfileEmailOtp(data, session);
      case "requestProfileChange": return handleRequestProfileChange(data, session);
      case "getProfileRequests":   return handleGetProfileRequests(session);
      case "decideProfileRequest": return handleDecideProfileRequest(data, session);
      case "getDirectory":    return handleGetDirectory(data, session);
      case "getNotify":       return handleGetNotify(session);
      case "saveNotify":      return handleSaveNotify(data, session);
      case "getShareSettings":return handleGetShareSettings(session);
      case "lockPost":        return handleLockPost(sheet, col, data, session);
      case "unlockPost":      return handleUnlockPost(sheet, col, data, session);
      case "getUploadSignature": return handleGetUploadSignature(data, session);
      case "reorderAlbum":    return handleReorderAlbum(sheet, col, data, session);
      case "submit":
      case undefined:
      case null:              return handleNewSubmission(sheet, col, data, session);
      default:                return json({ success: false, error: `Unknown action: ${data.action}` });
    }
  } catch (err) {
    console.error("doPost error:", err);
    return json({ success: false, error: "Something went wrong. Please try again." });
  }
}

// ---------------------------------------------------------------------------
// Post helpers: public id, Cloudinary type, ownership, edit lock
// ---------------------------------------------------------------------------
function extractPublicId_(url) {
  if (!url || String(url).indexOf("cloudinary.com") < 0) return "";
  const parts       = String(url).split("/");
  const uploadIndex = parts.indexOf("upload");
  if (uploadIndex !== -1 && parts.length > uploadIndex + 1) {
    const afterUpload = parts.slice(uploadIndex + 1);
    // skip the version segment (v1234567890) if present; the rest is folder/name
    const startIndex  = /^v\d+$/.test(afterUpload[0]) ? 1 : 0;
    return afterUpload.slice(startIndex).join("/").replace(/\.[^/.]+$/, "");
  }
  return "";
}

function cloudinaryTypeFromUrl_(url) {
  return /\/video\/upload\//.test(String(url || "")) ? "video" : "image";
}

function postOwnedBy_(row, col, session) {
  const sb = String(row[col["submitted_by"]] || "").trim();
  if (sb) return normalizePhone(sb) === session.phone;
  // older posts (before submitted_by existed): match on the author name
  const a = String(row[col["author"]] || "");
  return !!session.name && (a === session.name || a.indexOf(session.name + " —") === 0);
}

function readLock_(sheet, col, r) {
  const by    = String(sheet.getRange(r, col["editing_by"] + 1).getValue() || "");
  const until = Number(sheet.getRange(r, col["editing_until"] + 1).getValue()) || 0;
  return { by: by, until: until };
}
function lockedByOther_(lock, session) {
  return !!lock.by && lock.until > Date.now() && lock.by.split("|")[0] !== session.phone;
}
function lockOwnerName_(lock) { return lock.by.split("|")[1] || "someone else"; }
function setLock_(sheet, col, r, session) {
  sheet.getRange(r, col["editing_by"]    + 1).setValue(session.phone + "|" + (session.name || ""));
  sheet.getRange(r, col["editing_until"] + 1).setValue(Date.now() + LOCK_MS);
}
function clearLock_(sheet, col, r) {
  sheet.getRange(r, col["editing_by"]    + 1).setValue("");
  sheet.getRange(r, col["editing_until"] + 1).setValue("");
}
function validRow_(sheet, r) { return r >= 2 && r <= sheet.getLastRow(); }

// Users can edit a post only until 7 days after its scheduled date
function withinEditWindow_(dateStr) {
  const cutoff = Utilities.formatDate(new Date(Date.now() - 7 * 24 * 60 * 60 * 1000),
                                      Session.getScriptTimeZone(), "yyyy-MM-dd");
  return !!dateStr && String(dateStr) >= cutoff;
}

// The project list is kept for 5 minutes. It is cleared whenever a project is added, changed or removed.
function getProjectsList_() {
  const cache = CacheService.getScriptCache();
  const raw = cache.get("projects_v1");
  if (raw) { try { return JSON.parse(raw); } catch (e) {} }
  const data = ensureProjectsSheet().getDataRange().getValues();
  const list = data.slice(1).map((r, i) => ({
    row: i + 1, name: r[0] || "", status: (r[1] || "ongoing").toString().toLowerCase(), color: r[2] || "#b34e1a"
  })).filter(p => p.name);
  try { cache.put("projects_v1", JSON.stringify(list), 300); } catch (e) {}
  return list;
}
function clearProjectsCache_() { CacheService.getScriptCache().remove("projects_v1"); }

function getProjectStatus_(name) {
  const want = String(name).trim();
  const hit = getProjectsList_().find(p => String(p.name).trim() === want);
  return hit ? hit.status : null;   // null = not a known project
}

// ---------------------------------------------------------------------------
// Action handlers
// ---------------------------------------------------------------------------

function handlePublishToBuffer(sheet, col, data, session) {
  if (!canShare_(session)) return json({ success: false, error: "You don't have permission to share posts." });
  const r       = toSheetRow(data.row);
  if (!validRow_(sheet, r)) return json({ success: false, error: "Post not found." });
  if (String(sheet.getRange(r, col["status"] + 1).getValue()).toLowerCase() !== "approved")
    return json({ success: false, error: "Only approved posts can be shared." });
  const content = data.strippedContent !== undefined ? data.strippedContent : sheet.getRange(r, col["content"] + 1).getValue();
  const cover   = String(sheet.getRange(r, col["media"] + 1).getValue() || "");
  const album   = parseAlbum_(sheet.getRange(r, col["album_media"] + 1).getValue());
  const urls    = album.length ? album : (cover ? [cover] : []);
  const hasVideo = urls.some(isVideoUrl_);
  const mention  = String(PropertiesService.getScriptProperties().getProperty("STORY_MENTION") || "");
  const project  = String(sheet.getRange(r, col["title"] + 1).getValue() || "");

  const youtubeMeta = data.youtubeMeta || null;
  const postIds = [], succeeded = [], failed = [];

  for (const prof of (data.profiles || [])) {
    const token = getBufferToken(prof.orgType || "social");
    if (!token) { failed.push(`${prof.username || prof.id}: no token`); continue; }
    const channelId = prof.id || prof.profileId;
    const service   = (prof.serviceName || "").toLowerCase();
    const dest      = prof.destination === "story" ? "story" : "post";
    let text = (prof.text !== undefined && prof.text !== null && prof.text !== "") ? String(prof.text) : content;

    const who = prof.username || prof.id;

    // Instagram Stories can be switched off; Facebook Stories are never affected by this
    if (dest === "story" && service === "instagram" && !igStoriesOn_()) {
      failed.push(`${who}: Instagram Stories are turned off.`); continue;
    }

    let tiktokTitle = "";
    if (service === "tiktok" && dest === "post") {
      const max = hasVideo ? TIKTOK_VIDEO_TEXT_MAX : TIKTOK_PHOTO_TEXT_MAX;
      if (text.length > max) { failed.push(`${who}: text is ${text.length} characters, TikTok allows ${max}. Shorten it.`); continue; }
      if (!hasVideo) {   // photo post: editable title, default "Project Name | LeosofUOP"
        const given = prof.title !== undefined && prof.title !== null && String(prof.title).trim() !== "";
        tiktokTitle = cleanTitle_(given ? prof.title : fitTitle_(project, TIKTOK_TITLE_MAX));
        if (!tiktokTitle) { failed.push(`${who}: the TikTok title can't be empty.`); continue; }
        if (tiktokTitle.length > TIKTOK_TITLE_MAX) { failed.push(`${who}: the TikTok title is ${tiktokTitle.length} characters, TikTok allows ${TIKTOK_TITLE_MAX}. Shorten it.`); continue; }
      }
    }

    let ytMeta = null;
    if (service === "youtube") {
      ytMeta = Object.assign({}, youtubeMeta || {});
      ytMeta.title = cleanTitle_(ytMeta.title) || fitTitle_(project, YOUTUBE_TITLE_MAX);
      if (ytMeta.title.length > YOUTUBE_TITLE_MAX) { failed.push(`${who}: the YouTube title is ${ytMeta.title.length} characters, YouTube allows ${YOUTUBE_TITLE_MAX}. Shorten it.`); continue; }
      if (text.length > YOUTUBE_DESC_MAX) { failed.push(`${who}: the YouTube description is ${text.length} characters, YouTube allows ${YOUTUBE_DESC_MAX}. Shorten the text.`); continue; }
    }

    const result = createBufferPost(token, channelId, text, urls, service, ytMeta,
                                    { destination: dest, mention: mention, tiktokTitle: tiktokTitle });
    if (result.success) {
      postIds.push(result.postId);
      succeeded.push((prof.serviceName || prof.id) + (dest === "story" ? " (story)" : ""));
      if (result.partial) failed.push(`${prof.username || prof.id}: some story images failed (${result.partial})`);
    } else {
      failed.push(`${prof.username || prof.id}: ${result.error}`);
    }
  }

  if (succeeded.length > 0) {
    sheet.getRange(r, col["shared"]          + 1).setValue(true);
    sheet.getRange(r, col["buffer_post_ids"] + 1).setValue(postIds.join(","));
    sheet.getRange(r, col["platforms"]       + 1).setValue(succeeded.join(","));
    logActivity(session, "publishToBuffer", `row=${data.row} channels=${succeeded.join(",")}`);
    return json({ success: true, failed });
  }
  return json({ success: false, error: "Failed to publish to any selected channel.", details: failed });
}

function handleUpdateStatus(sheet, col, data, session) {
  if (!canApprove_(session)) return json({ success: false, error: "Unauthorized." });
  const status = String(data.status || "");
  if (["approved", "rejected", "pending"].indexOf(status) < 0) return json({ success: false, error: "Unknown status." });
  const r = toSheetRow(data.row);
  if (!validRow_(sheet, r)) return json({ success: false, error: "Post not found." });
  sheet.getRange(r, col["status"] + 1).setValue(status);
  if (status === "approved") {
    // Approving turns on automatic sharing at the slot time, so only people who may share can do that.
    // An Admin (Approve) approval leaves the post on manual: someone with sharing permission shares it.
    sheet.getRange(r, col["scheduled"]   + 1).setValue(canShare_(session) ? "true" : "false");
    sheet.getRange(r, col["approved_by"] + 1).setValue(session.name || session.phone);
  }
  if (status === "rejected") {
    sheet.getRange(r, col["rejected_by"] + 1).setValue(session.name || session.phone);
  }
  logActivity(session, "updateStatus", `row=${data.row} status=${data.status}`);
  return json({ success: true });
}

// Admin edit (admin page). Respects the edit lock so two people never edit at once.
function handleEditPost(sheet, col, data, session) {
  if (!session || session.role === "user") return json({ success: false, error: "Unauthorized." });
  const r = toSheetRow(data.row);
  if (!validRow_(sheet, r)) return json({ success: false, error: "Post not found." });
  const lock = readLock_(sheet, col, r);
  if (lockedByOther_(lock, session))
    return json({ success: false, locked: true, error: "This post is being edited by " + lockOwnerName_(lock) + "." });
  const upErr = data.media ? checkUploads_([data.media], session) : "";
  if (upErr) return json({ success: false, error: upErr });
  sheet.getRange(r, col["content"] + 1).setValue(safeCell_(data.content));
  if (data.media) sheet.getRange(r, col["media"] + 1).setValue(data.media);
  sheet.getRange(r, col["shared"]  + 1).setValue(false);
  clearLock_(sheet, col, r);
  if (data.media) consumeUploads_([data.media]);
  logActivity(session, "editPost", `row=${data.row}`);
  return json({ success: true });
}

function handleChangeSlot(sheet, col, data, session) {
  if (!canApprove_(session)) return json({ success: false, error: "Unauthorized." });
  const r          = toSheetRow(data.row);
  if (!validRow_(sheet, r)) return json({ success: false, error: "Post not found." });
  const targetDate = data.newDate || data.date;
  // the post being moved is skipped, every other post is checked
  if (slotTaken_(sheet, col, normDateServer(targetDate), normSlotServer(data.newSlot), Number(data.row)))
    return json({ success: false, error: "That slot is already taken on this date." });
  sheet.getRange(r, col["timeslot"] + 1).setNumberFormat("@").setValue(data.newSlot);
  if (data.newDate) sheet.getRange(r, col["date"] + 1).setNumberFormat("@").setValue(data.newDate);
  sheet.getRange(r, col["shared"]    + 1).setValue(false);
  if (canShare_(session)) sheet.getRange(r, col["scheduled"] + 1).setValue("true");   // automatic sharing needs sharing permission
  logActivity(session, "changeSlot", `row=${data.row} newSlot=${data.newSlot}`);
  return json({ success: true });
}

function handleDeleteRow(sheet, col, data, session) {
  if (!canApprove_(session)) return json({ success: false, error: "Unauthorized." });
  const r        = toSheetRow(data.row);
  if (!validRow_(sheet, r)) return json({ success: false, error: "Post not found." });   // also protects the header row
  const publicId = sheet.getRange(r, col["cloudinary_public_id"] + 1).getValue();
  const mediaUrl = sheet.getRange(r, col["media"] + 1).getValue();
  sheet.deleteRow(r);
  if (publicId) deleteCloudinaryAsset(publicId, cloudinaryTypeFromUrl_(mediaUrl));
  logActivity(session, "deleteRow", `row=${data.row}`);
  return json({ success: true });
}

function handleSetScheduled(sheet, col, data, session) {
  if (!canShare_(session)) return json({ success: false, error: "You don't have permission to share posts." });
  if (!validRow_(sheet, toSheetRow(data.row))) return json({ success: false, error: "Post not found." });
  sheet.getRange(toSheetRow(data.row), col["scheduled"] + 1)
    .setValue(data.scheduled === true || data.scheduled === "true" ? "true" : "false");
  logActivity(session, "setScheduled", `row=${data.row} scheduled=${data.scheduled}`);
  return json({ success: true });
}

// ── New post from the submit page ───────────────────────────────────────────
function isOwnMediaUrl_(u) {
  const cloud = String(PropertiesService.getScriptProperties().getProperty("CLOUDINARY_CLOUD_NAME") || "");
  return !!cloud && String(u || "").indexOf("https://res.cloudinary.com/" + cloud + "/") === 0;
}

function safeCell_(v) {
  v = String(v == null ? "" : v);
  return /^[=+\-@\t\r]/.test(v) ? "'" + v : v;
}

function handleNewSubmission(sheet, col, data, session) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return json({ success: false, error: "The server is busy. Please press Submit again." }); }
  try { return handleNewSubmissionInner_(sheet, col, data, session); }
  finally { lock.releaseLock(); }
}

function handleNewSubmissionInner_(sheet, col, data, session) {
  if (!session) return json({ success: false, error: "Please log in again." });

  // The page retries a save that lost its connection. If the first try did get through, answer success again.
  const clientId = String(data.clientId || "").replace(/[^A-Za-z0-9_-]/g, "").slice(0, 60);
  if (clientId && CacheService.getScriptCache().get("sub_" + clientId)) return json({ success: true, duplicate: true });

  const project = String(data.project || data.title || "").trim();
  if (!project) return json({ success: false, error: "Select a project." });

  // Access rules (admins and the owner can post to any project)
  if (project === BOD_PROJECT) {
    const u = findUser(session.phone);
    const hasClubRole = session.role === "owner" ||
      (u && resolveAccess(u.email, u.phone).clubRole);
    if (!hasClubRole) return json({ success: false, error: "You do not have access to BOD Post." });
  } else {
    const status = getProjectStatus_(project);
    if (status === null)        return json({ success: false, error: "Unknown project." });
    if (status === "completed") return json({ success: false, error: "This project is completed. You can't post to it." });
    if (session.role === "user") {
      const u = findUser(session.phone);
      if (!u || !u.projects.some(p => p.project === project))
        return json({ success: false, error: "You are not a member of this project." });
    }
  }

  const date = normDateServer(data.date);
  const slot = normSlotServer(data.timeslot);
  if (!date || !slot) return json({ success: false, error: "Select a date and a time slot." });

  // A slot closes 15 minutes after its start time
  const slotStart = slotToDate(date, slot, getTzOffsetHours());
  if (!slotStart) return json({ success: false, error: "Invalid time slot." });
  if (Date.now() - slotStart.getTime() > SLOT_GRACE_MIN * 60 * 1000)
    return json({ success: false, error: "That time slot has already closed. Please choose another." });

  if (slotTaken_(sheet, col, date, slot, -1)) return json({ success: false, error: "That time slot is already taken." });

  // Album: up to 10 images, images only
  const album = (Array.isArray(data.album) ? data.album : []).map(String).filter(Boolean);
  if (album.length > ALBUM_USER_MAX) return json({ success: false, error: "Maximum " + ALBUM_USER_MAX + " images." });
  if (album.length > 1 && album.some(isVideoUrl_)) return json({ success: false, error: "An album can only hold images." });
  const cover = album.length ? album[0] : String(data.media || "");
  const newMedia = album.concat(cover ? [cover] : []);
  const upErr = checkUploads_(newMedia, session);
  if (upErr) return json({ success: false, error: upErr });

  const me = findUser(session.phone);
  const myRole = project === BOD_PROJECT
    ? ((me && resolveAccess(me.email, me.phone).clubRole) || "Owner")
    : ((me && (me.projects.find(p => p.project === project) || {}).role) || "");
  const authorName = ((session.name || "") + (myRole ? " — " + myRole : "")).substring(0, 120);

  const width = sheet.getLastColumn();
  const row   = new Array(width).fill("");
  const put   = (name, val) => { if (col[name] !== undefined) row[col[name]] = val; };
  put("album_media", album.length > 1 ? JSON.stringify(album) : "");
  put("album_public_ids", album.length > 1 ? JSON.stringify(album.map(extractPublicId_)) : "");
  put("title", project);
  put("content", safeCell_(String(data.content || "")));
  put("author", safeCell_(authorName));
  put("timeslot", slot);   put("date", date);
  put("media", cover);
  put("status", "pending");
  put("cloudinary_public_id", extractPublicId_(cover));
  put("shared", false);    put("scheduled", "false");
  put("submitted_by", session.phone);
  put("submitted_at", new Date().toISOString());

  const target = sheet.getLastRow() + 1;
  sheet.getRange(target, col["timeslot"] + 1).setNumberFormat("@");
  sheet.getRange(target, col["date"] + 1).setNumberFormat("@");
  sheet.getRange(target, 1, 1, width).setValues([row]);
  logActivity(session, "submit", `project=${project} date=${date} slot=${slot}`);
  try { notifyNewPost_(project, date, slot, authorName); } catch (e) { console.error("notify error:", e); }
  consumeUploads_(newMedia);
  if (clientId) CacheService.getScriptCache().put("sub_" + clientId, "1", 600);
  return json({ success: true });
}

// ── Signed Cloudinary upload permission (logged-in users only) ──────────────
// Each permission is for ONE file: it carries a random id ("upl_...") that is remembered for the person who asked.
// When a post is saved, the server only accepts uploads whose id belongs to that person, and uses each id once.
// The signature also fixes the folder and the allowed file formats. It expires after about an hour.
const UPLOAD_MAX_IMAGE_MB = 10;    // keep the same as CONFIG.APP.MAX_IMAGE_SIZE_MB
const UPLOAD_MAX_VIDEO_MB = 100;   // keep the same as CONFIG.APP.MAX_VIDEO_SIZE_MB
const UPLOAD_FORMATS = "jpg,jpeg,png,webp,gif,heic,heif,mp4,mov,m4v,webm,ogv,3gp,mpeg,mpg";

function handleGetUploadSignature(data, session) {
  if (!session) return json({ success: false, error: "Please log in again." });
  const P = PropertiesService.getScriptProperties();
  const cloudName = P.getProperty("CLOUDINARY_CLOUD_NAME");
  const apiKey    = P.getProperty("CLOUDINARY_API_KEY");
  const apiSecret = P.getProperty("CLOUDINARY_API_SECRET");
  if (!cloudName || !apiKey || !apiSecret) return json({ success: false, error: "Uploads are not set up yet." });

  // Size and type as declared by the page (Cloudinary's own limits are the final guard)
  const isVid = /^video\//.test(String(data.type || ""));
  const maxMB = isVid ? UPLOAD_MAX_VIDEO_MB : UPLOAD_MAX_IMAGE_MB;
  if (Number(data.size) > maxMB * 1024 * 1024) return json({ success: false, error: "That file is too big. The limit is " + maxMB + " MB." });

  // At most 60 permissions per person per hour
  const cache = CacheService.getScriptCache();
  const key = "sig_" + session.phone;
  const n = Number(cache.get(key) || 0) + 1;
  if (n > 60) return json({ success: false, error: "Too many uploads. Please try again later." });
  cache.put(key, String(n), 3600);

  const folder  = String(P.getProperty("CLOUDINARY_UPLOAD_FOLDER") || "time-slot-scheduler");
  const preset  = String(P.getProperty("CLOUDINARY_SIGNED_PRESET") || "");   // optional signed preset
  const formats = String(P.getProperty("CLOUDINARY_ALLOWED_FORMATS") || UPLOAD_FORMATS);
  const nonce   = "upl_" + (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "").substring(0, 40);
  cache.put("up_" + nonce, String(session.phone), 21600);                   // valid for 6 hours

  const timestamp = Math.floor(Date.now() / 1000);
  const params = { allowed_formats: formats, folder: folder, public_id: nonce, timestamp: String(timestamp) };
  if (preset) params.upload_preset = preset;
  const toSign = Object.keys(params).sort().map(k => k + "=" + params[k]).join("&");
  const signature = toHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, toSign + apiSecret));
  return json({ success: true, cloudName: cloudName, apiKey: apiKey, timestamp: timestamp, signature: signature,
                folder: folder, publicId: nonce, allowedFormats: formats, uploadPreset: preset });
}

// The upload id sits at the end of the Cloudinary public id of a file uploaded through this page
function uploadNonce_(url) {
  const m = /(?:^|\/)(upl_[A-Za-z0-9]{16,64})$/.exec(extractPublicId_(url));
  return m ? m[1] : "";
}
// "" = fine, otherwise the message to show. Every file must be ours AND uploaded by this person with a permission still unused.
function checkUploads_(urls, session) {
  const cache = CacheService.getScriptCache();
  for (const u of urls) {
    if (!isOwnMediaUrl_(u)) return "Invalid media.";
    const n = uploadNonce_(u);
    if (!n || cache.get("up_" + n) !== String(session.phone)) return "This upload is not valid or has expired. Please upload the file again.";
  }
  return "";
}
// Called after the post is saved, so each upload can be used only once
function consumeUploads_(urls) {
  const cache = CacheService.getScriptCache();
  urls.forEach(u => { const n = uploadNonce_(u); if (n) cache.remove("up_" + n); });
}

// ── User history ────────────────────────────────────────────────────────────
function handleGetMyPosts(sheet, col, session) {
  if (!session) return json({ success: false, error: "Please log in again." });
  const range = sheet.getDataRange();
  const disp  = range.getDisplayValues();
  const raw   = range.getValues();
  const tz    = Session.getScriptTimeZone();
  const now   = Date.now();
  const out   = [];
  for (let i = 1; i < disp.length; i++) {
    const r = disp[i];
    if (!r[col["title"]] && !r[col["date"]]) continue;
    if (!postOwnedBy_(r, col, session)) continue;

    const dv = raw[i][col["date"]];
    const date = (dv instanceof Date) ? Utilities.formatDate(dv, tz, "yyyy-MM-dd") : (r[col["date"]] || "");
    const status = (r[col["status"]] || "pending").toString().toLowerCase();
    const shared = String(r[col["shared"]]).toUpperCase() === "TRUE";
    const lockBy = String(r[col["editing_by"]] || "");
    const lockUntil = Number(r[col["editing_until"]]) || 0;
    const lockedByOther = !!lockBy && lockUntil > now && lockBy.split("|")[0] !== session.phone;

    out.push({
      row: i, project: r[col["title"]] || "", content: r[col["content"]] || "",
      media: r[col["media"]] || "", album: parseAlbum_(r[col["album_media"]]),
      date: date, timeslot: normSlotServer(r[col["timeslot"]]),
      status: status, shared: shared,
      approvedBy: r[col["approved_by"]] || "", rejectedBy: r[col["rejected_by"]] || "",
      submittedAt: r[col["submitted_at"]] || "",
      lockedBy: lockedByOther ? (lockBy.split("|")[1] || "someone else") : "",
      canEdit: !shared && status !== "rejected" && !lockedByOther && withinEditWindow_(date)
    });
  }
  out.sort((a, b) => (b.date + b.timeslot).localeCompare(a.date + a.timeslot) || b.row - a.row);
  return json({ success: true, posts: out.slice(0, 100) });
}

// ── Edit lock (shared by the user page and the admin page) ──────────────────
function handleLockPost(sheet, col, data, session) {
  if (!session) return json({ success: false, error: "Please log in again." });
  const r = toSheetRow(data.row);
  if (!validRow_(sheet, r)) return json({ success: false, error: "Post not found." });
  if (session.role === "user") {
    const row = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    if (!postOwnedBy_(row, col, session)) return json({ success: false, error: "You can only edit your own posts." });
    if (!withinEditWindow_(normDateServer(sheet.getRange(r, col["date"] + 1).getValue())))
      return json({ success: false, error: "Posts can only be edited within 7 days." });
  }
  const lock = readLock_(sheet, col, r);
  if (lockedByOther_(lock, session))
    return json({ success: false, locked: true, error: "This post is being edited by " + lockOwnerName_(lock) + "." });
  setLock_(sheet, col, r, session);
  return json({ success: true });
}

function handleUnlockPost(sheet, col, data, session) {
  if (!session) return json({ success: false, error: "Please log in again." });
  const r = toSheetRow(data.row);
  if (!validRow_(sheet, r)) return json({ success: true });
  const lock = readLock_(sheet, col, r);
  if (lock.by && lock.by.split("|")[0] === session.phone) clearLock_(sheet, col, r);
  return json({ success: true });
}

// ── User edits their own post (until it is shared) ──────────────────────────
function handleEditMyPost(sheet, col, data, session) {
  if (!session) return json({ success: false, error: "Please log in again." });
  const r = toSheetRow(data.row);
  if (!validRow_(sheet, r)) return json({ success: false, error: "Post not found." });
  const row = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

  if (!postOwnedBy_(row, col, session)) return json({ success: false, error: "You can only edit your own posts." });
  if (String(row[col["shared"]]).toUpperCase() === "TRUE")
    return json({ success: false, error: "This post has already been shared, so it can't be edited." });
  if (String(row[col["status"]]).toLowerCase() === "rejected")
    return json({ success: false, error: "This post was rejected. Please submit a new post." });
  if (!withinEditWindow_(normDateServer(sheet.getRange(r, col["date"] + 1).getValue())))
    return json({ success: false, error: "Posts can only be edited within 7 days." });

  const lock = readLock_(sheet, col, r);
  if (lockedByOther_(lock, session))
    return json({ success: false, locked: true, error: "This post is being edited by " + lockOwnerName_(lock) + "." });

  const content = String(data.content || "").trim();
  if (!content) return json({ success: false, error: "Post content can't be empty." });
  const upErr = data.media ? checkUploads_([data.media], session) : "";
  if (upErr) return json({ success: false, error: upErr });

  sheet.getRange(r, col["content"] + 1).setValue(safeCell_(content));
  if (data.media) {
    const oldId  = row[col["cloudinary_public_id"]];
    const oldUrl = row[col["media"]];
    sheet.getRange(r, col["media"] + 1).setValue(data.media);
    sheet.getRange(r, col["cloudinary_public_id"] + 1).setValue(extractPublicId_(data.media));
    if (oldId) deleteCloudinaryAsset(oldId, cloudinaryTypeFromUrl_(oldUrl));
    parseAlbum_(row[col["album_public_ids"]]).forEach(id => { if (id && id !== oldId) deleteCloudinaryAsset(id, "image"); });
    sheet.getRange(r, col["album_media"] + 1).setValue("");
    sheet.getRange(r, col["album_public_ids"] + 1).setValue("");
  }
  sheet.getRange(r, col["reminder_sent"] + 1).setValue("");
  // After any edit the post must be approved again
  sheet.getRange(r, col["status"]      + 1).setValue("pending");
  sheet.getRange(r, col["approved_by"] + 1).setValue("");
  sheet.getRange(r, col["scheduled"]   + 1).setValue("false");
  sheet.getRange(r, col["shared"]      + 1).setValue(false);
  clearLock_(sheet, col, r);
  if (data.media) consumeUploads_([data.media]);
  logActivity(session, "editMyPost", `row=${data.row}`);
  return json({ success: true });
}

// ── Projects ────────────────────────────────────────────────────────────────
const PROJECT_STATUSES = ["ongoing", "tobe", "completed"];

function handleAddProject(data, session) {
  if (!session || session.role === "user") return json({ success: false, error: "Unauthorized." });
  const name = String(data.name || "").trim();
  if (!name) return json({ success: false, error: "Enter a project name." });
  if (name.toLowerCase() === BOD_PROJECT.toLowerCase()) return json({ success: false, error: "That name is reserved." });
  const status = PROJECT_STATUSES.includes(String(data.status)) ? data.status : "ongoing";
  ensureProjectsSheet().appendRow([name, status, data.color || "#b34e1a"]);
  clearProjectsCache_();
  logActivity(session, "addProject", `name=${name}`);
  return json({ success: true });
}

function handleRemoveProject(data, session) {
  if (!session || session.role === "user") return json({ success: false, error: "Unauthorized." });
  const sheet    = ensureProjectsSheet();
  const sheetRow = Number(data.row) + 1;
  const projectName = sheet.getRange(sheetRow, 1).getValue();
  if (sheetRow > 1) sheet.deleteRow(sheetRow);

  // also remove this project from every user's saved project list
  const usersSheet = ensureUsersSheet();
  const usersData  = usersSheet.getDataRange().getValues();
  for (let i = 1; i < usersData.length; i++) {
    let projects = [];
    try { projects = JSON.parse(usersData[i][3] || "[]"); } catch(e) { continue; }
    const filtered = projects.filter(p => p.project !== projectName);
    if (filtered.length !== projects.length) {
      usersSheet.getRange(i + 1, 4).setValue(JSON.stringify(filtered));
    }
  }
  clearProjectsCache_();
  logActivity(session, "removeProject", `row=${data.row} name=${projectName}`);
  return json({ success: true });
}

function handleUpdateProject(data, session) {
  if (!session || session.role === "user") return json({ success: false, error: "Unauthorized." });
  const sheet    = ensureProjectsSheet();
  const sheetRow = Number(data.row) + 1;
  const status   = PROJECT_STATUSES.includes(String(data.status)) ? data.status : "ongoing";
  const oldName  = String(sheet.getRange(sheetRow, 1).getValue() || "");
  const newName  = String(data.name || "").trim();
  if (!newName) return json({ success: false, error: "Enter a project name." });
  if (newName.toLowerCase() === BOD_PROJECT.toLowerCase()) return json({ success: false, error: "That name is reserved." });
  const all = sheet.getDataRange().getValues();
  for (let i = 1; i < all.length; i++)
    if (i + 1 !== sheetRow && String(all[i][0]).trim().toLowerCase() === newName.toLowerCase())
      return json({ success: false, error: "Another project already has that name." });

  sheet.getRange(sheetRow, 1).setValue(newName);
  sheet.getRange(sheetRow, 2).setValue(status);
  sheet.getRange(sheetRow, 3).setValue(data.color || "#b34e1a");

  // A rename also updates users' saved projects and every existing post
  if (oldName && oldName !== newName) {
    const us = ensureUsersSheet(), ud = us.getDataRange().getValues();
    for (let i = 1; i < ud.length; i++) {
      let ps = []; try { ps = JSON.parse(ud[i][3] || "[]"); } catch (e) { continue; }
      let ch = false; ps.forEach(p => { if (p.project === oldName) { p.project = newName; ch = true; } });
      if (ch) us.getRange(i + 1, 4).setValue(JSON.stringify(ps));
    }
    const db = ensureDatabaseSheet(); ensureColumns();
    const c = buildColMap(db.getRange(1, 1, 1, db.getLastColumn()).getValues()[0]);
    const vals = db.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) if (vals[i][c["title"]] === oldName) db.getRange(i + 1, c["title"] + 1).setValue(newName);
  }
  clearProjectsCache_();
  logActivity(session, "updateProject", `row=${data.row} ${oldName} → ${newName} status=${status}`);
  return json({ success: true });
}

// ── Admin whitelist (phone or email + club role) — owner only ───────────────
function applyAccessToAccount_(acct, access) {
  const sh = ensureUsersSheet();
  sh.getRange(acct.row, 3).setValue(access.role);
  sh.getRange(acct.row, 9).setValue(access.clubRole);
}

function handleGetAdminEntries(session) {
  if (!session || session.role !== "owner") return json({ success: false, error: "Unauthorized." });
  const entries = getAdminEntries_().map(en => {
    const isEmail = en.id.indexOf("@") >= 0;
    const acct = isEmail ? findUserByEmail(en.id) : findUser(en.id);
    return { id: en.id, type: isEmail ? "email" : "phone", role: en.role, canShare: en.canShare !== false,
             account: acct ? { name: acct.name, email: acct.email, phone: acct.phone } : null };
  });
  return json({ success: true, entries: entries, clubRoles: getRoleLists().clubRoles });
}

function handleSetAdminEntry(data, session) {
  if (!session || session.role !== "owner") return json({ success: false, error: "Unauthorized." });
  const id   = normId_(data.id);
  const role = String(data.role || "").trim();
  const canShare = data.canShare === true;
  if (!id) return json({ success: false, error: "Enter a phone number or email." });
  if (id.indexOf("@") >= 0) {
    if (!EMAIL_RE.test(id)) return json({ success: false, error: "That email isn't valid." });
  } else if (id.length < 9 || id.length > 15) {
    return json({ success: false, error: "That phone number isn't valid." });
  }
  if (!getRoleLists().clubRoles.includes(role)) return json({ success: false, error: "Choose a role from the list." });

  const entries = getAdminEntries_();
  const hit = entries.find(x => x.id === id);
  if (hit) { hit.role = role; hit.canShare = canShare; } else entries.push({ id: id, role: role, canShare: canShare });
  saveAdminEntries_(entries);

  const acct = id.indexOf("@") >= 0 ? findUserByEmail(id) : findUser(id);
  if (acct) applyAccessToAccount_(acct, { role: "admin", clubRole: role });
  logActivity(session, "setAdminEntry", `${id} → ${role} share=${canShare}`);
  return json({ success: true });
}

function handleRemoveAdminEntry(data, session) {
  if (!session || session.role !== "owner") return json({ success: false, error: "Unauthorized." });
  const id = normId_(data.id);
  saveAdminEntries_(getAdminEntries_().filter(x => x.id !== id));
  const acct = id.indexOf("@") >= 0 ? findUserByEmail(id) : findUser(id);
  if (acct) applyAccessToAccount_(acct, { role: "user", clubRole: "" });   // back to a normal user
  logActivity(session, "removeAdminEntry", id);
  return json({ success: true });
}

// Older admin-page whitelist box (kept so the current page keeps working until it is replaced)
function handleGetWhitelist(session) {
  if (!session || session.role !== "owner") return json({ success: false, error: "Unauthorized." });
  return json({ success: true, whitelist: getAdminEntries_().map(x => x.id) });
}

function handleSetWhitelist(data, session) {
  if (!session || session.role !== "owner") return json({ success: false, error: "Unauthorized." });
  const ids = (data.whitelist || []).map(normId_).filter(Boolean);
  const old = getAdminEntries_();
  const next = ids.map(id => old.find(x => x.id === id) || { id: id, role: "Admin" });
  saveAdminEntries_(next);
  logActivity(session, "setWhitelist", ids.join(","));
  return json({ success: true });
}

function handleGetSheetUrls(session) {
  if (!session || session.role !== "owner") return json({ success: false, error: "Unauthorized." });
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const baseUrl = ss.getUrl();
  const result  = {};
  ["Database", "Projects", "Users", "ActivityLog"].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) result[name] = `${baseUrl}#gid=${s.getSheetId()}`;
  });
  return json({ success: true, urls: result });
}

// ---------------------------------------------------------------------------
// Auth — Email signup (Gmail OTP) + password login
// ---------------------------------------------------------------------------

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function hashPassword(pw, salt) {
  let h = salt + pw;
  for (let i = 0; i < 300; i++) {
    h = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  }
  return h;
}

// Validates a project list. New (not previously saved) projects must exist and not be completed.
function cleanProjects_(arr, existing) {
  const out = [], seen = {};
  for (const p of (Array.isArray(arr) ? arr : []).slice(0, 30)) {
    const name = String((p && p.project) || "").trim();
    const role = String((p && p.role) || "").trim().substring(0, 40);
    if (!name || !role || seen[name]) continue;
    const st = getProjectStatus_(name);
    if (st === null) return { error: "Unknown project: " + name };
    if (st === "completed" && !(existing || []).some(e => e.project === name))
      return { error: name + " is completed and can't be added." };
    seen[name] = true;
    out.push({ project: name, role: role });
  }
  return { projects: out };
}

// 1) Send OTP to the user's email — new users only
function handleEmailSendOtp(data) {
  const email = String(data.email || "").trim().toLowerCase();
  if (!EMAIL_RE.test(email))
    return json({ success: false, error: "Please enter a valid email address." });
  if (findUserByEmail(email))
    return json({ success: false, error: "This email is already registered. Please log in." });

  const cache = CacheService.getScriptCache();
  if (cache.get("eotp_cd_" + email))
    return json({ success: false, error: "Please wait 60 seconds before requesting another code." });

  const code = String(Math.floor(100000 + Math.random() * 900000)); // 6-digit
  cache.put("eotp_" + email, JSON.stringify({ code: code, tries: 0 }), 600);
  cache.put("eotp_cd_" + email, "1", 60);

  try {
    MailApp.sendEmail({
      to: email,
      subject: "Your verification code: " + code,
      htmlBody: "<div style='font-family:Arial,sans-serif'><h2>Time Slot Scheduler</h2>" +
        "<p>Your verification code is</p>" +
        "<p style='font-size:32px;font-weight:bold;letter-spacing:6px;color:#b34e1a'>" + code + "</p>" +
        "<p>It expires in 10 minutes. If you didn't request this, ignore this email.</p></div>"
    });
  } catch (e) {
    return json({ success: false, error: "Could not send email: " + e.message });
  }
  logActivity(null, "emailSendOtp", "email=" + email);
  return json({ success: true });
}

// 2) Verify OTP → one-time signup ticket
function handleEmailVerifyOtp(data) {
  const email = String(data.email || "").trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const raw   = cache.get("eotp_" + email);
  if (!raw) return json({ success: false, error: "Code expired. Request a new one." });
  const o = JSON.parse(raw);
  if (o.tries >= 5) {
    cache.remove("eotp_" + email);
    return json({ success: false, error: "Too many attempts. Request a new code." });
  }
  if (String(data.code || "").trim() !== o.code) {
    o.tries++;
    cache.put("eotp_" + email, JSON.stringify(o), 600);
    return json({ success: false, error: "Wrong code." });
  }
  cache.remove("eotp_" + email);
  const ticket = Utilities.getUuid();
  cache.put("eticket_" + ticket, email, 1800); // 30 min to finish signup
  return json({ success: true, ticket: ticket });
}

// 3) Create account: password + name + phone number + projects
function handleRegisterAccount(data) {
  const cache = CacheService.getScriptCache();
  const email = cache.get("eticket_" + data.ticket);
  if (!email) return json({ success: false, error: "Verification expired. Start again." });

  const pw = String(data.password || "");
  if (pw.length < 8) return json({ success: false, error: "Password must be at least 8 characters." });
  const name = String(data.name || "").trim().substring(0, 60);
  if (!name) return json({ success: false, error: "Enter your name." });
  const phone = normalizePhone(data.phone || "");
  if (phone.length < 9 || phone.length > 15) return json({ success: false, error: "Enter a valid phone number." });
  const cp = cleanProjects_(data.projects, []);
  if (cp.error) return json({ success: false, error: cp.error });
  if (!cp.projects.length) return json({ success: false, error: "Select at least one project and role." });
  const reserved = getOwnerPhone_();
  if (reserved && phone === reserved) return json({ success: false, error: "This number is reserved." });

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (findUserByEmail(email)) return json({ success: false, error: "Email already registered." });
    if (findUser(phone))
      return json({ success: false, error: "This phone number already has an account. Ask the owner to delete the old row in the Users sheet, then sign up again." });

    const access = resolveAccess(email, phone);
    const salt = Utilities.getUuid();
    ensureUsersSheet().appendRow([phone, safeCell_(name), access.role, JSON.stringify(cp.projects),
      email, salt, hashPassword(pw, salt), new Date(), access.clubRole]);

    cache.remove("eticket_" + data.ticket);
    const token = issueToken(phone, access.role, name);
    logActivity({ phone, role: access.role, name }, "register", "email=" + email);
    return json({
      success: true, token, paused: isSystemPaused(),
      user: { email, phone, name, role: access.role, clubRole: access.clubRole, canShare: access.canShare, projects: cp.projects, isNew: false }
    });
  } finally {
    lock.releaseLock();
  }
}

// 4) Login with email OR phone number + password
function handleLoginPassword(data) {
  const id = String(data.identifier || "").trim();
  const pw = String(data.password || "");
  if (!id || !pw) return json({ success: false, error: "Enter username and password." });

  const isEmail = id.indexOf("@") >= 0;
  const key     = "elf_" + (isEmail ? id.toLowerCase() : normalizePhone(id));
  const cache   = CacheService.getScriptCache();
  const fails   = parseInt(cache.get(key) || "0", 10);
  if (fails >= 5) return json({ success: false, error: "Too many failed attempts. Try again in 10 minutes." });

  // ── Owner login: owner phone + OWNER_PASSWORD from Script Properties ──────
  const props      = PropertiesService.getScriptProperties();
  const ownerPhone = getOwnerPhone_();
  const ownerName  = props.getProperty("OWNER_NAME") || "Owner";
  if (ownerPhone && !isEmail && normalizePhone(id) === ownerPhone) {
    if (!ownerPasswordOk_(pw)) {
      cache.put(key, String(fails + 1), 600);
      return json({ success: false, error: "Wrong username or password." });
    }
    cache.remove(key);
    const ou = findUser(ownerPhone);
    const otoken = issueToken(ownerPhone, "owner", ownerName);
    logActivity({ phone: ownerPhone, role: "owner", name: ownerName }, "login", "owner password");
    return json({
      success: true, token: otoken, paused: isSystemPaused(),
      user: { email: "", phone: ownerPhone, name: ownerName, role: "owner", clubRole: "Owner", canShare: true,
              projects: ou ? ou.projects : [], isNew: false }
    });
  }

  const user = isEmail ? findUserByEmail(id) : findUser(normalizePhone(id));
  const ok   = user && user.hash && user.salt && hashPassword(pw, user.salt) === user.hash;
  if (!ok) {
    cache.put(key, String(fails + 1), 600);
    return json({ success: false, error: "Wrong username or password." });
  }
  cache.remove(key);

  // Access is re-checked on every login (whitelist changes apply immediately)
  const access = resolveAccess(user.email, user.phone);
  if (access.role !== user.role || access.clubRole !== user.clubRole) applyAccessToAccount_(user, access);

  const token = issueToken(user.phone, access.role, user.name);
  logActivity({ phone: user.phone, role: access.role, name: user.name }, "login", "password");
  return json({
    success: true, token, paused: isSystemPaused(),
    user: { email: user.email, phone: user.phone, name: user.name, role: access.role,
            clubRole: access.clubRole, canShare: access.canShare, projects: user.projects, isNew: false }
  });
}

// 5) Logged-in user saves their project list
function handleSaveAccountProjects(data, session) {
  if (!session) return json({ success: false, error: "Unauthorized." });
  const user = findUser(session.phone);
  const cp = cleanProjects_(data.projects, user ? user.projects : []);
  if (cp.error) return json({ success: false, error: cp.error });
  if (!user) {
    if (session.role !== "owner") return json({ success: false, error: "Account not found." });
    ensureUsersSheet().appendRow([session.phone, session.name || "Owner", "owner", JSON.stringify(cp.projects), "", "", "", new Date(), "Owner"]);
  } else {
    ensureUsersSheet().getRange(user.row, 4).setValue(JSON.stringify(cp.projects));
  }
  logActivity(session, "saveAccountProjects", JSON.stringify(cp.projects));
  return json({ success: true, projects: cp.projects });
}

// 6) Change password (logged-in users)
function handleChangePassword(data, session) {
  if (!session) return json({ success: false, error: "Please log in again." });
  const user = findUser(session.phone);
  if (!user || !user.hash)
    return json({ success: false, error: "The owner password is kept in the Script Properties (see setupOwnerPassword)." });

  const cache = CacheService.getScriptCache();
  const key   = "cpf_" + session.phone;
  const fails = parseInt(cache.get(key) || "0", 10);
  if (fails >= 5) return json({ success: false, error: "Too many attempts. Try again in 10 minutes." });

  const oldPw = String(data.oldPassword || ""), newPw = String(data.newPassword || "");
  if (hashPassword(oldPw, user.salt) !== user.hash) {
    cache.put(key, String(fails + 1), 600);
    return json({ success: false, error: "Your current password is wrong." });
  }
  if (newPw.length < 8)  return json({ success: false, error: "New password must be at least 8 characters." });
  if (newPw === oldPw)   return json({ success: false, error: "New password must be different." });

  const salt = Utilities.getUuid();
  const sh = ensureUsersSheet();
  sh.getRange(user.row, 6).setValue(salt);
  sh.getRange(user.row, 7).setValue(hashPassword(newPw, salt));
  cache.remove(key);
  logActivity(session, "changePassword", "");
  return json({ success: true });
}

// ── Forgot password: a code by email, then a new password ───────────────────
// The answer is the same whether or not an account exists, so nobody can use this page to find out who is registered.
// The owner password is not reset here (it lives in Script Properties: use setupOwnerPassword).
function resetCode_() {
  const hex = Utilities.getUuid().replace(/-/g, "").substring(0, 8);           // random, not guessable
  return ("000000" + (parseInt(hex, 16) % 1000000)).slice(-6);
}

// The account for an email or phone typed on the login page, or null
function forgotAccount_(id) {
  const user = id.indexOf("@") >= 0 ? findUserByEmail(id) : findUser(normalizePhone(id));
  if (!user || !user.hash || !user.email) return null;
  if (user.phone === getOwnerPhone_() || isOwnerEmail_(user.email)) return null;
  return user;
}

function handleForgotSendOtp(data) {
  const id = String(data.identifier || "").trim();
  if (!id) return json({ success: false, error: "Enter your email or phone number." });

  const key   = id.indexOf("@") >= 0 ? id.toLowerCase() : normalizePhone(id);
  const cache = CacheService.getScriptCache();
  if (cache.get("fp_cd_" + key)) return json({ success: false, error: "Please wait 60 seconds before requesting another code." });
  const n = Number(cache.get("fp_n_" + key) || 0) + 1;
  if (n > 5) return json({ success: false, error: "Too many requests. Try again in an hour." });
  cache.put("fp_cd_" + key, "1", 60);
  cache.put("fp_n_" + key, String(n), 3600);

  const user = forgotAccount_(id);
  if (user) {
    const code = resetCode_();
    cache.put("fp_" + user.phone, JSON.stringify({ code: code, tries: 0 }), 600);   // valid 10 minutes
    try {
      MailApp.sendEmail({
        to: user.email,
        subject: "Your password reset code: " + code,
        htmlBody: "<div style='font-family:Arial,sans-serif'><h2>Time Slot Scheduler</h2>" +
          "<p>Use this code to choose a new password:</p>" +
          "<p style='font-size:32px;font-weight:bold;letter-spacing:6px;color:#b34e1a'>" + code + "</p>" +
          "<p>It expires in 10 minutes. If you didn't ask for this, ignore this email. Your password has not changed.</p></div>"
      });
    } catch (e) { console.error("Reset email failed: " + e.message); }
    logActivity(null, "forgotSendOtp", "phone=" + user.phone);
  }
  return json({ success: true });
}

function handleForgotReset(data) {
  const id    = String(data.identifier || "").trim();
  const newPw = String(data.newPassword || "");
  if (!id) return json({ success: false, error: "Enter your email or phone number." });
  if (newPw.length < 8) return json({ success: false, error: "New password must be at least 8 characters." });

  const bad   = json({ success: false, error: "The code is wrong or has expired. Request a new one." });   // one message for every failure
  const user  = forgotAccount_(id);
  if (!user) return bad;
  const cache = CacheService.getScriptCache();
  const ck    = "fp_" + user.phone;
  const raw   = cache.get(ck);
  if (!raw) return bad;
  let o = null; try { o = JSON.parse(raw); } catch (e) { cache.remove(ck); return bad; }
  if (o.tries >= 5) { cache.remove(ck); return bad; }
  if (!safeEqual_(String(data.code || "").trim(), String(o.code))) {
    o.tries++;
    cache.put(ck, JSON.stringify(o), 600);
    return bad;
  }

  const salt = Utilities.getUuid();
  const sh   = ensureUsersSheet();
  sh.getRange(user.row, 6).setValue(salt);
  sh.getRange(user.row, 7).setValue(hashPassword(newPw, salt));
  cache.remove(ck);
  cache.removeAll(["elf_" + user.email, "elf_" + user.phone, "cpf_" + user.phone, "fp_n_" + user.email, "fp_n_" + user.phone]);   // clears login lock-outs

  try {
    MailApp.sendEmail({
      to: user.email,
      subject: "Your password was changed",
      htmlBody: "<div style='font-family:Arial,sans-serif'><h2>Time Slot Scheduler</h2>" +
        "<p>Your password was just changed. If this was not you, ask the owner for help straight away.</p></div>"
    });
  } catch (e) { console.error("Password-changed email failed: " + e.message); }
  logActivity({ phone: user.phone, role: user.role, name: user.name }, "forgotReset", "");
  return json({ success: true });
}

// Run ONCE from the editor to grant Gmail-sending permission, then redeploy.
function authorizeMail() {
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), "Mail permission OK", "Authorized.");
}

// ---------------------------------------------------------------------------
// Clean-up: 7 days after the post date the MEDIA is deleted (saves storage),
// the post text stays in the sheet so history can still show it.
// ---------------------------------------------------------------------------
function autoDeleteOldPosts() {
  const sheet = ensureDatabaseSheet();
  ensureColumns();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col     = buildColMap(headers);
  const data    = sheet.getDataRange().getValues();
  const tz      = Session.getScriptTimeZone();
  const cutoff  = Utilities.formatDate(new Date(Date.now() - 7 * 24 * 60 * 60 * 1000), tz, "yyyy-MM-dd");
  let done = 0;
  for (let i = 1; i < data.length && done < 30; i++) {
    const dateStr = normDateServer(data[i][col["date"]]);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr) || dateStr >= cutoff) continue;
    const publicId = data[i][col["cloudinary_public_id"]];
    const mediaUrl = data[i][col["media"]];
    if (!publicId && !mediaUrl) continue;                 // media already removed
    if (publicId) {
      const res = deleteCloudinaryAsset(publicId, cloudinaryTypeFromUrl_(mediaUrl));
      if (!res.success && !/not found/i.test(res.error || "")) continue;   // retry next run
    }
    parseAlbum_(data[i][col["album_public_ids"]]).forEach(id => { if (id && id !== publicId) deleteCloudinaryAsset(id, "image"); });
    sheet.getRange(i + 1, col["media"] + 1).setValue("");
    sheet.getRange(i + 1, col["cloudinary_public_id"] + 1).setValue("");
    sheet.getRange(i + 1, col["album_media"] + 1).setValue("");
    sheet.getRange(i + 1, col["album_public_ids"] + 1).setValue("");
    done++;
  }
  if (done) console.log("autoDeleteOldPosts: removed media from " + done + " old post(s).");
}

function setupAutoDeleteTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "autoDeleteOldPosts") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("autoDeleteOldPosts").timeBased().everyDays(1).atHour(2).create();
  console.log("✅ Clean-up trigger created: runs daily at 2 AM.");
}

// ---------------------------------------------------------------------------
// AUTO-PUBLISHER (kept, but switched off: admins share manually)
// ---------------------------------------------------------------------------
const AUTO_SHARE_ENABLED = false;

function autoSchedulePosts() {
  if (!AUTO_SHARE_ENABLED || isSystemPaused()) return;
  const sheet = ensureDatabaseSheet();
  ensureColumns();

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col     = buildColMap(headers);
  const data    = sheet.getDataRange().getValues();
  const now     = new Date();
  const tzOff   = getTzOffsetHours();

  const EARLY_MS  =  5 * 60 * 1000;
  const LATE_MS   = 10 * 60 * 1000;

  const defaultChannelsRaw = PropertiesService.getScriptProperties()
    .getProperty("BUFFER_DEFAULT_CHANNELS") || "";
  const defaultChannels = defaultChannelsRaw.split(",").map(s => s.trim()).filter(Boolean);

  for (let i = 1; i < data.length; i++) {
    const r = data[i];

    const status    = (r[col["status"]]    || "").toString().toLowerCase();
    const scheduled = (r[col["scheduled"]] || "").toString().toLowerCase();
    const shared    = (r[col["shared"]]    || "").toString().toUpperCase();

    if (status    !== "approved") continue;
    if (scheduled !== "true")     continue;
    if (shared    === "TRUE")     continue;

    const dateStr = normDateServer(r[col["date"]]);
    const slotStr = normSlotServer(r[col["timeslot"]]);
    if (!dateStr || !slotStr) continue;

    const slotDate = slotToDate(dateStr, slotStr, tzOff);
    if (!slotDate) continue;
    const msFromSlot = now.getTime() - slotDate.getTime();

    if (msFromSlot < -EARLY_MS) continue;
    if (msFromSlot > LATE_MS) {
      sheet.getRange(i + 1, col["scheduled"] + 1).setValue("false");
      console.log(`autoSchedulePosts: row ${i} slot passed → set to manual`);
      continue;
    }

    const content  = String(r[col["content"]] || "");
    const mediaUrl = String(r[col["media"]]   || "");
    const title    = String(r[col["title"]]   || "");
    console.log(`autoSchedulePosts: publishing row ${i} "${title}" (slot: ${slotStr} ${dateStr})`);

    let autoSocialTargets = [];
    let youtubeTitle = null;
    let youtubeCategoryId = null;
    try {
      const rawSocial = r[col["auto_social_targets"]];
      if (rawSocial && rawSocial !== "") autoSocialTargets = JSON.parse(rawSocial);
    } catch(e) { console.warn(`Failed to parse auto_social_targets row ${i}:`, e); }
    if (r[col["youtube_title"]]) youtubeTitle = r[col["youtube_title"]];
    if (r[col["youtube_category_id"]]) youtubeCategoryId = r[col["youtube_category_id"]];

    const postIds   = [];
    const succeeded = [];
    const failed    = [];

    if (autoSocialTargets.length > 0) {
      for (const profile of autoSocialTargets) {
        const token = getBufferToken(profile.orgType || "social");
        if (!token) { failed.push(`${profile.id}: no token`); continue; }
        const isYt = (profile.serviceName || "").toLowerCase() === "youtube";
        const youtubeMeta = isYt ? { title: youtubeTitle || fitTitle_(title, YOUTUBE_TITLE_MAX), categoryId: youtubeCategoryId || "22" } : null;
        const autoOpts = isVideoUrl_(mediaUrl) ? {} : { tiktokTitle: fitTitle_(title, TIKTOK_TITLE_MAX) };
        const result = createBufferPost(token, profile.id, content, mediaUrl, profile.serviceName, youtubeMeta, autoOpts);
        if (result.success) { postIds.push(result.postId); succeeded.push(profile.serviceName || profile.id); }
        else failed.push(`${profile.id}: ${result.error}`);
      }
    } else if (defaultChannels.length > 0) {
      for (const entry of defaultChannels) {
        const parts   = entry.split(":");
        const chId    = parts[0];
        const orgType = parts[1] || "social";
        const token   = getBufferToken(orgType);
        if (!token) { failed.push(`${chId}: no token`); continue; }
        const result = createBufferPost(token, chId, content, mediaUrl, null, null);
        if (result.success) { postIds.push(result.postId); succeeded.push(chId); }
        else failed.push(`${chId}: ${result.error}`);
      }
    } else {
      console.warn(`autoSchedulePosts: no social destinations for row ${i}`);
    }

    if (succeeded.length > 0) {
      const sheetRow = i + 1;
      sheet.getRange(sheetRow, col["shared"]          + 1).setValue(true);
      sheet.getRange(sheetRow, col["buffer_post_ids"] + 1).setValue(postIds.join(","));
      sheet.getRange(sheetRow, col["platforms"]       + 1).setValue(succeeded.join(","));
      console.log(`autoSchedulePosts: row ${i} published to: ${succeeded.join(", ")}`);
    }
    if (failed.length) console.warn(`autoSchedulePosts: failures for row ${i}:`, failed.join("; "));
  }
}

function setupAutoScheduleTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "autoSchedulePosts") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("autoSchedulePosts").timeBased().everyMinutes(10).create();
  console.log("✅ Auto-publish trigger created: runs every 10 minutes.");
}

function handleSaveAutoDestinations(sheet, col, data, session) {
  if (!canShare_(session)) return json({ success: false, error: "You don't have permission to share posts." });
  const r = toSheetRow(data.row);
  if (data.socialProfiles !== undefined) {
    sheet.getRange(r, col["auto_social_targets"] + 1).setValue(JSON.stringify(data.socialProfiles));
  }
  if (data.youtubeTitle !== undefined) {
    sheet.getRange(r, col["youtube_title"] + 1).setValue(data.youtubeTitle);
  }
  if (data.youtubeCategoryId !== undefined) {
    sheet.getRange(r, col["youtube_category_id"] + 1).setValue(data.youtubeCategoryId);
  }
  logActivity(session, "saveAutoDestinations", `row=${data.row}`);
  return json({ success: true });
}

// ---------------------------------------------------------------------------
// Buffer asset shape: ask Buffer's own schema how media must be sent (cached 6 h)
// ---------------------------------------------------------------------------
function getBufferAssetShape(token) {
  const cache = CacheService.getScriptCache();
  const hit   = cache.get("BUFFER_ASSET_SHAPE");
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  try {
    const typeFrag = "kind name ofType { kind name ofType { kind name ofType { kind name } } }";
    const query = `{
      ci: __type(name: "CreatePostInput") { inputFields { name type { ${typeFrag} } } }
      ai: __type(name: "AssetInput")      { inputFields { name type { ${typeFrag} } } }
    }`;
    const r = UrlFetchApp.fetch(BUFFER_GRAPHQL_URL, {
      method: "post", contentType: "application/json",
      headers: { Authorization: "Bearer " + token },
      payload: JSON.stringify({ query }), muteHttpExceptions: true
    });
    const body = JSON.parse(r.getContentText());
    const ci = body.data && body.data.ci, ai = body.data && body.data.ai;
    if (!ci || !ai) return null;
    const isList = (t) => { for (let x = t; x; x = x.ofType) if (x.kind === "LIST") return true; return false; };
    const assetsField = ci.inputFields.find(f => f.name === "assets");
    const names = ai.inputFields.map(f => f.name);
    const pick = (re) => {
      const f = ai.inputFields.find(x => re.test(x.name));
      return f ? { key: f.name, isList: isList(f.type) } : null;
    };
    const shape = {
      assetsIsList: assetsField ? isList(assetsField.type) : false,
      image: pick(/^images?$/i),
      video: pick(/^videos?$/i),
      fields: names
    };
    cache.put("BUFFER_ASSET_SHAPE", JSON.stringify(shape), 21600);
    return shape;
  } catch (e) {
    console.error("getBufferAssetShape error:", e);
    return null;
  }
}

// YouTube metadata: read Buffer's schema and fill every REQUIRED field (cached 6 h)
function getYoutubeMetaShape(token) {
  const cache = CacheService.getScriptCache();
  const hit   = cache.get("BUFFER_YT_SHAPE");
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  try {
    const gql = (query) => JSON.parse(UrlFetchApp.fetch(BUFFER_GRAPHQL_URL, {
      method: "post", contentType: "application/json",
      headers: { Authorization: "Bearer " + token },
      payload: JSON.stringify({ query }), muteHttpExceptions: true
    }).getContentText());

    const typeFrag = "kind name ofType { kind name ofType { kind name ofType { kind name } } }";
    const first = gql(`{ yt: __type(name: "YoutubePostMetadataInput") { inputFields { name type { ${typeFrag} } } } }`);
    const yt = first.data && first.data.yt;
    if (!yt) return null;

    const unwrap = (t) => { let x = t; while (x && (x.kind === "NON_NULL" || x.kind === "LIST")) x = x.ofType; return x || {}; };
    const shape = {};
    const enumTypes = {};
    yt.inputFields.forEach(f => {
      const base = unwrap(f.type);
      shape[f.name] = { required: f.type.kind === "NON_NULL", type: base.name, kind: base.kind, values: null };
      if (base.kind === "ENUM") enumTypes[base.name] = true;
    });
    const enumNames = Object.keys(enumTypes);
    if (enumNames.length) {
      const q = "{ " + enumNames.map((n, i) => `e${i}: __type(name: "${n}") { enumValues { name } }`).join(" ") + " }";
      const ev = gql(q).data || {};
      enumNames.forEach((n, i) => {
        const vals = ev["e" + i] ? ev["e" + i].enumValues.map(v => v.name) : [];
        Object.keys(shape).forEach(k => { if (shape[k].type === n) shape[k].values = vals; });
      });
    }
    cache.put("BUFFER_YT_SHAPE", JSON.stringify(shape), 21600);
    return shape;
  } catch (e) {
    console.error("getYoutubeMetaShape error:", e);
    return null;
  }
}

function buildYoutubeMetadata(token, base) {
  const shape = getYoutubeMetaShape(token);
  if (!shape) return base;
  const out = {};
  const pickEnum = (f, want) => {
    const vals = f.values || [];
    if (!vals.length) return want;
    const hit = vals.find(v => String(v).toLowerCase() === String(want).toLowerCase());
    return hit || vals.find(v => /^public$/i.test(v)) || vals[0];
  };
  Object.keys(base).forEach(k => {
    const YT_ALIASES = {
      madeForKids: ["selfDeclaredMadeForKids", "isMadeForKids"], notifySubscribers: ["notify"],
      embeddable: ["allowEmbedding"], license: ["licence"],
      aiGenerated: ["containsSyntheticMedia", "isAiGenerated", "syntheticMedia", "alteredContent"]
    };
    const name = [k, k.replace(/Id$/, ""), k + "Id"].concat(YT_ALIASES[k] || []).find(n => shape[n]);
    if (!name) return;
    const f = shape[name];
    let v = base[k];
    if (f.kind === "ENUM")                  v = pickEnum(f, v);
    else if (f.type === "Int")              v = Number(v) || 0;
    else if (f.type === "Boolean")          v = !!v;
    else if (f.kind === "SCALAR")           v = String(v);
    out[name] = v;
  });
  Object.keys(shape).forEach(name => {
    const f = shape[name];
    if (!f.required || out[name] !== undefined) return;
    if (f.kind === "ENUM")                  out[name] = pickEnum(f, "public");
    else if (f.type === "Boolean")          out[name] = /embed/i.test(name);
    else if (f.type === "Int" || f.type === "Float") out[name] = 0;
  });
  return out;
}

// ---------------------------------------------------------------------------
// Buffer API — publish immediately
// ---------------------------------------------------------------------------
// kind = "post" | "story".  urls = array of media URLs (an album = several images).
function sendBufferPost_(token, channelId, text, urls, service, youtubeMeta, kind, tiktokTitle) {
  const input = { text: text, channelId: channelId, schedulingType: "automatic", mode: "shareNow" };
  const isVid = urls.length === 1 && isVideoUrl_(urls[0]);

  let assetVariants = [null];
  if (urls.length) {
    const kindKey = isVid ? "video" : "image";
    const objs = urls.map(u => { const o = {}; o[kindKey] = { url: u }; return o; });
    assetVariants = [objs];
    if (urls.length === 1) assetVariants.push(objs[0], { [kindKey]: [{ url: urls[0] }] }, { [kindKey + "s"]: [{ url: urls[0] }] });
    else assetVariants.push({ images: urls.map(u => ({ url: u })) });
    const shape = getBufferAssetShape(token);
    const f = shape && (isVid ? shape.video : shape.image);
    if (f) {
      const shaped = urls.map(u => { const inner = { url: u }, o = {}; o[f.key] = f.isList ? [inner] : inner; return o; });
      if (shape.assetsIsList) assetVariants.unshift(shaped);
      else if (shaped.length === 1) assetVariants.unshift(shaped[0]);
    }
  }

  let metaVariants = [null];
  if (service === "facebook")  metaVariants = [{ facebook: { type: kind } }];
  if (service === "instagram") {
    metaVariants = kind === "story" ? [{ instagram: { type: "story", shouldShareToFeed: false } }, { instagram: { type: "story" } }]
      : isVid ? [{ instagram: { type: "reel", shouldShareToFeed: true } }, { instagram: { type: "post", shouldShareToFeed: true } }]
              : [{ instagram: { type: "post", shouldShareToFeed: true } }];
  }
  if (service === "tiktok") {
    // Photo posts: the title chosen on the Admin page. Videos: the start of the text. Never longer than 90.
    const tt = cleanTitle_(tiktokTitle) || cleanTitle_(text);
    input.metadata = { tiktok: { title: tt.substring(0, TIKTOK_TITLE_MAX).trim() } };
  }
  if (service === "youtube") {
    const y = youtubeMeta || {};
    const ytTitle = cleanTitle_(y.title ? y.title : text).substring(0, YOUTUBE_TITLE_MAX).trim();
    const base = { title: ytTitle, categoryId: y.categoryId || "22", privacy: y.privacy || "public" };
    ["madeForKids", "notifySubscribers", "embeddable", "license", "aiGenerated"].forEach(k => { if (y[k] !== undefined && y[k] !== "") base[k] = y[k]; });
    input.metadata = { youtube: buildYoutubeMetadata(token, base) };
  }

  const mutation = `
    mutation CreatePost($input: CreatePostInput!) {
      createPost(input: $input) {
        ... on PostActionSuccess { post { id } }
        ... on MutationError    { message }
      }
    }
  `;
  try {
    let lastError = "No post ID returned.";
    for (let mi = 0; mi < metaVariants.length; mi++) {
      if (metaVariants[mi]) input.metadata = metaVariants[mi];
      for (const assets of assetVariants) {
        if (assets) input.assets = assets; else delete input.assets;
        const response = UrlFetchApp.fetch(BUFFER_GRAPHQL_URL, {
          method: "post", contentType: "application/json",
          headers: { Authorization: `Bearer ${token}` },
          payload: JSON.stringify({ query: mutation, variables: { input } }),
          muteHttpExceptions: true
        });
        const result = JSON.parse(response.getContentText());
        console.log("Buffer response for", channelId, ":", JSON.stringify(result));
        if (result.data?.createPost?.post?.id) return { success: true, postId: result.data.createPost.post.id };
        lastError = result.errors ? result.errors[0].message
                  : (result.data?.createPost?.message || "No post ID returned.");
        if (!/AssetInput|input\.assets/i.test(lastError)) break;
      }
      const typeProblem = /InstagramPostType|FacebookPostType|invalid value|enum|does not exist|unknown/i.test(lastError)
                       && !/minute|second|longer|duration|size|aspect|resolution/i.test(lastError);
      if (!typeProblem) break;
    }
    return { success: false, error: lastError };
  } catch (err) {
    console.error("createBufferPost error:", err);
    return { success: false, error: err.toString() };
  }
}

// opts: { destination: "post"|"story", mention: "handle" }
function createBufferPost(token, channelId, text, mediaUrls, service, youtubeMeta, opts) {
  opts = opts || {};
  let urls = (typeof mediaUrls === "string" ? [mediaUrls] : (mediaUrls || [])).filter(Boolean);
  const hasVideo = urls.some(isVideoUrl_);
  if (hasVideo) urls = [urls.find(isVideoUrl_)];                 // albums are images only

  if (service === "youtube" && !hasVideo && !ytPhotosOn_())
    return { success: false, error: "YouTube only accepts videos (Buffer does not support image posts there)." };

  if (opts.destination === "story") {
    // Second check on the server: Instagram Stories stay off unless the switch is on (Facebook Stories are not affected)
    if (service === "instagram" && !igStoriesOn_()) return { success: false, error: "Instagram Stories are turned off." };
    if (service !== "instagram" && service !== "facebook") return { success: false, error: "Stories are only for Instagram and Facebook." };
    if (!urls.length) return { success: false, error: "A story needs an image or a video." };
    const ids = [], errs = [];
    urls.forEach(u => {                                          // one story per image
      const r = sendBufferPost_(token, channelId, text, [withStoryMention_(u, opts.mention)], service, null, "story");
      if (r.success) ids.push(r.postId); else errs.push(r.error);
    });
    return ids.length ? { success: true, postId: ids.join("+"), partial: errs.length ? errs[0] : "" } : { success: false, error: errs[0] || "Story failed." };
  }

  if (service === "instagram" && !hasVideo && urls.length > IG_ALBUM_MAX) urls = urls.slice(0, IG_ALBUM_MAX);   // Instagram gets the first 10
  return sendBufferPost_(token, channelId, text, urls, service, youtubeMeta, "post", opts.tiktokTitle);
}

// ---------------------------------------------------------------------------
// Cloudinary — delete asset (images AND videos)
// ---------------------------------------------------------------------------
function deleteCloudinaryAsset(publicId, resourceType) {
  if (!publicId) return { success: false, error: "No public ID." };
  const props     = PropertiesService.getScriptProperties();
  const cloudName = props.getProperty("CLOUDINARY_CLOUD_NAME");
  const apiKey    = props.getProperty("CLOUDINARY_API_KEY");
  const apiSecret = props.getProperty("CLOUDINARY_API_SECRET");
  if (!cloudName || !apiKey || !apiSecret) return { success: false, error: "Cloudinary credentials not set." };
  const type         = resourceType === "video" ? "video" : "image";
  const timestamp    = Math.floor(Date.now() / 1000);
  const stringToSign = `public_id=${publicId}&timestamp=${timestamp}${apiSecret}`;
  const signature    = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, stringToSign)
    .map(b => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
  try {
    const response = UrlFetchApp.fetch(`https://api.cloudinary.com/v1_1/${cloudName}/${type}/destroy`, {
      method: "post", payload: { public_id: publicId, timestamp: String(timestamp), api_key: apiKey, signature },
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());
    if (result.result === "ok")        return { success: true };
    if (result.result === "not found") return { success: false, error: "not found" };
    return { success: false, error: result.error?.message || "Unknown." };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ---------------------------------------------------------------------------
// Diagnostics: quick scan (admins, read-only) and full scan (owner only)
// ---------------------------------------------------------------------------
function fetchBufferChannels_(orgType) {
  const t = getBufferToken(orgType);
  if (!t) return { ok: false, missing: true, error: "No token set." };
  try {
    const r = UrlFetchApp.fetch(BUFFER_GRAPHQL_URL, {
      method: "post", contentType: "application/json",
      headers: { Authorization: "Bearer " + t },
      payload: JSON.stringify({ query: `{ channels(input: { organizationId: "${getBufferOrgId(orgType)}" }) { id name service } }` }),
      muteHttpExceptions: true
    });
    const body = JSON.parse(r.getContentText());
    if (body.errors) return { ok: false, error: String(body.errors[0].message || "Buffer error").substring(0, 160) };
    const ch = (body.data && body.data.channels) || [];
    return { ok: true, count: ch.length, services: Array.from(new Set(ch.map(c => String(c.service || "?").toLowerCase()))) };
  } catch (e) { return { ok: false, error: String(e).substring(0, 160) }; }
}

function handleScan_(session, full) {
  if (!session || session.role === "user") return json({ success: false, error: "Unauthorized." });
  if (full && session.role !== "owner")    return json({ success: false, error: "Only the owner can run the full scan." });

  const started = Date.now();
  const res = [];
  const add = (group, name, level, detail) => res.push({ group: group, name: name, level: level, detail: detail || "" });
  const P   = PropertiesService.getScriptProperties();
  const has = (k) => !!String(P.getProperty(k) || "").trim();
  const ss  = SpreadsheetApp.getActiveSpreadsheet();

  // ── Settings ──────────────────────────────────────────────────────────────
  add("Settings", "Login token secret", has("TOKEN_SECRET") ? "ok" : "error",
      has("TOKEN_SECRET") ? "Set." : "TOKEN_SECRET is missing, so logins use an insecure default. Add it in Script Properties.");
  const ownerPwSet = has("OWNER_PASSWORD_HASH") || has("OWNER_PASSWORD");
  add("Settings", "Owner login", (ownerPwSet && has("OWNER_PHONE")) ? "ok" : (ownerPwSet ? "warn" : "error"),
      ownerPwSet ? (has("OWNER_PHONE") ? "Owner phone and password are set." : "OWNER_PHONE is not set, so the owner cannot log in.")
                 : "The owner password is missing: the owner can't log in with the phone number.");
  add("Settings", "Owner password stored as a hash", has("OWNER_PASSWORD_HASH") ? "ok" : "warn",
      has("OWNER_PASSWORD_HASH") ? "Yes." : "The owner password is still plain text in Script Properties. Run setupOwnerPassword() once.");
  add("Settings", "Owner name", has("OWNER_NAME") ? "ok" : "warn", has("OWNER_NAME") ? "Set." : "OWNER_NAME is not set, so the welcome text says \"Owner\".");
  add("Settings", "Time zone", has("TZ_OFFSET_HOURS") ? "ok" : "warn",
      "Script time zone " + Session.getScriptTimeZone() + ", slot offset UTC" + (getTzOffsetHours() >= 0 ? "+" : "") + getTzOffsetHours() +
      (has("TZ_OFFSET_HOURS") ? "." : " (default; add TZ_OFFSET_HOURS to change)."));
  add("Settings", "Buffer token (Facebook / Instagram / LinkedIn)", has("BUFFER_TOKEN_SOCIAL") ? "ok" : "warn",
      has("BUFFER_TOKEN_SOCIAL") ? "Set." : "BUFFER_TOKEN_SOCIAL is missing: those channels can't be shared to.");
  add("Settings", "Buffer token (YouTube / TikTok)", has("BUFFER_TOKEN_VIDEO") ? "ok" : "warn",
      has("BUFFER_TOKEN_VIDEO") ? "Set." : "BUFFER_TOKEN_VIDEO is missing: YouTube and TikTok can't be shared to.");
  const cloudOk = has("CLOUDINARY_CLOUD_NAME") && has("CLOUDINARY_API_KEY") && has("CLOUDINARY_API_SECRET");
  add("Settings", "Cloudinary keys", cloudOk ? "ok" : "warn", cloudOk ? "Set." : "Cloudinary keys are missing: old media can't be deleted.");

  // ── System state ──────────────────────────────────────────────────────────
  add("System", "Pause switch", isSystemPaused() ? "warn" : "ok",
      isSystemPaused() ? "The system is PAUSED. Normal users see the maintenance message." : "System is running.");
  scanExtras_(add);
  const triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  add("System", "Media clean-up trigger", triggers.includes("autoDeleteOldPosts") ? "ok" : "warn",
      triggers.includes("autoDeleteOldPosts") ? "Runs daily." : "Not set up. Run setupAutoDeleteTrigger once, or old media will pile up.");
  let quota = null;
  try { quota = MailApp.getRemainingDailyQuota(); } catch (e) {}
  if (quota !== null) add("System", "Email (signup codes)", quota > 10 ? "ok" : "warn", quota + " emails left today.");

  // ── Sheets: tabs and columns ──────────────────────────────────────────────
  const wanted = {
    "Database":    DB_HEADERS,
    "Projects":    ["name", "status", "color"],
    "Users":       ["phone", "name", "role", "projects", "email", "salt", "hash", "createdat", "club_role"],
    "ActivityLog": ["timestamp", "phone", "name", "role", "action", "detail"]
  };
  Object.keys(wanted).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { add("Sheets", name + " tab", "error", "The tab is missing."); return; }
    const head = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase()) : [];
    const missing = wanted[name].filter(h => !head.includes(h.toLowerCase()));
    add("Sheets", name + " tab", missing.length ? "error" : "ok",
        missing.length ? "Missing columns: " + missing.join(", ") : (sh.getLastRow() - 1) + " rows, all columns present.");
  });

  if (full) {
    // ── Sheets: write + read back test (hidden _diagnostics tab) ─────────────
    try {
      let ds = ss.getSheetByName("_diagnostics");
      if (!ds) { ds = ss.insertSheet("_diagnostics"); ds.appendRow(["time", "token"]); ds.hideSheet(); }
      const tok = Utilities.getUuid();
      ds.appendRow([new Date().toISOString(), tok]);
      const last = ds.getLastRow();
      const back = String(ds.getRange(last, 2).getValue());
      ds.deleteRow(last);
      add("Sheets", "Write and read-back test", back === tok ? "ok" : "error",
          back === tok ? "Wrote a test row, read it back and removed it." : "The text read back did not match what was written.");
    } catch (e) { add("Sheets", "Write and read-back test", "error", String(e).substring(0, 160)); }

    // ── Data health ──────────────────────────────────────────────────────────
    try {
      ensureColumns();
      const db = ensureDatabaseSheet();
      const head = db.getRange(1, 1, 1, db.getLastColumn()).getValues()[0];
      const col  = buildColMap(head);
      const rng  = db.getDataRange(), vals = rng.getValues(), disp = rng.getDisplayValues();
      let bad = 0, dup = 0, oldMedia = 0, noProject = 0;
      const seen = {};
      const cutoff = Utilities.formatDate(new Date(Date.now() - 8 * 24 * 60 * 60 * 1000), Session.getScriptTimeZone(), "yyyy-MM-dd");
      for (let i = 1; i < vals.length; i++) {
        const d = normDateServer(vals[i][col["date"]]), sl = normSlotServer(disp[i][col["timeslot"]]);
        if (!disp[i][col["title"]] && !d) continue;
        if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || !/^\d{1,2}:\d{2} (AM|PM)$/.test(sl)) bad++;
        if (!String(disp[i][col["title"]] || "").trim()) noProject++;
        if (String(vals[i][col["status"]]).toLowerCase() !== "rejected") {
          const k = d + "|" + sl; if (seen[k]) dup++; seen[k] = true;
        }
        if (d && d < cutoff && String(vals[i][col["media"]] || "")) oldMedia++;
      }
      add("Data", "Posts: date and time slot format", bad ? "warn" : "ok", bad ? bad + " post(s) have an unreadable date or slot." : "All posts readable.");
      add("Data", "Posts: double-booked slots", dup ? "warn" : "ok", dup ? dup + " slot(s) hold more than one post." : "No double bookings.");
      add("Data", "Posts: project name", noProject ? "warn" : "ok", noProject ? noProject + " older post(s) have no project name (shown as Untitled)." : "All posts have a project name.");
      add("Data", "Old media clean-up", oldMedia ? "warn" : "ok", oldMedia ? oldMedia + " post(s) older than 7 days still have media. The daily clean-up should remove it." : "No old media left behind.");

      const users = ensureUsersSheet().getDataRange().getValues();
      let noEmail = 0, noPw = 0;
      for (let i = 1; i < users.length; i++) { if (!users[i][0]) continue; if (!users[i][4]) noEmail++; if (!users[i][6]) noPw++; }
      add("Users", "Accounts", noPw ? "warn" : "ok", (users.length - 1) + " accounts" + (noPw ? "; " + noPw + " old account(s) have no password and can't log in (ask them to sign up again)." : "."));
      const clubRoles = getRoleLists().clubRoles;
      const entries = getAdminEntries_();
      const badRole = entries.filter(e => !clubRoles.includes(e.role) && e.role !== "Admin");
      add("Users", "Admin whitelist", badRole.length ? "warn" : "ok",
          entries.length + " admin entr" + (entries.length === 1 ? "y" : "ies") + (badRole.length ? "; " + badRole.length + " have a role that is no longer in the role list." : "."));
      const projects = ensureProjectsSheet().getDataRange().getValues().slice(1).filter(r => r[0]);
      add("Data", "Projects", projects.length ? "ok" : "warn", projects.length ? projects.length + " projects." : "No projects yet.");
    } catch (e) { add("Data", "Data checks", "error", String(e).substring(0, 160)); }

    // ── Buffer channels ──────────────────────────────────────────────────────
    ["social", "video"].forEach(orgType => {
      const label = orgType === "social" ? "Buffer: Facebook / Instagram / LinkedIn" : "Buffer: YouTube / TikTok";
      const r = fetchBufferChannels_(orgType);
      if (r.missing) add("Social media", label, "warn", "No token set.");
      else if (!r.ok) add("Social media", label, "error", r.error);
      else add("Social media", label, r.count ? "ok" : "warn", r.count ? r.count + " channel(s): " + r.services.join(", ") : "Connected, but no channels were found.");
    });

    // ── Cloudinary ───────────────────────────────────────────────────────────
    if (cloudOk) {
      try {
        const r = UrlFetchApp.fetch("https://api.cloudinary.com/v1_1/" + P.getProperty("CLOUDINARY_CLOUD_NAME") + "/ping", {
          headers: { Authorization: "Basic " + Utilities.base64Encode(P.getProperty("CLOUDINARY_API_KEY") + ":" + P.getProperty("CLOUDINARY_API_SECRET")) },
          muteHttpExceptions: true
        });
        add("Storage", "Cloudinary keys work", r.getResponseCode() === 200 ? "ok" : "error",
            r.getResponseCode() === 200 ? "Cloudinary accepted the keys." : "Cloudinary replied " + r.getResponseCode() + ". Check the three CLOUDINARY_ properties.");
      } catch (e) { add("Storage", "Cloudinary keys work", "error", String(e).substring(0, 160)); }
    }
  }

  if (session.role !== "owner") {
    const worst = res.some(r => r.level === "error") ? "error" : res.some(r => r.level === "warn") ? "warn" : "ok";
    res.length = 0;
    res.push({ group: "System", name: "System", level: worst, detail: worst === "ok" ? "" : "There is a system problem. Please tell the owner." });
  }

  return json({ success: true, full: !!full, results: res, serverOffset: getTzOffsetHours(), scriptTz: Session.getScriptTimeZone(), ms: Date.now() - started });
}

// Removes the tiny test file the diagnostics page uploads
function handleDiagDeleteAsset(data, session) {
  if (!session || session.role !== "owner") return json({ success: false, error: "Unauthorized." });
  const id = String(data.publicId || "");
  if (!/^[\w\-\/]+$/.test(id)) return json({ success: false, error: "Bad id." });
  return json(deleteCloudinaryAsset(id, "image"));
}

// ===========================================================================
// UPDATE 3 — albums, stories, TikTok/YouTube options, profile edits,
// in-app directory views, email reminders
// ===========================================================================
const ALBUM_USER_MAX  = 10;   // lowest platform limit (Instagram)
const ALBUM_ADMIN_MAX = 10;   // same limit for admins: Buffer does not accept more
const IG_ALBUM_MAX    = 10;
const TIKTOK_PHOTO_TEXT_MAX = 4000;
const TIKTOK_VIDEO_TEXT_MAX = 2200;
const TIKTOK_TITLE_MAX   = 90;     // TikTok photo-post title
const YOUTUBE_TITLE_MAX  = 100;
const YOUTUBE_DESC_MAX   = 5000;   // the post text is sent as the YouTube description
const TITLE_SUFFIX       = " | LeosofUOP";

// Script Properties switches (Project settings → Script properties):
//   INSTAGRAM_STORIES_ENABLED = "true"  → Instagram Stories allowed. Missing or anything else = OFF. Facebook Stories are not affected.
//   YOUTUBE_PHOTO_POSTS       = "true"  → try image posts on YouTube (Buffer may refuse them). Missing = OFF.
function igStoriesOn_() { return PropertiesService.getScriptProperties().getProperty("INSTAGRAM_STORIES_ENABLED") === "true"; }
function ytPhotosOn_()  { return PropertiesService.getScriptProperties().getProperty("YOUTUBE_PHOTO_POSTS") === "true"; }

// Titles: no < >, one line, never longer than the platform allows.
function cleanTitle_(s) { return String(s == null ? "" : s).replace(/[<>]/g, "").replace(/\s+/g, " ").trim(); }
// "Project Name | LeosofUOP". If it is too long the project name is shortened and the suffix is kept.
function fitTitle_(name, max) {
  let n = cleanTitle_(name);
  if ((n + TITLE_SUFFIX).length > max) n = n.substring(0, Math.max(0, max - TITLE_SUFFIX.length)).trim();
  return n + TITLE_SUFFIX;
}
const REMINDER_MIN_BEFORE   = 10;   // second email 10 min before the slot (11:50 for 12:00)
const DEFAULT_NOTIFY_ROLES  = ["Chief Media Officer", "Chief PR Officer"];

function isVideoUrl_(u) { u = String(u || ""); return /\.(mp4|mov|webm|avi|m4v|3gp|mpeg|ogg)(\?|$)/i.test(u) || u.indexOf("/video/upload/") >= 0; }

function parseAlbum_(v) {
  if (!v) return [];
  try { const a = JSON.parse(v); return Array.isArray(a) ? a.map(String).filter(Boolean) : []; } catch (e) { return []; }
}

// Burn the Instagram handle into a story image (Buffer cannot add a clickable mention automatically)
function withStoryMention_(url, handle) {
  handle = String(handle || "").trim().replace(/^@/, "");
  if (!handle || isVideoUrl_(url) || url.indexOf("/image/upload/") < 0) return url;
  const layer = "l_text:Arial_60_bold:" + encodeURIComponent("@" + handle) + ",co_white,b_rgb:00000080,g_south,y_140/";
  return url.replace("/image/upload/", "/image/upload/" + layer);
}

// ── Notification settings (owner) ───────────────────────────────────────────
function getNotifyRoles_() { return readList_("NOTIFY_ROLES", DEFAULT_NOTIFY_ROLES); }
function getNotifyExtra_() {
  try { const a = JSON.parse(PropertiesService.getScriptProperties().getProperty("NOTIFY_EXTRA") || "[]"); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}

function handleGetNotify(session) {
  if (!session || session.role !== "owner") return json({ success: false, error: "Unauthorized." });
  const on = getNotifyRoles_();
  return json({ success: true,
    roles: getRoleLists().clubRoles.map(r => ({ role: r, on: on.indexOf(r) >= 0 })),
    extra: getNotifyExtra_(),
    storyMention: String(PropertiesService.getScriptProperties().getProperty("STORY_MENTION") || ""),
    quota: MailApp.getRemainingDailyQuota() });
}

function handleSaveNotify(data, session) {
  if (!session || session.role !== "owner") return json({ success: false, error: "Unauthorized." });
  const P = PropertiesService.getScriptProperties();
  if (Array.isArray(data.roles)) P.setProperty("NOTIFY_ROLES", JSON.stringify(data.roles.map(String)));
  if (Array.isArray(data.extra)) {
    const emails = Array.from(new Set(data.extra.map(s => String(s).trim().toLowerCase()).filter(Boolean)));
    const bad = emails.find(e => !EMAIL_RE.test(e));
    if (bad) return json({ success: false, error: "Not a valid email: " + bad });
    P.setProperty("NOTIFY_EXTRA", JSON.stringify(emails.slice(0, 30)));
  }
  if (data.storyMention !== undefined) P.setProperty("STORY_MENTION", String(data.storyMention).trim().replace(/^@/, "").substring(0, 60));
  logActivity(session, "saveNotify", "");
  return handleGetNotify(session);
}

function handleGetShareSettings(session) {
  if (!canShare_(session)) return json({ success: false, error: "You don't have permission to share posts." });
  return json({ success: true,
    storyMention: String(PropertiesService.getScriptProperties().getProperty("STORY_MENTION") || ""),
    instagramStories: igStoriesOn_(), youtubePhotos: ytPhotosOn_() });
}

function officerEmails_() {
  const roles = getNotifyRoles_(), out = [];
  getAdminEntries_().forEach(en => {
    if (roles.indexOf(en.role) < 0) return;
    let email = "";
    if (en.id.indexOf("@") >= 0) email = en.id;
    else { const u = findUser(en.id); email = u ? u.email : ""; }
    if (email && EMAIL_RE.test(email)) out.push(email);
  });
  return Array.from(new Set(out));
}

function sendMailSafe_(to, subject, html) {
  to = Array.from(new Set(to.filter(Boolean)));
  if (!to.length) return false;
  try {
    if (MailApp.getRemainingDailyQuota() < to.length) { console.error("Email quota too low"); return false; }
    MailApp.sendEmail({ to: to.join(","), subject: subject, htmlBody: html });
    return true;
  } catch (e) { console.error("sendMailSafe_ error:", e); return false; }
}

function esc_(s) { return String(s == null ? "" : s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
function postMailBody_(heading, project, date, slot, author, note) {
  return "<div style='font-family:Arial,sans-serif'><h3 style='color:#b34e1a'>" + esc_(heading) + "</h3>" +
    "<p><b>Project:</b> " + esc_(project) + "<br><b>Date:</b> " + esc_(date) + "<br><b>Time slot:</b> " + esc_(slot) +
    "<br><b>Posted by:</b> " + esc_(author || "-") + "</p><p>" + esc_(note) + "</p>" +
    "<p><a href='" + SITE_URL + "' style='display:inline-block;background:#b34e1a;color:#ffffff;text-decoration:none;" +
    "padding:10px 20px;border-radius:20px;font-weight:bold'>Open the site</a></p>" +
    "<p style='font-size:12px;color:#777'>Or copy this link: " + SITE_URL + "</p></div>";
}

// Email 1: sent the moment a post is uploaded
function notifyNewPost_(project, date, slot, author) {
  sendMailSafe_(officerEmails_(), "New post waiting for approval: " + project + " (" + date + " " + slot + ")",
    postMailBody_("New post to approve", project, date, slot, author, "Please open the Admin Panel, check it and approve it."));
}

// Email 2: runs every 5 minutes; fires once per post, ~10 minutes before the slot
function checkPendingReminders() {
  const sheet = ensureDatabaseSheet();
  ensureColumns();
  const col  = buildColMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
  const data = sheet.getDataRange().getValues();
  const disp = sheet.getDataRange().getDisplayValues();
  const off  = getTzOffsetHours();
  const now  = Date.now();
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][col["status"]] || "pending").toLowerCase();
    if (status !== "pending") continue;
    if (String(disp[i][col["shared"]]).toUpperCase() === "TRUE") continue;
    if (String(disp[i][col["reminder_sent"]] || "")) continue;
    const date = normDateServer(data[i][col["date"]]);
    const slot = normSlotServer(disp[i][col["timeslot"]]);
    const start = slotToDate(date, slot, off);
    if (!start) continue;
    const minsLeft = (start.getTime() - now) / 60000;
    if (minsLeft > REMINDER_MIN_BEFORE || minsLeft < -SLOT_GRACE_MIN) continue;
    const to = officerEmails_().concat(getNotifyExtra_());
    const ok = sendMailSafe_(to, "NOT APPROVED YET: " + data[i][col["title"]] + " goes out at " + slot,
      postMailBody_("This post is still not approved", data[i][col["title"]], date, slot, data[i][col["author"]],
        "It is scheduled for " + slot + " and nobody has approved it yet. Please approve or reject it now."));
    if (ok) sheet.getRange(i + 1, col["reminder_sent"] + 1).setValue(new Date().toISOString());
  }
}

// Run ONCE from the Apps Script editor, allow the permissions, then redeploy
function setupReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "checkPendingReminders") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("checkPendingReminders").timeBased().everyMinutes(5).create();
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), "Reminder trigger OK", "The reminder trigger runs every 5 minutes.");
  console.log("Reminder trigger created (every 5 minutes).");
}

// ── Albums: admin adds more photos (up to 10) ───────────────────────────────
function handleAddAlbumPhotos(sheet, col, data, session) {
  if (!session || session.role === "user") return json({ success: false, error: "Unauthorized." });
  const r = toSheetRow(data.row);
  if (!validRow_(sheet, r)) return json({ success: false, error: "Post not found." });
  const add = (Array.isArray(data.urls) ? data.urls : []).map(String).filter(Boolean);
  if (!add.length) return json({ success: false, error: "No photos to add." });
  const upErr = checkUploads_(add, session);
  if (upErr) return json({ success: false, error: upErr });
  if (add.some(isVideoUrl_)) return json({ success: false, error: "Albums can only hold images." });
  const cover = String(sheet.getRange(r, col["media"] + 1).getValue() || "");
  if (isVideoUrl_(cover)) return json({ success: false, error: "This post is a video, so it can't become an album." });
  let album = parseAlbum_(sheet.getRange(r, col["album_media"] + 1).getValue());
  if (!album.length && cover) album = [cover];
  if (album.length + add.length > ALBUM_ADMIN_MAX)
    return json({ success: false, error: "An album can hold at most " + ALBUM_ADMIN_MAX + " photos (it has " + album.length + ")." });
  const ids = parseAlbum_(sheet.getRange(r, col["album_public_ids"] + 1).getValue());
  if (!ids.length && cover) ids.push(extractPublicId_(cover));
  album = album.concat(add);
  add.forEach(u => ids.push(extractPublicId_(u)));
  sheet.getRange(r, col["album_media"] + 1).setValue(JSON.stringify(album));
  sheet.getRange(r, col["album_public_ids"] + 1).setValue(JSON.stringify(ids));
  consumeUploads_(add);
  logActivity(session, "addAlbumPhotos", "row=" + data.row + " +" + add.length);
  return json({ success: true, count: album.length, album: album });
}

// ── Albums: admin changes the photo order before the post is shared ─────────
// The order saved here is the order Buffer receives. The first photo is the cover.
function handleReorderAlbum(sheet, col, data, session) {
  if (!session || session.role === "user") return json({ success: false, error: "Unauthorized." });
  const r = toSheetRow(data.row);
  if (!validRow_(sheet, r)) return json({ success: false, error: "Post not found." });
  const row = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (String(row[col["shared"]]).toUpperCase() === "TRUE")
    return json({ success: false, error: "This post was already shared, so the photo order can't change." });

  const album = parseAlbum_(row[col["album_media"]]);
  if (album.length < 2) return json({ success: false, error: "This post is not an album." });
  if (album.length > ALBUM_ADMIN_MAX) return json({ success: false, error: "An album can hold at most " + ALBUM_ADMIN_MAX + " photos." });

  // The new order must contain exactly the same photos, nothing added, removed or repeated
  const order = (Array.isArray(data.order) ? data.order : []).map(String);
  if (order.length !== album.length || order.slice().sort().join("\n") !== album.slice().sort().join("\n"))
    return json({ success: false, error: "The photos changed. Please reload the page and try again." });

  const ids = parseAlbum_(row[col["album_public_ids"]]);
  const idOf = {};
  album.forEach((u, i) => { idOf[u] = ids[i] || extractPublicId_(u); });

  sheet.getRange(r, col["album_media"] + 1).setValue(JSON.stringify(order));
  sheet.getRange(r, col["album_public_ids"] + 1).setValue(JSON.stringify(order.map(u => idOf[u] || extractPublicId_(u))));
  sheet.getRange(r, col["media"] + 1).setValue(order[0]);                                   // the cover
  sheet.getRange(r, col["cloudinary_public_id"] + 1).setValue(extractPublicId_(order[0]));
  logActivity(session, "reorderAlbum", "row=" + data.row);
  return json({ success: true, album: order });
}

// ── Editable profile (changes wait for admin approval) ─────────────────────
function handleGetMyProfile(session) {
  if (!session) return json({ success: false, error: "Please log in again." });
  const u = findUser(session.phone);
  if (!u) return json({ success: true, name: session.name || "", email: "", phone: session.phone, pending: null, readOnly: true });
  let pending = null;
  try { const raw = ensureUsersSheet().getRange(u.row, 10).getValue(); pending = raw ? JSON.parse(raw) : null; } catch (e) {}
  return json({ success: true, name: u.name, email: u.email, phone: u.phone, pending: pending });
}

function handleProfileEmailOtp(data, session) {
  if (!session) return json({ success: false, error: "Please log in again." });
  const email = String(data.email || "").trim().toLowerCase();
  if (!EMAIL_RE.test(email)) return json({ success: false, error: "Please enter a valid email address." });
  const other = findUserByEmail(email);
  if (other && other.phone !== session.phone) return json({ success: false, error: "That email is already used by another account." });
  const cache = CacheService.getScriptCache();
  if (cache.get("peotp_cd_" + session.phone)) return json({ success: false, error: "Please wait 60 seconds before requesting another code." });
  const code = String(Math.floor(100000 + Math.random() * 900000));
  cache.put("peotp_" + session.phone, JSON.stringify({ email: email, code: code, tries: 0 }), 600);
  cache.put("peotp_cd_" + session.phone, "1", 60);
  if (!sendMailSafe_([email], "Your verification code: " + code,
      "<div style='font-family:Arial,sans-serif'><h2>Time Slot Scheduler</h2><p>Code to confirm your new email:</p>" +
      "<p style='font-size:32px;font-weight:bold;letter-spacing:6px;color:#b34e1a'>" + code + "</p><p>It expires in 10 minutes.</p></div>"))
    return json({ success: false, error: "Could not send the email. Try again later." });
  return json({ success: true });
}

function handleRequestProfileChange(data, session) {
  if (!session) return json({ success: false, error: "Please log in again." });
  const u = findUser(session.phone);
  if (!u) return json({ success: false, error: "Profile editing is not available for this account." });
  const req = { requestedAt: new Date().toISOString() };
  const name = String(data.name || "").trim();
  if (name && name !== u.name) { if (name.length < 2 || name.length > 80) return json({ success: false, error: "Enter a valid name." }); req.name = name; }

  if (data.phone) {
    const ph = normalizePhone(data.phone);
    if (ph !== u.phone) {
      if (ph.length < 9 || ph.length > 15) return json({ success: false, error: "That phone number isn't valid." });
      if (findUser(ph)) return json({ success: false, error: "That phone number is already used by another account." });
      req.phone = ph;
    }
  }
  const email = String(data.email || "").trim().toLowerCase();
  if (email && email !== u.email) {
    if (!EMAIL_RE.test(email)) return json({ success: false, error: "That email isn't valid." });
    const other = findUserByEmail(email);
    if (other && other.phone !== u.phone) return json({ success: false, error: "That email is already used by another account." });
    const cache = CacheService.getScriptCache();
    let entry = null; try { entry = JSON.parse(cache.get("peotp_" + session.phone) || "null"); } catch (e) {}
    if (!entry || entry.email !== email) return json({ success: false, error: "Verify the new email with the code first." });
    if (entry.tries >= 5) return json({ success: false, error: "Too many wrong codes. Request a new one." });
    if (String(data.emailCode || "").trim() !== entry.code) {
      entry.tries++; cache.put("peotp_" + session.phone, JSON.stringify(entry), 600);
      return json({ success: false, error: "That code is wrong." });
    }
    cache.remove("peotp_" + session.phone);
    req.email = email;
  }
  if (!req.name && !req.phone && !req.email) return json({ success: false, error: "Nothing was changed." });
  ensureUsersSheet().getRange(u.row, 10).setValue(JSON.stringify(req));
  logActivity(session, "requestProfileChange", Object.keys(req).filter(k => k !== "requestedAt").join(","));
  return json({ success: true, pending: req });
}

function handleGetProfileRequests(session) {
  if (!session || session.role === "user") return json({ success: false, error: "Unauthorized." });
  const data = ensureUsersSheet().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][9]) continue;
    let req; try { req = JSON.parse(data[i][9]); } catch (e) { continue; }
    out.push({ phone: normalizePhone(String(data[i][0])), name: data[i][1] || "", email: String(data[i][4] || ""), request: req });
  }
  return json({ success: true, requests: out });
}

function handleDecideProfileRequest(data, session) {
  if (!session || session.role === "user") return json({ success: false, error: "Unauthorized." });
  const u = findUser(normalizePhone(data.phone));
  if (!u) return json({ success: false, error: "Account not found." });
  const sh = ensureUsersSheet();
  let req; try { req = JSON.parse(sh.getRange(u.row, 10).getValue() || "null"); } catch (e) { req = null; }
  if (!req) return json({ success: false, error: "There is no pending request." });
  if (data.approve) {
    if (req.phone && findUser(req.phone)) return json({ success: false, error: "That phone number was taken meanwhile. Reject this request." });
    if (req.email) { const o = findUserByEmail(req.email); if (o && o.phone !== u.phone) return json({ success: false, error: "That email was taken meanwhile. Reject this request." }); }
    if (req.name)  sh.getRange(u.row, 2).setValue(safeCell_(req.name));
    if (req.email) sh.getRange(u.row, 5).setValue(req.email);
    if (req.phone) {
      sh.getRange(u.row, 1).setValue(req.phone);
      // keep the person's posts and admin entry attached to the new number
      const db = ensureDatabaseSheet(); ensureColumns();
      const c = buildColMap(db.getRange(1, 1, 1, db.getLastColumn()).getValues()[0]);
      const vals = db.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) if (normalizePhone(String(vals[i][c["submitted_by"]])) === u.phone) db.getRange(i + 1, c["submitted_by"] + 1).setValue(req.phone);
      const entries = getAdminEntries_(); let ch = false;
      entries.forEach(en => { if (en.id === u.phone) { en.id = req.phone; ch = true; } });
      if (ch) saveAdminEntries_(entries);
    }
  }
  sh.getRange(u.row, 10).setValue("");
  logActivity(session, data.approve ? "approveProfileChange" : "rejectProfileChange", u.phone);
  return json({ success: true });
}

// ── In-app views (owner only): projects, users, posts, activity log ─────────
function handleGetDirectory(data, session) {
  if (!session || session.role !== "owner") return json({ success: false, error: "Unauthorized." });
  const kind = String(data.kind || ""), q = String(data.query || "").trim().toLowerCase();
  const hit = (obj) => !q || JSON.stringify(obj).toLowerCase().indexOf(q) >= 0;
  const usersRaw = ensureUsersSheet().getDataRange().getValues();
  const users = [];
  for (let i = 1; i < usersRaw.length; i++) users.push(userFromRow(usersRaw[i], i));

  if (kind === "users") {
    return json({ success: true, items: users.map(u => ({ name: u.name, phone: u.phone, email: u.email, role: u.role, clubRole: u.clubRole,
      projects: u.projects })).filter(hit).slice(0, 300) });
  }
  if (kind === "projects") {
    const ps = ensureProjectsSheet().getDataRange().getValues().slice(1).filter(r => r[0]);
    const items = ps.map(r => ({ name: String(r[0]), status: String(r[1] || "ongoing"), color: r[2] || "#b34e1a",
      members: users.filter(u => u.projects.some(p => p.project === r[0]))
                    .map(u => ({ name: u.name, phone: u.phone, role: (u.projects.find(p => p.project === r[0]) || {}).role || "" })) }))
      .filter(p => !q || p.name.toLowerCase().indexOf(q) >= 0 || p.members.some(m => (m.name + m.role).toLowerCase().indexOf(q) >= 0));
    return json({ success: true, items: items });
  }
  if (kind === "projectPosts") {
    const sheet = ensureDatabaseSheet(); ensureColumns();
    const range = sheet.getDataRange(), disp = range.getDisplayValues(), raw = range.getValues();
    const col = buildColMap(disp[0]), out = [];
    for (let i = 1; i < disp.length; i++) {
      if (disp[i][col["title"]] !== data.project) continue;
      out.push({ date: normDateServer(raw[i][col["date"]]), timeslot: normSlotServer(disp[i][col["timeslot"]]),
        status: disp[i][col["status"]] || "pending", shared: String(disp[i][col["shared"]]).toUpperCase() === "TRUE",
        author: disp[i][col["author"]] || "", content: String(disp[i][col["content"]] || "").substring(0, 160) });
    }
    out.sort((a, b) => (b.date + b.timeslot).localeCompare(a.date + a.timeslot));
    return json({ success: true, items: out.slice(0, 100) });
  }
  if (kind === "posts") {
    const sheet = ensureDatabaseSheet(); ensureColumns();
    const range = sheet.getDataRange(), disp = range.getDisplayValues(), raw = range.getValues();
    const col = buildColMap(disp[0]), out = [];
    for (let i = 1; i < disp.length; i++) {
      if (!disp[i][col["title"]]) continue;
      out.push({ project: disp[i][col["title"]], date: normDateServer(raw[i][col["date"]]), timeslot: normSlotServer(disp[i][col["timeslot"]]),
        status: disp[i][col["status"]] || "pending", shared: String(disp[i][col["shared"]]).toUpperCase() === "TRUE",
        author: disp[i][col["author"]] || "", platforms: disp[i][col["platforms"]] || "", content: String(disp[i][col["content"]] || "").substring(0, 160) });
    }
    out.sort((a, b) => (b.date + b.timeslot).localeCompare(a.date + a.timeslot));
    return json({ success: true, items: out.filter(hit).slice(0, 200) });
  }
  if (kind === "activity") {
    const rows = ensureActivityLogSheet().getDataRange().getValues().slice(1).reverse();
    const items = rows.map(r => ({ time: String(r[0]), phone: String(r[1]), name: String(r[2]), role: String(r[3]), action: String(r[4]), detail: String(r[5]) }))
                      .filter(hit).slice(0, 200);
    return json({ success: true, items: items });
  }
  return json({ success: false, error: "Unknown list." });
}

// Extra scan checks (called from handleScan_)
function scanExtras_(add) {
  const triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  add("System", "Email reminder trigger", triggers.indexOf("checkPendingReminders") >= 0 ? "ok" : "warn",
      triggers.indexOf("checkPendingReminders") >= 0 ? "Runs every 5 minutes."
        : "Not set up: reminder emails will not be sent. Run setupReminderTrigger once in Apps Script.");
  try {
    const left = MailApp.getRemainingDailyQuota();
    add("System", "Email quota today", left < 10 ? "error" : left < 30 ? "warn" : "ok", left + " recipients left today (Gmail limit is about 100 per day).");
  } catch (e) { add("System", "Email quota today", "warn", "Cannot read the quota. Run authorizeMail() once."); }
  add("System", "Email recipients", officerEmails_().length ? "ok" : "warn",
      officerEmails_().length ? officerEmails_().length + " officer email(s) found." : "No officer email found for the selected roles (Owner Tools → Email notifications).");
}