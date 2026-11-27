/**
 * User/Script.js
 * Submit form logic: date picker, slot picker, media upload, form submit.
 * Also contains _showMainApp() which must be defined before Auth.js runs.
 *
 * Load order in Submit.html:
 *   1. Config.js
 *   2. Script.js   ← defines _showMainApp before Auth runs
 *   3. Auth.js     ← calls initAuth() immediately
 */

// ---------------------------------------------------------------------------
// Session helpers (fallback if Auth.js not yet loaded)
// ---------------------------------------------------------------------------
if (typeof getSession === "undefined") {
  window.getSession   = function() { try { return JSON.parse(sessionStorage.getItem("auth")); } catch(e) { return null; } };
  window.clearSession = function() { sessionStorage.removeItem("auth"); };
}

// ---------------------------------------------------------------------------
// Loading overlay: spinner (loading) or a filling circle with % (uploading)
// ---------------------------------------------------------------------------
const _overlay = document.getElementById("loadingOverlay");
const RING_LEN = 2 * Math.PI * 52;

function _ringMode(on) {
  document.getElementById("loadingSpinner")?.classList.toggle("hidden", on);
  document.getElementById("progressRing")?.classList.toggle("hidden", !on);
}
function showLoading(msg) {
  _ringMode(false);
  const t = document.getElementById("loadingText");
  if (t) t.textContent = msg || "Please wait…";
  _overlay?.classList.add("active");
}
function showProgress(pct, msg) {
  _ringMode(true);
  const p = Math.max(0, Math.min(100, Math.round(pct || 0)));
  const fg = document.getElementById("ringFg");
  if (fg) fg.style.strokeDashoffset = String(RING_LEN * (1 - p / 100));
  const tx = document.getElementById("ringPct");
  if (tx) tx.textContent = p + "%";
  const t = document.getElementById("loadingText");
  if (t) t.textContent = msg || "Uploading…";
  _overlay?.classList.add("active");
}
function hideLoading() {
  _overlay?.classList.remove("active");
  setTimeout(() => _ringMode(false), 350);
}

// ---------------------------------------------------------------------------
// Backend call with the logged-in token. Shows the maintenance screen if the
// system is paused for a normal user.
// ---------------------------------------------------------------------------
async function apiCall(payload) {
  const session = getSession();
  const res = await fetch(CONFIG.SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(Object.assign({ token: session?.token || "" }, payload))
  });
  const data = await res.json();
  if (data && data.paused && session && session.role === "user") showMaintenance();
  return data;
}

// ---------------------------------------------------------------------------
// Roles: dropdown with an "Other" option that opens a text box
// ---------------------------------------------------------------------------
let _rolesCache = null;
async function loadRoleLists() {
  if (_rolesCache) return _rolesCache;
  try {
    const res = await fetch(`${CONFIG.SCRIPT_URL}?action=getRoles`);
    const d = await res.json();
    if (d && Array.isArray(d.projectRoles) && d.projectRoles.length) { _rolesCache = d; return d; }
  } catch (e) { /* fall back below */ }
  return { projectRoles: CONFIG.ROLES.PROJECT.slice(), clubRoles: CONFIG.ROLES.CLUB.slice() };
}

async function fillRoleSelect(selectEl, otherEl, selected) {
  if (!selectEl) return;
  const lists = await loadRoleLists();
  const roles = lists.projectRoles;
  selectEl.innerHTML = '<option value="">-- Role --</option>' +
    roles.map(r => `<option value="${_escHtml(r)}">${_escHtml(r)}</option>`).join("") +
    '<option value="__other__">Other…</option>';
  if (selected) {
    if (roles.includes(selected)) selectEl.value = selected;
    else { selectEl.value = "__other__"; if (otherEl) otherEl.value = selected; }
  }
  const sync = () => otherEl?.classList.toggle("hidden", selectEl.value !== "__other__");
  selectEl.onchange = sync;
  sync();
}

function readRole(selectEl, otherEl) {
  if (!selectEl) return "";
  if (selectEl.value === "__other__") return (otherEl?.value || "").trim();
  return selectEl.value.trim();
}

// ---------------------------------------------------------------------------
// Screens
// ---------------------------------------------------------------------------
const _SCREEN_IDS = ["loginScreen","signupEmailScreen","otpScreen","passwordScreen","registerScreen",
  "projectPickerScreen","mainSubmitForm","adminChoiceScreen","historyScreen","maintenanceScreen"];

function _hideAllScreens() {
  _SCREEN_IDS.forEach(id => { const el = document.getElementById(id); if (el) el.classList.add("hidden"); });
  const nav = document.getElementById("postLoginNav");
  if (nav) nav.style.display = "none";
  document.getElementById("profileMenu")?.classList.add("hidden");
}

function showMaintenance() {
  _hideAllScreens();
  document.getElementById("maintenanceScreen")?.classList.remove("hidden");
  loadContactLinks();
}

document.getElementById("maintLogoutBtn")?.addEventListener("click", () => { clearSession(); location.reload(); });
document.getElementById("maintRetryBtn")?.addEventListener("click", async () => {
  showLoading("Checking…");
  try {
    const r = await fetch(`${CONFIG.SCRIPT_URL}?action=getSystemPaused`);
    const d = await r.json();
    hideLoading();
    if (!d.paused) { location.reload(); return; }
    alert("The system is still under maintenance. Please try again later.");
  } catch (e) { hideLoading(); alert("Could not check. Please try again."); }
});

// ---------------------------------------------------------------------------
// Project picker (shared by the login flow and "Switch Project")
// BOD Post is shown first for people who hold a club role.
// ---------------------------------------------------------------------------
async function renderProjectPicker(container, session, onPick) {
  container.innerHTML = "";
  const status = {};
  try {
    const res = await fetch(`${CONFIG.SCRIPT_URL}?action=getProjects`);
    (await res.json()).forEach(p => { status[p.name] = p.status; });
  } catch (e) { /* show everything if the list can't be loaded */ }

  const rows = [];
  if (session?.clubRole) rows.push({ project: CONFIG.BOD_PROJECT, role: session.clubRole, bod: true });
  (session?.projects || []).forEach(p => {
    if (status[p.project] === "completed") return;       // completed projects are closed
    rows.push({ project: p.project, role: p.role });
  });

  if (!rows.length) {
    container.innerHTML = '<p class="slot-note" style="justify-content:center;">You have no active projects. Tap "Add more projects" below.</p>';
    return;
  }
  rows.forEach(r => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "project-picker-item" + (r.bod ? " pp-bod" : "");
    btn.innerHTML = `<i class="fas ${r.bod ? "fa-bullhorn" : "fa-folder"}"></i>` +
      `<span class="pp-name">${_escHtml(r.project)}</span><span class="pp-role">${_escHtml(r.role)}</span>`;
    btn.addEventListener("click", () => onPick(r.project, r.role));
    container.appendChild(btn);
  });
}

async function renderPickerScreen() {
  const session = getSession();
  const container = document.getElementById("projectPickerList");
  if (container) await renderProjectPicker(container, session, pickProject);
  const sel = document.getElementById("pickerAddProjectSelect");
  if (sel) _populateProjectSelect(sel, "");
  fillRoleSelect(document.getElementById("pickerAddRoleSelect"),
                 document.getElementById("pickerAddRoleInput"), "");

  // Add-project form: always minimized; the toggle opens it
  const wrap = document.querySelector(".picker-add-project-wrap");
  if (wrap) { wrap.classList.add("collapsed"); wrap.classList.remove("open"); }
  document.getElementById("pickerAddToggle")?.classList.remove("hidden");

  // Admin Panel shortcut: admins and the owner only
  const isStaff = session?.role === "admin" || session?.role === "owner";
  document.getElementById("pickerAdminBtn")?.classList.toggle("hidden", !isStaff);

  // Back button at the top: to the admin page, or to the submit form
  const back = document.getElementById("pickerTopBack");
  if (back) back.classList.toggle("hidden", !(sessionStorage.getItem("fromAdmin") || session?.activeProject));
}

document.getElementById("pickerAddToggle")?.addEventListener("click", () => {
  document.querySelector(".picker-add-project-wrap")?.classList.toggle("open");
});
document.getElementById("pickerTopBack")?.addEventListener("click", () => goBackFromScreen());
document.getElementById("pickerAdminBtn")?.addEventListener("click", () => { location.href = "../Admin/Admin.html"; });

function pickProject(project, role) {
  const session = getSession();
  if (!session) return;
  sessionStorage.removeItem("fromAdmin");
  session.activeProject = { project: project, role: role };
  sessionStorage.setItem("auth", JSON.stringify(session));
  if (typeof activeSession !== "undefined" && activeSession) activeSession.activeProject = session.activeProject;
  _showMainApp(session);
}

// ---------------------------------------------------------------------------
// _showMainApp — called after login / project selection
// ---------------------------------------------------------------------------
let _navBound = false;

function _bindNavOnce() {
  if (_navBound) return;
  _navBound = true;
  const menu = document.getElementById("profileMenu");
  document.getElementById("profileBtn")?.addEventListener("click", (e) => {
    e.stopPropagation();
    menu?.classList.toggle("hidden");
  });
  document.addEventListener("click", (e) => {
    if (menu && !menu.classList.contains("hidden") && !menu.contains(e.target)) menu.classList.add("hidden");
  });
  document.getElementById("navLogoutBtn")?.addEventListener("click", () => { clearSession(); location.reload(); });
  document.getElementById("navSwitchProjectBtn")?.addEventListener("click", () => {
    _hideAllScreens();
    document.getElementById("projectPickerScreen")?.classList.remove("hidden");
    renderPickerScreen();
  });
  document.getElementById("navHistoryBtn")?.addEventListener("click", () => { menu?.classList.add("hidden"); openHistory(); });
  document.getElementById("navPasswordBtn")?.addEventListener("click", () => { menu?.classList.add("hidden"); openPasswordModal(); });
  document.getElementById("navProfileBtn")?.addEventListener("click", () => { menu?.classList.add("hidden"); openProfileModal(); });
}

function _showMainApp(session) {
  _hideAllScreens();
  document.getElementById("mainSubmitForm").classList.remove("hidden");

  const nav = document.getElementById("postLoginNav");
  if (nav) nav.style.display = "flex";

  const nameEl = document.getElementById("navUserName");
  if (nameEl) nameEl.textContent = session.name || "User";
  const menuName = document.getElementById("profileMenuName");
  if (menuName) menuName.textContent = session.name || "User";

  const proj = session.activeProject;
  const badgeEl = document.getElementById("navProjectBadge");
  if (badgeEl && proj) badgeEl.textContent = proj.project + (proj.role ? " — " + proj.role : "");

  // Admin Panel button: admins and the owner only (unchanged)
  const adminBtn = document.getElementById("navAdminBtn");
  if (adminBtn) adminBtn.style.display = (session.role === "admin" || session.role === "owner") ? "inline-flex" : "none";

  // Pre-fill hidden author field
  const authorEl = document.getElementById("author");
  if (authorEl) authorEl.value = (session.name || "") + (proj && proj.role ? " — " + proj.role : "");

  // Lock project dropdown to the active project
  if (proj && proj.project) {
    const sel = document.getElementById("projectSelect");
    if (sel) {
      const lockProject = () => {
        sel.value = proj.project;
        if (!sel.value) {
          const opt = document.createElement("option");
          opt.value = proj.project; opt.textContent = proj.project;
          sel.appendChild(opt);
          sel.value = proj.project;
        }
        sel.disabled = true;
      };
      lockProject();
      setTimeout(lockProject, 1500);
    }
  }

  _bindNavOnce();
  initDatePicker();
  loadProjects();
  _restoreDraft();
  _showFlash();
  loadContactLinks();
}

async function _populateProjectSelect(sel, selectedVal) {
  try {
    const res      = await fetch(`${CONFIG.SCRIPT_URL}?action=getProjects`);
    const projects = await res.json();
    sel.innerHTML  = '<option value="">-- Project --</option>' +
      projects.filter(p => p.status !== "completed")
              .map(p => `<option value="${_escHtml(p.name)}">${_escHtml(p.name)}</option>`).join("");
    if (selectedVal) sel.value = selectedVal;
  } catch(e) { console.error("_populateProjectSelect error", e); }
}

function _escHtml(str) {
  return String(str).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;");
}

// ---------------------------------------------------------------------------
// Date picker
// ---------------------------------------------------------------------------

// Taken slots are fetched once and reused when the person taps another date (kept for 60 seconds)
let _takenCache = null;     // { at, slots }
async function getTakenSlotsCached(force) {
  if (!force && _takenCache && Date.now() - _takenCache.at < 60000) return _takenCache.slots;
  const r = await apiCall({ action: "getTakenSlots" });
  _takenCache = { at: Date.now(), slots: r.slots || [] };
  return _takenCache.slots;
}

function initDatePicker() {
  const container = document.getElementById("datePicker");
  if (!container) return;
  container.innerHTML = "";

  const today = new Date();
  const btns  = [];
  for (let i = 0; i < 14; i++) {
    const d = new Date(today);
    d.setDate(today.getDate() + i);
    const btn = document.createElement("button");
    btn.type         = "button";
    btn.className    = "date-btn" + (i === 0 ? " today" : "");
    btn.dataset.date = formatDateYMD(d);
    btn.innerHTML = `
      <span class="day-name">${d.toLocaleDateString(undefined,{weekday:"short"})}</span>
      <span class="day-num">${d.getDate()}</span>
      <span class="month-name">${d.toLocaleDateString(undefined,{month:"short"})}</span>
    `;
    btn.addEventListener("click", () => selectDate(btn, btn.dataset.date));
    container.appendChild(btn);
    btns.push(btn);
  }

  // Auto-select: find first day that has at least one available (future, not taken) slot
  autoSelectBestDate(btns);
}

async function autoSelectBestDate(btns) {
  let takenMap = {};
  try {
    const all = await getTakenSlotsCached(true);       // fresh when the page opens
    all.forEach(p => {
      const key = formatDateYMD(toLocalMidnight(p.date));
      if (!takenMap[key]) takenMap[key] = [];
      takenMap[key].push(p.timeslot);
    });
  } catch(e) { /* fall back to first date */ }

  const allSlots = Object.values(CONFIG.SLOT_GROUPS).flat();

  for (const btn of btns) {
    const dateStr = btn.dataset.date;
    const taken   = takenMap[dateStr] || [];
    const hasAvailable = allSlots.some(slotKey => {
      const label = CONFIG.SLOT_VALUES[slotKey] || slotKey;
      return !taken.includes(label) && !isSlotPastWithBuffer(dateStr, label);
    });
    if (hasAvailable) {
      selectDate(btn, dateStr, taken);
      // Auto-select the nearest available slot
      autoSelectNearestSlot(dateStr, taken);
      return;
    }
  }
  if (btns.length) selectDate(btns[0], btns[0].dataset.date, []);
}

function autoSelectNearestSlot(dateStr, takenSlots) {
  // Wait for slots to render then click the first available one
  setTimeout(() => {
    const slotBtns = document.querySelectorAll(".slot-btn:not(:disabled)");
    if (slotBtns.length > 0) {
      slotBtns[0].click();
    }
  }, 400);
}

function selectDate(btn, dateStr, preloadedTaken) {
  document.querySelectorAll(".date-btn").forEach(b => b.classList.remove("selected"));
  btn.classList.add("selected");
  document.getElementById("selectedDate").value = dateStr;
  loadSlots(dateStr, preloadedTaken);
}

// ---------------------------------------------------------------------------
// Slot picker
// ---------------------------------------------------------------------------

async function loadSlots(dateStr, preloadedTaken) {
  const container = document.getElementById("slotPicker");
  if (!container) return;
  container.innerHTML = '<div style="text-align:center;color:#888;grid-column:1/-1;">Loading slots…</div>';

  let takenSlots = preloadedTaken || [];
  if (!preloadedTaken) {
    try {
      const all = await getTakenSlotsCached(false);     // no new request when the list is recent
      const targetTime = toLocalMidnight(dateStr).getTime();
      takenSlots = all.filter(p => toLocalMidnight(p.date).getTime() === targetTime).map(p => p.timeslot);
    } catch(e) {
      console.error("loadSlots fetch error:", e);
    }
  }

  container.innerHTML = "";

  for (const [groupName, slots] of Object.entries(CONFIG.SLOT_GROUPS)) {
    const groupLabel = document.createElement("div");
    groupLabel.className = "slot-group-label";
    groupLabel.textContent = groupName;
    container.appendChild(groupLabel);

    for (const slotLabel of slots) {
      const slotValue = CONFIG.SLOT_VALUES[slotLabel];
      const isTaken   = takenSlots.includes(slotLabel) || takenSlots.includes(slotValue);
      const isPast    = isSlotPastWithBuffer(dateStr, slotLabel);

      const btn = document.createElement("button");
      btn.type      = "button";
      btn.className = "slot-btn";
      btn.textContent = CONFIG.UI.SLOT_NO_LEADING_ZERO
        ? slotLabel.replace(/^0/, "")
        : slotLabel;
      btn.disabled = isTaken || isPast;
      if (isTaken) btn.title = "Already booked";
      if (isPast)  btn.title = "This slot has closed";

      btn.addEventListener("click", () => {
        document.querySelectorAll(".slot-btn").forEach(b => b.classList.remove("selected"));
        btn.classList.add("selected");
        document.getElementById("selectedTimeslot").value = slotLabel;
      });
      container.appendChild(btn);
    }
  }
}

// ---------------------------------------------------------------------------
// Time helpers
// ---------------------------------------------------------------------------

function formatDateYMD(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function toLocalMidnight(dateStr) {
  const [y, m, d] = dateStr.split("-").map(Number);
  return new Date(y, m - 1, d);
}

/** Returns true once the slot's start time + 15 minutes has passed (slot is then closed) */
function isSlotPastWithBuffer(dateStr, slotLabel) {
  if (!dateStr || !slotLabel) return false;
  const match = slotLabel.match(/^(\d+):(\d+)\s*(AM|PM)$/i);
  if (!match) return false;
  let hour   = parseInt(match[1]);
  const mins  = parseInt(match[2]);
  const period = match[3].toUpperCase();
  if (period === "PM" && hour !== 12) hour += 12;
  if (period === "AM" && hour === 12) hour = 0;
  const slotTime = toLocalMidnight(dateStr);
  slotTime.setHours(hour, mins, 0, 0);
  const graceMs = (CONFIG.SLOT_GRACE_MINUTES || 15) * 60 * 1000;
  return (slotTime.getTime() + graceMs) < Date.now();
}

// ---------------------------------------------------------------------------
// Load projects into dropdown
// ---------------------------------------------------------------------------

async function loadProjects() {
  const sel = document.getElementById("projectSelect");
  if (!sel) return;
  const session = getSession();
  const proj    = session?.activeProject;
  try {
    const res      = await fetch(`${CONFIG.SCRIPT_URL}?action=getProjects`);
    const projects = await res.json();
    let names = projects.filter(p => p.status !== "completed" || (proj && p.name === proj.project)).map(p => p.name);
    // Normal users can only pick projects assigned to them (admins and the owner can use any, as on the server)
    const staff    = session?.role === "admin" || session?.role === "owner";
    const assigned = (session?.projects || []).map(p => p.project);
    if (!staff) names = names.filter(n => assigned.includes(n));
    if (proj && proj.project && !names.includes(proj.project)) names.unshift(proj.project);
    if (proj && proj.project === CONFIG.BOD_PROJECT && !names.includes(CONFIG.BOD_PROJECT)) names.unshift(CONFIG.BOD_PROJECT);
    // Nothing assigned: show the Add More Project route
    document.getElementById("addMoreProjectBtn")?.classList.toggle("hidden", names.length > 0);
    sel.innerHTML  = '<option value="">— Select a project —</option>' +
      names.map(n => `<option value="${_escHtml(n)}">${_escHtml(n)}</option>`).join("");
    if (proj && proj.project) {
      sel.value    = proj.project;
      sel.disabled = true;
    }
  } catch(e) { console.error("loadProjects error:", e); }
}

// ---------------------------------------------------------------------------
// Media upload
// ---------------------------------------------------------------------------

const mediaFile    = document.getElementById("mediaFile");
const mediaBox     = document.getElementById("mediaUploadBox");
const mediaUI      = document.getElementById("mediaUploadUI");
const mediaPreview = document.getElementById("mediaUploadPreview");
const mediaPreviewInner = document.getElementById("mediaPreviewInner");
const mediaFileName     = document.getElementById("mediaFileName");
const mediaClearBtn     = document.getElementById("mediaClearBtn");

let selectedFiles = [];      // one video OR up to ALBUM_MAX_IMAGES images
let mediaBusy     = false;   // true while the preview is loading
let _previewToken = 0;
const MAX_ALBUM   = (CONFIG.APP && CONFIG.APP.ALBUM_MAX_IMAGES) || 10;
const _submitBtn  = () => document.querySelector('#submitForm button[type="submit"]');

function _fmtSize(b) { return b >= 1048576 ? (b / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(b / 1024)) + " KB"; }
function _mediaMsg(html, kind) {
  const el = document.getElementById("mediaStatus");
  if (!el) return;
  el.className = "media-status " + kind;
  el.innerHTML = html;
}
function _setBusy(on) { mediaBusy = on; const b = _submitBtn(); if (b) b.disabled = on; }

mediaFile?.addEventListener("change", () => {
  const files = Array.from(mediaFile.files || []);
  mediaFile.value = "";                       // lets the same file be picked again
  handleMediaSelect(files);
});

mediaBox?.addEventListener("dragover", (e) => { e.preventDefault(); mediaBox.classList.add("drag-over"); });
mediaBox?.addEventListener("dragleave", () => mediaBox.classList.remove("drag-over"));
mediaBox?.addEventListener("drop", (e) => {
  e.preventDefault();
  mediaBox.classList.remove("drag-over");
  handleMediaSelect(Array.from(e.dataTransfer.files || []));
});
mediaClearBtn?.addEventListener("click", (e) => { e.stopPropagation(); clearMedia(); });

function handleMediaSelect(input) {
  const files = (Array.isArray(input) ? input : [input]).filter(Boolean);
  if (!files.length) return;
  const isVid = (f) => f.type.startsWith("video/");
  const isImg = (f) => f.type.startsWith("image/");
  if (files.some(f => !isVid(f) && !isImg(f))) { _mediaMsg("Please choose images or a video.", "err"); return; }
  if (files.some(isVid) && (files.length > 1 || selectedFiles.length > 0 && !selectedFiles.some(isVid) && false)) {
    _mediaMsg("An album can only contain images. Choose several images, or just one video.", "err"); return;
  }
  for (const f of files) {
    const maxMB = isVid(f) ? CONFIG.APP.MAX_VIDEO_SIZE_MB : CONFIG.APP.MAX_IMAGE_SIZE_MB;
    if (f.size > maxMB * 1024 * 1024) { _mediaMsg(`"${_escHtml(f.name)}" is too large. Max ${maxMB} MB for ${isVid(f) ? "video" : "image"}.`, "err"); return; }
  }
  let next;
  if (isVid(files[0])) next = [files[0]];                       // a video replaces everything
  else {
    const keep = selectedFiles.filter(isImg);
    next = keep.slice();
    files.forEach(f => { if (!next.some(x => x.name === f.name && x.size === f.size)) next.push(f); });
    if (next.length > MAX_ALBUM) { _mediaMsg(`Maximum ${MAX_ALBUM} images. You selected ${next.length}.`, "err"); return; }
  }
  selectedFiles = next;
  _syncJobs();            // uploads start now, in the background
  buildPreviews();
}

function _makeTile(f, url) {
  return new Promise((resolve, reject) => {
    const tile = document.createElement("div");
    tile.className = "album-tile";
    tile._file = f; tile._url = url;
    const timer = setTimeout(() => reject(new Error("timeout")), 25000);
    const done = () => { clearTimeout(timer); resolve(tile); };
    const fail = () => { clearTimeout(timer); reject(new Error("preview failed")); };
    if (f.type.startsWith("video/")) {
      const v = document.createElement("video");
      v.controls = true; v.muted = true; v.playsInline = true; v.preload = "metadata";
      v.onloadeddata = done; v.onerror = fail;
      v.src = url + "#t=0.1";
      tile.appendChild(v);
    } else {
      const im = document.createElement("img");
      im.decoding = "async"; im.alt = f.name;
      im.onload = done; im.onerror = fail;
      im.src = url;
      tile.appendChild(im);
    }
  });
}

function _revokeTiles() {
  mediaPreviewInner?.querySelectorAll(".album-tile").forEach(t => { try { URL.revokeObjectURL(t._url); } catch (e) {} });
}

function _updateSummary() {
  const total = selectedFiles.reduce((s, f) => s + f.size, 0);
  const vid = selectedFiles.some(f => f.type.startsWith("video/"));
  if (mediaFileName) mediaFileName.textContent = vid
    ? `${selectedFiles[0].name} · ${_fmtSize(total)}`
    : `${selectedFiles.length} image${selectedFiles.length === 1 ? "" : "s"} (max ${MAX_ALBUM}) · ${_fmtSize(total)}` +
      (selectedFiles.length > 1 ? " · use ◀ ▶ to set the order (the first photo is the Cover)" : "");
}

// Move a photo earlier or later. The saved order is the order the post, the admin and Buffer use.
function _moveFile(f, d) {
  if (mediaBusy) return;                                     // wait until every preview is loaded
  const i = selectedFiles.indexOf(f), k = i + d;
  if (i < 0 || k < 0 || k >= selectedFiles.length) return;
  [selectedFiles[i], selectedFiles[k]] = [selectedFiles[k], selectedFiles[i]];
  const grid = mediaPreviewInner.querySelector(".album-grid");
  const tiles = Array.from(grid.querySelectorAll(".album-tile"));
  selectedFiles.forEach(file => { const t = tiles.find(x => x._file === file); if (t) grid.appendChild(t); });
  _refreshOrderUI();
}

function _refreshOrderUI() {
  const n = selectedFiles.length;
  mediaPreviewInner?.querySelectorAll(".album-tile").forEach(t => {
    const i = selectedFiles.indexOf(t._file);
    t.querySelector(".album-cover")?.classList.toggle("hidden", !(n > 1 && i === 0));
    t.querySelector(".album-order")?.classList.toggle("hidden", n < 2);
    t.querySelector(".album-x")?.classList.toggle("hidden", n < 2);
    const num = t.querySelector(".album-num"); if (num) num.textContent = (i + 1) + "/" + n;
    t.querySelector(".order-left")?.toggleAttribute("disabled", i <= 0);
    t.querySelector(".order-right")?.toggleAttribute("disabled", i >= n - 1);
  });
}

async function buildPreviews() {
  const my = ++_previewToken;
  _revokeTiles();
  mediaPreviewInner.innerHTML = "";
  mediaUI.classList.add("hidden");
  mediaPreview.classList.remove("hidden");
  _setBusy(true);
  const list = selectedFiles.map(f => `${_escHtml(f.name)} (${_fmtSize(f.size)})`).join("<br>");
  _mediaMsg(`<span class="mini-spin"></span> <b>File selected.</b> Please wait, your preview is loading. Don't select it again.<div class="media-names">${list}</div>`, "wait");
  _updateSummary();

  const grid = document.createElement("div");
  grid.className = "album-grid" + (selectedFiles.length === 1 ? " single" : "");
  mediaPreviewInner.appendChild(grid);
  try {
    for (const f of selectedFiles.slice()) {
      const url = URL.createObjectURL(f);
      const tile = await _makeTile(f, url);
      if (my !== _previewToken) { URL.revokeObjectURL(url); return; }      // a newer selection took over
      if (selectedFiles.length > 1) {
        // Photo order: first photo = Cover, ◀ ▶ move a photo earlier or later
        const cover = document.createElement("span");
        cover.className = "album-cover"; cover.textContent = "Cover";
        tile.appendChild(cover);
        const ctl = document.createElement("div");
        ctl.className = "album-order";
        const lb = document.createElement("button");
        lb.type = "button"; lb.className = "order-left"; lb.innerHTML = "&#9664;"; lb.setAttribute("aria-label", "Move earlier");
        const num = document.createElement("span");
        num.className = "album-num";
        const rb = document.createElement("button");
        rb.type = "button"; rb.className = "order-right"; rb.innerHTML = "&#9654;"; rb.setAttribute("aria-label", "Move later");
        lb.addEventListener("click", (ev) => { ev.stopPropagation(); _moveFile(tile._file, -1); });
        rb.addEventListener("click", (ev) => { ev.stopPropagation(); _moveFile(tile._file, 1); });
        ctl.append(lb, num, rb);
        tile.appendChild(ctl);

        const x = document.createElement("button");
        x.type = "button"; x.className = "album-x"; x.innerHTML = "&times;"; x.setAttribute("aria-label", "Remove this image");
        x.addEventListener("click", (ev) => {
          ev.stopPropagation();
          selectedFiles = selectedFiles.filter(s => s !== tile._file);
          _syncJobs();      // a removed file's upload is cancelled and ignored
          try { URL.revokeObjectURL(tile._url); } catch (e) {}
          tile.remove();
          if (!selectedFiles.length) clearMedia(); else { _updateSummary(); _refreshOrderUI(); }
        });
        tile.appendChild(x);
      }
      grid.appendChild(tile);
      _paintTile(tile._file);
      _refreshOrderUI();
    }
    if (my !== _previewToken) return;
    _mediaMsg(`<i class="fas fa-check-circle"></i> ${selectedFiles.length > 1 ? selectedFiles.length + " images ready" : "Your file is ready"}.` +
              (selectedFiles.some(f => f.type.startsWith("image/")) && selectedFiles.length < MAX_ALBUM ? " You can add more images (up to " + MAX_ALBUM + ")." : ""), "ok");
    _setBusy(false);
    _updateUploadSummary();
  } catch (err) {
    if (my !== _previewToken) return;
    selectedFiles = [];
    _syncJobs();
    _revokeTiles();
    mediaPreviewInner.innerHTML = "";
    mediaPreview.classList.add("hidden");
    mediaUI.classList.remove("hidden");
    _mediaMsg('<i class="fas fa-exclamation-circle"></i> Please select the file again.', "err");
    _setBusy(false);
  }
}

function clearMedia() {
  ++_previewToken;
  _revokeTiles();
  selectedFiles = [];
  _syncJobs();
  if (mediaFile) mediaFile.value = "";
  if (mediaPreviewInner) mediaPreviewInner.innerHTML = "";
  if (mediaFileName) mediaFileName.textContent = "";
  mediaUI?.classList.remove("hidden");
  mediaPreview?.classList.add("hidden");
  const el = document.getElementById("mediaStatus"); if (el) { el.className = "media-status hidden"; el.innerHTML = ""; }
  _setBusy(false);
}

// ---------------------------------------------------------------------------
// Automatic background upload (Module 04)
// Each chosen file starts uploading at once. Upload.js retries a weak connection 3 times
// and sends big videos in pieces. A failed file gets a Retry button; the page never reloads.
// ---------------------------------------------------------------------------
const _jobs = new Map();     // File -> { status: "uploading" | "done" | "failed", pct, url, error, note, ctrl, promise }
let _submitId = "";          // one id per save attempt, so a retried save can never post twice

function _syncJobs() {
  for (const [f, job] of _jobs) {                      // removed or replaced files: cancel and ignore
    if (!selectedFiles.includes(f)) { job.ctrl.abort(); _jobs.delete(f); }
  }
  selectedFiles.forEach(f => { if (!_jobs.has(f)) _startJob(f); });
}

function _startJob(f) {
  const job = { status: "uploading", pct: 0, url: "", error: "", note: "", ctrl: new AbortController() };
  _jobs.set(f, job);
  const live = () => _jobs.get(f) === job;             // false once the file was removed, replaced or retried
  job.promise = uploadMedia(f, {
    signal: job.ctrl.signal,
    onProgress: (p) => { if (!live()) return; job.pct = p; job.note = ""; _paintTile(f); _updateUploadSummary(); },
    onRetry: (n) => { if (!live()) return; job.note = `Connection problem, trying again (${n}/3)…`; _paintTile(f); }
  }).then(url => {
    if (!live()) return;
    job.status = "done"; job.url = url; job.pct = 100;
    _paintTile(f); _updateUploadSummary();
  }).catch(err => {
    if (!live() || (err && err.aborted)) return;
    job.status = "failed"; job.error = (err && err.message) || "The upload failed.";
    _paintTile(f); _updateUploadSummary();
  });
  _paintTile(f);
}

function _retryJob(f) {
  const old = _jobs.get(f);
  if (old) old.ctrl.abort();
  _jobs.delete(f);
  _startJob(f);
  _updateUploadSummary();
}

function _paintTile(f) {
  const job = _jobs.get(f);
  const tile = Array.from(mediaPreviewInner?.querySelectorAll(".album-tile") || []).find(t => t._file === f);
  if (!job || !tile) return;
  let box = tile.querySelector(".tile-prog");
  if (!box) { box = document.createElement("div"); tile.appendChild(box); }
  box.className = "tile-prog " + job.status;
  box.textContent = "";
  if (job.status === "done") {
    box.innerHTML = '<i class="fas fa-check-circle"></i><span>Uploaded</span>';
  } else if (job.status === "failed") {
    const msg = document.createElement("span"); msg.textContent = job.error;
    const btn = document.createElement("button");
    btn.type = "button"; btn.className = "tile-retry"; btn.textContent = "Retry";
    btn.addEventListener("click", (e) => { e.stopPropagation(); _retryJob(f); });
    box.append(msg, btn);
  } else {
    const bar = document.createElement("div"); bar.className = "tile-bar";
    const fill = document.createElement("div"); fill.style.width = Math.round(job.pct) + "%";
    bar.appendChild(fill);
    const txt = document.createElement("span"); txt.textContent = job.note || (Math.round(job.pct) + "%");
    box.append(bar, txt);
  }
}

function _overallPct(files) {
  if (!files.length) return 0;
  return files.reduce((s, f) => s + ((_jobs.get(f) || {}).pct || 0), 0) / files.length;
}

function _updateUploadSummary() {
  if (mediaBusy) return;                               // the preview is still loading and shows its own message
  const jobs = selectedFiles.map(f => _jobs.get(f)).filter(Boolean);
  if (!jobs.length) return;
  const done = jobs.filter(j => j.status === "done").length;
  const failed = jobs.filter(j => j.status === "failed").length;
  if (failed) {
    _mediaMsg(`<i class="fas fa-exclamation-circle"></i> ${failed} file${failed > 1 ? "s" : ""} did not upload. Tap Retry on the file.`, "err");
  } else if (done === jobs.length) {
    _mediaMsg(`<i class="fas fa-check-circle"></i> ${jobs.length > 1 ? "All " + jobs.length + " files are" : "Your file is"} uploaded. Finish the other steps and press Submit.` +
              (selectedFiles.some(f => f.type.startsWith("image/")) && selectedFiles.length < MAX_ALBUM ? " You can add more images (up to " + MAX_ALBUM + ")." : ""), "ok");
  } else {
    _mediaMsg(`<span class="mini-spin"></span> Uploading ${done} of ${jobs.length}… You can keep filling in the form.`, "wait");
  }
}

// ---------------------------------------------------------------------------
// Draft / flash message (survive the automatic refresh after a failed upload)
// ---------------------------------------------------------------------------
const DRAFT_KEY = "draftPost";
const FLASH_KEY = "flashMsg";

function _saveDraft(content) {
  try { sessionStorage.setItem(DRAFT_KEY, JSON.stringify({ content: content || "" })); } catch (e) {}
}
function _restoreDraft() {
  try {
    const d = JSON.parse(sessionStorage.getItem(DRAFT_KEY) || "null");
    const el = document.getElementById("content");
    if (d && d.content && el && !el.value) el.value = d.content;
  } catch (e) {}
}
function _clearDraft() { sessionStorage.removeItem(DRAFT_KEY); }
function _showFlash() {
  const raw = sessionStorage.getItem(FLASH_KEY);
  if (!raw) return;
  sessionStorage.removeItem(FLASH_KEY);
  try { const f = JSON.parse(raw); showStatus(f.msg, f.type, 12000); } catch (e) {}
}

// ---------------------------------------------------------------------------
// Form submit
// ---------------------------------------------------------------------------

function _failAndRefresh(reason, content) {
  hideLoading();
  const nice = /failed to fetch|networkerror|load failed|network request/i.test(reason) ? "connection problem" : reason;
  _saveDraft(content);
  sessionStorage.setItem(FLASH_KEY, JSON.stringify({
    type: "error",
    msg: "Not uploaded — " + nice + ". The page was refreshed. Please upload your post again."
  }));
  showStatus("Not uploaded — " + nice + ". Refreshing… please upload again.", "error", 4000);
  setTimeout(() => location.reload(), 2200);
}

document.getElementById("submitForm")?.addEventListener("submit", async (e) => {
  e.preventDefault();

  const date     = document.getElementById("selectedDate").value;
  const slot     = document.getElementById("selectedTimeslot").value;
  const content  = document.getElementById("content").value.trim();
  const project  = document.getElementById("projectSelect").value;
  const author   = document.getElementById("author").value;

  if (mediaBusy) { showStatus("Your file is still loading. Please wait a moment.", "error"); return; }
  _clearErrors();
  const problems = [];   // in step order: Project, Media, Caption, Date, Time slot
  if (!project) problems.push(["err-project", "Please select a project."]);
  if (!selectedFiles.length) problems.push(["err-media", "Please attach media."]);
  else if (selectedFiles.length > MAX_ALBUM) problems.push(["err-media", `Maximum ${MAX_ALBUM} images.`]);
  if (!content) problems.push(["err-content", "Please write a caption."]);
  if (!date)    problems.push(["err-date", "Please select a date."]);
  if (!slot)    problems.push(["err-slot", "Please select a time slot."]);
  if (problems.length) {
    problems.forEach(([id, msg]) => _showErr(id, msg));
    document.getElementById(problems[0][0])?.closest(".form-step")?.scrollIntoView({ behavior: "smooth", block: "center" });
    showStatus("Please complete the highlighted steps.", "error");
    return;
  }

  const files = selectedFiles.slice();
  _syncJobs();                                          // makes sure every file has an upload running
  const pctNow = () => _overallPct(files);
  showProgress(pctNow(), "Finishing upload…");
  const tick = setInterval(() => showProgress(pctNow(), "Finishing upload…"), 300);
  try {
    await Promise.all(files.map(f => (_jobs.get(f) || {}).promise));
    clearInterval(tick);

    if (files.some(f => (_jobs.get(f) || {}).status !== "done")) {
      hideLoading();
      showStatus("A file did not upload. Tap Retry on it, then press Submit again. Your form is kept.", "error", 9000);
      document.getElementById("step2")?.scrollIntoView({ behavior: "smooth", block: "center" });
      return;
    }
    const urls = files.map(f => _jobs.get(f).url);

    showLoading("Saving your post…");
    if (!_submitId) _submitId = "s" + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
    const payload = { action: "submit", date, timeslot: slot, content, project, author, media: urls[0], clientId: _submitId };
    if (urls.length > 1) payload.album = urls;

    // If saving fails, try the SAVE again with the links we already have. The files are not uploaded twice.
    let result = null;
    for (let i = 0; i < 3 && !result; i++) {
      try { result = await apiCall(payload); }
      catch (e) {
        if (i < 2) { showLoading(`Connection problem. Trying to save again (${i + 2}/3)…`); await new Promise(r => setTimeout(r, 1500 * (i + 1))); }
      }
    }
    if (!result) {
      hideLoading();
      showStatus("Could not save your post (connection problem). Your files are already uploaded. Press Submit to try again.", "error", 10000);
      return;
    }

    if (result.success) {
      _clearDraft();
      _submitId = "";
      _showSuccess();
      document.getElementById("submitForm").reset();
      clearMedia();
      document.querySelectorAll(".date-btn").forEach(b => b.classList.remove("selected"));
      document.querySelectorAll(".slot-btn").forEach(b => b.classList.remove("selected"));
      document.getElementById("selectedDate").value    = "";
      document.getElementById("selectedTimeslot").value = "";
      document.getElementById("slotPicker").innerHTML  = '<div style="text-align:center;color:#888;">Select a date first</div>';
      const proj = getSession()?.activeProject;
      if (proj) {
        const sel = document.getElementById("projectSelect");
        if (sel) { sel.value = proj.project; sel.disabled = true; }
      }
      initDatePicker();     // reload slots so the one just used shows as taken
    } else if (result.paused) {
      /* maintenance screen is already shown */
    } else {
      _submitId = "";       // the server answered, so this attempt is finished
      _takenCache = null;   // the slot list may have changed, load it again on the next date tap
      _saveDraft(content);
      showStatus(result.error || "Could not save your post.", "error", 9000);
    }
  } catch (err) {
    showStatus((err && err.message) || "Something went wrong. Your form is kept. Please try again.", "error", 9000);
  } finally {
    clearInterval(tick);
    hideLoading();
  }
});

// ── Step validation + success screen (Module 03) ────────────────────────────
function _showErr(id, msg) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = msg;
  el.classList.add("show");
  el.closest(".form-step")?.classList.add("has-error");
}
function _clearErrors() {
  document.querySelectorAll(".field-error").forEach(e => { e.textContent = ""; e.classList.remove("show"); });
  document.querySelectorAll(".form-step.has-error").forEach(e => e.classList.remove("has-error"));
}
function _clearErr(id) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = ""; el.classList.remove("show");
  el.closest(".form-step")?.classList.remove("has-error");
}
document.getElementById("projectSelect")?.addEventListener("change", () => _clearErr("err-project"));
document.getElementById("mediaFile")?.addEventListener("change", () => _clearErr("err-media"));
document.getElementById("content")?.addEventListener("input", () => _clearErr("err-content"));
document.getElementById("datePicker")?.addEventListener("click", () => _clearErr("err-date"));
document.getElementById("slotPicker")?.addEventListener("click", () => _clearErr("err-slot"));
document.getElementById("addMoreProjectBtn")?.addEventListener("click", () => document.getElementById("navSwitchProjectBtn")?.click());

function _showSuccess() {
  document.getElementById("successScreen")?.classList.remove("hidden");
}
document.getElementById("successAnotherBtn")?.addEventListener("click", () => {
  document.getElementById("successScreen")?.classList.add("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
});
document.getElementById("successHistoryBtn")?.addEventListener("click", () => {
  document.getElementById("successScreen")?.classList.add("hidden");
  openHistory();
});

function showStatus(msg, type, ms) {
  const el = document.getElementById("statusMsg");
  if (!el) return;
  el.textContent = msg;
  el.className   = "message " + type;
  el.style.display = "block";
  clearTimeout(showStatus._t);
  showStatus._t = setTimeout(() => { el.style.display = "none"; }, ms || 5000);
  el.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

// ---------------------------------------------------------------------------
// History: the user's own posts (text only) + editing until the post is shared
// ---------------------------------------------------------------------------
function _fmtDate(ymd) {
  if (!ymd) return "";
  const [y, m, d] = String(ymd).split("-").map(Number);
  if (!y) return ymd;
  return new Date(y, m - 1, d).toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" });
}
function _fmtDateTime(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  if (isNaN(d)) return String(iso);
  return d.toLocaleString(undefined, { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" });
}
function _statusInfo(p) {
  if (p.status === "rejected") return { cls: "st-rejected", label: "Rejected" };
  if (p.shared)                return { cls: "st-shared",   label: "Published" };
  if (p.status === "approved") return { cls: "st-approved", label: "Approved" };
  return { cls: "st-pending", label: "Pending approval" };
}

async function openHistory() {
  _hideAllScreens();
  const nav = document.getElementById("postLoginNav");
  if (nav) nav.style.display = "flex";
  document.getElementById("historyScreen")?.classList.remove("hidden");
  await loadHistory();
}

document.getElementById("historyBackBtn")?.addEventListener("click", () => goBackFromScreen());

function goBackFromScreen() {
  if (sessionStorage.getItem("fromAdmin")) { sessionStorage.removeItem("fromAdmin"); location.href = "../Admin/Admin.html"; return; }
  const s = getSession();
  if (s && s.activeProject) _showMainApp(s);
  else { _hideAllScreens(); document.getElementById("projectPickerScreen")?.classList.remove("hidden"); renderPickerScreen(); }
}

async function loadHistory(notice) {
  const list = document.getElementById("historyList");
  if (!list) return;
  list.innerHTML = '<p class="history-note">Loading…</p>';
  let data;
  try { data = await apiCall({ action: "getMyPosts" }); }
  catch (e) { list.innerHTML = '<p class="u-msg err">Could not load your posts. Check your connection and try again.</p>'; return; }
  if (!data.success) { list.innerHTML = `<p class="u-msg err">${_escHtml(data.error || "Could not load your posts.")}</p>`; return; }

  list.innerHTML = "";
  if (notice) {
    const n = document.createElement("div");
    n.className = "message success"; n.style.display = "block"; n.textContent = notice;
    list.appendChild(n);
  }
  if (!data.posts.length) {
    list.insertAdjacentHTML("beforeend", '<p class="history-note">You have no posts yet.</p>');
    return;
  }
  data.posts.forEach(p => {
    const st = _statusInfo(p);
    const card = document.createElement("div");
    card.className = "hist-card clickable " + st.cls;
    const who = p.status === "rejected"
      ? (p.rejectedBy ? `<div class="hist-line"><i class="fas fa-user-times"></i> Rejected by ${_escHtml(p.rejectedBy)}</div>` : "")
      : (p.approvedBy ? `<div class="hist-line"><i class="fas fa-user-check"></i> Approved by ${_escHtml(p.approvedBy)}</div>` : "");
    const why = p.canEdit ? "Edit this post"
      : p.shared ? "Shared, so it can't be edited"
      : p.lockedBy ? "Being edited by " + p.lockedBy
      : p.status === "rejected" ? "Rejected. Please submit a new post"
      : "Older than 7 days, so it can't be edited";
    const preview = (p.content || "").length > 160 ? p.content.slice(0, 160) + "…" : (p.content || "");
    const count = (p.album && p.album.length > 1) ? ` · ${p.album.length} images` : "";
    card.innerHTML = `
      <div class="hist-top">
        <div class="hist-project">${_escHtml(p.project || "Post")}</div>
        <span class="st-badge ${st.cls}">${st.label}</span>
        <button type="button" class="hist-pen${p.canEdit ? "" : " off"}" title="${_escHtml(why)}" aria-label="${_escHtml(why)}"><i class="fas ${p.canEdit ? "fa-pen" : "fa-lock"}"></i></button>
      </div>
      <div class="hist-line"><i class="fas fa-calendar-day"></i> ${_escHtml(_fmtDate(p.date))} · ${_escHtml(p.timeslot)}${count}</div>
      <div class="hist-line"><i class="fas fa-clock"></i> Uploaded ${_escHtml(_fmtDateTime(p.submittedAt))}</div>
      ${who}
      <div class="hist-line"><i class="fas fa-share-alt"></i> Published: ${p.shared ? "Yes" : "No"}</div>
      <div class="hist-content">${_escHtml(preview)}</div>
      <div class="hist-tap">Tap the card to see the full post</div>`;
    card.querySelector(".hist-pen").addEventListener("click", (ev) => {
      ev.stopPropagation();
      if (p.canEdit) openEditPost(p); else alert(why + ".");
    });
    card.addEventListener("click", () => openHistDetail(p));
    list.appendChild(card);
  });
}

// ── Edit a post (locked so only one person edits at a time) ────────────────
let _editing = null;

async function openEditPost(p) {
  showLoading("Opening…");
  let r;
  try { r = await apiCall({ action: "lockPost", row: p.row }); }
  catch (e) { hideLoading(); alert("Connection problem. Please try again."); return; }
  hideLoading();
  if (!r.success) {
    if (!r.paused) alert(r.error || "You can't edit this post right now.");
    loadHistory();
    return;
  }
  _editing = { row: p.row };
  document.getElementById("editPostInfo").textContent =
    `${p.project} · ${_fmtDate(p.date)} · ${p.timeslot}`;
  document.getElementById("editPostContent").value = p.content || "";
  document.getElementById("editPostMedia").value = "";
  document.getElementById("editPostMsg").textContent = "";
  document.getElementById("editPostModal").classList.remove("hidden");
}

async function closeEditPost(unlock) {
  document.getElementById("editPostModal")?.classList.add("hidden");
  if (unlock && _editing) { try { await apiCall({ action: "unlockPost", row: _editing.row }); } catch (e) {} }
  _editing = null;
}

document.getElementById("editPostCancelBtn")?.addEventListener("click", async () => { await closeEditPost(true); loadHistory(); });

document.getElementById("editPostSaveBtn")?.addEventListener("click", async () => {
  if (!_editing) return;
  const msg = document.getElementById("editPostMsg");
  const content = document.getElementById("editPostContent").value.trim();
  const file = document.getElementById("editPostMedia").files[0];
  msg.className = "u-msg err";
  if (!content) { msg.textContent = "Post content can't be empty."; return; }
  if (file) {
    const isVideo = file.type.startsWith("video/");
    const maxMB = isVideo ? CONFIG.APP.MAX_VIDEO_SIZE_MB : CONFIG.APP.MAX_IMAGE_SIZE_MB;
    if (file.size > maxMB * 1024 * 1024) { msg.textContent = `File too large. Max ${maxMB} MB for ${isVideo ? "video" : "image"}.`; return; }
  }
  try {
    let mediaUrl = "";
    if (file) {
      showProgress(0, "Uploading…");
      mediaUrl = await uploadToCloudinary(file, (pct) => showProgress(pct, "Uploading…"));
    }
    showLoading("Saving…");
    const r = await apiCall({ action: "editMyPost", row: _editing.row, content, media: mediaUrl });
    hideLoading();
    if (r.success) {
      await closeEditPost(false);
      loadHistory("Post updated. It must be approved again.");
    } else if (!r.paused) {
      msg.textContent = r.error || "Could not save. Please try again.";
    }
  } catch (err) {
    hideLoading();
    msg.textContent = "Could not save (" + (err.message || "connection problem") + "). Please try again.";
  }
});

// release the edit lock if the page is closed while editing
window.addEventListener("pagehide", () => {
  if (_editing) {
    try {
      const s = getSession();
      navigator.sendBeacon(CONFIG.SCRIPT_URL, JSON.stringify({ action: "unlockPost", row: _editing.row, token: s?.token || "" }));
    } catch (e) {}
  }
});

// ---------------------------------------------------------------------------
// Change password
// ---------------------------------------------------------------------------
function openPasswordModal() {
  ["cpOld", "cpNew", "cpConfirm"].forEach(id => { const el = document.getElementById(id); if (el) el.value = ""; });
  const m = document.getElementById("cpMsg"); if (m) { m.textContent = ""; m.className = "u-msg"; }
  document.getElementById("passwordModal")?.classList.remove("hidden");
}
document.getElementById("cpCancelBtn")?.addEventListener("click", () => document.getElementById("passwordModal").classList.add("hidden"));
document.getElementById("cpSaveBtn")?.addEventListener("click", async () => {
  const oldPw = document.getElementById("cpOld").value;
  const newPw = document.getElementById("cpNew").value;
  const conf  = document.getElementById("cpConfirm").value;
  const msg   = document.getElementById("cpMsg");
  msg.className = "u-msg err";
  if (!oldPw)            { msg.textContent = "Enter your current password."; return; }
  if (newPw.length < 8)  { msg.textContent = "New password must be at least 8 characters."; return; }
  if (newPw !== conf)    { msg.textContent = "The new passwords don't match."; return; }
  showLoading("Saving…");
  try {
    const r = await apiCall({ action: "changePassword", oldPassword: oldPw, newPassword: newPw });
    hideLoading();
    if (r.success) {
      msg.className = "u-msg ok"; msg.textContent = "Password changed.";
      setTimeout(() => document.getElementById("passwordModal").classList.add("hidden"), 1200);
    } else if (!r.paused) {
      msg.textContent = r.error || "Could not change the password.";
    }
  } catch (e) { hideLoading(); msg.textContent = "Connection problem. Please try again."; }
});

// ---------------------------------------------------------------------------
// Back buttons at the top of every window
// ---------------------------------------------------------------------------
document.getElementById("editPostBack")?.addEventListener("click", async () => { await closeEditPost(true); loadHistory(); });
document.getElementById("cpBack")?.addEventListener("click", () => document.getElementById("passwordModal").classList.add("hidden"));
document.getElementById("detailBack")?.addEventListener("click", () => document.getElementById("histDetailModal").classList.add("hidden"));
document.getElementById("profBack")?.addEventListener("click", () => document.getElementById("profileModal").classList.add("hidden"));

// ---------------------------------------------------------------------------
// Contact link (the number comes from the server, only for logged-in people)
// ---------------------------------------------------------------------------
var _contactUrl = "";
async function loadContactLinks() {
  if (!_contactUrl) {
    try {
      const r = await apiCall({ action: "getContact" });
      if (r && r.whatsapp) _contactUrl = "https://wa.me/" + r.whatsapp;
    } catch (e) {}
  }
  document.querySelectorAll(".contact-link").forEach(a => {
    if (_contactUrl) { a.href = _contactUrl; a.target = "_blank"; a.rel = "noopener"; }
    else a.removeAttribute("href");
  });
}

// ---------------------------------------------------------------------------
// History details: full text + media (media exists for 7 days after the post date)
// ---------------------------------------------------------------------------
function openHistDetail(p) {
  const st = _statusInfo(p);
  const body = document.getElementById("detailBody");
  let media = "";
  const album = (p.album && p.album.length > 1) ? p.album : (p.media ? [p.media] : []);
  if (album.length) {
    media = '<div class="album-grid' + (album.length === 1 ? " single" : "") + '">' + album.map(u => {
      const isV = /\.(mp4|mov|webm|avi|m4v|3gp|mpeg|ogg)(\?|$)/i.test(u) || u.indexOf("/video/upload/") >= 0;
      return '<div class="album-tile">' + (isV ? `<video controls preload="metadata" src="${_escHtml(u)}"></video>` : `<img loading="lazy" src="${_escHtml(u)}" alt="">`) + "</div>";
    }).join("") + "</div>";
  } else {
    media = '<p class="history-note"><i class="fas fa-image"></i> The media was removed 7 days after the post date. Only the text is kept.</p>';
  }
  body.innerHTML = `
    <div class="hist-top"><div class="hist-project">${_escHtml(p.project || "Post")}</div><span class="st-badge ${st.cls}">${st.label}</span></div>
    <div class="hist-line"><i class="fas fa-calendar-day"></i> ${_escHtml(_fmtDate(p.date))} · ${_escHtml(p.timeslot)}</div>
    <div class="hist-line"><i class="fas fa-clock"></i> Uploaded ${_escHtml(_fmtDateTime(p.submittedAt))}</div>
    <div class="hist-line"><i class="fas fa-share-alt"></i> Published: ${p.shared ? "Yes" : "No"}</div>
    ${media}
    <div class="detail-text">${_escHtml(p.content || "")}</div>`;
  const act = document.getElementById("detailEditBtn");
  act.classList.toggle("hidden", !p.canEdit);
  act.onclick = () => { document.getElementById("histDetailModal").classList.add("hidden"); openEditPost(p); };
  document.getElementById("histDetailModal").classList.remove("hidden");
}

// ---------------------------------------------------------------------------
// Edit profile (name, email, phone): waits for admin approval
// ---------------------------------------------------------------------------
let _profOrig = { name: "", email: "", phone: "" };

async function openProfileModal() {
  const msg = document.getElementById("profMsg");
  msg.textContent = ""; msg.className = "u-msg";
  document.getElementById("profCodeRow").classList.add("hidden");
  document.getElementById("profCode").value = "";
  document.getElementById("profPending").classList.add("hidden");
  showLoading("Loading…");
  let r;
  try { r = await apiCall({ action: "getMyProfile" }); }
  catch (e) { hideLoading(); alert("Connection problem. Please try again."); return; }
  hideLoading();
  if (!r.success) { if (!r.paused) alert(r.error || "Could not load your profile."); return; }
  _profOrig = { name: r.name || "", email: r.email || "", phone: r.phone || "" };
  document.getElementById("profName").value  = _profOrig.name;
  document.getElementById("profEmail").value = _profOrig.email;
  document.getElementById("profPhone").value = _profOrig.phone;
  const ro = !!r.readOnly;
  ["profName", "profEmail", "profPhone"].forEach(id => document.getElementById(id).disabled = ro);
  document.getElementById("profSaveBtn").disabled = ro;
  if (r.pending) {
    const parts = [];
    if (r.pending.name)  parts.push("name → " + r.pending.name);
    if (r.pending.email) parts.push("email → " + r.pending.email);
    if (r.pending.phone) parts.push("phone → " + r.pending.phone);
    const box = document.getElementById("profPending");
    box.textContent = "Waiting for admin approval: " + parts.join(", ") + ". Your old details stay until it is approved.";
    box.classList.remove("hidden");
  }
  document.getElementById("profileModal").classList.remove("hidden");
}

document.getElementById("profEmail")?.addEventListener("input", () => {
  const changed = document.getElementById("profEmail").value.trim().toLowerCase() !== _profOrig.email.toLowerCase();
  document.getElementById("profCodeRow").classList.toggle("hidden", !changed);
});

document.getElementById("profSendCodeBtn")?.addEventListener("click", async () => {
  const msg = document.getElementById("profMsg");
  const email = document.getElementById("profEmail").value.trim();
  msg.className = "u-msg err";
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { msg.textContent = "Enter a valid email address."; return; }
  showLoading("Sending code…");
  try {
    const r = await apiCall({ action: "profileEmailOtp", email });
    hideLoading();
    if (r.success) { msg.className = "u-msg ok"; msg.textContent = "Code sent to " + email + "."; }
    else if (!r.paused) msg.textContent = r.error || "Could not send the code.";
  } catch (e) { hideLoading(); msg.textContent = "Connection problem. Please try again."; }
});

document.getElementById("profSaveBtn")?.addEventListener("click", async () => {
  const msg = document.getElementById("profMsg");
  msg.className = "u-msg err";
  const name = document.getElementById("profName").value.trim();
  const email = document.getElementById("profEmail").value.trim();
  const phone = document.getElementById("profPhone").value.trim();
  const emailChanged = email.toLowerCase() !== _profOrig.email.toLowerCase();
  const code = document.getElementById("profCode").value.trim();
  if (!name) { msg.textContent = "Enter your name."; return; }
  if (emailChanged && !/^\d{6}$/.test(code)) { msg.textContent = "Press “Send code”, then enter the 6-digit code from your new email."; return; }
  showLoading("Sending for approval…");
  try {
    const r = await apiCall({ action: "requestProfileChange", name, email, phone, emailCode: code });
    hideLoading();
    if (r.success) {
      msg.className = "u-msg ok";
      msg.textContent = "Sent. An admin must approve it. Your old details stay until then.";
      setTimeout(() => document.getElementById("profileModal").classList.add("hidden"), 2200);
    } else if (!r.paused) msg.textContent = r.error || "Could not send the request.";
  } catch (e) { hideLoading(); msg.textContent = "Connection problem. Please try again."; }
});