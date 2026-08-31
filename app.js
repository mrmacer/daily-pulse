/* ── Daily Pulse — standalone teacher + administrator app ────────────────
 * Orchestrates auth → authorization → roster → today's-completion →
 * status screen → sequential entry screen → submit. UI/scale/submit-payload
 * logic is extracted from MAC-Walkthrough app.js (renderPulse and friends,
 * ~app.js:2456-3145) and restructured into a strictly sequential one-
 * student-at-a-time flow per the product brief. See
 * docs/daily-pulse-extraction-report.md for what changed and why.
 */

function escHtml(str) {
  return String(str ?? "").replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[c]));
}

// Safe, non-sensitive support/debugging diagnostic block, shown alongside
// a login/authorization denial. Uses only AUTH state already held in
// memory — makes no additional Graph calls. Never shows anything beyond
// the signed-in user's own name/username (already visible to them in
// their own Microsoft account) plus a found/not-found summary: no tokens,
// no Graph payloads, no SharePoint IDs, no student or roster data.
// identityFound: "Not Found" | "Found". appPermission (optional):
// e.g. "Denied" — only passed once an app-permission check has actually
// run (identity already resolved), per the "IEP_Users2 identity failure
// is the priority case" instruction.
function loginDiagnosticHtml({ identityFound, appPermission } = {}) {
  const signedInAs = escHtml(AUTH.account?.name || AUTH.displayName || "(name unavailable)");
  const accountUsername = escHtml(AUTH.account?.username || "(unavailable)");
  const permissionLine = appPermission
    ? `<p class="diagnostic-line">App permission: <strong>${escHtml(appPermission)}</strong></p>`
    : "";
  return `<div class="login-diagnostic">
    <p class="diagnostic-line">Signed in as: <strong>${signedInAs}</strong></p>
    <p class="diagnostic-line">Account: <strong>${accountUsername}</strong></p>
    <p class="diagnostic-line">Identity record: <strong>${escHtml(identityFound || "")}</strong></p>
    ${permissionLine}
  </div>`;
}

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

function formatDateLong(dateStr) {
  const d = new Date(`${dateStr}T12:00:00`);
  return d.toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric" });
}

function greetingWord() {
  const h = new Date().getHours();
  if (h < 12) return "Good morning";
  if (h < 17) return "Good afternoon";
  return "Good evening";
}

function showToast(msg, kind = "success") {
  const el = document.getElementById("toast");
  if (!el) return;
  el.textContent = msg;
  el.className = `toast toast-${kind} visible`;
  clearTimeout(showToast._t);
  showToast._t = setTimeout(() => el.classList.remove("visible"), 3200);
}

const APP = {
  demoMode:        false,
  screen:          "loading", // loading | denied | error | teacher-select | status | entry | complete
  currentUser:     null,      // actual signed-in actor { id, name, email, role, isAdmin }
  effectiveTeacher:null,      // teacher whose roster/record Teacher value is in use
  teacherDirectory: [],       // live identities derived from eligible roster rows
  pendingTeacherName: "",
  roster:          [],        // eligible students for effectiveTeacher
  completedToday:  new Set(),
  todayEntries:    [],        // for the read-only "Today's Submissions" list
  draft:           null,
  _inFlight:       false,

  async init() {
    this.demoMode = DEMO.isActive();
    this._bindStaticControls();

    if (this.demoMode) {
      document.getElementById("demoBanner")?.classList.remove("hidden");
      this.currentUser = DEMO.currentUser();
      this.effectiveTeacher = DAILY_PULSE_ACCESS.teacherIdentity(this.currentUser.teacherName || this.currentUser.name);
      await this._loadRosterAndStatus();
      return;
    }

    this._showScreen("loading");
    try {
      await this._loadMsal();
      await AUTH.init();
    } catch (err) {
      console.error("AUTH.init failed:", err);
      this._showLoginError("Authentication service failed to load. Check your connection and try again.");
      return;
    }

    if (!AUTH.account) {
      this._showScreen("login");
      return;
    }

    if (AUTH.isUnauthorized) {
      this._showLoginError(
        AUTH.lookupError ||
        "You signed in successfully, but your account has not been added to IEP Skook yet. Please contact an administrator."
      );
      return;
    }

    // PATCH B: authorization now runs through IEP_App_Users (email-keyed),
    // not the legacy IEP_Users2 Teacher/Administrator role check alone. See
    // access.js#resolveDailyPulseAccess and iep-app-users.js for the
    // migration-mode contract (no row -> legacy fallback; lookup error ->
    // always denied).
    const appUsersOutcome = await this._resolveAppUsersOutcome();
    const decision = DAILY_PULSE_ACCESS.resolveDailyPulseAccess(AUTH.role, appUsersOutcome);
    if (!decision.allowed) {
      this.screen = "denied";
      this._render();
      return;
    }

    this.currentUser = this._resolveAuthenticatedIdentity();
    if (AUTH.isAdmin) {
      await this._prepareAdminTeacherSelection();
      return;
    }

    const teacherName = DAILY_PULSE_ACCESS.resolveTeacherRosterKey(this.currentUser, appUsersOutcome);
    this.effectiveTeacher = DAILY_PULSE_ACCESS.teacherIdentity(teacherName);
    await this._loadRosterAndStatus();
  },

  // Looks up IEP_App_Users by the signed-in email and normalizes the result
  // into the { found | failed } shape resolveDailyPulseAccess() expects —
  // see APP_USERS in iep-app-users.js for the reader itself.
  async _resolveAppUsersOutcome() {
    const row = await APP_USERS.resolve(AUTH.account?.username || "");
    if (APP_USERS.lookupError) return { failed: true };
    return row ? { found: true, row } : { found: false };
  },

  // Loaded lazily, only on the real-auth path — never in Demo Mode, so
  // Demo Mode fetches zero bytes from any Microsoft-owned domain (not just
  // zero Graph/SharePoint calls). See docs/demo-verification.md.
  _loadMsal() {
    if (window.msal) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2/lib/msal-browser.min.js";
      s.onload = resolve;
      s.onerror = () => reject(new Error("Could not load the Microsoft sign-in library."));
      document.head.appendChild(s);
    });
  },

  _bindStaticControls() {
    document.getElementById("loginBtn")?.addEventListener("click", () => AUTH.login());
    document.getElementById("demoModeBtn")?.addEventListener("click", () => {
      DEMO.activate();
      location.reload();
    });
    document.getElementById("exitDemoBtn")?.addEventListener("click", () => {
      DEMO.deactivate();
      location.reload();
    });
    document.getElementById("signOutBtn")?.addEventListener("click", () => {
      if (this.demoMode) { DEMO.deactivate(); location.reload(); }
      else AUTH.logout();
    });
  },

  // The actor is always the authenticated IEP_Users2 identity. For a
  // teacher, it also supplies effectiveTeacher automatically. For an admin,
  // it remains the Submitted By identity while roster ownership comes from
  // an explicit teacher selection.
  _resolveAuthenticatedIdentity() {
    return DAILY_PULSE_ACCESS.resolveActor(AUTH);
  },

  async _prepareAdminTeacherSelection() {
    this._showScreen("loading");
    try {
      await STUDENT_ROSTER.refresh();
      if (STUDENT_ROSTER.error) throw new Error(STUDENT_ROSTER.error);
      this.teacherDirectory = STUDENT_ROSTER.getDailyPulseTeachers();
      this._resetTeacherContext();
      this.screen = "teacher-select";
      this._render();
    } catch (err) {
      console.error("Failed to load Daily Pulse teacher directory:", err.message || String(err));
      this._showError("Could not load the Daily Pulse teacher directory. Please refresh or contact an administrator.");
    }
  },

  _resetTeacherContext() {
    this.effectiveTeacher = null;
    this.pendingTeacherName = "";
    this.roster = [];
    this.completedToday = new Set();
    this.todayEntries = [];
    this.draft = null;
  },

  async selectAdminTeacher(teacherName) {
    if (!this.currentUser?.isAdmin) return false;
    const target = String(teacherName || "").trim().toLowerCase();
    const teacher = this.teacherDirectory.find(item => item.teacherName.toLowerCase() === target);
    if (!teacher) return false;

    // Switching contexts always discards prior teacher-specific completion
    // and draft state before the next roster is loaded.
    this._resetTeacherContext();
    this.effectiveTeacher = { ...teacher };
    await this._loadRosterAndStatus();
    return true;
  },

  chooseAnotherTeacher() {
    if (!this.currentUser?.isAdmin) return;
    if (DAILY_PULSE_ACCESS.hasMeaningfulDraft(this.draft) &&
        !window.confirm("This student has unsaved Daily Pulse information. Choose another teacher and discard it?")) {
      return;
    }
    this._resetTeacherContext();
    this.screen = "teacher-select";
    this._render();
  },

  async _loadRosterAndStatus() {
    this._showScreen("loading");
    try {
      if (this.demoMode) {
        this.roster = DEMO.getRoster().filter(s => s.active && s.dailyPulseEnabled);
      } else {
        if (!STUDENT_ROSTER.loaded) await STUDENT_ROSTER.refresh();
        if (STUDENT_ROSTER.error) throw new Error(STUDENT_ROSTER.error);
        this.roster = STUDENT_ROSTER.getDailyPulseForTeacher(this.effectiveTeacher?.teacherName);
      }

      const date = todayStr();
      if (this.demoMode) {
        this.completedToday = DEMO.getTodayCompleted(date);
        this.todayEntries   = DEMO.getTodayEntries(date);
      } else {
        const entries = await GRAPH.getTodayDailyPulseEntries(this.effectiveTeacher?.teacherName, date);
        this.completedToday = new Set(entries.map(e => e["Student"]));
        this.todayEntries   = entries.map(e => ({
          student: e["Student"], status: e["Status"], attendanceStatus: e["Attendance Status"]
        }));
      }
    } catch (err) {
      console.error("Failed to load Daily Pulse status:", err.message || String(err));
      this._showError("Could not load your student roster. Please refresh or contact an administrator.");
      return;
    }

    this._goToStatusOrComplete();
  },

  _remainingRoster() {
    return this.roster.filter(s => !this.completedToday.has(s.name));
  },

  _goToStatusOrComplete() {
    this.draft = null;
    this.screen = this._remainingRoster().length === 0 && this.roster.length > 0 ? "complete" : "status";
    this._render();
  },

  /* ── Entry flow ─────────────────────────────────────────────────────── */

  _initDraft(student) {
    this.draft = {
      studentName:        student.name,
      attendanceStatus:   null,
      overallStatus:      "",
      selectedCategories: [],
      categoryResponses:  {},
      quickNote:          "",
      date:               todayStr(),
      recordId:           `pulse-${crypto.randomUUID()}`
    };
  },

  startOrContinue() {
    const remaining = this._remainingRoster();
    if (remaining.length === 0) { this._goToStatusOrComplete(); return; }
    this._initDraft(remaining[0]);
    this.screen = "entry";
    this._render();
  },

  jumpToStudent(name) {
    const student = this.roster.find(s => s.name === name);
    if (!student) return;
    this._initDraft(student);
    this.screen = "entry";
    this._render();
  },

  backToStatus() {
    this.draft = null;
    this.screen = "status";
    this._render();
  },

  /* ── Submit ─────────────────────────────────────────────────────────── */

  async submitAbsent() {
    if (this._inFlight) { console.warn("Duplicate Daily Pulse submission blocked."); return; }
    this._inFlight = true;
    const btn = document.getElementById("confirmAbsentBtn");
    if (btn) { btn.disabled = true; btn.textContent = "Saving…"; }

    const draft = this.draft;
    try {
      const payload = this._buildPayload(draft, "absent");
      await this._save(payload, draft);
      showToast(`Marked ${draft.studentName} absent.`);
      this._afterSuccessfulSave(draft);
    } catch (err) {
      console.error("Daily Pulse save failed:", err.message || String(err));
      this._showEntryError(err.message || "Save failed. Please try again.");
    } finally {
      this._inFlight = false;
      if (btn) { btn.disabled = false; btn.textContent = "Confirm Absent"; }
    }
  },

  async submitPresent() {
    if (this._inFlight) { console.warn("Duplicate Daily Pulse submission blocked."); return; }
    const draft = this.draft;

    if (!draft.overallStatus) { this._showEntryError("Please select how today went."); return; }
    for (const catKey of draft.selectedCategories) {
      const resp = draft.categoryResponses[catKey];
      if (!resp?.rating) {
        this._showEntryError(`Please select a response for ${DAILY_PULSE_SCALES[catKey]?.label || catKey}.`);
        document.querySelector(`[data-scale-panel="${catKey}"]`)?.scrollIntoView({ behavior: "smooth", block: "start" });
        return;
      }
    }

    this._inFlight = true;
    const btn = document.getElementById("submitBtn");
    if (btn) { btn.disabled = true; btn.textContent = "Saving…"; }

    try {
      const payload = this._buildPayload(draft, "present");
      await this._save(payload, draft);
      showToast("Daily Pulse saved!");
      this._afterSuccessfulSave(draft);
    } catch (err) {
      console.error("Daily Pulse save failed:", err.message || String(err));
      this._showEntryError(err.message || "Save failed. Please try again.");
    } finally {
      this._inFlight = false;
      if (btn) { btn.disabled = false; btn.textContent = "Submit & Next"; }
    }
  },

  _buildPayload(draft, attendanceStatus) {
    return DAILY_PULSE_ACCESS.buildPayload(
      draft,
      attendanceStatus,
      this.effectiveTeacher,
      this.currentUser
    );
  },

  // Local copy first (resilience layer), then the authoritative SharePoint
  // write. On failure the local copy stays, the draft is left untouched
  // (same recordId → an idempotent retry), and nothing advances — matching
  // source's failure-preservation behavior (docs/daily-pulse-source-
  // audit.md §20).
  // The DB (local-history) resilience layer is production-only — Demo Mode
  // has its own fully-isolated persistence (DEMO.savePulse, a separate
  // localStorage key). Writing demo entries through DB too would mix fake
  // students into the same key a real teacher's resilience layer uses.
  async _save(payload, draft) {
    if (this.demoMode) {
      await DEMO.savePulse(payload);
      return;
    }

    const localEntry = DB.addDailyPulse({
      id:                draft.recordId,
      student:           draft.studentName,
      date:              draft.date,
      attendanceStatus:  payload.AttendanceStatus,
      pulseStatus:       payload.Status,
      categories:        payload.Categories,
      categoryResponses: payload.CategoryResponses,
      note:              payload.QuickNote,
      teacherName:       this.effectiveTeacher?.teacherName || "",
      submittedById:     this.currentUser.id,
      submittedByName:   this.currentUser.name
    });

    const result = await GRAPH.saveDailyPulse(payload);
    if (!result.duplicatePrevented) DB.markSynced(localEntry.id);
  },

  _afterSuccessfulSave(draft) {
    this.completedToday.add(draft.studentName);
    const statusLabel = draft.attendanceStatus === "absent"
      ? ""
      : ({ great: "Great Day", struggles: "Some Struggles", concern: "Significant Concern" }[draft.overallStatus] || draft.overallStatus);
    this.todayEntries.push({
      student:          draft.studentName,
      status:           statusLabel,
      attendanceStatus: draft.attendanceStatus === "absent" ? "Absent" : "Present"
    });
    const remaining = this._remainingRoster();
    if (remaining.length === 0) {
      this._goToStatusOrComplete();
    } else {
      this._initDraft(remaining[0]);
      this.screen = "entry";
      this._render();
      document.getElementById("main")?.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  },

  /* ── Rendering ──────────────────────────────────────────────────────── */

  _showScreen(name) { this.screen = name; this._render(); },
  _showError(msg) { this.screen = "error"; this._errorMsg = msg; this._render(); },
  _showLoginError(msg) {
    document.getElementById("loginScreen")?.classList.remove("hidden");
    const el = document.getElementById("loginError");
    if (el) { el.textContent = msg; el.classList.remove("hidden"); }
    this._renderLoginDiagnostic();
  },

  // Only ever shown once Microsoft sign-in itself succeeded (AUTH.account
  // set) — hidden for the earlier "authentication service failed to load"
  // case, where there's no account yet to report. See loginDiagnosticHtml().
  _renderLoginDiagnostic() {
    const el = document.getElementById("loginDiagnostic");
    if (!el) return;
    if (!AUTH.account) {
      el.classList.add("hidden");
      el.innerHTML = "";
      return;
    }
    el.innerHTML = loginDiagnosticHtml({ identityFound: AUTH.isUnauthorized ? "Not Found" : "Found" });
    el.classList.remove("hidden");
  },
  _showEntryError(msg) {
    const el = document.getElementById("entryError");
    if (el) {
      el.textContent = msg;
      el.classList.remove("hidden");
      el.scrollIntoView({ behavior: "smooth", block: "nearest" });
    } else {
      showToast(msg, "error");
    }
  },

  _render() {
    const loginScreen = document.getElementById("loginScreen");
    if (this.screen === "login") {
      loginScreen?.classList.remove("hidden");
      const main0 = document.getElementById("main");
      if (main0) main0.innerHTML = "";
      return;
    }
    loginScreen?.classList.add("hidden");
    const main = document.getElementById("main");
    if (!main) return;

    if (this.screen === "loading") {
      main.innerHTML = `<div class="status-card"><p class="muted">Loading Daily Pulse…</p></div>`;
      return;
    }
    if (this.screen === "error") {
      main.innerHTML = `<div class="status-card"><p class="error-text">${escHtml(this._errorMsg || "Something went wrong.")}</p>
        <button class="btn btn-secondary" onclick="location.reload()">Retry</button></div>`;
      return;
    }
    if (this.screen === "denied") {
      main.innerHTML = this._accessDeniedHtml();
      return;
    }
    if (this.screen === "teacher-select") {
      main.innerHTML = this._teacherSelectHtml();
      this._bindTeacherSelectEvents(main);
      return;
    }
    if (this.screen === "status")   { main.innerHTML = this._statusHtml();   this._bindStatusEvents(main);   return; }
    if (this.screen === "complete") { main.innerHTML = this._completeHtml(); this._bindStatusEvents(main);   return; }
    if (this.screen === "entry")    { main.innerHTML = this._entryHtml();    this._bindEntryEvents(main);    return; }
  },

  // Consistent IEP Skook access-denied template (same wording used across
  // the other three apps in the suite). The header's "Sign out" button
  // stays available for switching accounts — no duplicate control needed
  // here. Never renders roster/student data behind this screen.
  _accessDeniedHtml() {
    // Reached only after the IEP_Users2 identity lookup already succeeded
    // (see init()) — so the diagnostic always reports Identity record:
    // Found here; the app-permission check is what failed.
    return `<div class="status-card access-denied">
      <h2>IEP Skook</h2>
      <p>You are signed in, but your account does not currently have access to this application.</p>
      ${loginDiagnosticHtml({ identityFound: "Found", appPermission: "Denied" })}
      <p class="muted">Contact an IEP Skook administrator if you believe this is incorrect.</p>
    </div>`;
  },

  _statusHtml() {
    const total     = this.roster.length;
    const completed = this.roster.filter(s => this.completedToday.has(s.name)).length;
    const remaining = total - completed;
    const name      = (this.currentUser.name || "").split(" ")[0] || this.currentUser.name;

    if (total === 0) {
      return `${this._adminContextHtml()}<div class="status-card">
        <h1 class="status-title">Daily Pulse</h1>
        <p class="muted">${formatDateLong(todayStr())}</p>
        <p class="empty-state">No Daily Pulse-enabled students are assigned to this teacher right now.</p>
        <p class="muted small">Need this changed? Contact an administrator.</p>
      </div>`;
    }

    return `${this._adminContextHtml()}<div class="status-card">
      <h1 class="status-title">Daily Pulse</h1>
      <p class="status-greeting">${greetingWord()}, ${escHtml(name)}</p>
      <p class="muted">${formatDateLong(todayStr())}</p>
      <div class="status-counts">
        <div class="count-item"><span class="count-num">${total}</span><span class="count-label">Students</span></div>
        <div class="count-item"><span class="count-num">${completed}</span><span class="count-label">Complete</span></div>
        <div class="count-item"><span class="count-num">${remaining}</span><span class="count-label">Remaining</span></div>
      </div>
      <button class="btn btn-primary btn-lg btn-full" id="continueBtn">Continue Daily Pulse</button>
      ${this._todaySubmissionsHtml()}
    </div>`;
  },

  _completeHtml() {
    return `${this._adminContextHtml()}<div class="status-card">
      <div class="complete-icon">&#10003;</div>
      <h1 class="status-title">Daily Pulse Complete</h1>
      <p class="status-greeting">${this.roster.length} of ${this.roster.length} students submitted</p>
      <p class="muted">${formatDateLong(todayStr())}</p>
      ${this._todaySubmissionsHtml()}
    </div>`;
  },

  _todaySubmissionsHtml() {
    if (this.todayEntries.length === 0) return "";
    const rows = this.todayEntries.map(e => {
      const label = String(e.attendanceStatus || "").toLowerCase() === "absent" ? "Absent" : (e.status || "—");
      return `<li><span class="ts-student">${escHtml(e.student)}</span><span class="ts-status">${escHtml(label)}</span></li>`;
    }).join("");
    return `<div class="today-submissions">
      <h2 class="today-submissions-title">Today's Submissions</h2>
      <ul class="today-submissions-list">${rows}</ul>
      <p class="muted small">Need to correct an entry? Contact an administrator.</p>
    </div>`;
  },

  _bindStatusEvents(main) {
    this._bindAdminContext(main);
    main.querySelector("#continueBtn")?.addEventListener("click", () => this.startOrContinue());
  },

  _teacherSelectHtml() {
    const choices = this.teacherDirectory.map(teacher => `
      <button type="button" class="teacher-choice" data-teacher="${escHtml(teacher.teacherName)}" aria-pressed="false">
        ${escHtml(teacher.teacherName)}
      </button>`).join("");
    const empty = this.teacherDirectory.length === 0
      ? `<p class="empty-state">No teachers currently have active, Daily Pulse-enabled students.</p>`
      : "";

    return `<div class="status-card teacher-select-card">
      <h1 class="status-title">Daily Pulse</h1>
      <p class="teacher-select-prompt">Completing Daily Pulse for:</p>
      <label class="field-label" for="teacherSearch">Choose a teacher</label>
      <input class="form-input teacher-search-input" id="teacherSearch" type="search"
        placeholder="Search teacher directory" autocomplete="off">
      <div class="teacher-choice-list" id="teacherChoiceList">${choices}</div>
      <p class="muted small teacher-no-results hidden" id="teacherNoResults">No matching teachers.</p>
      ${empty}
      <div id="teacherSelectError" class="entry-error hidden"></div>
      <button type="button" class="btn btn-primary btn-lg btn-full" id="teacherContinueBtn" disabled>Continue</button>
    </div>`;
  },

  _bindTeacherSelectEvents(main) {
    const input = main.querySelector("#teacherSearch");
    const continueBtn = main.querySelector("#teacherContinueBtn");
    const buttons = [...main.querySelectorAll(".teacher-choice")];
    const noResults = main.querySelector("#teacherNoResults");

    const select = teacherName => {
      const exact = this.teacherDirectory.find(
        teacher => teacher.teacherName.toLowerCase() === String(teacherName || "").trim().toLowerCase()
      );
      this.pendingTeacherName = exact?.teacherName || "";
      buttons.forEach(button => {
        const selected = button.dataset.teacher === this.pendingTeacherName;
        button.classList.toggle("is-selected", selected);
        button.setAttribute("aria-pressed", String(selected));
      });
      if (continueBtn) continueBtn.disabled = !this.pendingTeacherName;
    };

    input?.addEventListener("input", () => {
      const term = input.value.trim().toLowerCase();
      let visible = 0;
      buttons.forEach(button => {
        const matches = button.dataset.teacher.toLowerCase().includes(term);
        button.classList.toggle("hidden", !matches);
        if (matches) visible += 1;
      });
      noResults?.classList.toggle("hidden", visible > 0 || buttons.length === 0);
      select(input.value);
    });

    buttons.forEach(button => button.addEventListener("click", () => {
      if (input) input.value = button.dataset.teacher;
      select(button.dataset.teacher);
    }));

    continueBtn?.addEventListener("click", async () => {
      continueBtn.disabled = true;
      const selected = await this.selectAdminTeacher(this.pendingTeacherName);
      if (!selected) {
        const error = main.querySelector("#teacherSelectError");
        if (error) {
          error.textContent = "Please choose a teacher from the directory.";
          error.classList.remove("hidden");
        }
        continueBtn.disabled = false;
      }
    });
  },

  _adminContextHtml() {
    if (!this.currentUser?.isAdmin || !this.effectiveTeacher) return "";
    return `<aside class="admin-context" aria-label="Administrator teacher context">
      <span>Completing Daily Pulse for</span>
      <strong>${escHtml(this.effectiveTeacher.teacherName)}</strong>
      <button type="button" class="btn btn-secondary" id="chooseTeacherBtn">Choose Another Teacher</button>
    </aside>`;
  },

  _bindAdminContext(main) {
    main.querySelector("#chooseTeacherBtn")?.addEventListener("click", () => this.chooseAnotherTeacher());
  },

  _entryHtml() {
    const draft      = this.draft;
    const total      = this.roster.length;
    const idx        = this.roster.findIndex(s => s.name === draft.studentName);
    const num        = idx >= 0 ? idx + 1 : this.roster.length - this._remainingRoster().length + 1;
    const dots       = this.roster.map(s => this.completedToday.has(s.name) ? "done" : (s.name === draft.studentName ? "current" : "pending"));
    const showPresent = draft.attendanceStatus === "present";
    const showAbsentConfirm = draft.attendanceStatus === "absent";

    return `
      ${this._adminContextHtml()}
      <div class="entry-header">
        <button class="link-btn" id="backToStatusBtn">&larr; Status</button>
        <div class="entry-progress-text">Student ${num} of ${total}</div>
      </div>
      <div class="progress-dots">${dots.map(d => `<span class="dot dot-${d}"></span>`).join("")}</div>

      <div class="form-card">
        <h2 class="student-name">${escHtml(draft.studentName)}</h2>
      </div>

      <div class="form-card">
        <div class="field-label">Was the student present? <span class="req">*</span></div>
        <div class="attendance-options">
          <button type="button" class="attendance-option attendance-present${draft.attendanceStatus === "present" ? " is-selected" : ""}" data-attendance="present">Present</button>
          <button type="button" class="attendance-option attendance-absent${draft.attendanceStatus === "absent" ? " is-selected" : ""}" data-attendance="absent">Absent</button>
        </div>
        <div id="absentConfirm" class="${showAbsentConfirm ? "absent-confirm" : "hidden"}">
          <p>Mark <strong>${escHtml(draft.studentName)}</strong> absent and continue?</p>
          <div class="absent-confirm-btns">
            <button type="button" class="btn btn-danger" id="confirmAbsentBtn">Confirm Absent</button>
            <button type="button" class="btn btn-secondary" id="cancelAbsentBtn">Cancel</button>
          </div>
        </div>
      </div>

      <div id="presenceSections" class="${showPresent ? "" : "hidden"}">
        <div class="form-card">
          <div class="field-label">How was the student's day? <span class="req">*</span></div>
          <div class="option-cards">
            <label class="option-card${draft.overallStatus === "great" ? " selected" : ""}" data-pulse="great">
              <input type="radio" name="pulseStatus" value="great"${draft.overallStatus === "great" ? " checked" : ""}>
              <span class="status-emoji">&#128522;</span><span>Great Day</span>
            </label>
            <label class="option-card${draft.overallStatus === "struggles" ? " selected" : ""}" data-pulse="struggles">
              <input type="radio" name="pulseStatus" value="struggles"${draft.overallStatus === "struggles" ? " checked" : ""}>
              <span class="status-emoji">&#128528;</span><span>Some Struggles</span>
            </label>
            <label class="option-card${draft.overallStatus === "concern" ? " selected" : ""}" data-pulse="concern">
              <input type="radio" name="pulseStatus" value="concern"${draft.overallStatus === "concern" ? " checked" : ""}>
              <span class="status-emoji">&#9785;&#65039;</span><span>Significant Concern</span>
            </label>
          </div>
        </div>

        <div class="form-card">
          <div class="field-label">Focus Areas <span class="optional-tag">optional</span></div>
          <p class="section-help">Select one or more to add detail.</p>
          <div class="category-grid">
            ${Object.entries(DAILY_PULSE_SCALES).map(([key, scale]) => {
              const isSel  = draft.selectedCategories.includes(key);
              const isDone = !!(draft.categoryResponses[key]?.rating);
              return `<button type="button" class="category-btn${isSel ? " is-selected" : ""}${isDone ? " is-complete" : ""}" data-category-key="${key}">
                ${escHtml(scale.label)}${isDone ? " &#10003;" : ""}
              </button>`;
            }).join("")}
          </div>
          <div id="scalePanels">${draft.selectedCategories.map(key => this._scalePanelHtml(key)).join("")}</div>
        </div>

        <div class="form-card">
          <div class="field-label">Quick Note <span class="optional-tag">optional</span></div>
          <textarea class="form-textarea" id="quickNote" rows="2" placeholder="Any additional observations or context.">${escHtml(draft.quickNote)}</textarea>
        </div>

        <div id="entryError" class="entry-error hidden"></div>

        <div class="save-bar">
          <button type="button" class="btn btn-primary btn-lg btn-full" id="submitBtn">Submit &amp; Next</button>
        </div>
      </div>`;
  },

  _scalePanelHtml(categoryKey) {
    const scale = DAILY_PULSE_SCALES[categoryKey];
    if (!scale) return "";
    const response  = this.draft.categoryResponses[categoryKey] || { rating: "", multiSelect: {}, details: {} };
    const choicesHtml = scale.choices.map(c => {
      const isSel = response.rating === c.value;
      return `<button type="button" class="scale-choice-btn scale-choice-${c.tone}${isSel ? " is-selected" : ""}" data-category="${categoryKey}" data-value="${escHtml(c.value)}">${escHtml(c.label)}</button>`;
    }).join("");
    return `<section class="scale-panel" data-scale-panel="${categoryKey}">
      <header class="scale-header"><h3>${escHtml(scale.label)}</h3><p>${escHtml(scale.prompt)}</p></header>
      <div class="scale-choice-grid">${choicesHtml}</div>
      <div class="scale-detail-wrapper">${this._scaleDetailHtml(scale, response, categoryKey)}</div>
    </section>`;
  },

  _scaleDetailHtml(scale, response, categoryKey) {
    let html = "";
    if (scale.multiSelectFields) {
      scale.multiSelectFields.forEach(field => {
        const selected = response.multiSelect?.[field.key] || [];
        html += `<div class="scale-detail-section">
          <div class="scale-detail-label">${escHtml(field.label)}</div>
          <div class="chip-grid">
            ${field.options.map(opt => {
              const isSel = selected.includes(opt);
              return `<button type="button" class="detail-chip${isSel ? " is-selected" : ""}" data-category="${categoryKey}" data-field="${escHtml(field.key)}" data-option="${escHtml(opt)}">${escHtml(opt)}</button>`;
            }).join("")}
          </div>
        </div>`;
      });
    }
    if (scale.detailFields) {
      scale.detailFields.forEach(field => {
        if (field.showForValues && !field.showForValues.includes(response.rating)) return;
        if (field.showWhenOption) {
          const anySelected = Object.values(response.multiSelect || {}).some(arr => Array.isArray(arr) && arr.includes(field.showWhenOption));
          if (!anySelected) return;
        }
        const val = response.details?.[field.key] || "";
        const tag = field.type === "textarea" ? "textarea" : "input";
        html += `<div class="scale-detail-section">
          <div class="scale-detail-label">${escHtml(field.label)} <span class="optional-hint">(optional)</span></div>
          ${tag === "textarea"
            ? `<textarea class="scale-text-field form-textarea" data-category="${categoryKey}" data-field="${escHtml(field.key)}" rows="2">${escHtml(val)}</textarea>`
            : `<input type="text" class="scale-text-field form-input" data-category="${categoryKey}" data-field="${escHtml(field.key)}" value="${escHtml(val)}">`}
        </div>`;
      });
    }
    return html;
  },

  _bindEntryEvents(main) {
    this._bindAdminContext(main);
    main.querySelector("#backToStatusBtn")?.addEventListener("click", () => this.backToStatus());

    main.querySelectorAll(".attendance-option").forEach(btn => {
      btn.addEventListener("click", () => {
        this.draft.attendanceStatus = btn.dataset.attendance;
        this._render();
      });
    });

    main.querySelector("#confirmAbsentBtn")?.addEventListener("click", () => this.submitAbsent());
    main.querySelector("#cancelAbsentBtn")?.addEventListener("click", () => {
      this.draft.attendanceStatus = null;
      this._render();
    });

    main.querySelectorAll('input[name="pulseStatus"]').forEach(radio => {
      radio.addEventListener("change", () => { this.draft.overallStatus = radio.value; this._render(); });
    });

    main.querySelectorAll(".category-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const key = btn.dataset.categoryKey;
        const isSel = this.draft.selectedCategories.includes(key);
        this.draft.selectedCategories = isSel
          ? this.draft.selectedCategories.filter(k => k !== key)
          : [...this.draft.selectedCategories, key];
        this._render();
      });
    });

    main.querySelectorAll(".scale-choice-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const categoryKey = btn.dataset.category;
        const response = this.draft.categoryResponses[categoryKey] || { rating: "", multiSelect: {}, details: {} };
        response.rating = btn.dataset.value;
        this.draft.categoryResponses[categoryKey] = response;
        this._render();
      });
    });

    main.querySelectorAll(".detail-chip").forEach(chip => {
      chip.addEventListener("click", () => {
        const categoryKey = chip.dataset.category, fieldKey = chip.dataset.field, optionVal = chip.dataset.option;
        const response = this.draft.categoryResponses[categoryKey] || { rating: "", multiSelect: {}, details: {} };
        if (!response.multiSelect) response.multiSelect = {};
        if (!response.multiSelect[fieldKey]) response.multiSelect[fieldKey] = [];
        const arr = response.multiSelect[fieldKey];
        const i = arr.indexOf(optionVal);
        if (i >= 0) arr.splice(i, 1); else arr.push(optionVal);
        this.draft.categoryResponses[categoryKey] = response;
        this._render();
      });
    });

    main.querySelectorAll(".scale-text-field").forEach(field => {
      field.addEventListener("input", () => {
        const categoryKey = field.dataset.category, fieldKey = field.dataset.field;
        const response = this.draft.categoryResponses[categoryKey] || { rating: "", multiSelect: {}, details: {} };
        if (!response.details) response.details = {};
        response.details[fieldKey] = field.value;
        this.draft.categoryResponses[categoryKey] = response;
      });
    });

    main.querySelector("#quickNote")?.addEventListener("input", e => { this.draft.quickNote = e.target.value; });
    main.querySelector("#submitBtn")?.addEventListener("click", () => this.submitPresent());
  }
};

document.addEventListener("DOMContentLoaded", () => APP.init());
