/* ── Daily Pulse — standalone teacher app ────────────────────────────────
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
  screen:          "loading", // loading | denied | error | status | entry | complete
  currentUser:     null,      // { id, name, teacherName, email }
  roster:          [],        // eligible students for this teacher
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

    // Daily Pulse is a teacher tool — see docs/daily-pulse-extraction-
    // report.md "Teacher Identification" for why administrators are
    // pointed to MAC-Walkthrough instead of being let in here.
    if (!AUTH.isTeacher) {
      this.screen = "denied";
      this._render();
      return;
    }

    this.currentUser = this._resolveTeacherIdentity();
    await this._loadRosterAndStatus();
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

  // Live/general identity path only (extracted from MAC-Walkthrough
  // app.js:379-389 fallback branch) — the hardcoded 2-teacher
  // TEACHER_IDENTITIES dict and admin "view as teacher" override are
  // intentionally not extracted. See docs/daily-pulse-source-audit.md §7,
  // §26.
  _resolveTeacherIdentity() {
    const name = AUTH.pilotUser?.Name || AUTH.displayName || "";
    return {
      id:          AUTH.pilotUserId || slugifyTeacherId(name),
      name,
      teacherName: name,
      email:       AUTH.account?.username || ""
    };
  },

  async _loadRosterAndStatus() {
    this._showScreen("loading");
    try {
      if (this.demoMode) {
        this.roster = DEMO.getRoster().filter(s => s.active && s.dailyPulseEnabled);
      } else {
        await STUDENT_ROSTER.refresh();
        if (STUDENT_ROSTER.error) throw new Error(STUDENT_ROSTER.error);
        this.roster = STUDENT_ROSTER.getDailyPulseForTeacher(this.currentUser.teacherName);
      }

      const date = todayStr();
      if (this.demoMode) {
        this.completedToday = DEMO.getTodayCompleted(date);
        this.todayEntries   = DEMO.getTodayEntries(date);
      } else {
        const entries = await GRAPH.getTodayDailyPulseEntries(this.currentUser.teacherName, date);
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
    const present = attendanceStatus === "present";
    const statusLabel = present
      ? ({ great: "Great Day", struggles: "Some Struggles", concern: "Significant Concern" }[draft.overallStatus] || draft.overallStatus)
      : "";
    return {
      RecordID:          draft.recordId,
      Date:              draft.date,
      Teacher:           this.currentUser.teacherName,
      Student:           draft.studentName,
      AttendanceStatus:  attendanceStatus,
      Present:           present,
      Status:            statusLabel,
      Categories:        present ? [...draft.selectedCategories] : [],
      CategoryResponses: present ? structuredClone(draft.categoryResponses) : {},
      QuickNote:         present ? (draft.quickNote || "") : "",
      SubmittedBy:       this.currentUser.name,
      SubmittedByEmail:  this.currentUser.email,
      SubmittedAt:       new Date().toISOString()
    };
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
      teacherName:       this.currentUser.teacherName,
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
      main.innerHTML = `<div class="status-card">
        <p>Daily Pulse is a teacher tool. Your account role isn't set up as Teacher in IEP Skook.</p>
        <p class="muted">Administrators: use MAC-Walkthrough for the full toolset.</p>
      </div>`;
      return;
    }
    if (this.screen === "status")   { main.innerHTML = this._statusHtml();   this._bindStatusEvents(main);   return; }
    if (this.screen === "complete") { main.innerHTML = this._completeHtml(); this._bindStatusEvents(main);   return; }
    if (this.screen === "entry")    { main.innerHTML = this._entryHtml();    this._bindEntryEvents(main);    return; }
  },

  _statusHtml() {
    const total     = this.roster.length;
    const completed = this.roster.filter(s => this.completedToday.has(s.name)).length;
    const remaining = total - completed;
    const name      = (this.currentUser.name || "").split(" ")[0] || this.currentUser.name;

    if (total === 0) {
      return `<div class="status-card">
        <h1 class="status-title">Daily Pulse</h1>
        <p class="muted">${formatDateLong(todayStr())}</p>
        <p class="empty-state">No Daily Pulse-enabled students are assigned to you right now.</p>
        <p class="muted small">Need this changed? Contact an administrator.</p>
      </div>`;
    }

    return `<div class="status-card">
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
    return `<div class="status-card">
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
    main.querySelector("#continueBtn")?.addEventListener("click", () => this.startOrContinue());
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
