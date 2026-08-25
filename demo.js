/* ── Demo Mode ────────────────────────────────────────────────────────────
 * New capability — MAC-Walkthrough has no demo mode to extract (see
 * docs/daily-pulse-source-audit.md §23). Built to be PHYSICALLY incapable
 * of reaching production: app.js checks `APP.demoMode` before ever calling
 * AUTH.init(), STUDENT_ROSTER.refresh(), or any GRAPH.* method, and takes
 * the DEMO.* facade branch instead — no MSAL script even runs its login
 * flow, no fetch() to graph.microsoft.com or login.microsoftonline.com is
 * reachable from this code path. Verified in docs/demo-verification.md.
 *
 * DEMO mirrors just the shape app.js needs from STUDENT_ROSTER + GRAPH so
 * the render/submit code in app.js does not need separate demo branches.
 */

const DEMO_TEACHER = {
  id:          "demo-teacher",
  name:        "Demo Teacher",
  teacherName: "Demo Teacher",
  email:       "demo.teacher@example.invalid",
  isAdmin:     false
};

const DEMO_STUDENTS = [
  { id: "demo-1", firstName: "Riley",   lastName: "Fox",     name: "Riley Fox",     teacherName: "Demo Teacher", classroom: "Demo Room", active: true, dailyPulseEnabled: true },
  { id: "demo-2", firstName: "Sam",     lastName: "Ortiz",   name: "Sam Ortiz",     teacherName: "Demo Teacher", classroom: "Demo Room", active: true, dailyPulseEnabled: true },
  { id: "demo-3", firstName: "Jordan",  lastName: "Blake",   name: "Jordan Blake",  teacherName: "Demo Teacher", classroom: "Demo Room", active: true, dailyPulseEnabled: true },
  { id: "demo-4", firstName: "Casey",   lastName: "Nguyen",  name: "Casey Nguyen",  teacherName: "Demo Teacher", classroom: "Demo Room", active: true, dailyPulseEnabled: true },
  { id: "demo-5", firstName: "Avery",   lastName: "Patel",   name: "Avery Patel",   teacherName: "Demo Teacher", classroom: "Demo Room", active: true, dailyPulseEnabled: true }
];

const DEMO = {
  currentUser() { return DEMO_TEACHER; },

  getRoster() { return DEMO_STUDENTS; },

  _pulses() {
    try {
      return JSON.parse(localStorage.getItem(CONFIG.STORAGE_KEYS.DEMO_PULSES)) || [];
    } catch {
      return [];
    }
  },
  _savePulses(list) {
    localStorage.setItem(CONFIG.STORAGE_KEYS.DEMO_PULSES, JSON.stringify(list));
  },

  getTodayCompleted(dateStr) {
    return new Set(
      this._pulses()
        .filter(p => p.date === dateStr)
        .map(p => p.student)
    );
  },

  getTodayEntries(dateStr) {
    return this._pulses().filter(p => p.date === dateStr);
  },

  // Simulated network latency + the same duplicate-by-RecordID guard as the
  // real GRAPH.saveDailyPulse, entirely inside localStorage.
  async savePulse(entry) {
    await new Promise(r => setTimeout(r, 300));
    const list = this._pulses();
    if (list.some(p => p.recordId === entry.RecordID)) {
      return { duplicatePrevented: true };
    }
    list.push({
      recordId: entry.RecordID,
      date:     entry.Date,
      student:  entry.Student,
      status:   entry.Status,
      attendanceStatus: entry.AttendanceStatus,
      submittedAt: entry.SubmittedAt
    });
    this._savePulses(list);
    return { ok: true };
  },

  reset() {
    localStorage.removeItem(CONFIG.STORAGE_KEYS.DEMO_PULSES);
  },

  isActive() {
    return sessionStorage.getItem(CONFIG.STORAGE_KEYS.DEMO_MODE) === "1";
  },
  activate() {
    sessionStorage.setItem(CONFIG.STORAGE_KEYS.DEMO_MODE, "1");
  },
  deactivate() {
    sessionStorage.removeItem(CONFIG.STORAGE_KEYS.DEMO_MODE);
  }
};
