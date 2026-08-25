/* ── Daily Pulse Configuration ───────────────────────────────────────────────
 * Standalone teacher app, extracted from MAC-Walkthrough. This file only
 * ever knows HOW to reach the live SharePoint list — never WHO is on it or
 * what a submission contains beyond field names.
 */
const CONFIG = {
  VERSION:  "0.1.0",
  APP_NAME: "Daily Pulse",
  ORG_NAME: "IU29",

  // Same production list MAC-Walkthrough reads, so a Daily Pulse Enabled
  // flip in SharePoint takes effect here with zero code changes. See
  // docs/daily-pulse-source-audit.md §8-10 for the exact column contract.
  STUDENT_ROSTER_LIST: "IEP_Students_2026_27",

  // Same list MAC-Walkthrough writes Daily Pulse entries to. The admin app
  // must be able to read what this app creates with no translation step.
  DAILY_PULSE_LIST: "IEP_Daily_Pulse",

  // Same authorization list MAC-Walkthrough uses.
  USERS_LIST: "IEP_Users2",

  // localStorage keys — namespaced distinctly from MAC-Walkthrough's
  // "skookWalkthrough_*" keys so the two apps' local state never collides,
  // even if someone ever runs both against the same origin during dev.
  STORAGE_KEYS: {
    LOCAL_HISTORY: "dailyPulseApp_localHistory",
    DEMO_TEACHERS: "dailyPulseApp_demoTeachers",
    DEMO_STUDENTS: "dailyPulseApp_demoStudents",
    DEMO_PULSES:   "dailyPulseApp_demoPulses",
    DEMO_MODE:     "dailyPulseApp_demoMode"
  }
};
