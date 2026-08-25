# DAILY PULSE SOURCE AUDIT

Audit date: 2026-08-25
Source inspected (read-only): `/Users/greg_macer/Projects/01-IU29/MAC-Walkthrough` (git-tracked, Vercel-linked; `~/tries/MAC-Walkthrough` does not exist — location confirmed with the user before inspection).

## 1. Main files involved
`app.js` (Daily Pulse UI/state/submit, ~295KB, DAILY_PULSE_SCALES + APP.render/submit methods), `graph.js` (`GRAPH.saveDailyPulse`, generic list/schema/create helpers), `auth.js` (MSAL + `IEP_Users2` lookup), `config.js` (`STUDENT_ROSTER_LIST`, storage keys), `data.js` (local draft/history helpers), `index.html` (script load order, MSAL CDN), `styles.css` (`.pulse-*`, `.attendance-*`, `.scale-*`, `.walk-*`, `.form-card`, `.btn*` rules).

## 2. Daily Pulse route/view
Hash route `#pulse` → `APP.navigate("pulse")` → `APP.renderPulse()` into `#page-pulse`. It is the default landing page for authenticated teachers (`app.js:1199`).

## 3. UI implementation
DOM-string rendering (`renderPulse()`, `app.js:2513`). Student picked from a `<select>` (not auto-advanced) showing a ✓ for already-completed students. Sections: Student → Attendance (Present/Absent buttons) → (if absent) confirm dialog → (if present) Overall Day (3 option-cards: Great Day 😊 / Some Struggles 😐 / Significant Concern ☹️) → optional "Focus Areas" (7 toggleable detail scales, see below) → Quick Note textarea → Submit button.

## 4. State implementation
Single module-level `APP._pulseDraft` object (not persisted to localStorage as a draft — lost on reload, by design, since a confirmed SharePoint save is what "counts"). Shape: `{ studentName, attendanceStatus, overallStatus, selectedCategories[], categoryResponses{7 keys}, quickNote, date, recordId }`. `recordId = "pulse-" + crypto.randomUUID()`, generated once per `_initPulseDraft()` call and reused across a retry.

## 5. Authentication implementation
MSAL Browser 2.x (CDN, `index.html:8`). `clientId: 145a3fc7-5cff-4d03-96c7-577e17980110`, tenant authority `https://login.microsoftonline.com/3276761c-22db-462b-a930-172d155bd795`, `redirectUri: window.location.origin`, localStorage cache. Scopes `User.Read`, `Sites.ReadWrite.All`. Login/logout via redirect flow (`auth.js`).

## 6. Authorization implementation
After MSAL sign-in, `GRAPH.findUserByEmail()` scans **`IEP_Users2`** for the signed-in email (checks `Email`/`email`/`EMail`/`Email0`/`EmailAddress`/`Email_x0020_Address`/`Email Address`). `AUTH.loadPilotUserFromSharePoint()` then reads `Title`→userId, `field_1`/`Name`→name, `field_2`/`Role`→role, `field_3`/`Active`→active (Yes/true/1/boolean-true all accepted). No match or inactive → `AUTH.isUnauthorized` → login screen shows a denial message. Role is normalized lowercase (`normalizeUserRole`); `AUTH.isTeacher` = role === "teacher", `AUTH.isAdmin` = role === "administrator".

## 7. Teacher identity mapping
Two layers exist. (a) **Legacy/hardcoded**: `TEACHER_IDENTITIES` (`app.js:348`) maps exactly 2 pilot `userId`s to a name/email/classroom — dead weight for a general teacher population. (b) **Live/general fallback** (`app.js:379-389`, used whenever a teacher isn't one of those 2 pilot IDs — i.e. essentially always in production): builds `{ userId: AUTH.pilotUserId, teacherName: AUTH.pilotUser.Name, displayName: AUTH.pilotUser.Name, email: AUTH.account.username }` directly from the authenticated `IEP_Users2` record. `getCurrentUser()` (`app.js:1249`) wraps this into `{ id, name, teacherName, role, isAdmin }`. **The standalone app should use path (b) only** and must not copy `TEACHER_IDENTITIES` or the "view as teacher" admin override — see §26.

## 8. Student filtering implementation
`STUDENT_ROSTER` (`app.js:215`), a live provider over **`IEP_Students_2026_27`** (via `CONFIG.STUDENT_ROSTER_LIST`). Normalizes rows to `{ id (SharePoint item id), firstName, lastName, name, teacherName, teacherId (slugified), classroom, active, paceEnabled, dailyPulseEnabled }`. Teacher match is case/whitespace-insensitive exact string match on `Teacher` display name (`_matchesTeacher`) — no ID-based linkage exists in the schema.

## 9. Active filtering
`normalizeRosterBoolean(row["Active"], true)` — defaults **true** when the column is blank (fail-open, preserves existing rows without the column set). `getActive()` filters on this.

## 10. Daily Pulse Enabled filtering
`normalizeRosterBoolean(row["Daily Pulse Enabled"], false)` — defaults **false** when blank (fail-closed/opt-in, unlike Active). `getDailyPulseForTeacher(teacherName)` = active ∧ dailyPulseEnabled ∧ teacher match. Admin variant `getDailyPulseEnabled()` skips the teacher match.

## 11. Attendance handling
Two buttons, Present/Absent (`app.js:2592`). Absent shows an inline confirm ("Mark X absent and continue?") before submitting, to guard against a mis-tap.

## 12. Absent workflow
Confirming Absent calls `_submitPulseAbsent()` (`app.js:2930`), which **skips Overall Day, all 7 focus-area scales, and Quick Note entirely** — it submits immediately with `pulseStatus: "", categories: [], categoryResponses: {}, note: ""`. No fabricated "N/A" values are written; fields are just empty strings/arrays, matching existing behavior.

## 13. Day-rating values
Exactly 3, stored as short codes with display labels applied only at submit time: `great` → "Great Day", `struggles` → "Some Struggles", `concern` → "Significant Concern" (`app.js:3043`). These display strings are what's written to the `Status` SharePoint column; `AttendanceStatus`/internal state keeps the short code.

## 14. Additional detail fields
7 optional "Focus Area" scales in `DAILY_PULSE_SCALES` (`app.js:1001`): Academic Progress, Behavior, Communication, Social Skills, Self-Regulation, Transition Difficulty, IEP Goal Progress. Each has a 4–5 option single-select "rating", and some have multi-select chip fields (e.g. Behavior → "Behaviors observed") and free-text fields (some conditional on rating value or on an "Other" chip being selected). None are required — a teacher can submit with zero focus areas expanded.

## 15. Validation
Required: student selected, attendance status selected, and — only when Present — Overall Day selected. If a focus-area scale was expanded, its rating becomes required (`app.js:3017`); its multi-select/text sub-fields stay optional. First violation shown inline (`_showPulseError`), scrolled into view; no Graph call is made until validation passes.

## 16. Sequential progression
**Not automatic on load** — the source shows a plain dropdown with a ✓ marker for done students; nothing pre-selects the first incomplete student, and nothing prevents re-selecting an already-done student (it just shows a "will create a second entry" warning). Progression *after a submit* **is** automatic: `_advancePulseToNextStudent()` (`app.js:3110`) walks to the next roster entry by array index and auto-selects it. The brief's "select first incomplete on load/resume" is a gap versus source behavior — see §16 in the Extraction Report for how this is closed.

## 17. Today's-date logic
`draft.date = new Date().toISOString().slice(0,10)`, set once per draft init; not recomputed mid-session (a pulse started right before midnight keeps that date).

## 18. Existing-submission lookup
**Local-only in source.** `completedToday` is a `Set` built from `DB.getDailyPulses()` — plain `localStorage`, keyed by `CONFIG.STORAGE_KEYS.DAILY_PULSES`. There is **no SharePoint query** in source for "who did I already do today" — it purely reflects this browser's local history. See Extraction Report §"Today's Completion Detection" for the adaptation needed since the standalone app won't share that localStorage with MAC-Walkthrough.

## 19. Duplicate protection
Per-submission, not per-student-per-day: `GRAPH.saveDailyPulse()` looks up the target list by display field **"Record ID"** (`entry.RecordID`, the same stable UUID minted once by `_initPulseDraft`) via `findListItemByDisplayField`. If found, returns `{ duplicatePrevented: true }` without a second create. A module-level `APP._pulseInFlight` boolean blocks a second submit while one is in progress (checked at the top of both `_submitPulsePresent` and `_submitPulseAbsent`).

## 20. Failed-save protection
On a Graph error, the local copy (`DB.addDailyPulse`) already happened, but the draft (`recordId`, `studentName`, `attendanceStatus`, all answers) is left untouched, the button is re-enabled, and `showSyncError()` displays the failure — the teacher can retry with the same `recordId` (idempotent) without re-entering anything or creating a duplicate. No advance happens on failure.

## 21. SharePoint list used
**`IEP_Daily_Pulse`**, in site `siu29.sharepoint.com:/sites/IEP_Skook:`. Roster list is the separate **`IEP_Students_2026_27`**; auth list is **`IEP_Users2`**.

## 22. Exact field/payload mapping
`GRAPH.saveDailyPulse()` (`graph.js:313`) writes these **display-name** fields to `IEP_Daily_Pulse` via the dynamic schema mapper (`mapFields`/`createMappedListItem`, which resolves display→internal names at runtime and silently skips any column that doesn't exist — never guesses internal names):

| Display field | Source |
|---|---|
| Record ID | stable per-draft UUID (dup-detection key) |
| Date | `entry.Date` (YYYY-MM-DD) |
| Teacher | roster student's `teacherName`, else current user's teacherName/name |
| Student | `draft.studentName` |
| Attendance Status | `"Absent"` or `"Present"` |
| Present | boolean |
| Status | display label ("Great Day"/"Some Struggles"/"Significant Concern") or `""` if absent |
| Category | comma-joined `selectedCategories` keys |
| Category Responses | `JSON.stringify(categoryResponses)` |
| Academic Progress Rating / Behavior Rating / Communication Rating / Social Skills Rating / Self-Regulation Rating / Transition Rating / IEP Goal Rating | `responses.<key>.rating` per scale, else `""` |
| Quick Note | `draft.quickNote` |
| Submitted By | resolved teacher display name (falls back to the *real* signed-in user's name if an admin were viewing-as, N/A for standalone) |
| Submitted At | ISO timestamp from the local `DB.addDailyPulse` entry |

## 23. Demo-mode architecture
No demo mode exists in source at all — MAC-Walkthrough is production-only. The standalone app's Demo Mode is therefore a **new, additive** capability, built from scratch on top of the same UI/state code but swapping `STUDENT_ROSTER`/`GRAPH` for fixture data + a `localStorage`-backed fake save — see Extraction Report.

## 24. PWA/service-worker architecture
None in source (`index.html` has no manifest link or SW registration; no `manifest.webmanifest`/`sw.js` files exist in the repo). The standalone app's PWA layer is also new, modeled loosely after `~/tries/PACE Room Tracker/sw.js` and `manifest.webmanifest` (a genuinely standalone sibling app) but with its own cache name/scope so it can never intersect MAC-Walkthrough's origin.

## 25. Dependencies Daily Pulse truly requires
MSAL config/flow + `IEP_Users2` lookup (`auth.js` in full); Graph transport, site/list discovery, schema mapping, `saveDailyPulse`, `findListItemByDisplayField`, `createMappedListItem`, `getListItems`/`getListSchema` (`graph.js`, most of it — walkthrough/PACE-specific save functions excluded); `STUDENT_ROSTER` provider + `normalizeRosterBoolean` + `slugifyTeacherId` (`app.js:195-344`); `DAILY_PULSE_SCALES` constant (`app.js:1001-1125`); the pulse render/state/submit methods (`app.js:2456-3145`); `CONFIG.STUDENT_ROSTER_LIST` and the pulse `STORAGE_KEYS` entry (`config.js`); the local draft-history helpers `getDailyPulses`/`addDailyPulse` (`data.js`, kept only as a resilience/local-history layer per the brief, not as the source of truth); the `.pulse-*`/`.attendance-*`/`.scale-*`/`.walk-*`/`.form-card`/`.btn*`/auth-screen CSS rules from `styles.css`.

## 26. Components that must NOT be extracted
`TEACHER_IDENTITIES` and all of `PILOT_TEACHERS`/`PILOT_STUDENTS`/`PILOT_ADMIN` (`app.js:89-186`) — **these contain real, hardcoded student and staff names** and the source file itself labels them legacy/deprecated for Daily Pulse; `USER_CONTEXT.isViewingAsTeacher` / admin "view as teacher" override; `AUTH.isAdmin`-gated routes and all admin rendering (`renderTeacherDashboard`, `renderPaceAdmin`, reports, setup, storage settings, Form Lab, dashboard-sync.js entirely); `SETUP_LISTS` beyond `users`; walkthrough/PACE/Student-Check-In save paths in `graph.js` (`saveWalkthrough`, `saveWalkthroughV2`, `savePaceVisit`, `getWhoAreYouVisiting`); `api/teachers.js` (unrelated Vercel function reading a `TEACHER_LIST` env var, not used by Daily Pulse).

---
No destructive operation or unresolved backend dependency was found that requires stopping — continuing to the extraction per the standing instruction.
