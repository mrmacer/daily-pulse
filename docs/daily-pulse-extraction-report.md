# DAILY PULSE EXTRACTION REPORT

## Existing Source Architecture

MAC-Walkthrough is a single-page vanilla-JS app (no build step) authenticating
teachers/admins via MSAL against a shared Entra tenant, authorizing them
against the `IEP_Users2` SharePoint list, and reading/writing several other
SharePoint lists via a generic Microsoft Graph wrapper (`graph.js`). Daily
Pulse is one route (`#pulse`) inside that app, reading students from
`IEP_Students_2026_27` and writing to `IEP_Daily_Pulse`. Full detail in
[daily-pulse-source-audit.md](daily-pulse-source-audit.md).

## Source Files Inspected

`app.js`, `auth.js`, `graph.js`, `config.js`, `data.js`, `index.html`,
`styles.css`, `dashboard-sync.js` (confirmed unused by Daily Pulse),
`api/teachers.js` (confirmed unrelated), `.vercel/project.json` (project name
only), `.env.local` (variable **names** only — never its values).

## Reused Modules

`auth.js` — copied verbatim (MSAL config, `IEP_Users2` lookup, role helpers).
`graph.js` core transport (`_get`/`_post`, `getSiteId`/`getListId`,
`getListItems`/`getAllListItems`/`getListItemsByDisplayName`,
`findListItemByDisplayField`, `findUserByEmail`, `getListSchema`/`mapFields`/
`createMappedListItem`, `isWritableSharePointField`) — copied verbatim.
`saveDailyPulse`'s exact field mapping — copied verbatim. `DAILY_PULSE_SCALES`
(all 7 focus-area scales, choices, multi-select/detail fields) — copied
verbatim. `STUDENT_ROSTER`'s normalization/filtering logic
(`normalizeRosterBoolean`, `slugifyTeacherId`, Active/Daily-Pulse-Enabled
default-value behavior) — copied verbatim, PACE-only getters dropped.

## Adapted Modules

- **Effective teacher identity** (`access.js` and `app.js`): teachers resolve
  directly from their authenticated `IEP_Users2` record. Administrators use
  an explicit teacher selection derived from the live roster's eligible
  teacher values. No hardcoded teacher identities are used.
- **Today's-completion detection** (`graph.js` → `getTodayDailyPulseEntries()`):
  source computes this purely from local `localStorage`
  (`DB.getDailyPulses()`), which never leaves the browser it was written in.
  A standalone app that a teacher might open on a different device/browser
  needs the shared backend as the source of truth, so this reads
  `IEP_Daily_Pulse` directly (via the same generic `getListItemsByDisplayName`
  reader already in `graph.js`) filtered by teacher + today's date. This is
  the source's own reading infrastructure, just pointed at a query the
  source didn't need.
- **Entry flow structure** (`app.js` render/submit methods): the underlying
  validation rules, payload construction, absent-skip behavior,
  duplicate-tap guard, and failure-preservation behavior are unchanged: only
  the layout changed from source's single-page-with-a-dropdown into a
  strictly sequential one-student-per-screen flow with a status screen in
  front of it, per the product brief.

## Reimplemented Logic

- **Demo Mode** (`demo.js`) — source has no demo mode at all (MAC-Walkthrough
  is production-only). Built from scratch, mirroring only the shape `app.js`
  needs from `STUDENT_ROSTER`/`GRAPH` so the render/submit code needs no
  demo-specific branches.
- **PWA layer** (`manifest.webmanifest`, `sw.js`, `icons/`) — source has none
  to extract (no manifest link, no `sw.js`, no registration code in
  `index.html`). Built new, with its own cache name (`daily-pulse-cache-v2`)
  and an explicit network-only bypass for any Microsoft-owned domain.
- **"First incomplete student" auto-selection** — source's dropdown-based
  picker does not auto-select on load, only after a submit (and even then
  just walks to the next array index, not the next *incomplete* one). The
  brief requires resuming directly into the first incomplete student, so
  `startOrContinue()`/`_afterSuccessfulSave()` compute `_remainingRoster()`
  (roster minus `completedToday`) and select its first entry, rather than
  reusing source's plain index-increment.
- **Status screen** (`APP._statusHtml()`) — source has no separate
  status/landing screen for Daily Pulse (it renders straight into the
  dropdown-and-form view); MAC-Walkthrough's `renderTeacherDashboard` is an
  *admin* view intentionally excluded here. Built new, following the brief's
  mockup, from data already loaded for the entry flow.

## New Files

`config.js` (trimmed/renamed from source's), `demo.js`, `styles.css`,
`manifest.webmanifest`, `sw.js`, `index.html` (new shell, not copied),
`docs/*.md`, `README.md`, `.gitignore`, `icons/*` (derived from the source's
`assets/iu29-logo.png`, copied then resized with `sips` — the original
source file was only read/copied, never modified).

## New Standalone Workflow

Sign in (or Demo Mode) → status screen (`Good morning/afternoon/evening,
<Teacher>`, today's date, students/complete/remaining counts, Today's
Submissions if any exist yet) → **Continue Daily Pulse** selects the first
incomplete student → sequential entry screen (Present/Absent → day rating →
optional focus areas → optional quick note → **Submit & Next**, or Absent →
confirm → next) → repeats until none remain → completion screen. Reopening
the app at any point re-derives status from the same backend query and never
restarts a completed student. An authorized administrator first selects a
teacher, uses the same status and entry flow, and can choose another teacher
without signing out.

## Teacher Identification

`AUTH.loadPilotUserFromSharePoint()` resolves the signed-in Microsoft account
against `IEP_Users2`. Active roles `"teacher"` and `"administrator"` are
admitted; inactive users and every other role remain denied. Teachers become
their own effective teacher automatically. Administrators must select an
effective teacher from the distinct teacher names present on active,
Daily-Pulse-enabled roster rows before student data loads.

The signed-in actor and effective teacher are intentionally separate. The
effective teacher controls roster filtering, completion lookup, and the saved
`Teacher` field. The signed-in actor controls `Submitted By` (and the in-memory
`SubmittedByEmail` value), preserving source reporting semantics.

## Student Filtering

`STUDENT_ROSTER.getDailyPulseForTeacher(teacherName)` — unchanged: Active
(default true) ∧ Daily Pulse Enabled (default **false**) ∧ case/whitespace-
insensitive exact match on the roster row's `Teacher` display name.

## Today's Completion Detection

`GRAPH.getTodayDailyPulseEntries(teacherName, dateStr)` (new, see Adapted
Modules) reads `IEP_Daily_Pulse` directly — this is the actual production
backend, not a local cache, so it survives reloads across devices/browsers.
Demo Mode's equivalent (`DEMO.getTodayCompleted`) reads the same shape from
`localStorage` under a `dailyPulseApp_demo*` key.

## Daily Pulse Data Mapping

Identical to source (`GRAPH.saveDailyPulse`, list `IEP_Daily_Pulse`):

| Display field | Value |
|---|---|
| Record ID | stable per-draft UUID (`pulse-<uuid>`), duplicate-detection key |
| Date | today, `YYYY-MM-DD` |
| Teacher | effective teacher (authenticated teacher or administrator selection) |
| Student | selected student's display name |
| Attendance Status | `"Present"` / `"Absent"` |
| Present | boolean |
| Status | `"Great Day"` / `"Some Struggles"` / `"Significant Concern"` / `""` (absent) |
| Category | comma-joined selected focus-area keys |
| Category Responses | `JSON.stringify(...)` of all focus-area answers |
| Academic Progress / Behavior / Communication / Social Skills / Self-Regulation / Transition / IEP Goal Rating | each scale's `rating`, else `""` |
| Quick Note | free text, `""` if absent |
| Submitted By / Submitted At | actual signed-in user's name / ISO timestamp |

## SharePoint Compatibility

Same list (`IEP_Daily_Pulse`), same site, same dynamic display-name→internal
schema mapper, same duplicate-detection field (`Record ID`). MAC-Walkthrough
requires zero changes to read entries this app creates.

## Save / Duplicate Protection

Identical contract to source: a stable per-draft `Record ID` is checked
against the list before create; a module-level in-flight flag blocks a
second concurrent submit; on failure the local copy is kept, the draft
(including `Record ID`) is left untouched for an idempotent retry, and
nothing advances until a save is confirmed. Verified live — see
[demo-verification.md](demo-verification.md).

## Features Intentionally Excluded

Dashboard, Reports, Walkthrough (V1/V2), PACE Log/PACE admin, Form Lab,
Setup, Storage Settings, admin "view as teacher", Student Check-In,
analytics/history beyond today, notification configuration — none of this
code was extracted.

## Demo Verification

See [demo-verification.md](demo-verification.md) — full functional pass plus
confirmed zero network requests to any Microsoft-owned domain (not just zero
successful Graph/SharePoint calls).

## Device Verification

Verified in the Claude Browser pane at desktop (1280×720-class) and mobile
(375×812) viewports: no horizontal scroll, single-column stacking below
420px, ≥44px touch targets throughout, clear focus states, adequate
contrast. iPad/Chromebook-class viewports were not independently tested
(no such device available in this environment) — MANUAL VERIFICATION
REQUIRED on an actual iPad/Chromebook, though the CSS uses the same
relative/flex/grid rules validated at both tested sizes.

## PWA / Cache Isolation

`sw.js` uses its own cache name (`daily-pulse-cache-v2`), only ever
`caches.match`/`caches.put`s same-origin GET requests, and explicitly bypasses
(never intercepts) any request to a `graph.microsoft.com`,
`login.microsoftonline.com`, or `microsoftonline.com` host. MAC-Walkthrough
has no service worker at all, so there is nothing to collide with. **Service
worker registration itself could not be confirmed in this sandboxed preview
browser** (`navigator.serviceWorker.register` failed with a generic sandbox
error on every attempt, before ever reaching the network) — code was
reviewed and is syntactically valid (`node --check sw.js` passes); actual
registration/offline/update behavior needs MANUAL VERIFICATION REQUIRED in a
normal browser.

## Entra / Redirect URI Requirements

See [entra-redirect-uris.md](entra-redirect-uris.md). Reuses MAC-Walkthrough's
existing Entra app registration (no new app registered). Add
`http://localhost:5173` (or whatever port you standardize local dev on) and
`https://daily-pulse.vercel.app` as SPA redirect URIs. **I did not modify
Entra** — this needs to be done by a tenant administrator.

## Manual Actions Required

- Add the redirect URIs above to the Entra SPA registration.
- Sign-in flow, unauthorized-account handling, live roster filtering
  (Active/Daily-Pulse-Enabled/teacher match against real `IEP_Students_2026_27`
  data), and live today's-completion lookup against real `IEP_Daily_Pulse`
  data all need an authorized Microsoft session to verify — untestable in
  this environment. Everything downstream of a successful sign-in (entry
  flow, validation, submit, duplicate/failure protection, completion,
  resume) was verified end-to-end in Demo Mode using the *same* code path.
- Service worker registration/update behavior in a real (non-sandboxed)
  browser.
- Physical iPad/Chromebook check.
- Confirm `IEP_Daily_Pulse` actually has all the display columns listed
  above — `mapFields()` silently skips (with a console warning) any column
  that doesn't exist, so a missing column degrades gracefully but should
  still be checked once against the real list.

## Known Limitations

- Two bugs were found and fixed during local verification: (1) `_render()`
  unconditionally hid the login screen and had no branch for it, so the app
  got stuck on "Loading Daily Pulse…" before sign-in; (2) `styles.css`'s
  header comment contained a stray `*/` that closed the comment early and
  corrupted the rest of the stylesheet parse, so no CSS custom properties
  (colors, etc.) applied. Both are fixed in the current files and verified
  visually.
- The administrator directory is intentionally limited to teacher values that
  currently have at least one active, Daily-Pulse-enabled roster row. It is not
  a teacher-management directory and does not invent or hardcode names.
- `getTodayDailyPulseEntries` reads the whole `IEP_Daily_Pulse` list
  (paginated) and filters client-side, same pattern as source's admin
  history readers — fine at current record volumes; if that list grows very
  large, a narrower server-side query would be worth revisiting.

## Suggested Vercel Project Name

`daily-pulse`

## Local Launch Commands

```bash
cd ~/tries/"Daily Pulse"
python3 -m http.server 5173
# open http://localhost:5173
```

(A server is already running from this session at the same address; the
command above is what you'd run again after closing it.)

## Commit Commands

```bash
cd ~/tries/"Daily Pulse"
git init
git add -A
git commit -m "Extract standalone Daily Pulse teacher app from MAC-Walkthrough

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

No push — no remote configured, and none should be added without your
explicit say-so.

## Deployment Commands

Not run. When you're ready:

```bash
cd ~/tries/"Daily Pulse"
npx vercel link      # choose/create the "daily-pulse" project
npx vercel           # preview deploy — sanity-check it first
npx vercel --prod    # production deploy
```

Register the resulting origin(s) as Entra SPA redirect URIs first (see
above) or sign-in will fail on the new origin.

## Source Project Integrity

**MAC-Walkthrough (`~/Projects/01-IU29/MAC-Walkthrough`) was not modified.**
`git status` shows only the same single pre-existing uncommitted change that
was already there before this session started (a `.gitignore` line adding
`.env*`, documented in this repo's own prior `docs/source-extraction-audit.md`
in the sibling `Admin Walkthrough` project) — nothing from this session
altered it. Every read of that repo in this session used read-only tools
(`Read`, `grep`/`find` via `Bash`); no `Write`/`Edit` call ever targeted it,
and `.env.local` was only checked for variable **names**, never opened for
its values.
