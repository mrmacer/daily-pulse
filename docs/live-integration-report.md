# DAILY PULSE LIVE INTEGRATION REPORT

Verification date: 2026-08-25. Local server: `python3 -m http.server 5173` at
`http://localhost:5173`, source MAC-Walkthrough at
`~/Projects/01-IU29/MAC-Walkthrough` (read-only).

## Microsoft Authentication
**PARTIAL PASS / MANUAL ACTION REQUIRED.** Config identifiers verified
identical to source by direct diff (clientId, tenant authority, SharePoint
site — see below). Clicked **Sign in with Microsoft** in a live browser: MSAL
redirected to `login.microsoftonline.com` and rendered a normal, IU29-branded
sign-in prompt (not an `AADSTS50011` redirect-URI-mismatch error) — this
confirms `http://localhost:5173` is **already registered** as a valid
redirect URI for this app registration. I stopped there and entered no
credentials — I have no IU29 account and should not attempt to obtain or use
one. Completing an actual sign-in (to verify `IEP_Users2` lookup, roster,
today's-completion, and a real submission) needs you to sign in yourself;
see "Manual Actions Required" below for exactly how.

## IEP_Users2 Authorization
**MANUAL ACTION REQUIRED** — same reason: needs a real authenticated
session. **Verified by code inspection instead**: `APP.init()`
(`app.js:62-95`) has an explicit early `return` after each of three gates —
no account → login screen; `AUTH.isUnauthorized` (no `IEP_Users2` match) →
denial message; `!AUTH.isTeacher` → `"denied"` screen. `_loadRosterAndStatus()`
— the only method that calls `STUDENT_ROSTER.refresh()` or any `GRAPH.*`
roster/pulse method — is reached only after all three gates pass. There is
no code path that reaches roster or pulse data without a matching, active,
`role = "teacher"` `IEP_Users2` record.

## Teacher Identification
Verified by code inspection: `_resolveTeacherIdentity()` builds
`{ id, name, teacherName, email }` entirely from the authenticated
`IEP_Users2` record (`AUTH.pilotUser.Name`, `AUTH.pilotUserId`,
`AUTH.account.username`) — no hardcoded teacher list is consulted. Cannot be
exercised end-to-end without a real signed-in session (see above).

## Roster Filtering
Verified by code inspection (identical to source, see
[daily-pulse-source-audit.md](daily-pulse-source-audit.md) §8-10):
`Active` (default true) ∧ `Daily Pulse Enabled` (default **false**) ∧
case/whitespace-insensitive `Teacher` name match. No actual `IEP_Students_2026_27`
rows were fetched in this session — doing so requires a real authenticated
Graph token. **MANUAL ACTION REQUIRED** to confirm live counts; no student
data can be reported here because none was retrieved.

## Today's Completion Lookup
**PASS (Demo Mode, identical code path) / MANUAL ACTION REQUIRED (live data).**
`GRAPH.getTodayDailyPulseEntries()` is new code (see extraction report) —
it hadn't been exercised end-to-end before this pass. I re-verified the
*state-management* contract it feeds into is correct: submitting a student in
Demo Mode adds them to `completedToday` in memory, and reloading a fully
completed demo session (5/5) correctly returned to the completion screen
without restarting anyone. The live variant swaps in
`GRAPH.getTodayDailyPulseEntries(teacherName, date)`, which reuses the same
already-proven `getListItemsByDisplayName` reader `dashboard-sync.js` uses
for its own live Daily Pulse stats (see MAC-Walkthrough Compatibility below)
— structurally sound, but the actual live SharePoint round-trip is untested.

## Controlled SharePoint Submission
**NOT PERFORMED.** Requires a real authenticated session; I have none and
should not seek one. Did not create, and could not have created, any
production `IEP_Daily_Pulse` record this session.

## Exact SharePoint Compatibility
**PASS**, verified by code inspection of MAC-Walkthrough's own live-read
path (not just its writer): `dashboard-sync.js`'s `loadDailyPulseStats()`
reads `IEP_Daily_Pulse` live from SharePoint and filters/aggregates on the
display fields `"Student"`, `"Date"`, `"Status"`, `"Category"` — exactly the
display fields both source's `saveDailyPulse()` and Daily Pulse's
`GRAPH.saveDailyPulse()` write. No translation layer needed. (Note,
unrelated to this extraction: that same function also reads a `"Support
Level"` field that *neither* source's writer nor Daily Pulse's writer ever
populates — a pre-existing gap in MAC-Walkthrough itself, not something this
extraction introduced or should fix.)

One MAC-Walkthrough admin view will **not** show entries from any Daily
Pulse source, including its own: `renderTeacherDashboard` ("My Classroom
Dashboard", reached only via the admin's "view as teacher" feature) reads
the local-only `DB.getDailyPulses()` and matches students against the
hardcoded `PILOT_TEACHERS` list — it's a legacy/pilot-only view, already
disconnected from live data before this extraction existed. The main admin
Dashboard (`dashboard-sync.js`) is the live one and is fully compatible.

## Resume After Reload
**PASS (Demo Mode)** — see Today's Completion Lookup above. Same code path
as production; live-data confirmation is part of the same MANUAL ACTION
REQUIRED item.

## Privacy / Security Check
- **No secrets committed**: the project is not yet a git repo (`git init`
  hasn't been run), so nothing has been committed at all. No `.env` file of
  any kind exists in the project. `clientId`/tenant authority/SharePoint
  site are public, client-side MSAL identifiers (same ones already visible
  in MAC-Walkthrough's own shipped source) — not secrets.
- **`.gitignore` covers `.env`/`.env.local`**: confirmed present.
- **No production roster in source**: confirmed — `PILOT_TEACHERS`/
  `PILOT_STUDENTS`/`TEACHER_IDENTITIES` were deliberately not extracted (see
  extraction report "Components that must NOT be extracted").
- **Demo Mode is fake-only**: `demo.js` contains 5 fictional students
  (Riley Fox, Sam Ortiz, Jordan Blake, Casey Nguyen, Avery Patel) and one
  fictional teacher.
- **Service worker never caches Microsoft traffic**: `sw.js`'s
  `isNetworkOnly()` bypasses (never `caches.match`/`caches.put`) any request
  to `graph.microsoft.com`, `login.microsoftonline.com`, or
  `microsoftonline.com`.
- **Token storage**: MSAL's `cacheLocation: "localStorage"` is unchanged
  from source — same risk posture as the app you already run in production,
  not something this extraction altered.
- **Demo/production isolation — one real gap found and fixed this pass**:
  `_save()` was unconditionally writing every submission (including demo
  ones) through the production-only local-history helper (`DB.addDailyPulse`,
  key `dailyPulseApp_localHistory`), which would have mixed fake demo
  student names into the same storage key a real teacher's failed-save
  resilience layer uses. Fixed to skip that call entirely in Demo Mode
  (`app.js`) and re-verified live: after a demo submission,
  `dailyPulseApp_localHistory` stayed `[]` while only
  `dailyPulseApp_demoPulses` received the entry.

## Entra Redirect URI
`http://localhost:5173` — **confirmed already registered** (Entra returned a
normal sign-in prompt, not a redirect-URI-mismatch error). The production
URI (`https://daily-pulse.vercel.app`, once deployed) still needs to be
added before that origin can sign in — see
[entra-redirect-uris.md](entra-redirect-uris.md). I did not modify Entra.

## Issues Found
1. **Fixed this pass**: Demo Mode submissions were writing into the
   production local-history storage key (see Privacy/Security Check above).
2. **Informational, not a Daily Pulse issue**: MAC-Walkthrough's own "My
   Classroom Dashboard" admin view is pre-existing legacy/pilot-only and
   disconnected from live SharePoint data — unrelated to this extraction,
   not fixed (source is read-only), noted here only because it's relevant
   context for "will MAC-Walkthrough see these records."

## Changes Made
One change, entirely inside `~/tries/Daily Pulse`:
[app.js](../app.js) `_save()` — Demo Mode submissions now go only through
`DEMO.savePulse()` and skip the production `DB.addDailyPulse()`/
`GRAPH.saveDailyPulse()` calls entirely.

## Source Integrity
**`~/Projects/01-IU29/MAC-Walkthrough` was not modified.** `git status`
still shows only the same single pre-existing uncommitted `.gitignore` line
present before this session began — re-confirmed at the end of this pass.

## Deployment Readiness

**NOT READY FOR DEPLOYMENT**

Blocking items (all require your direct participation — real IU29
credentials, which I don't have and shouldn't seek):
- Complete an actual Microsoft sign-in and confirm `IEP_Users2`
  authorization resolves your account as a teacher.
- Confirm your live roster (Active ∧ Daily Pulse Enabled ∧ your name on
  `IEP_Students_2026_27`) comes back as expected.
- Perform one controlled real submission (ideally against a designated
  pilot/test student) and confirm it appears once in `IEP_Daily_Pulse` and
  is picked up as "completed" on reload.
- Add `https://daily-pulse.vercel.app` as an Entra SPA redirect URI before
  that origin is used (localhost is already fine).

Nothing else is blocking — everything reachable without a real Microsoft
session (config parity, redirect URI, authorization code paths, SharePoint
field/compatibility structure, demo/production isolation) checked out.
