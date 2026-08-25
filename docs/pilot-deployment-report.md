# DAILY PULSE PILOT DEPLOYMENT REPORT

Deployment date: 2026-08-25. Deployed from `~/tries/Daily Pulse` using the
Vercel CLI already authenticated on this machine (`gmacer01-4422`).

## Pre-deployment checks (all passed)

No secrets committed (project had no `.git` history until this pass); no
`.env*` files existed in source; `.gitignore` covers `.env`/`.env.local`/`.env*`;
no secret-pattern matches (`password`, `client_secret`, `api_key`, private
keys); no real student data or hardcoded roster in source (confirmed in the
extraction audit — `PILOT_TEACHERS`/`PILOT_STUDENTS` were never extracted);
Demo Mode contains only 5 fictional students; `clientId`/tenant
authority/SharePoint site/list names re-diffed identical to
MAC-Walkthrough; service worker cache name (`daily-pulse-cache-v1`) is
unique to this project; `sw.js` bypasses (never caches) any
`graph.microsoft.com`/`login.microsoftonline.com`/`microsoftonline.com`
request; no build step needed (plain static files — confirmed via a real
Vercel deploy, no framework detected, deployed as-is); only "localhost"
reference in application code is an explanatory comment in `auth.js`, not a
hardcoded URL (`redirectUri` is `window.location.origin`, resolved at
runtime).

## Git Status

Repository initialized in `~/tries/Daily Pulse` (did not previously exist).
Two commits:

| Commit | Message |
|---|---|
| `2d741c4832924b65c1f229401915ee8fc54d0cb4` | Initial standalone Daily Pulse teacher application |
| (see `git log`) | Broaden .gitignore for .env* (added automatically by `vercel link`) |

## Vercel Project

**Actual project name: `daily-pulse`** (the name you suggested was
available — no fallback needed). Scope: `gmacer01-4422s-projects` (your
existing personal account/team — same one MAC-Walkthrough's own Vercel
project lives under, but a **separate, independent project**: different
`projectId` — `prj_oUHrmWWaT5LGjOZK0WCBdGG5OCSe` for Daily Pulse vs.
MAC-Walkthrough's own; not attached to or sharing anything with it).

## Production URL

**`https://daily-pulse-blue.vercel.app`**

The exact short domain `daily-pulse.vercel.app` was already claimed by a
different, unrelated Vercel account (`.vercel.app` short names are global,
first-come-first-served) — Vercel auto-assigned `daily-pulse-blue.vercel.app`
instead. A second, longer alias also resolves to the same deployment:
`https://daily-pulse-gmacer01-4422s-projects.vercel.app`. This was the
project's first deployment, which Vercel always assigns directly to
**production** — there was no separate "promote" step to run.

If you'd prefer a cleaner URL later, a custom domain (e.g. one under
`iu29.org`, if you have DNS access) can be added to this same project at any
time without redeploying.

## Anonymous Production Verification

**PASS.** Checked in a real browser against the live URL: `/`,
`/manifest.webmanifest`, `/sw.js`, and `/app.js` all return `200`; login
screen renders correctly (IU29 logo, "Sign in with Microsoft" / "Try Demo
Mode").

## Demo Mode Production Verification

**PASS.** On the live production URL: entered Demo Mode, completed one fake
submission (Riley Fox, Present, Great Day) — advanced correctly, status
counts updated (5 students / 1 complete / 4 remaining), "Today's
Submissions" showed it. Confirmed `window.msal` stayed `"undefined"` the
entire time (the library is never even fetched in Demo Mode), and the
browser's network log showed **zero requests** matching
`graph.microsoft|microsoftonline|jsdelivr` across the whole session.
`dailyPulseApp_localHistory` (the production-only resilience store) stayed
empty while `dailyPulseApp_demoPulses` correctly received the one entry —
confirms the isolation fix from the prior verification pass holds on the
real deployment too. Demo state was reset and Demo Mode deactivated before
finishing.

## PWA Verification

**PASS** — and better than the earlier localhost check: on this real HTTPS
origin, the service worker actually **registered and went active**
(`scope: https://daily-pulse-blue.vercel.app/`), something the sandboxed
local preview browser couldn't do at all. Its cache
(`daily-pulse-cache-v1`) holds only the app shell (`index.html`, `/`, the
JS/CSS files, icons, manifest) — no Microsoft traffic, confirmed by
inspecting the cache contents directly. Manifest fetches `200` and parses
with the expected name and 3 icons.

## Entra Redirect URI Required

**`https://daily-pulse-blue.vercel.app`** — exact origin, no path or
trailing slash (same pattern as the already-working
`http://localhost:5173`). This has **not** been added to Entra yet — I did
not modify Entra, and a real sign-in attempt from this production URL will
fail (`AADSTS50011`) until it is. If you also want the longer alias to work
for sign-in, add `https://daily-pulse-gmacer01-4422s-projects.vercel.app`
too, though the shorter one above is the one to hand to teachers.

## Real Login Status

**MANUAL TEST REQUIRED.** I did not attempt a real Microsoft sign-in — per
your instruction, and because I have no IU29 credentials regardless. Add the
redirect URI above to Entra first, or "Sign in with Microsoft" will fail
with a redirect-URI-mismatch error.

## Pilot Teacher Test Plan

Give each pilot teacher this short list (5–10 minutes):

1. Open `https://daily-pulse-blue.vercel.app` on the device they'll
   actually use (iPad/Chromebook/laptop).
2. Sign in with their normal IU29 Microsoft account.
3. Confirm the status screen shows *their* name and a plausible student
   count for their classroom.
4. Tap **Continue Daily Pulse** — confirm it lands on a real student from
   their roster, not someone else's.
5. Complete one student as **Present** with a day rating, expand one Focus
   Area, add a Quick Note, submit — confirm it advances and the count
   updates.
6. Complete one student as **Absent** — confirm no day-rating/focus
   questions were asked.
7. Close the tab (or reload) and reopen — confirm the students they already
   did are *not* asked again, and the remaining count is correct.
8. If comfortable, try installing it (browser's "Add to Home Screen" /
   install-app prompt) and confirm it opens standalone.
9. Report anything confusing, slow, or wrong — especially anything that
   looks like the wrong student, wrong teacher, or a stuck "Saving…" button.

Ask them **not** to use Demo Mode for this pass (it's for your own
exploration, not pilot testing) and not to worry about "Today's Submissions"
being a permanent record — real corrections go through an administrator, as
the app itself says.

## Rollback

Nothing on the SharePoint/Entra/MAC-Walkthrough side needs undoing to roll
back this app — it's a static frontend with no destructive writes, so
rollback is entirely a Vercel-side action:

- **Fastest (seconds, no rebuild)**: `npx vercel rollback` from this
  directory (or the "Promote to Production" button on any prior deployment
  in the Vercel dashboard) instantly repoints the production alias at the
  last-known-good deployment.
- **If you just want it offline immediately**: remove the domain alias or
  pause the project from the Vercel dashboard (Project → Settings) — the
  code and every deployment stay intact, nothing is deleted.
- **If a bad commit needs to be un-shipped for good**: `git revert <bad
  commit>` (or `git reset --hard <last-good commit>`) in this repo, then
  `npx vercel --prod` to ship the reverted state.

Since there is only one deployment so far, there's nothing to roll back to
yet — this note is for after a second deployment exists.

## Source Integrity

**`~/Projects/01-IU29/MAC-Walkthrough` was not modified.** Re-confirmed via
`git status` at the end of this pass: only the same single pre-existing
uncommitted `.gitignore` line present before any of this work began.
