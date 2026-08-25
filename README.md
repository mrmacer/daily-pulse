# Daily Pulse

Standalone teacher app for daily student check-ins, extracted from
[MAC-Walkthrough](../../../Projects/01-IU29/MAC-Walkthrough)'s Daily Pulse
feature. Plain HTML/CSS/JS — no build step, no framework, no bundler.

See [docs/daily-pulse-source-audit.md](docs/daily-pulse-source-audit.md) and
[docs/daily-pulse-extraction-report.md](docs/daily-pulse-extraction-report.md)
for the full extraction writeup.

## Run locally

Any static file server works. From this directory:

```bash
python3 -m http.server 5173
```

Then open <http://localhost:5173>. Use **Try Demo Mode** to explore the app
with fake data and zero network calls, or **Sign in with Microsoft** for the
real IU29/IEP Skook flow (requires the local origin to be registered as an
Entra SPA redirect URI first — see
[docs/entra-redirect-uris.md](docs/entra-redirect-uris.md)).

## Deploy (not yet done — see extraction report)

```bash
npx vercel link          # first time only, creates/links the "daily-pulse" project
npx vercel                # preview deploy
npx vercel --prod         # production deploy
```

## Project layout

- `index.html` — shell (login screen + app root)
- `config.js` — SharePoint list names, localStorage key names
- `auth.js` — MSAL + IEP_Users2 authorization (extracted from source)
- `graph.js` — Microsoft Graph/SharePoint transport + Daily Pulse save (extracted + adapted)
- `roster.js` — live student roster provider + the 7 Daily Pulse focus-area scales (extracted)
- `data.js` — local resilience/history layer (adapted)
- `demo.js` — Demo Mode fixtures + fake save (new)
- `app.js` — status/entry screens, state, submit orchestration (extracted UI logic, restructured flow)
- `styles.css` — new stylesheet
- `manifest.webmanifest`, `sw.js`, `icons/` — PWA layer (new)
