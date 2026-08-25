# Demo Mode Verification

Tested locally at `http://localhost:5173` via the Claude Browser pane on
2026-08-25. Full run: sign-in screen → Try Demo Mode → status screen →
5 fake students (Riley Fox, Sam Ortiz, Jordan Blake, Casey Nguyen, Avery
Patel) → Present + Great Day, Absent, Present + Some Struggles + a Focus
Area, Present + Great Day (with a simulated save failure + retry), Present +
Great Day → all-complete screen → reload/resume → mobile layout.

## Network isolation

- `MSAL` is not even loaded from the CDN in Demo Mode — [index.html](../index.html)
  no longer has a static `<script src=".../msal-browser...">` tag; `app.js`
  (`_loadMsal()`) injects that script tag only inside the real-auth branch of
  `init()`, which Demo Mode never reaches.
- `AUTH.init()`, `STUDENT_ROSTER.refresh()`, and every `GRAPH.*` method are
  never called on the Demo Mode code path — `app.js` branches to the `DEMO`
  facade before any of them would run (`init()`, `_loadRosterAndStatus()`,
  `_save()`).
- Checked via the browser's network log (`mcp__Claude_Browser__read_network_requests`)
  across the entire demo session, filtered for `graph.microsoft|microsoftonline|jsdelivr`:
  **zero requests recorded** — not just zero *successful* requests, zero
  requests attempted at all.
- All demo data lives in `localStorage`/`sessionStorage` under
  `dailyPulseApp_demo*` keys (see [config.js](../config.js)) and is
  fabricated in [demo.js](../demo.js) — no real student or teacher data.

## Functional checks (all passed)

| Behavior | Result |
|---|---|
| Status screen shows correct counts (5 students / 0 complete / 5 remaining → updates live) | ✅ |
| Continue selects the first incomplete student | ✅ |
| Present + Great Day submits, toasts, advances to next student | ✅ |
| Absent confirm dialog, skips day-rating/focus fields entirely | ✅ |
| Overall Day required — blocked with inline error when missing | ✅ |
| Expanding a Focus Area makes its rating required — blocked with inline error when missing | ✅ |
| Multi-select chips and conditional "Other" text field render/toggle correctly | ✅ |
| Rapid double-click on Submit creates exactly **one** entry (duplicate-tap guard) | ✅ |
| Simulated save failure: draft preserved (`overallStatus` intact), error shown, no advance, retry succeeds with no duplicate | ✅ |
| All 5 complete → "Daily Pulse Complete" screen, correct "Today's Submissions" list (including the Absent entry, labeled "Absent" not a status) | ✅ |
| Reload after completing all students → still shows the complete state (does not restart) | ✅ |
| Mobile viewport (375×812) → single-column layout, no horizontal scroll, large touch targets | ✅ |

Two real bugs were found and fixed during this verification pass — see
[daily-pulse-extraction-report.md](daily-pulse-extraction-report.md) "Known
Limitations" for what they were.
