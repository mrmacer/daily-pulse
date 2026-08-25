# Entra SPA Redirect URIs

Daily Pulse reuses MAC-Walkthrough's existing Entra app registration
(`clientId: 145a3fc7-5cff-4d03-96c7-577e17980110`, tenant
`3276761c-22db-462b-a930-172d155bd795` — see [auth.js](../auth.js)). It does
**not** register a new Entra app.

`MSAL_CONFIG.auth.redirectUri` is `window.location.origin` — the same
pattern MAC-Walkthrough uses. That means the redirect URI is always just the
**origin** (scheme + host + port), never a path or trailing slash. Whatever
origin this app is served from must be added to the Entra app's **Single-page
application** platform redirect URIs (MSAL 2.x redirect-flow requires the SPA
platform, not Web).

## URIs to register

| Environment | Redirect URI to add |
|---|---|
| Local dev (this session's server) | `http://localhost:5173` |
| Future Vercel production (suggested project name `daily-pulse`) | `https://daily-pulse.vercel.app` |
| Any Vercel preview deployment you want to sign in from | `https://<preview-url>.vercel.app` (each unique preview URL needs its own entry, or test auth against production/a stable preview alias only) |

If you standardize local dev on a different port, register that exact
`http://localhost:<port>` instead/as well — there's nothing else
port-specific to change in code.

**I did not modify Entra.** This is a request for you (or whoever
administers the `IU29`/`IEP_Skook` Entra tenant) to add these redirect URIs
to the existing app registration's SPA platform before sign-in will work
from a new origin.
