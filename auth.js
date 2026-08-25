/* ── Authentication (extracted verbatim from MAC-Walkthrough auth.js) ───────
 * Same Entra app registration, same IEP_Users2 authorization contract, so
 * an authorized user's existing organizational account works here with no separate
 * login system and no new secrets. See docs/daily-pulse-source-audit.md §5-6.
 *
 * redirectUri is window.location.origin (no path), so this works unchanged
 * on localhost during dev and on the eventual Vercel origin — each origin
 * just needs to be added as an Entra SPA redirect URI. See
 * docs/entra-redirect-uris.md for the exact values to register.
 */
const MSAL_CONFIG = {
  auth: {
    clientId: "145a3fc7-5cff-4d03-96c7-577e17980110",
    authority: "https://login.microsoftonline.com/3276761c-22db-462b-a930-172d155bd795",
    redirectUri: window.location.origin
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

function normalizeUserRole(value) {
  const raw = Array.isArray(value) ? value[0] : value;
  return String(raw || "").trim().toLowerCase();
}

const AUTH = {
  _client:     null,
  account:     null,
  pilotUserId: null,
  pilotUser:   null,
  role:        null,
  lookupError: null,

  async init() {
    this._client = new msal.PublicClientApplication(MSAL_CONFIG);

    try {
      const response = await this._client.handleRedirectPromise();
      if (response?.account) {
        this.account = response.account;
      } else {
        const accounts = this._client.getAllAccounts();
        if (accounts.length > 0) this.account = accounts[0];
      }
    } catch (err) {
      console.error("MSAL error:", err);
    }

    if (this.account) {
      await this.loadPilotUserFromSharePoint();
    }
  },

  async loadPilotUserFromSharePoint() {
    const email = (this.account.username || "").toLowerCase().trim();
    try {
      const user = await GRAPH.findUserByEmail(email);
      if (!user) {
        this.pilotUserId = null;
        this.pilotUser   = null;
        this.role        = null;
        return null;
      }

      const userId  = user.Title || user.UserID || user.User_x0020_ID || null;
      const name    = user.field_1 || user.Name || user.Title || "";
      const rawRole = user.field_2 || user.Role || "";
      const role    = Array.isArray(rawRole) ? rawRole[0] : rawRole;
      const active  = user.field_3 ?? user.Active;

      const isActive =
        active === true   ||
        active === "Yes"  ||
        active === "true" ||
        active === 1;

      if (!isActive) {
        this.pilotUserId = null;
        this.pilotUser   = null;
        this.role        = null;
        this.lookupError = "Your IEP Skook account is inactive.";
        return null;
      }

      this.pilotUserId  = userId;
      this.role         = role;
      this.lookupError  = null;
      this.pilotUser    = {
        ...user,
        UserID: userId,
        Name:   name,
        Role:   role,
        Active: isActive,
        Email:  user.Email
      };
      return this.pilotUser;
    } catch (err) {
      console.error("Failed to load pilot user from SharePoint:", err);
      this.lookupError = err.message || "Unable to verify your IEP Skook account.";
      return null;
    }
  },

  get isAuthenticated() {
    return this.account !== null && this.pilotUserId !== null;
  },

  get isUnauthorized() {
    return this.account !== null && this.pilotUserId === null;
  },

  get displayName() {
    return this.account?.name || this.account?.username || "";
  },

  async acquireGraphToken() {
    const request = {
      scopes: ["User.Read", "Sites.ReadWrite.All"],
      account: this.account
    };
    try {
      const resp = await this._client.acquireTokenSilent(request);
      return resp.accessToken;
    } catch (err) {
      await this._client.acquireTokenRedirect(request);
      return null;
    }
  },

  login() {
    this._client.loginRedirect({
      scopes: ["User.Read", "Sites.ReadWrite.All"]
    });
  },

  logout() {
    this._client.logoutRedirect({ account: this.account });
  },

  get isAdmin()   { return normalizeUserRole(this.role) === "administrator"; },
  get isTeacher() { return normalizeUserRole(this.role) === "teacher"; },
  get isDailyPulseAuthorized() {
    return this.isAuthenticated && (this.isTeacher || this.isAdmin);
  }
};
