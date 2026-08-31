/* ─────────────────────────────────────────────────────────────────────────
   Identity diagnostic patch — safe login/access-denial debugging block

   Covers loginDiagnosticHtml() directly, plus a full APP.init() E2E run
   (real auth path, DOM actually rendered — unlike patch-b-authorization's
   harness, _render() is NOT stubbed out here) proving:
     - an authenticated-but-unmatched-in-IEP_Users2 account sees the
       diagnostic (Identity record: Not Found) on the login screen;
     - a hostile display name is escaped, not executed;
     - no token/Graph-payload text ever lands in the rendered HTML;
     - the authorization decision itself is unchanged (still denied);
     - a fully valid Teacher sees no denial screen and no diagnostic block
       at all.
   ───────────────────────────────────────────────────────────────────────── */

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const ROOT = path.resolve(__dirname, "..");

function runFile(context, filename) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, filename), "utf8"), context, { filename });
}

function makeDomElement() {
  return {
    _hidden: true,
    textContent: "",
    innerHTML: "",
    classList: {
      _hidden: true,
      add(cls) { if (cls === "hidden") this._hidden = true; },
      remove(cls) { if (cls === "hidden") this._hidden = false; },
      contains(cls) { return cls === "hidden" ? this._hidden : false; }
    },
    querySelector() { return null; },
    querySelectorAll() { return []; },
    addEventListener() {},
    scrollIntoView() {},
    appendChild() {}
  };
}

function makeContext(overrides = {}) {
  const elements = {};
  const ids = ["loginScreen", "loginError", "loginDiagnostic", "main", "signOutBtn", "demoBanner"];
  ids.forEach(id => { elements[id] = makeDomElement(); });

  const document = {
    addEventListener() {},
    getElementById(id) { return elements[id] || null; },
    querySelector() { return null; },
    querySelectorAll() { return []; },
    createElement() { return makeDomElement(); },
    head: { appendChild() {} }
  };
  const context = vm.createContext({
    console,
    structuredClone,
    crypto: crypto.webcrypto,
    setTimeout,
    clearTimeout,
    document,
    location: { reload() {} },
    navigator: {},
    ...overrides
  });
  context.window = context;
  context.window.location = { origin: "http://localhost:5173" };
  context.window.confirm = () => true;
  context.slugifyTeacherId = value => String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  context._elements = elements;
  return context;
}

function loadEscHtmlAndDiagnostic(context) {
  // loginDiagnosticHtml() and escHtml() are plain top-level functions in
  // app.js with no DOM dependency at parse time — but app.js as a whole
  // needs AUTH to exist as a global before any function referencing AUTH is
  // *called* (not merely declared), which is fine since we only call
  // loginDiagnosticHtml() after setting up AUTH below.
  runFile(context, "app.js");
}

/* ── loginDiagnosticHtml() in isolation ─────────────────────────────── */

test("diagnostic shows Not Found + signed-in identity when AUTH.account exists but is unauthorized", () => {
  const context = makeContext({
    AUTH: { account: { name: "Test Teacher", username: "test.teacher@iu29.org" }, isUnauthorized: true, displayName: "Test Teacher" }
  });
  loadEscHtmlAndDiagnostic(context);
  const html = vm.runInContext("loginDiagnosticHtml", context)({ identityFound: "Not Found" });
  assert.match(html, /Signed in as:.*<strong>Test Teacher<\/strong>/s);
  assert.match(html, /Account:.*<strong>test\.teacher@iu29\.org<\/strong>/s);
  assert.match(html, /Identity record:.*<strong>Not Found<\/strong>/s);
  assert.doesNotMatch(html, /App permission/);
});

test("diagnostic can show Found + a permission decision when both are known", () => {
  const context = makeContext({
    AUTH: { account: { name: "Test Admin", username: "test.admin@iu29.org" }, isUnauthorized: false, displayName: "Test Admin" }
  });
  loadEscHtmlAndDiagnostic(context);
  const html = vm.runInContext("loginDiagnosticHtml", context)({ identityFound: "Found", appPermission: "Denied" });
  assert.match(html, /Identity record:.*<strong>Found<\/strong>/s);
  assert.match(html, /App permission:.*<strong>Denied<\/strong>/s);
});

test("a hostile display name is escaped, never executed", () => {
  const context = makeContext({
    AUTH: { account: { name: "<script>alert(1)</script>", username: "attacker@iu29.org" }, isUnauthorized: true, displayName: "" }
  });
  loadEscHtmlAndDiagnostic(context);
  const html = vm.runInContext("loginDiagnosticHtml", context)({ identityFound: "Not Found" });
  assert.doesNotMatch(html, /<script>/);
  assert.match(html, /&lt;script&gt;alert\(1\)&lt;\/script&gt;/);
});

test("the diagnostic never contains anything resembling a token or Graph payload", () => {
  const context = makeContext({
    AUTH: { account: { name: "Test Teacher", username: "test.teacher@iu29.org" }, isUnauthorized: true, displayName: "" }
  });
  loadEscHtmlAndDiagnostic(context);
  const html = vm.runInContext("loginDiagnosticHtml", context)({ identityFound: "Not Found" });
  for (const forbidden of ["accessToken", "Bearer ", "@odata", "idToken", "SharePoint list", "graph.microsoft.com"]) {
    assert.equal(html.includes(forbidden), false, `diagnostic must never contain "${forbidden}"`);
  }
});

/* ── End-to-end APP.init() (real render path) ───────────────────────── */

function appInitHarness({ appUsersRows = [], iepUsersRow, account }) {
  const context = makeContext({
    APP_MODE: undefined,
    msal: { PublicClientApplication: class { async handleRedirectPromise() { return null; } getAllAccounts() { return [account]; } } },
    GRAPH: {
      async findUserByEmail(email) {
        if (!iepUsersRow) throw new Error("No matching IEP_Users2 Email found for the signed-in account.");
        return iepUsersRow;
      },
      async getListItemsByDisplayName(listName) {
        if (listName === "IEP_App_Users") return appUsersRows;
        return [];
      },
      async getTodayDailyPulseEntries() { return []; }
    },
    DEMO: { isActive: () => false },
    DB: { addDailyPulse: v => v, markSynced() {} },
    STUDENT_ROSTER: {
      loaded: true, error: null, _students: [],
      async refresh() {},
      getDailyPulseForTeacher() { return []; },
      getDailyPulseTeachers() { return []; }
    }
  });
  context.window.msal = context.msal;
  runFile(context, "config.js");
  runFile(context, "auth.js");
  runFile(context, "iep-app-users.js");
  runFile(context, "access.js");
  context.DAILY_PULSE_ACCESS = vm.runInContext("DAILY_PULSE_ACCESS", context);
  runFile(context, "app.js");
  const app = vm.runInContext("APP", context);
  app._loadMsal = () => Promise.resolve();
  const auth = vm.runInContext("AUTH", context);
  auth.account = account;
  return { app, auth, context };
}

test("E2E: authenticated account with no matching IEP_Users2 row shows the diagnostic with Not Found", async () => {
  const account = { username: "ghost@iu29.org", name: "Ghost Account" };
  const { app, context } = appInitHarness({ account, iepUsersRow: null });
  await app.init();

  const loginScreenEl = context._elements.loginScreen;
  const loginErrorEl = context._elements.loginError;
  const diagnosticEl = context._elements.loginDiagnostic;

  assert.equal(loginScreenEl.classList._hidden, false, "login screen must be visible");
  assert.equal(loginErrorEl.classList._hidden, false, "the human-friendly error must be visible");
  assert.match(loginErrorEl.textContent, /No matching IEP_Users2 Email/);

  assert.equal(diagnosticEl.classList._hidden, false, "diagnostic block must be visible");
  assert.match(diagnosticEl.innerHTML, /Signed in as:.*<strong>Ghost Account<\/strong>/s);
  assert.match(diagnosticEl.innerHTML, /Account:.*<strong>ghost@iu29\.org<\/strong>/s);
  assert.match(diagnosticEl.innerHTML, /Identity record:.*<strong>Not Found<\/strong>/s);
});

test("E2E: authorization decision remains denied — no access is granted alongside the diagnostic", async () => {
  const account = { username: "ghost@iu29.org", name: "Ghost Account" };
  const { app, auth } = appInitHarness({ account, iepUsersRow: null });
  await app.init();
  assert.equal(auth.isUnauthorized, true);
  assert.equal(auth.isAuthenticated, false);
  assert.notEqual(app.screen, "status");
  assert.notEqual(app.screen, "entry");
});

test("E2E: a valid, authorized Teacher sees no denial screen and no diagnostic block at all", async () => {
  const account = { username: "khiggins@iu29.org", name: "Kelly Higgins" };
  const { app, context } = appInitHarness({
    account,
    iepUsersRow: { Title: "u1", field_1: "Kelly Higgins", field_2: "Teacher", field_3: true, Email: account.username },
    appUsersRows: [{ id: "1", Email: account.username, "Daily Pulse": "Yes" }]
  });
  await app.init();
  assert.notEqual(app.screen, "denied");
  const diagnosticEl = context._elements.loginDiagnostic;
  assert.equal(diagnosticEl.classList._hidden, true, "an authorized user must never see the diagnostic block");
  assert.equal(diagnosticEl.innerHTML, "");
});

test("E2E: authentication-service-failed case (no account yet) never shows the diagnostic", async () => {
  const context = makeContext({
    APP_MODE: undefined,
    GRAPH: {},
    DEMO: { isActive: () => false }
  });
  runFile(context, "config.js");
  runFile(context, "auth.js");
  runFile(context, "iep-app-users.js");
  runFile(context, "access.js");
  context.DAILY_PULSE_ACCESS = vm.runInContext("DAILY_PULSE_ACCESS", context);
  runFile(context, "app.js");
  const app = vm.runInContext("APP", context);
  app._loadMsal = () => Promise.reject(new Error("Could not load the Microsoft sign-in library."));
  await app.init();

  const diagnosticEl = context._elements.loginDiagnostic;
  assert.equal(diagnosticEl.classList._hidden, true, "with no account at all, the diagnostic must stay hidden");
});
