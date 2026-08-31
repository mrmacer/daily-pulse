/* ─────────────────────────────────────────────────────────────────────────
   PATCH B — Daily Pulse authorization via IEP_App_Users

   Covers access.js's new decision helpers directly, plus an end-to-end
   APP.init() run (real auth path) exercising the full gate + teacher-roster
   resolution against a mocked GRAPH/AUTH.
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

function makeContext(overrides = {}) {
  const document = {
    addEventListener() {},
    getElementById() { return null; },
    querySelector() { return null; },
    createElement() { return {}; },
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
  return context;
}

function accessHarness() {
  const context = makeContext({ CONFIG: { APP_USERS_LIST: "IEP_App_Users" } });
  runFile(context, "iep-app-users.js");
  runFile(context, "access.js");
  return vm.runInContext("DAILY_PULSE_ACCESS", context);
}

/* ── resolveDailyPulseAccess ─────────────────────────────────────────── */

test("app-users row is authoritative: PACE-style explicit No overrides a supported role", () => {
  const access = accessHarness();
  const decision = access.resolveDailyPulseAccess("Teacher", { found: true, row: { "Daily Pulse": "No" } });
  assert.equal(decision.allowed, false);
  assert.equal(decision.reason, "app-users-denied");
});

test("app-users row Yes allows even a role the legacy check would deny", () => {
  const access = accessHarness();
  const decision = access.resolveDailyPulseAccess("Behavior Specialist", { found: true, row: { "Daily Pulse": "Yes" } });
  assert.equal(decision.allowed, true);
  assert.equal(decision.reason, "app-users-permission");
});

test("no matching row falls back to the legacy Teacher/Administrator check (migration mode)", () => {
  const access = accessHarness();
  assert.equal(access.resolveDailyPulseAccess("Teacher", { found: false }).allowed, true);
  assert.equal(access.resolveDailyPulseAccess("Administrator", { found: false }).allowed, true);
  assert.equal(access.resolveDailyPulseAccess("Behavior Specialist", { found: false }).allowed, false);
});

test("a lookup failure is always denied, never falls back", () => {
  const access = accessHarness();
  const decision = access.resolveDailyPulseAccess("Administrator", { failed: true });
  assert.equal(decision.allowed, false);
  assert.equal(decision.reason, "app-users-lookup-failed");
});

/* ── resolveTeacherRosterKey ─────────────────────────────────────────── */

test("Teacher/Classroom from IEP_App_Users is preferred when populated", () => {
  const access = accessHarness();
  const key = access.resolveTeacherRosterKey(
    { name: "Kelly H." },
    { found: true, row: { "Teacher/Classroom": "Kelly Higgins" } }
  );
  assert.equal(key, "Kelly Higgins");
});

test("falls back to the signed-in person's own name when Teacher/Classroom is blank or no row", () => {
  const access = accessHarness();
  assert.equal(access.resolveTeacherRosterKey({ name: "Kelly Higgins" }, { found: true, row: { "Teacher/Classroom": "" } }), "Kelly Higgins");
  assert.equal(access.resolveTeacherRosterKey({ name: "Kelly Higgins" }, { found: false }), "Kelly Higgins");
});

/* ── End-to-end APP.init() (real-auth path) ─────────────────────────── */

function appInitHarness({ appUsersRows = [], appUsersError = null, iepUsersRow, account }) {
  const students = [
    { name: "A Enabled", teacherName: "Teacher A", active: true, dailyPulseEnabled: true }
  ];
  const context = makeContext({
    APP_MODE: undefined,
    msal: { PublicClientApplication: class { async handleRedirectPromise() { return null; } getAllAccounts() { return [account]; } } },
    GRAPH: {
      async findUserByEmail() { return iepUsersRow; },
      async getListItemsByDisplayName(listName) {
        if (listName === "IEP_App_Users") {
          if (appUsersError) throw new Error(appUsersError);
          return appUsersRows;
        }
        return [];
      },
      async getTodayDailyPulseEntries() { return []; },
      async saveDailyPulse() { return { duplicatePrevented: false }; }
    },
    DEMO: { isActive: () => false },
    DB: { addDailyPulse: v => v, markSynced() {} },
    STUDENT_ROSTER: {
      loaded: true,
      error: null,
      _students: students,
      async refresh() {},
      getDailyPulseForTeacher(name) { return context.DAILY_PULSE_ACCESS.filterRoster(students, name); },
      getDailyPulseTeachers() { return context.DAILY_PULSE_ACCESS.eligibleTeachers(students); }
    }
  });
  runFile(context, "config.js");
  runFile(context, "auth.js");
  runFile(context, "iep-app-users.js");
  // roster.js is NOT loaded — STUDENT_ROSTER is supplied as a mock above
  // (matching appHarness() in administrator-access.test.js) so this file
  // stays about the authorization gate, not roster-loading mechanics.
  runFile(context, "access.js");
  context.DAILY_PULSE_ACCESS = vm.runInContext("DAILY_PULSE_ACCESS", context);
  runFile(context, "app.js");
  const app = vm.runInContext("APP", context);
  app._render = () => {};
  app._loadMsal = () => Promise.resolve();
  const auth = vm.runInContext("AUTH", context);
  auth.account = account;
  return { app, auth, context };
}

test("E2E: Teacher with Daily Pulse=Yes app-users row is allowed and gets Teacher/Classroom roster", async () => {
  const account = { username: "khiggins@iu29.org", name: "Kelly Higgins" };
  const { app } = appInitHarness({
    account,
    iepUsersRow: { Title: "u1", field_1: "Kelly Higgins", field_2: "Teacher", field_3: true, Email: account.username },
    appUsersRows: [{ id: "1", Email: account.username, "Daily Pulse": "Yes", "Teacher/Classroom": "Teacher A" }]
  });
  await app.init();
  assert.equal(app.screen, "status");
  assert.equal(app.effectiveTeacher.teacherName, "Teacher A");
});

test("E2E: Teacher with Daily Pulse=No app-users row is denied even though role would allow it", async () => {
  const account = { username: "khiggins@iu29.org", name: "Kelly Higgins" };
  const { app } = appInitHarness({
    account,
    iepUsersRow: { Title: "u1", field_1: "Kelly Higgins", field_2: "Teacher", field_3: true, Email: account.username },
    appUsersRows: [{ id: "1", Email: account.username, "Daily Pulse": "No" }]
  });
  await app.init();
  assert.equal(app.screen, "denied");
});

test("E2E: Administrator with Daily Pulse=Yes app-users row is allowed and keeps teacher-selection UI", async () => {
  const account = { username: "admin@iu29.org", name: "Actual Administrator" };
  const { app } = appInitHarness({
    account,
    iepUsersRow: { Title: "u2", field_1: "Actual Administrator", field_2: "Administrator", field_3: true, Email: account.username },
    appUsersRows: [{ id: "2", Email: account.username, "Daily Pulse": "Yes" }]
  });
  await app.init();
  assert.equal(app.screen, "teacher-select");
});

test("E2E: no matching IEP_App_Users row falls back to safe legacy Teacher/Administrator behavior", async () => {
  const account = { username: "khiggins@iu29.org", name: "Kelly Higgins" };
  const { app } = appInitHarness({
    account,
    iepUsersRow: { Title: "u1", field_1: "Kelly Higgins", field_2: "Teacher", field_3: true, Email: account.username },
    appUsersRows: [] // no row for this email
  });
  await app.init();
  assert.equal(app.screen, "status", "legacy fallback must still let an active Teacher in");
  assert.equal(app.effectiveTeacher.teacherName, "Kelly Higgins", "falls back to their own IEP_Users2 display name");
});

test("E2E: an IEP_App_Users lookup failure denies access, never falls back", async () => {
  const account = { username: "khiggins@iu29.org", name: "Kelly Higgins" };
  const { app } = appInitHarness({
    account,
    iepUsersRow: { Title: "u1", field_1: "Kelly Higgins", field_2: "Teacher", field_3: true, Email: account.username },
    appUsersRows: [],
    appUsersError: "Graph 403: insufficient privileges"
  });
  await app.init();
  assert.equal(app.screen, "denied", "a lookup failure must deny, not fall back, even for an otherwise-valid Teacher");
});

test("E2E: existing submission flow remains intact end-to-end for an allowed Teacher", async () => {
  const account = { username: "khiggins@iu29.org", name: "Kelly Higgins" };
  const { app } = appInitHarness({
    account,
    iepUsersRow: { Title: "u1", field_1: "Kelly Higgins", field_2: "Teacher", field_3: true, Email: account.username },
    appUsersRows: [{ id: "1", Email: account.username, "Daily Pulse": "Yes", "Teacher/Classroom": "Teacher A" }]
  });
  await app.init();
  app.startOrContinue();
  app.draft.overallStatus = "great";
  await app.submitAbsent();
  assert.equal(app.completedToday.has("A Enabled"), true);
});
