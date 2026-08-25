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
  return context;
}

async function authorizedAuth(role, active = true) {
  const context = makeContext({
    msal: { PublicClientApplication: class {} },
    GRAPH: {
      async findUserByEmail() {
        return {
          Title: "user-1",
          field_1: `${role} Person`,
          field_2: role,
          field_3: active,
          Email: `${role.toLowerCase()}@example.test`
        };
      }
    }
  });
  runFile(context, "auth.js");
  const auth = vm.runInContext("AUTH", context);
  auth.account = { username: `${role.toLowerCase()}@example.test`, name: `${role} Person` };
  await auth.loadPilotUserFromSharePoint();
  return auth;
}

function accessHarness() {
  const context = makeContext({
    slugifyTeacherId(value) {
      return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
    }
  });
  runFile(context, "access.js");
  return vm.runInContext("DAILY_PULSE_ACCESS", context);
}

function appHarness() {
  const students = [
    { name: "A Enabled", teacherName: "Teacher A", active: true, dailyPulseEnabled: true },
    { name: "A Disabled", teacherName: "Teacher A", active: true, dailyPulseEnabled: false },
    { name: "A Inactive", teacherName: "Teacher A", active: false, dailyPulseEnabled: true },
    { name: "B Enabled", teacherName: "Teacher B", active: true, dailyPulseEnabled: true }
  ];
  const context = makeContext({
    DEMO: { isActive: () => false },
    AUTH: {},
    DAILY_PULSE_SCALES: {},
    DB: { addDailyPulse: value => value, markSynced() {} },
    GRAPH: {
      async getTodayDailyPulseEntries() { return []; },
      async saveDailyPulse() { return { ok: true }; }
    },
    STUDENT_ROSTER: {
      loaded: true,
      error: null,
      _students: students,
      async refresh() {},
      getDailyPulseForTeacher(name) {
        return context.DAILY_PULSE_ACCESS.filterRoster(this._students, name);
      },
      getDailyPulseTeachers() {
        return context.DAILY_PULSE_ACCESS.eligibleTeachers(this._students);
      }
    }
  });
  context.slugifyTeacherId = value => String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  runFile(context, "access.js");
  context.DAILY_PULSE_ACCESS = vm.runInContext("DAILY_PULSE_ACCESS", context);
  runFile(context, "app.js");
  const app = vm.runInContext("APP", context);
  app._render = () => {};
  return { app, context, students };
}

test("1. active Teacher is authorized", async () => {
  const auth = await authorizedAuth("Teacher");
  assert.equal(auth.isDailyPulseAuthorized, true);
});

test("2. active Administrator is authorized", async () => {
  const auth = await authorizedAuth("Administrator");
  assert.equal(auth.isDailyPulseAuthorized, true);
});

test("3. inactive users remain denied", async () => {
  const auth = await authorizedAuth("Teacher", false);
  assert.equal(auth.isUnauthorized, true);
  assert.equal(auth.isDailyPulseAuthorized, false);
});

test("4. unsupported roles remain denied", async () => {
  const auth = await authorizedAuth("Supervisor");
  assert.equal(auth.isAuthenticated, true);
  assert.equal(auth.isDailyPulseAuthorized, false);
});

test("5. Teacher automatically receives own roster", async () => {
  const { app, context } = appHarness();
  app.currentUser = { id: "teacher-a", name: "Teacher A", email: "a@example.test", role: "teacher", isAdmin: false };
  app.effectiveTeacher = context.DAILY_PULSE_ACCESS.effectiveTeacherForActor(app.currentUser);
  await app._loadRosterAndStatus();
  assert.deepEqual(app.roster.map(student => student.name), ["A Enabled"]);
});

test("6. Teacher does not see teacher selector", () => {
  const { app, context } = appHarness();
  app.currentUser = { name: "Teacher A", role: "Teacher", isAdmin: false };
  app.effectiveTeacher = context.DAILY_PULSE_ACCESS.effectiveTeacherForActor(app.currentUser);
  app.roster = [{ name: "A Enabled" }];
  assert.equal(context.DAILY_PULSE_ACCESS.requiresTeacherSelection("Teacher"), false);
  assert.equal(app.effectiveTeacher.teacherName, "Teacher A");
  assert.doesNotMatch(app._statusHtml(), /id="teacherSearch"/);
});

test("7. Administrator sees teacher selector", () => {
  const { app, context } = appHarness();
  app.currentUser = { name: "Actual Administrator", role: "administrator", isAdmin: true };
  app.teacherDirectory = context.DAILY_PULSE_ACCESS.eligibleTeachers(context.STUDENT_ROSTER._students);
  assert.equal(context.DAILY_PULSE_ACCESS.requiresTeacherSelection(app.currentUser.role), true);
  assert.match(app._teacherSelectHtml(), /id="teacherSearch"/);
  assert.match(app._teacherSelectHtml(), /Teacher A/);
});

test("8. Administrator can select Teacher A", async () => {
  const { app, context } = appHarness();
  app.currentUser = { name: "Actual Administrator", role: "administrator", isAdmin: true };
  app.teacherDirectory = context.DAILY_PULSE_ACCESS.eligibleTeachers(context.STUDENT_ROSTER._students);
  assert.equal(await app.selectAdminTeacher("Teacher A"), true);
  assert.equal(app.effectiveTeacher.teacherName, "Teacher A");
});

test("9. only Teacher A's active Daily Pulse-enabled students load", async () => {
  const { app, context } = appHarness();
  app.currentUser = { name: "Actual Administrator", role: "administrator", isAdmin: true };
  app.teacherDirectory = context.DAILY_PULSE_ACCESS.eligibleTeachers(context.STUDENT_ROSTER._students);
  await app.selectAdminTeacher("Teacher A");
  assert.deepEqual(app.roster.map(student => student.name), ["A Enabled"]);
});

test("10. Administrator can switch to Teacher B", async () => {
  const { app, context } = appHarness();
  app.currentUser = { name: "Actual Administrator", role: "administrator", isAdmin: true };
  app.teacherDirectory = context.DAILY_PULSE_ACCESS.eligibleTeachers(context.STUDENT_ROSTER._students);
  await app.selectAdminTeacher("Teacher A");
  await app.selectAdminTeacher("Teacher B");
  assert.equal(app.effectiveTeacher.teacherName, "Teacher B");
  assert.deepEqual(app.roster.map(student => student.name), ["B Enabled"]);
});

test("11. Teacher A state does not leak into Teacher B", async () => {
  const { app, context } = appHarness();
  app.currentUser = { name: "Actual Administrator", role: "administrator", isAdmin: true };
  app.teacherDirectory = context.DAILY_PULSE_ACCESS.eligibleTeachers(context.STUDENT_ROSTER._students);
  await app.selectAdminTeacher("Teacher A");
  app.draft = { studentName: "A Enabled", quickNote: "Teacher A draft" };
  app.completedToday.add("A Enabled");
  app.todayEntries.push({ student: "A Enabled" });
  await app.selectAdminTeacher("Teacher B");
  assert.equal(app.draft, null);
  assert.equal(app.completedToday.has("A Enabled"), false);
  assert.deepEqual(app.todayEntries, []);
});

test("12. Daily Pulse save payload remains source-compatible", () => {
  const access = accessHarness();
  const draft = {
    recordId: "pulse-123",
    date: "2026-08-25",
    studentName: "A Enabled",
    overallStatus: "great",
    selectedCategories: ["behavior"],
    categoryResponses: { behavior: { rating: "positive" } },
    quickNote: "Good day"
  };
  const payload = access.buildPayload(
    draft,
    "present",
    { teacherName: "Teacher A" },
    { name: "Actual Administrator", email: "admin@example.test" },
    new Date("2026-08-25T15:30:00.000Z")
  );
  assert.deepEqual(Object.keys(payload), [
    "RecordID", "Date", "Teacher", "Student", "AttendanceStatus", "Present",
    "Status", "Categories", "CategoryResponses", "QuickNote", "SubmittedBy",
    "SubmittedByEmail", "SubmittedAt"
  ]);
  assert.equal(payload.Status, "Great Day");
  assert.deepEqual(Array.from(payload.Categories), ["behavior"]);
});

test("13. actual administrator identity is retained as submitter", async () => {
  const access = accessHarness();
  const payload = access.buildPayload(
    {
      recordId: "pulse-124", date: "2026-08-25", studentName: "A Enabled",
      overallStatus: "", selectedCategories: [], categoryResponses: {}, quickNote: ""
    },
    "absent",
    { teacherName: "Teacher A" },
    { name: "Actual Administrator", email: "admin@example.test" },
    new Date("2026-08-25T15:30:00.000Z")
  );
  assert.equal(payload.Teacher, "Teacher A");
  assert.equal(payload.SubmittedBy, "Actual Administrator");
  assert.equal(payload.SubmittedByEmail, "admin@example.test");

  const graphContext = makeContext({
    CONFIG: { USERS_LIST: "IEP_Users2", DAILY_PULSE_LIST: "Daily Pulse" },
    AUTH: {},
    fetch: async () => { throw new Error("unexpected network request"); }
  });
  runFile(graphContext, "graph.js");
  const graph = vm.runInContext("GRAPH", graphContext);
  let mappedFields;
  graph.findListItemByDisplayField = async () => null;
  graph.createMappedListItem = async (_list, fields) => { mappedFields = fields; return { id: "new-item" }; };
  await graph.saveDailyPulse(payload);
  assert.equal(mappedFields.Teacher, "Teacher A");
  assert.equal(mappedFields["Submitted By"], "Actual Administrator");
});

test("14. duplicate protection still prevents a second SharePoint create", async () => {
  let creates = 0;
  const context = makeContext({
    CONFIG: { USERS_LIST: "IEP_Users2", DAILY_PULSE_LIST: "Daily Pulse" },
    AUTH: {},
    fetch: async () => { throw new Error("unexpected network request"); }
  });
  runFile(context, "graph.js");
  const graph = vm.runInContext("GRAPH", context);
  graph.findListItemByDisplayField = async () => ({ id: "existing-item" });
  graph.createMappedListItem = async () => { creates += 1; return { id: "new-item" }; };
  const result = await graph.saveDailyPulse({ RecordID: "pulse-existing" });
  assert.equal(result.duplicatePrevented, true);
  assert.equal(creates, 0);
});
