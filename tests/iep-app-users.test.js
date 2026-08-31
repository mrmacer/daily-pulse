/* ─────────────────────────────────────────────────────────────────────────
   PATCH A — IEP_App_Users reader coverage

   Covers only iep-app-users.js in isolation. No app.js/auth.js behavior
   changes yet in PATCH A, so there is nothing end-to-end to test here —
   that arrives with the Daily Pulse gate patch.
   ───────────────────────────────────────────────────────────────────────── */

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");

function makeContext(overrides = {}) {
  const context = vm.createContext({
    console,
    CONFIG: { APP_USERS_LIST: "IEP_App_Users" },
    ...overrides
  });
  return context;
}

function loadModule(context) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, "iep-app-users.js"), "utf8"), context, { filename: "iep-app-users.js" });
}

test("normalizeAppUsersBoolean handles Yes/No/true/false/1/0 and blank defaults", () => {
  const context = makeContext();
  loadModule(context);
  const normalize = vm.runInContext("normalizeAppUsersBoolean", context);
  for (const truthy of ["Yes", "yes", " YES ", "true", "1", true, 1]) {
    assert.equal(normalize(truthy), true);
  }
  for (const falsy of ["No", "no", "false", "0", false, 0]) {
    assert.equal(normalize(falsy), false);
  }
  assert.equal(normalize(undefined, true), true);
  assert.equal(normalize("", false), false);
  assert.equal(normalize(null), false);
});

test("PATCH A ships in migration mode, not enforced mode", () => {
  const context = makeContext();
  loadModule(context);
  assert.equal(vm.runInContext("ACCESS_MODEL_CONFIG", context).ENFORCE_APP_USERS, false);
});

test("resolve() finds a matching row by email, case/whitespace-insensitive", async () => {
  let calledArgs = null;
  const context = makeContext({
    GRAPH: {
      async getListItemsByDisplayName(...args) {
        calledArgs = args;
        return [
          { id: "1", "Display Name": "Sharon Morgan", "Role": "Behavior Specialist", "Email": "Sharon.Morgan@iu29.org", "PACE": "Yes", "Daily Pulse": "No", "Walkthrough": "No", "Admin Panel": "No", "Teacher/Classroom": "" },
          { id: "2", "Display Name": "Kelly Higgins", "Role": "Teacher", "Email": "khiggins@iu29.org", "PACE": "No", "Daily Pulse": "Yes", "Walkthrough": "No", "Admin Panel": "No", "Teacher/Classroom": "Kelly Higgins" }
        ];
      }
    }
  });
  loadModule(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  const row = await appUsers.resolve("  SHARON.MORGAN@IU29.ORG  ");
  assert.deepEqual(calledArgs, ["IEP_App_Users", true]);
  assert.ok(row);
  assert.equal(appUsers.found, true);
  assert.equal(appUsers.lookupError, null);
  assert.deepEqual(JSON.parse(JSON.stringify(appUsers.permissions)), { dailyPulse: false, pace: true, walkthrough: false, adminPanel: false });
});

test("resolve() with a teacher row exposes Teacher/Classroom for roster matching", async () => {
  const context = makeContext({
    GRAPH: {
      async getListItemsByDisplayName() {
        return [{ id: "2", "Display Name": "Kelly Higgins", "Role": "Teacher", "Email": "khiggins@iu29.org", "Daily Pulse": "Yes", "PACE": "No", "Walkthrough": "No", "Admin Panel": "No", "Teacher/Classroom": "Kelly Higgins" }];
      }
    }
  });
  loadModule(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  await appUsers.resolve("khiggins@iu29.org");
  assert.equal(appUsers.teacherClassroom, "Kelly Higgins");
});

test("resolve() with no matching row returns null cleanly (not an error)", async () => {
  const context = makeContext({
    GRAPH: { async getListItemsByDisplayName() { return [{ id: "1", Email: "someone-else@iu29.org" }]; } }
  });
  loadModule(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  const row = await appUsers.resolve("nobody@iu29.org");
  assert.equal(row, null);
  assert.equal(appUsers.found, false);
  assert.equal(appUsers.permissions, null);
  assert.equal(appUsers.lookupError, null);
});

test("resolve() fails closed on a Graph/list error", async () => {
  const context = makeContext({
    GRAPH: { async getListItemsByDisplayName() { throw new Error("Graph 403: insufficient privileges"); } }
  });
  loadModule(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  const row = await appUsers.resolve("khiggins@iu29.org");
  assert.equal(row, null, "a failed lookup must never resolve to a match");
  assert.match(appUsers.lookupError, /insufficient privileges/);
});
