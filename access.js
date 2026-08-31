/* ── Daily Pulse role/effective-teacher helpers ─────────────────────────
 * Keeps the actual signed-in actor separate from the teacher whose roster
 * is being completed. Teachers are their own effective teacher; an
 * administrator must explicitly select a teacher identity taken from the
 * existing Active + Daily Pulse Enabled roster data.
 */

const DAILY_PULSE_ACCESS = {
  normalizeRole(value) {
    const raw = Array.isArray(value) ? value[0] : value;
    return String(raw || "").trim().toLowerCase();
  },

  isSupportedRole(value) {
    return ["teacher", "administrator"].includes(this.normalizeRole(value));
  },

  requiresTeacherSelection(value) {
    return this.normalizeRole(value) === "administrator";
  },

  resolveActor(auth) {
    const name = auth.pilotUser?.Name || auth.displayName || "";
    return {
      id:      auth.pilotUserId || slugifyTeacherId(name),
      name,
      email:   auth.account?.username || "",
      role:    this.normalizeRole(auth.role),
      isAdmin: this.normalizeRole(auth.role) === "administrator"
    };
  },

  teacherIdentity(name) {
    const normalized = String(name || "").trim();
    if (!normalized) return null;
    return {
      id:          slugifyTeacherId(normalized),
      name:        normalized,
      teacherName: normalized
    };
  },

  effectiveTeacherForActor(actor) {
    if (!actor || this.normalizeRole(actor.role) !== "teacher") return null;
    return this.teacherIdentity(actor.name);
  },

  /* ── PATCH B: IEP_App_Users-based authorization ────────────────────────
   * Combines the existing IEP_Users2 role gate (isSupportedRole, above)
   * with the new IEP_App_Users "Daily Pulse" flag. AUTH.isUnauthorized
   * (no active IEP_Users2 record at all) is still checked first by app.js,
   * unchanged — this only decides the ADDITIONAL app-permission gate.
   *
   * appUsersOutcome is one of:
   *   { found: true, row }  — a matching IEP_App_Users row exists; its
   *                           "Daily Pulse" flag is authoritative, even if
   *                           it disagrees with the legacy role check.
   *   { found: false }      — the lookup succeeded but no row matched this
   *                           email. MIGRATION MODE (see iep-app-users.js):
   *                           fall back to the existing Teacher/Admin role
   *                           check so nobody is locked out before
   *                           IEP_App_Users is populated.
   *   { failed: true }      — the lookup itself errored (Graph/schema/
   *                           network). Always denied — a failed lookup is
   *                           never treated as "no row found."
   */
  resolveDailyPulseAccess(role, appUsersOutcome) {
    if (appUsersOutcome?.failed) {
      return { allowed: false, reason: "app-users-lookup-failed" };
    }
    if (appUsersOutcome?.found) {
      const allowed = normalizeAppUsersBoolean(appUsersOutcome.row["Daily Pulse"]);
      return { allowed, reason: allowed ? "app-users-permission" : "app-users-denied" };
    }
    if (typeof ACCESS_MODEL_CONFIG !== "undefined" && ACCESS_MODEL_CONFIG.ENFORCE_APP_USERS) {
      return { allowed: false, reason: "enforced-no-row" };
    }
    const allowed = this.isSupportedRole(role);
    return { allowed, reason: allowed ? "legacy-fallback" : "legacy-denied" };
  },

  // Teacher/Classroom (IEP_App_Users) is preferred when a matching row
  // exists and it's populated — same literal string IEP_Students_2026_27
  // uses for Teacher, no new teacher ID invented. Otherwise falls back to
  // the signed-in person's own IEP_Users2 display name, exactly as before
  // PATCH B. Administrators keep using the existing teacher-selection UI
  // (app.js branches on AUTH.isAdmin before ever calling this).
  resolveTeacherRosterKey(actor, appUsersOutcome) {
    const fromAppUsers = appUsersOutcome?.found
      ? String(appUsersOutcome.row["Teacher/Classroom"] || "").trim()
      : "";
    return fromAppUsers || actor?.name || "";
  },

  eligibleTeachers(students) {
    const seen = new Set();
    return (students || [])
      .filter(student => student?.active && student?.dailyPulseEnabled && String(student.teacherName || "").trim())
      .map(student => this.teacherIdentity(student.teacherName))
      .filter(teacher => {
        const key = teacher.teacherName.toLowerCase();
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      })
      .sort((a, b) => a.teacherName.localeCompare(b.teacherName));
  },

  filterRoster(students, teacherName) {
    const target = String(teacherName || "").trim().toLowerCase();
    if (!target) return [];
    return (students || []).filter(student =>
      student?.active &&
      student?.dailyPulseEnabled &&
      String(student.teacherName || "").trim().toLowerCase() === target
    );
  },

  hasMeaningfulDraft(draft) {
    if (!draft) return false;
    return Boolean(
      draft.attendanceStatus || draft.overallStatus ||
      (draft.selectedCategories || []).length ||
      Object.keys(draft.categoryResponses || {}).length ||
      String(draft.quickNote || "").trim()
    );
  },

  buildPayload(draft, attendanceStatus, effectiveTeacher, actor, now = new Date()) {
    const present = attendanceStatus === "present";
    const statusLabel = present
      ? ({ great: "Great Day", struggles: "Some Struggles", concern: "Significant Concern" }[draft.overallStatus] || draft.overallStatus)
      : "";
    return {
      RecordID:          draft.recordId,
      Date:              draft.date,
      Teacher:           effectiveTeacher?.teacherName || "",
      Student:           draft.studentName,
      AttendanceStatus:  attendanceStatus,
      Present:           present,
      Status:            statusLabel,
      Categories:        present ? [...draft.selectedCategories] : [],
      CategoryResponses: present ? structuredClone(draft.categoryResponses) : {},
      QuickNote:         present ? (draft.quickNote || "") : "",
      SubmittedBy:       actor?.name || "",
      SubmittedByEmail:  actor?.email || "",
      SubmittedAt:       now.toISOString()
    };
  }
};

