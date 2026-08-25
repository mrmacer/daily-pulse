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

