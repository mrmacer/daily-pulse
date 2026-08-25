/* ── Local History (adapted from MAC-Walkthrough data.js) ───────────────────
 * A resilience layer only — never the source of truth. A confirmed
 * SharePoint save is what marks a student complete; this just keeps a local
 * copy so a submission survives a tab crash before the retry, and gives the
 * "Today's Submissions" read-only list something to render immediately
 * without waiting on a fresh Graph read. See docs/daily-pulse-source-
 * audit.md §4, §18, §20.
 */
const DB = {
  _get(key) {
    try {
      return JSON.parse(localStorage.getItem(key)) || [];
    } catch {
      return [];
    }
  },
  _set(key, value) {
    localStorage.setItem(key, JSON.stringify(value));
  },
  _id(prefix) {
    return `${prefix}-${crypto.randomUUID()}`;
  },

  getDailyPulses() { return this._get(CONFIG.STORAGE_KEYS.LOCAL_HISTORY); },

  addDailyPulse(data) {
    const list = this.getDailyPulses();
    const entry = {
      // Reuse a caller-supplied stable id (the same one used as the
      // SharePoint duplicate-detection key) so a retry after a failed sync
      // is recognized as the same submission, not a fresh one.
      id:                data.id || this._id("pulse"),
      timestamp:         new Date().toISOString(),
      date:              data.date || new Date().toISOString().slice(0, 10),
      student:           data.student || "",
      studentId:         data.studentId || "",
      attendanceStatus:  data.attendanceStatus || "present",
      pulseStatus:       data.pulseStatus || "",
      categories:        data.categories || [],
      categoryResponses: data.categoryResponses || {},
      note:              data.note || "",
      teacherId:         data.teacherId || "",
      teacherName:       data.teacherName || "",
      submittedById:     data.submittedById || "",
      submittedByName:   data.submittedByName || "",
      synced:            data.synced === true
    };
    list.push(entry);
    this._set(CONFIG.STORAGE_KEYS.LOCAL_HISTORY, list);
    return entry;
  },

  markSynced(id) {
    const list = this.getDailyPulses().map(e => e.id === id ? { ...e, synced: true } : e);
    this._set(CONFIG.STORAGE_KEYS.LOCAL_HISTORY, list);
  }
};
