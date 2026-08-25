/* ── Microsoft Graph / SharePoint (extracted + adapted from MAC-Walkthrough
 * graph.js) ──────────────────────────────────────────────────────────────
 * Kept verbatim: transport, site/list discovery, schema mapping, the
 * writable-field guard, the generic duplicate-lookup helper, and
 * saveDailyPulse's exact field mapping (docs/daily-pulse-source-audit.md
 * §21-22). Removed: saveWalkthrough, saveWalkthroughV2, savePaceVisit,
 * getWhoAreYouVisiting — none of those lists are touched by this app.
 * Added: getTodayDailyPulseEntries(), a small read used to make "students
 * already completed today" resilient across reloads/devices instead of
 * relying on this browser's localStorage — see docs/daily-pulse-source-
 * audit.md §18 for why source didn't need this and standalone does.
 */

function isWritableSharePointField(internalName) {
  const blocked = new Set([
    "id",
    "ContentType",
    "Modified",
    "Created",
    "Author",
    "Editor",
    "AuthorLookupId",
    "EditorLookupId",
    "_UIVersionString",
    "Attachments",
    "Edit",
    "LinkTitle",
    "LinkTitleNoMenu",
    "ItemChildCount",
    "FolderChildCount",
    "_ComplianceFlags",
    "_ComplianceTag",
    "_ComplianceTagWrittenTime",
    "_ComplianceTagUserId"
  ]);
  if (!internalName) return false;
  if (blocked.has(internalName)) return false;
  if (internalName.startsWith("_")) return false;
  return true;
}

const GRAPH = {
  _BASE:        "https://graph.microsoft.com/v1.0",
  _SITE:        "siu29.sharepoint.com:/sites/IEP_Skook:",
  _siteId:      null,
  _listIdCache: {},
  _schemaCache: {},
  _lists:       { users: CONFIG.USERS_LIST },

  async _get(path) {
    const token = await AUTH.acquireGraphToken();
    const resp  = await fetch(`${this._BASE}/${path}`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      console.error("Graph GET failed", path, resp.status, body);
      throw new Error(`Graph ${resp.status}: ${body}`);
    }
    return resp.json();
  },

  async getSiteId() {
    if (this._siteId) return this._siteId;
    const data = await this._get(`sites/${this._SITE}`);
    this._siteId = data.id;
    return data.id;
  },

  async getListId(listName) {
    if (this._listIdCache[listName]) return this._listIdCache[listName];
    const siteId = await this.getSiteId();
    const data   = await this._get(`sites/${siteId}/lists?$select=id,name,displayName`);
    const match  = (data.value || []).find(list => {
      const names = [list.name, list.displayName]
        .map(v => String(v || "").toLowerCase().trim());
      return names.includes(String(listName).toLowerCase().trim());
    });
    if (!match) throw new Error(`SharePoint list not found: ${listName}`);
    this._listIdCache[listName] = match.id;
    return match.id;
  },

  async getListItems(listName) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const data   = await this._get(`sites/${siteId}/lists/${listId}/items?$expand=fields`);
    // Never log `items` itself here — this is a generic reader used for
    // users, students, and Daily Pulse entries alike.
    return (data.value || []).map(item => ({ id: item.id, ...item.fields }));
  },

  // Daily Pulse entries accumulate across the year beyond Graph's default
  // page size — use this all-pages reader for "today's completions" so a
  // busy list doesn't silently hide older-in-the-page duplicates.
  async getAllListItems(listName) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    let path  = `sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`;
    let items = [];
    while (path) {
      const data = await this._get(path);
      items = items.concat(data.value || []);
      const nextLink = data["@odata.nextLink"];
      path = nextLink ? nextLink.replace(`${this._BASE}/`, "") : null;
    }
    return items.map(item => ({ id: item.id, ...item.fields }));
  },

  // Normalizes internal SharePoint field keys back to display names.
  async getListItemsByDisplayName(listName, allPages = false) {
    const [items, schema] = await Promise.all([
      allPages ? this.getAllListItems(listName) : this.getListItems(listName),
      this.getListSchema(listName)
    ]);
    const displayByInternal = {};
    Object.entries(schema).forEach(([displayName, internalName]) => {
      displayByInternal[internalName] = displayName;
    });
    return items.map(item => {
      const row = { id: item.id };
      Object.entries(item).forEach(([key, value]) => {
        if (key !== "id") row[displayByInternal[key] || key] = value;
      });
      return row;
    });
  },

  async findListItemByDisplayField(listName, displayFieldName, value) {
    const [items, schema] = await Promise.all([
      this.getListItems(listName),
      this.getListSchema(listName)
    ]);
    const internalName = schema[displayFieldName];
    if (!internalName) throw new Error(`Column not found in ${listName}: ${displayFieldName}`);
    return items.find(item =>
      String(item[internalName] || "").trim() === String(value || "").trim()
    ) || null;
  },

  async findUserByEmail(email) {
    const users  = await this.getListItems(this._lists.users);
    const target = String(email || "").toLowerCase().trim();
    const match = users.find(user => {
      const possibleEmails = [
        user.Email,
        user.email,
        user.EMail,
        user.Email0,
        user.EmailAddress,
        user.Email_x0020_Address,
        user["Email Address"]
      ];
      return possibleEmails.some(value =>
        String(value || "").toLowerCase().trim() === target
      );
    }) || null;
    if (!match) {
      throw new Error("No matching IEP_Users2 Email found for the signed-in account.");
    }
    return match;
  },

  async _post(path, body) {
    const token = await AUTH.acquireGraphToken();
    const resp  = await fetch(`${this._BASE}/${path}`, {
      method:  "POST",
      headers: {
        Authorization:  `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    });
    if (!resp.ok) {
      const text = await resp.text().catch(() => "");
      console.error("Graph POST failed", path, resp.status, text);
      throw new Error(`Graph POST ${resp.status}: ${text}`);
    }
    return resp.json();
  },

  async createListItem(listName, fields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    try {
      return await this._post(`sites/${siteId}/lists/${listId}/items`, { fields });
    } catch (err) {
      console.error("createListItem failed:", { listName, listId, siteId, fieldNames: Object.keys(fields || {}), error: err.message });
      throw err;
    }
  },

  async getListColumns(listName) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const data   = await this._get(
      `sites/${siteId}/lists/${listId}/columns?$select=name,displayName,text,choice,dateTime,boolean,number`
    );
    return data.value || [];
  },

  async getListSchema(listName) {
    if (this._schemaCache[listName]) return this._schemaCache[listName];
    const columns = await this.getListColumns(listName);
    const schema  = {};
    columns.forEach(col => {
      if (col.displayName && col.name) schema[col.displayName] = col.name;
    });
    this._schemaCache[listName] = schema;
    return schema;
  },

  async mapFields(listName, displayFields) {
    const schema = await this.getListSchema(listName);
    const mapped = {};
    Object.entries(displayFields).forEach(([displayName, value]) => {
      if (value === undefined || value === null) return;
      const internalName = schema[displayName];
      if (!internalName) {
        console.warn(`UNMAPPED FIELD — display: "${displayName}", internalName: (none)`);
        return;
      }
      if (!isWritableSharePointField(internalName)) {
        console.warn(`SKIPPED (read-only) — display: "${displayName}", internalName: "${internalName}"`);
        return;
      }
      mapped[internalName] = value;
    });
    return mapped;
  },

  async createMappedListItem(listName, displayFields) {
    const fields = await this.mapFields(listName, displayFields);
    return this.createListItem(listName, fields);
  },

  // Same exact payload/duplicate-detection contract as MAC-Walkthrough's
  // saveDailyPulse — see docs/daily-pulse-source-audit.md §19 and §22.
  // MAC-Walkthrough must be able to read what this app writes with zero
  // translation.
  async saveDailyPulse(entry) {
    const responses = entry.CategoryResponses || {};
    const cats = Array.isArray(entry.Categories)
      ? entry.Categories.join(", ")
      : (entry.Categories || entry.Category || "");
    const recordId = entry.RecordID || `pulse-${Date.now()}`;
    const existing = await this.findListItemByDisplayField(
      CONFIG.DAILY_PULSE_LIST,
      "Record ID",
      recordId
    ).catch(() => null);

    if (existing) {
      console.warn("Daily Pulse entry already exists in SharePoint:", recordId);
      return { duplicatePrevented: true, existingItem: existing };
    }

    return this.createMappedListItem(CONFIG.DAILY_PULSE_LIST, {
      "Record ID":                recordId,
      "Date":                     entry.Date || new Date().toISOString().slice(0, 10),
      "Teacher":                  entry.Teacher || "",
      "Student":                  entry.Student || entry.StudentName || "",
      "Attendance Status":        entry.AttendanceStatus === "absent" ? "Absent" : "Present",
      "Present":                  entry.Present === true,
      "Status":                   entry.Status || entry.PulseStatus || "",
      "Category":                 cats,
      "Category Responses":       JSON.stringify(responses),
      "Academic Progress Rating": responses.academicProgress?.rating   || "",
      "Behavior Rating":          responses.behavior?.rating           || "",
      "Communication Rating":     responses.communication?.rating      || "",
      "Social Skills Rating":     responses.socialSkills?.rating       || "",
      "Self-Regulation Rating":   responses.selfRegulation?.rating     || "",
      "Transition Rating":        responses.transitionDifficulty?.rating || "",
      "IEP Goal Rating":          responses.iepGoalProgress?.rating    || "",
      "Quick Note":               entry.QuickNote || "",
      "Submitted By":             entry.SubmittedBy || "",
      "Submitted At":             entry.SubmittedAt || new Date().toISOString()
    });
  },

  // NEW (adapted, not present in source — see docs/daily-pulse-source-
  // audit.md §18): reads today's Daily Pulse entries for a teacher directly
  // from SharePoint, so "already completed today" survives a reload on a
  // different device/browser instead of depending on local history.
  async getTodayDailyPulseEntries(teacherName, dateStr) {
    const rows = await this.getListItemsByDisplayName(CONFIG.DAILY_PULSE_LIST, true);
    const targetTeacher = String(teacherName || "").trim().toLowerCase();
    return rows.filter(row => {
      const rowDate    = String(row["Date"] || "").slice(0, 10);
      const rowTeacher = String(row["Teacher"] || "").trim().toLowerCase();
      return rowDate === dateStr && (!targetTeacher || rowTeacher === targetTeacher);
    });
  }
};
