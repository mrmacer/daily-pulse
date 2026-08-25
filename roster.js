/* ── Student Roster Provider (extracted + trimmed from MAC-Walkthrough
 * app.js:195-344) ───────────────────────────────────────────────────────
 * Live, SharePoint-backed. This file only ever knows HOW to load the
 * roster, never WHO is on it until it fetches from SharePoint at runtime.
 * PACE-specific getters were dropped (not used here); everything else,
 * including the exact Active/Daily-Pulse-Enabled default-value behavior,
 * is unchanged from source. See docs/daily-pulse-source-audit.md §8-10.
 */

function slugifyTeacherId(name) {
  const slug = String(name || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
  return slug ? `teacher-${slug}` : "";
}

// Normalizes a SharePoint Yes/No column to a boolean. `defaultValue` applies
// only when the column is absent/blank.
function normalizeRosterBoolean(value, defaultValue) {
  if (value === undefined || value === null || value === "") return defaultValue;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;
  const s = String(value).trim().toLowerCase();
  return s === "yes" || s === "true" || s === "1";
}

const STUDENT_ROSTER = {
  _students: [],
  loading:   false,
  loaded:    false,
  error:     null,

  // Production schema (IEP_Students_2026_27) has separate first/last name
  // columns and no "Student ID" column.
  _normalize(row) {
    const firstName = String(row["Student First Name"] || "").trim();
    const lastName  = String(row["Student Last Name"]  || "").trim();
    const name = `${firstName} ${lastName}`.trim();
    if (!name) return null;

    const teacherName = String(row["Teacher"] || "").trim();

    return {
      // The SharePoint list item's own stable item id — never derived from
      // the student's name, which can change or collide across students.
      id:                row.id != null ? String(row.id) : name,
      firstName,
      lastName,
      name,
      teacherName,
      teacherId:         slugifyTeacherId(teacherName),
      classroom:         String(row["Classroom"] || "").trim(),
      // Fail-safe: a blank/missing Daily Pulse Enabled cell defaults to NOT
      // eligible, unlike Active (which keeps its existing default-true
      // behavior for compatibility).
      active:            normalizeRosterBoolean(row["Active"],             true),
      dailyPulseEnabled: normalizeRosterBoolean(row["Daily Pulse Enabled"], false)
    };
  },

  async refresh() {
    this.loading = true;
    this.error   = null;
    try {
      const [raw, schema] = await Promise.all([
        GRAPH.getListItems(CONFIG.STUDENT_ROSTER_LIST),
        GRAPH.getListSchema(CONFIG.STUDENT_ROSTER_LIST)
      ]);
      const inv = {};
      Object.entries(schema).forEach(([display, internal]) => { inv[internal] = display; });
      this._students = raw
        .map(item => {
          const row = {};
          Object.entries(item).forEach(([k, v]) => { row[inv[k] || k] = v; });
          row.id = item.id;
          return this._normalize(row);
        })
        .filter(Boolean);
      this.loaded = true;
      console.log(`Student roster loaded: ${this._students.length} record(s).`);
    } catch (err) {
      console.error("Student roster load failed:", err.message || String(err));
      this.error = err.message || "Unable to load the student roster.";
    } finally {
      this.loading = false;
    }
    return this._students;
  },

  getAll()    { return this._students; },
  getActive() { return this._students.filter(s => s.active); },
  findByName(name) { return this._students.find(s => s.name === name) || null; },

  // Case/whitespace-insensitive; a record with no Teacher value never
  // matches any teacher.
  _matchesTeacher(student, teacherName) {
    const target = String(teacherName || "").trim().toLowerCase();
    if (!target || !student.teacherName) return false;
    return student.teacherName.trim().toLowerCase() === target;
  },

  // Daily Pulse (teacher view): active + Daily-Pulse-enabled + assigned to
  // this specific teacher.
  getDailyPulseForTeacher(teacherName) {
    return this.getActive().filter(s => s.dailyPulseEnabled && this._matchesTeacher(s, teacherName));
  }
};

/* ── Daily Pulse Scales (extracted verbatim from MAC-Walkthrough
 * app.js:1001-1125) ────────────────────────────────────────────────────── */
const DAILY_PULSE_SCALES = {
  academicProgress: {
    label:  "Academic Progress",
    prompt: "How did the student participate or demonstrate skills in reading and/or math?",
    choices: [
      { value: "not-addressed",       label: "Not addressed today",                          tone: "neutral"  },
      { value: "significant-support", label: "Participated but required significant support", tone: "negative" },
      { value: "some-support",        label: "Demonstrated skills with some support",         tone: "positive" },
      { value: "independent",         label: "Demonstrated skills independently",             tone: "positive" },
      { value: "noticeable-progress", label: "Made noticeable progress",                      tone: "positive" }
    ],
    detailFields: [
      { key: "skillArea", label: "Skill, activity, or area of progress", type: "text", optional: true }
    ]
  },
  behavior: {
    label:  "Behavior",
    prompt: "Which statement best describes the student's behavior today?",
    choices: [
      { value: "excellent", label: "Excellent day — no behavioral concerns",                              tone: "positive" },
      { value: "minor",     label: "Minor concerns — redirection was effective",                          tone: "negative" },
      { value: "moderate",  label: "Moderate concerns — repeated behaviors or additional support needed", tone: "negative" },
      { value: "major",     label: "Major concerns — unsafe behavior occurred",                           tone: "negative" }
    ],
    multiSelectFields: [
      { key: "behaviorsObserved", label: "Behaviors observed", options: [
        "Hitting", "Biting", "Scratching", "Kicking", "Elopement / attempted elopement",
        "Property destruction", "Verbal aggression", "Refusal / noncompliance", "Other"
      ]}
    ],
    detailFields: [
      { key: "behaviorContext", label: "What occurred, possible trigger, and staff response", type: "textarea", optional: true },
      { key: "otherBehavior",  label: "Other behavior", type: "text", optional: true, showWhenOption: "Other" }
    ]
  },
  communication: {
    label:  "Communication",
    prompt: "Consider verbal communication, gestures, signs, picture symbols, or an AAC device.",
    choices: [
      { value: "did-not-communicate",   label: "Did not communicate needs today",             tone: "negative" },
      { value: "significant-prompting", label: "Communicated with significant prompting",      tone: "negative" },
      { value: "occasional-prompting",  label: "Communicated with occasional prompting",       tone: "positive" },
      { value: "independent",           label: "Communicated independently and appropriately", tone: "positive" },
      { value: "not-observed",          label: "Not observed / not applicable",                tone: "neutral"  }
    ],
    multiSelectFields: [
      { key: "communicationMethods", label: "Communication method", options: [
        "Verbal", "Gestures / sign language", "Picture symbols", "AAC / communication device", "Other"
      ]}
    ],
    detailFields: [
      { key: "communicationComments",    label: "Comments or observations",   type: "textarea", optional: true },
      { key: "otherCommunicationMethod", label: "Other communication method", type: "text", optional: true, showWhenOption: "Other" }
    ]
  },
  socialSkills: {
    label:  "Social Skills",
    prompt: "How did the student interact with peers or adults?",
    choices: [
      { value: "avoided",             label: "Avoided or did not interact with others",        tone: "negative" },
      { value: "significant-support", label: "Interacted only with significant adult support", tone: "negative" },
      { value: "some-support",        label: "Had some appropriate interactions with support", tone: "positive" },
      { value: "independent",         label: "Interacted appropriately and independently",     tone: "positive" },
      { value: "not-observed",        label: "Not observed / not applicable",                  tone: "neutral"  }
    ],
    detailFields: [
      { key: "interactionNotes", label: "Interaction with peers or adults", type: "textarea", optional: true }
    ]
  },
  selfRegulation: {
    label:  "Self-Regulation",
    prompt: "How well did the student manage emotions, sensory needs, frustration, and changes?",
    choices: [
      { value: "regulated",          label: "Remained calm and regulated throughout the day",         tone: "positive" },
      { value: "recovered-minimal",  label: "Became dysregulated but recovered with minimal support", tone: "positive" },
      { value: "repeated-support",   label: "Required repeated adult support or coping strategies",   tone: "negative" },
      { value: "significant-dysreg", label: "Significant dysregulation affected learning or safety",  tone: "negative" },
      { value: "not-observed",       label: "Not observed / not applicable",                          tone: "neutral"  }
    ],
    multiSelectFields: [
      { key: "regulationStrategies", label: "Strategies used", options: [
        "Break", "Quiet / calming area", "Sensory tool", "Movement",
        "Breathing / calming strategy", "Adult reassurance", "Visual schedule", "Other"
      ]}
    ],
    detailFields: [
      { key: "regulationComments",      label: "Comments or observations", type: "textarea", optional: true },
      { key: "otherRegulationStrategy", label: "Other strategy",           type: "text",     optional: true, showWhenOption: "Other" }
    ]
  },
  transitionDifficulty: {
    label:  "Transition Difficulty",
    prompt: "How did the student manage transitions today?",
    choices: [
      { value: "none",         label: "No difficulty — transitioned independently",                                 tone: "positive" },
      { value: "minor",        label: "Minor difficulty — needed one or two prompts",                               tone: "negative" },
      { value: "some",         label: "Some difficulty — needed repeated prompting or support",                     tone: "negative" },
      { value: "major",        label: "Major difficulty — refusal, aggression, elopement, or significant distress", tone: "negative" },
      { value: "not-observed", label: "Not observed / not applicable",                                              tone: "neutral"  }
    ],
    detailFields: [
      { key: "difficultTransition", label: "Location or transition that was difficult", type: "textarea", optional: true, showForValues: ["minor", "some", "major"] }
    ]
  },
  iepGoalProgress: {
    label:  "IEP Goal Progress",
    prompt: "How did the student progress toward an IEP goal today?",
    choices: [
      { value: "not-addressed",        label: "Goal was not addressed today",            tone: "neutral"  },
      { value: "no-progress",          label: "Goal addressed — no progress observed",   tone: "negative" },
      { value: "some-progress",        label: "Goal addressed — some progress observed", tone: "positive" },
      { value: "significant-progress", label: "Goal addressed — significant progress",   tone: "positive" },
      { value: "independent",          label: "Goal completed independently",            tone: "positive" }
    ],
    multiSelectFields: [
      { key: "goalAreas", label: "Goal area addressed", options: [
        "Academic", "Behavior", "Communication", "Social skills",
        "Functional / life skills", "Related service", "Other"
      ]}
    ],
    detailFields: [
      { key: "otherGoalArea", label: "Other goal area", type: "text", optional: true, showWhenOption: "Other" }
    ]
  }
};
