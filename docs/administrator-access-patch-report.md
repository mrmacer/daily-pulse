# Administrator Access Patch Report

## Authorization

Daily Pulse continues to use MSAL/Entra authentication and the `IEP_Users2`
authorization record. An account is admitted only when its record is active
and its normalized role is `Teacher` or `Administrator`. Inactive accounts,
missing records, and unsupported roles remain denied.

## Effective Teacher

- Teacher: the authenticated `IEP_Users2` name becomes the effective teacher
  automatically; no selector is shown.
- Administrator: the user must select a teacher before any student workflow
  loads. The searchable, touch-sized choices are distinct teacher values from
  active, Daily-Pulse-enabled roster rows in `IEP_Students_2026_27`.

The effective teacher scopes the roster and today's completion query. Student
filtering still requires `Active = true`, `Daily Pulse Enabled = true`, and an
exact case/whitespace-normalized teacher match.

## Save Semantics

The SharePoint mapping and field names are unchanged. `Teacher` contains the
effective teacher. `Submitted By` contains the actual signed-in user, including
an administrator acting for a teacher. The payload continues to carry
`SubmittedByEmail`, although the existing source-compatible SharePoint mapper
does not write a separate email display field.

Switching teachers clears the draft, completed-student set, today's entries,
and roster before loading the new context. If a draft contains entered data,
the **Choose Another Teacher** action asks for confirmation first.

## Verification

The local Node test suite covers all 14 requested behaviors: role admission,
inactive/unsupported denial, teacher auto-scoping, selector visibility,
administrator selection and switching, state isolation, payload compatibility,
submitter identity, and duplicate prevention. No production SharePoint writes,
deployments, or Entra changes were made.

Live verification is still required for an authorized administrator account,
the real roster teacher values, `IEP_Daily_Pulse` reads/writes, and physical
iPad/Chromebook interaction.
