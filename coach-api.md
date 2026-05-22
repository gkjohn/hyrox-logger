# HYROX Logger — Coach API Spec

This document describes how a **Coach** (human or AI) can write structured
revisions to the athlete's training plan stored in Google Sheets.

The athlete logs actuals against the plan via the [HYROX Logger
app](https://github.com/gkjohn/hyrox-logger). The plan itself lives in a
Google Sheet (sheets: `Dashboard`, `RunLog`, `KBLog`, `StationLog`,
`Benchmarks`, `HRVLog`). The Coach reads everything in the Sheet and writes
**overrides** into dedicated columns. The app displays the original
prescription with the override layered on top (struck-through original +
new value + reason + date).

## Identity & access

The Coach is expected to have **read-only Google Sheets access** (e.g.
via Claude's MCP tooling). It reads athlete data freely but **cannot
write to the Sheet directly**.

For writes, the Coach calls the **HYROX logger's Apps Script web app
endpoint** via HTTP POST. The Apps Script runs as the athlete (sheet
owner) and applies the writes on the Coach's behalf. This is the only
supported write path.

- **Read path**: Direct Sheets API. Read any sheet, any column.
- **Write path**: HTTPS POST to the Apps Script web app URL. Schema below.
- **Author identity**: Coach includes `"author": "Claude Coach"` in the
  request body; the endpoint writes this into `OverrideBy`.

## Date convention

- All dates in the sheet use `yyyy-MM-dd` (e.g. `2026-05-22`).
- Timezone is the spreadsheet's configured timezone (Asia/Kolkata).
- For dates being written via Apps Script or a script with `Utilities.formatDate`,
  use the spreadsheet timezone, not UTC.

## Source of truth

- **Planned values** (PlannedDist, TargetPace, HRCap, PlannedSetsRepsKg,
  Movement) are populated by the app's `setupSheets()` and reflect the
  V6.1 program. **Do not overwrite these directly.**
- **Override values** (the `*Override` columns) are how the Coach makes
  changes. Original planned values are preserved for audit and the
  athlete can always see what changed.
- **Actuals** (ActualDist, ActualTime, ActualSets, etc.) are the
  athlete's logged results. **Do not write to these.**
- **Skipped / SkipReason** are the athlete's own intentional skips. **Do
  not write to these.** A coach who wants a session removed should use
  `OverrideReason` + the appropriate override columns (e.g.
  `PlannedDistOverride = "skip — coach-directed rest"`).

## Sheet & column schema

The columns marked **NEW** are added by the
`migrateAddOverrideColumns()` Apps Script function and do not exist on
older installs.

### Sheet: `RunLog` — Tuesday / Thursday / Saturday running

| Column | Type | Read/Write | Notes |
|---|---|---|---|
| Week | int | read | 1..13 |
| Day | text | read | `Tuesday` / `Thursday` / `Saturday` |
| Date | yyyy-MM-dd | read | Set by athlete when logging |
| SessionType | text | read | e.g. `Easy Run`, `Fast 8-4-2s`, `5km TIME TRIAL` |
| Role | text | read | `RECOVERY`, `HIGH COST`, `CALIBRATION` |
| PlannedDist | text | read | e.g. `6.5km easy`. Original program prescription. |
| TargetPace | text | read | e.g. `5:00-5:30/km` |
| HRCap | int or text | read | e.g. `148` |
| ActualDist..MaxHR..Notes | various | read only | Athlete logs |
| Skipped, SkipReason | text | **do not write** | Athlete's own skips |
| **PlannedDistOverride** | text | write | NEW. Replaces PlannedDist for display. |
| **TargetPaceOverride** | text | write | NEW. Replaces TargetPace for display. |
| **HRCapOverride** | int or text | write | NEW. Replaces HRCap for display. |
| **OverrideBy** | text | write | NEW. e.g. `"Claude Coach"`. |
| **OverrideDate** | yyyy-MM-dd | write | NEW. Date the override was set. |
| **OverrideReason** | text | write | NEW. 1–2 sentences. Why this change? |

**Row identification**: a unique row in RunLog is `(Week, Day)`.

### Sheet: `KBLog` — Monday / Friday KB sessions

| Column | Type | Read/Write | Notes |
|---|---|---|---|
| Week | int | read | |
| Day | text | read | `Monday` / `Friday` |
| Date | yyyy-MM-dd | read | |
| Role | text | read | `SUPPORT` / `ACCESSORY` |
| SessionLabel | text | read | e.g. `Foundation - Mon KB` |
| Movement | text | read | e.g. `2H Swing`, `KB Goblet Squat`, `TGU`. **One row per movement.** |
| PlannedSetsRepsKg | text | read | e.g. `10×10 @ 24kg`, `3×6e @ 20-24kg` |
| ActualSets..Notes | various | read only | Athlete logs |
| Skipped, SkipReason | text | **do not write** | |
| **MovementOverride** | text | write | NEW. Swap the movement entirely. e.g. `2H Swing` to replace `1H Swing`. |
| **PlannedSetsRepsKgOverride** | text | write | NEW. e.g. `5×8 @ 24kg`. |
| **OverrideBy** | text | write | NEW. |
| **OverrideDate** | yyyy-MM-dd | write | NEW. |
| **OverrideReason** | text | write | NEW. Applies to all overrides on this row. |

**Row identification**: a unique row in KBLog is `(Week, Day, Movement)` —
or `(Week, Day, row position in the day's session)` if the same movement
appears twice on the same day (rare). When writing a Movement override,
preserve the original Movement column — only set `MovementOverride`.

### Sheet: `StationLog` — Wednesday stations

| Column | Type | Read/Write | Notes |
|---|---|---|---|
| Week | int | read | |
| Date | yyyy-MM-dd | read | |
| SessionType | text | read | e.g. `⭐ Benchmark TT + Sled Pull Race Load` |
| Role | text | read | `RECOVERY` / `SUPPORT` / `HIGH COST` / `SIM` |
| SkiErg..WallBalls..Notes | various | read only | Athlete logs |
| SledPullRaceLoad, SledPullTimes | various | read only | |
| Skipped, SkipReason | text | **do not write** | |
| **SessionTypeOverride** | text | write | NEW. Replace the whole session, e.g. `Easy Run 7km`. |
| **StationsOverride** | text or JSON | write | NEW. Either free-text description (`"SkiErg 2×200m easy, no sled pull this week"`) or JSON map (`{"SkiErg":"2×200m easy","SledPull":"skip","Row":"2×200m easy"}`). Free-text is fine. |
| **OverrideBy** | text | write | NEW. |
| **OverrideDate** | yyyy-MM-dd | write | NEW. |
| **OverrideReason** | text | write | NEW. |

**Row identification**: a unique row in StationLog is `Week`.

### Sheets NOT to write to

- `Dashboard` — read-only summary; updated by the app.
- `Benchmarks` — athlete's TT data.
- `HRVLog` — daily HRV data. The Coach SHOULD read this to inform
  decisions.

## Examples

### Example 1: Tuesday tempo pace adjustment

Coach Claude notices HRV has dipped after W4 hard week. Wants W5
Tuesday's Rolling 400s prescribed at 6:00-6:30/km instead of the
original 5:30-6:00/km.

**Locate row**: RunLog where `Week == 5` and `Day == "Tuesday"`.

**Write these columns**:

```
TargetPaceOverride  = "6:00-6:30/km"
OverrideBy          = "Claude Coach"
OverrideDate        = "2026-05-22"
OverrideReason      = "Avg HRV 23ms (down from 28ms baseline). Back off pace 30 sec/km to keep this an aerobic stim rather than threshold."
```

Leave all other columns alone. `TargetPace` still reads `5:30-6:00/km`.
`PlannedDist` still reads `6×400m + 200m jog (6km)`.

### Example 2: Swap a KB exercise

Left wrist flared up. Replace W7 Monday's `1H Swing` with `2H Swing`
at the same volume.

**Locate row**: KBLog where `Week == 7` and `Day == "Monday"` and
`Movement == "1H Swing"`.

**Write**:

```
MovementOverride            = "2H Swing"
PlannedSetsRepsKgOverride   = "5×8 @ 24kg"
OverrideBy                  = "Claude Coach"
OverrideDate                = "2026-05-22"
OverrideReason              = "Wrist soreness reported on May 21. Substitute bilateral swing to reduce unilateral strain. Volume preserved."
```

### Example 3: Reduce KB volume only

Carrying cumulative load — drop a set from W6 Monday's `KB Press`.

```
PlannedSetsRepsKgOverride   = "2×8 @ 24kg"
OverrideBy                  = "Claude Coach"
OverrideDate                = "2026-05-22"
OverrideReason              = "Cumulative shoulder load high after W5 ownership block. Drop one set, hold weight."
```

Leave `MovementOverride` empty.

### Example 4: Reroute the whole Wednesday

W8 Hard Brick is too aggressive given current Oura readiness.
Substitute with an easy outdoor run.

```
SessionTypeOverride = "Easy Run 7km (Wed substitute)"
StationsOverride    = "skip stations this week"
OverrideBy          = "Claude Coach"
OverrideDate        = "2026-05-22"
OverrideReason      = "3 consecutive nights of readiness below 70. Skip the hard brick, do an easy run instead. Resume hard sessions when readiness > 75 for 2 nights."
```

### Example 5: Clearing an override

Coach reverses a previous decision. Just set the override field(s) back
to empty string `""`. The app falls back to the original PlannedX value.

## Write API — HTTPS POST to the Apps Script endpoint

### Endpoint

```
POST https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec
Content-Type: application/json
```

The deployment URL is the same one the athlete uses for the logger
web app. The endpoint is shared between the human-facing UI (`doGet`)
and the Coach API (`doPost`). Athletes share this URL with their Coach.

### Authentication

Each request must include a shared secret token. The athlete sets the
token once in Apps Script's `PropertiesService` (instructions in repo
README). The token is passed in the request body, not a header (Apps
Script `doPost` doesn't see custom headers reliably).

### Request body

```json
{
  "token": "<shared-secret>",
  "author": "Claude Coach",
  "reason": "HRV dipped after W5 — back off pace this week. See HRVLog rows for May 19-22.",
  "overrides": [
    {
      "sheet": "RunLog",
      "match": {"Week": 5, "Day": "Tuesday"},
      "set": {
        "TargetPaceOverride": "6:00-6:30/km"
      }
    },
    {
      "sheet": "KBLog",
      "match": {"Week": 7, "Day": "Monday", "Movement": "1H Swing"},
      "set": {
        "MovementOverride": "2H Swing",
        "PlannedSetsRepsKgOverride": "5x8 @ 24kg"
      }
    }
  ]
}
```

**Field semantics:**

- `token` (string, required): shared secret. Request rejected with 401
  if missing/wrong.
- `author` (string, required): goes into the `OverrideBy` column for
  every row touched in this request. Free text.
- `reason` (string, required): goes into the `OverrideReason` column
  for every row touched. A single reason covers all overrides in the
  request — keep changes thematically grouped per request.
- `overrides` (array): one entry per row being modified.
  - `sheet`: one of `"RunLog"`, `"KBLog"`, `"StationLog"`. Other
    sheets rejected.
  - `match`: object whose keys are column names in that sheet, used
    to find the unique row. RunLog needs `Week` + `Day`; KBLog needs
    `Week` + `Day` + `Movement`; StationLog needs `Week`.
  - `set`: object whose keys are column names to write. Only
    `*Override` columns are accepted; writes to `Actual*`,
    `Skipped*`, original `Planned*` / `Movement` / `SessionType`
    are rejected.

The endpoint stamps `OverrideBy`, `OverrideDate` (today in IST), and
`OverrideReason` automatically on every row touched — the Coach
doesn't need to set these explicitly.

### Response

```json
{
  "status": "ok",
  "applied": 2,
  "rows": [
    {"sheet": "RunLog", "match": {"Week": 5, "Day": "Tuesday"}, "row": 17},
    {"sheet": "KBLog", "match": {"Week": 7, "Day": "Monday", "Movement": "1H Swing"}, "row": 64}
  ]
}
```

Or on error:

```json
{
  "status": "error",
  "error": "row not found for match {Week: 99, Day: Tuesday} in RunLog",
  "applied": 0
}
```

If any single override fails, the whole request is rejected (no
partial writes).

### Clearing overrides

Send `set` with empty-string values:

```json
{
  "sheet": "RunLog",
  "match": {"Week": 5, "Day": "Tuesday"},
  "set": {"TargetPaceOverride": ""}
}
```

Use `reason: "Reverting earlier override — HRV recovered"` (or
similar) for the audit trail.

### Idempotency

Sending the same request twice is safe — the endpoint just rewrites
the same values. Same `OverrideDate` if same day. If you need to
detect "is this already applied?", read the Sheet first (which Coach
already has read access to).

### Locating cells by header, not index

Internally the endpoint uses `headers.indexOf("ColumnName")` to find
columns, so the Coach should never need to know which column index
holds `TargetPaceOverride`. Just reference columns by name.

### Rate limits

Apps Script has soft limits (~20k API calls/day, 6 min per execution).
For this use case (one Coach making a few writes per week), the limit
is not a concern. Avoid bulk-rewriting the whole program in a single
request — split by theme.

## App display behaviour (for athlete reference)

When the app loads a session:

- If any `*Override` column is non-empty, that value is shown as the
  primary display value.
- The original planned value is shown beneath, struck-through and in
  muted grey.
- The session card on the Week view shows a `📋 Coach update (Date)`
  badge.
- The form view shows a yellow-tinted panel above the inputs with the
  `OverrideReason` text.
- Setting an override does NOT affect the athlete's actuals or skip
  flags.

## Invariants the Coach must respect

1. Never write to `Actual*` columns.
2. Never write to `Skipped` or `SkipReason`.
3. Never modify the original `PlannedX` / `Movement` / `SessionType`
   values — only their `*Override` siblings.
4. Always set `OverrideBy`, `OverrideDate`, `OverrideReason` when
   setting any override field. Empty reasons aren't useful.
5. Clearing overrides is done by setting the override fields back to
   `""` (and ideally setting `OverrideDate` + `OverrideReason` to
   document the reversal).
6. The HRVLog, Dashboard, Benchmarks sheets are out of scope for writes.
   Read freely.

## Versioning

This spec applies to **app v6.1.7 and above** — when both the column
migration AND the `doPost` write endpoint are deployed. Earlier
versions don't have the override columns or the write endpoint;
attempts to call the API will return 404 or no-op.

## Athlete setup checklist (one-time, for v6.1.7)

1. Run `migrateAddOverrideColumns()` once from the Apps Script editor.
2. Generate a shared-secret token (any random string, e.g.
   `openssl rand -hex 16`).
3. In Apps Script editor → Project Settings → Script Properties → add
   `COACH_TOKEN` = `<your-token>`.
4. Re-deploy as a new version (`Deploy → Manage deployments → New
   version`).
5. Share the deployment URL + token with the Coach via a secure
   channel.
