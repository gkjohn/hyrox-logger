# HYROX Logger — Coach API Spec

This document describes how a **Coach** (human or AI) collaborates with
the athlete to revise the training plan stored in the athlete's
Google Sheet.

The athlete uses the [HYROX Logger app](https://github.com/gkjohn/hyrox-logger)
to log actuals against the plan. The plan itself lives in the Sheet
(`Dashboard`, `RunLog`, `KBLog`, `StationLog`, `Benchmarks`, `HRVLog`).
The Coach reads everything in the Sheet, decides on revisions, and
emits a **structured JSON payload**. The athlete pastes that payload
into the HYROX app, previews the diff, and clicks Apply.

The app writes the overrides to dedicated `*Override` columns; the
original prescription is preserved. The display layer surfaces the
override with a yellow banner + struck-through original (see v6.1.7a).

## Why paste-and-apply (not direct write)

Coach Claude — the typical AI coach — runs in the Claude chat
interface with **read-only** Sheets access. It can't directly mutate
the Sheet. The paste-and-apply pattern works regardless of what tools
the Coach has, keeps the athlete in the loop (preview before apply),
and needs no shared secrets or HTTP endpoints. One paste per coach
review (usually weekly).

## Identity

- **Author**: Coach includes `"author"` in the JSON payload — written
  to the `OverrideBy` column. Free text. Conventional values:
  `"Claude Coach"` for the AI; the human coach's name otherwise.
- **Date stamping**: The app stamps `OverrideDate` automatically with
  today (in IST, the athlete's local timezone). Coach doesn't set it.

## The JSON payload contract

Coach emits a single JSON object in its chat response. The athlete
copies the block (just the JSON, no surrounding text) and pastes it
into the HYROX app's **Data → Apply Coach Update** panel.

### Minimal shape

```json
{
  "author": "Claude Coach",
  "reason": "Why these changes — 1-3 sentences. Shown to athlete in every banner.",
  "overrides": [
    {
      "sheet": "<RunLog | KBLog | StationLog>",
      "match": { "<keyColumn>": "<value>", ... },
      "set":   { "<*Override column>": "<new value>", ... }
    }
  ]
}
```

### Field semantics

- `author` (string, required): goes into `OverrideBy` for every row
  touched.
- `reason` (string, required): goes into `OverrideReason` for every
  row touched. One reason covers the whole payload — group thematically
  related changes into a single payload, send separate payloads if
  reasons differ.
- `overrides` (array, required): one entry per row being modified.

### Each override entry

- `sheet`: exactly one of `"RunLog"`, `"KBLog"`, `"StationLog"`. Other
  sheet names rejected.
- `match`: object whose keys are column names used to locate the
  unique row. Required keys by sheet:
  - **RunLog**: `Week` (integer 1–13), `Day` (`"Tuesday"` /
    `"Thursday"` / `"Saturday"`).
  - **KBLog**: `Week`, `Day` (`"Monday"` / `"Friday"`), `Movement`
    (the original movement name as written in the sheet).
  - **StationLog**: `Week`.

  The match must locate **exactly one row**. Zero or multiple matches
  is rejected with an error naming the bad match.

- `set`: object whose keys are column names to write. Only the
  whitelisted `*Override` columns are accepted. Writes to other
  columns (`Actual*`, `Skipped*`, originals like `PlannedDist`,
  `Movement`, `SessionType`) are **rejected wholesale** — the payload
  is rejected without writes.

## Whitelisted columns Coach can set

These are the only columns acceptable in `set`. The app stamps
`OverrideBy`, `OverrideDate`, `OverrideReason` automatically — do not
include them.

### RunLog
- `PlannedDistOverride` — text (e.g. `"7km easy"`)
- `TargetPaceOverride` — text (e.g. `"6:40/km"`)
- `HRCapOverride` — number or text (e.g. `150`)

### KBLog
- `MovementOverride` — text (swap the movement, e.g. `"2H Swing"` to
  replace `"1H Swing"`)
- `PlannedSetsRepsKgOverride` — text (e.g. `"5x8 @ 24kg"`). Use `x`
  not `×` for compatibility.

### StationLog
- `SessionTypeOverride` — text (e.g. `"Easy Run 7km (Wed substitute)"`)
- `StationsOverride` — text describing changed stations, free form
  (e.g. `"skip stations this week"` or
  `"SkiErg 2x200m easy, Row 2x200m easy, no sled"`)

## Worked examples

### Example 1 — Tuesday pace adjustment

HRV trending down; back off W5 Tuesday rolling 400s pace.

```json
{
  "author": "Claude Coach",
  "reason": "Avg HRV 23ms (down from 28ms baseline last week). Back off pace 30 sec/km to keep this aerobic rather than threshold. Resume original targets next week if HRV recovers.",
  "overrides": [
    {
      "sheet": "RunLog",
      "match": { "Week": 5, "Day": "Tuesday" },
      "set": { "TargetPaceOverride": "6:00-6:30/km" }
    }
  ]
}
```

### Example 2 — KB exercise swap

Left wrist sore; substitute W7 Monday 1H Swing with bilateral 2H Swing.

```json
{
  "author": "Claude Coach",
  "reason": "Wrist soreness reported May 21. Substitute bilateral swing to reduce unilateral strain. Total volume preserved.",
  "overrides": [
    {
      "sheet": "KBLog",
      "match": { "Week": 7, "Day": "Monday", "Movement": "1H Swing" },
      "set": {
        "MovementOverride": "2H Swing",
        "PlannedSetsRepsKgOverride": "5x8 @ 24kg"
      }
    }
  ]
}
```

### Example 3 — Whole-Wednesday reroute

W8 Hard Brick out of bounds given readiness; substitute easy run.

```json
{
  "author": "Claude Coach",
  "reason": "Oura readiness below 70 for 3 consecutive nights. Skip the hard brick — do an easy outdoor run instead. Resume hard sessions when readiness > 75 for 2 nights.",
  "overrides": [
    {
      "sheet": "StationLog",
      "match": { "Week": 8 },
      "set": {
        "SessionTypeOverride": "Easy Run 7km (Wed substitute)",
        "StationsOverride": "skip stations this week"
      }
    }
  ]
}
```

### Example 4 — Multiple changes in one payload

Coach reviews a week and adjusts two sessions for the same underlying
reason. Group both in one payload (same reason applies).

```json
{
  "author": "Claude Coach",
  "reason": "Recovery markers soft — back off W5 quality, hold KB at last week's loads instead of progressing.",
  "overrides": [
    {
      "sheet": "RunLog",
      "match": { "Week": 5, "Day": "Tuesday" },
      "set": { "TargetPaceOverride": "6:00-6:30/km" }
    },
    {
      "sheet": "KBLog",
      "match": { "Week": 5, "Day": "Monday", "Movement": "KB Press" },
      "set": { "PlannedSetsRepsKgOverride": "3x6e @ 20kg" }
    },
    {
      "sheet": "KBLog",
      "match": { "Week": 5, "Day": "Monday", "Movement": "KB Row" },
      "set": { "PlannedSetsRepsKgOverride": "3x6e @ 24-28kg" }
    }
  ]
}
```

### Example 5 — Clearing a previous override

To revert, set the override field to an empty string. Use the reason
to document the reversal.

```json
{
  "author": "Claude Coach",
  "reason": "HRV recovered (avg 27ms over last 3 nights). Reverting Tuesday pace back to original.",
  "overrides": [
    {
      "sheet": "RunLog",
      "match": { "Week": 5, "Day": "Tuesday" },
      "set": { "TargetPaceOverride": "" }
    }
  ]
}
```

## What the athlete does

1. Asks Coach Claude something like: *"Review my last 7 days. Any
   prescription changes for next week? Output as a JSON override block."*
2. Coach Claude reads the Sheet + Strava + Oura, reasons, emits the
   JSON in chat.
3. Athlete copies the JSON.
4. Athlete opens the HYROX app → **Data** tab → **Apply Coach Update**
   panel → pastes the JSON → taps **Preview**.
5. Preview expands to a plain-English diff (rows + columns that will
   change, old → new). Athlete sanity-checks.
6. Athlete taps **Apply**. Overrides land in the Sheet. An entry is
   recorded in the `CoachLog` sheet for audit.
7. Affected session cards in the Week view immediately show the
   yellow Coach badge and the override is visible in the form.

## Validation Coach should anticipate

The app validates before applying. These all reject the entire
payload (no partial writes):

- JSON doesn't parse → `"json parse error: <message>"`
- Missing required top-level keys → `"missing field: <name>"`
- `sheet` not in allowlist → `"unknown sheet: <value>"`
- `match` keys missing or unknown → `"match needs: <keys>"`
- `match` finds zero rows → `"no row found for match: <criteria>"`
- `match` finds multiple rows → `"match ambiguous (N rows): <criteria>"`
- `set` references a non-override column → `"cannot write to '<col>' — only *Override columns are writable"`
- `set` empty → `"nothing to set"`

If the athlete pastes an invalid payload, the Preview shows the error
verbatim. They can ask Coach Claude to fix and try again.

## Audit log

Every applied payload appends a row to the `CoachLog` sheet:

| Timestamp | Author | Reason | AppliedCount | RowsAffected | PayloadJSON |
|---|---|---|---|---|---|
| 2026-05-22 18:34:01 | Claude Coach | HRV down — back off pace | 1 | RunLog W5 Tuesday | `{...full JSON...}` |

This is just an audit trail — the app doesn't read from it. Useful
for retrospective review ("when did we change the W5 pace, and why?").

## Coach Claude prompt template

Suggested system / message prompt the athlete can drop into a coach
chat to get well-formed JSON output:

> You are my training coach for the HYROX 13-week V6.1 program. You
> have read access to my training Sheet, Strava, and Oura.
>
> When I ask for recommendations, you will:
> 1. Read the most recent week's actuals + HRV + Oura data.
> 2. Identify any prescription changes worth making — adjust paces,
>    swap movements, reduce volume, or skip sessions.
> 3. Emit your recommendations as a single JSON code block matching
>    the schema in `coach-api.md` (sheet, match, set; whitelisted
>    *Override columns only; one shared reason per payload).
> 4. Briefly explain what you changed and why in prose above the
>    JSON block, but the JSON itself should be standalone and
>    paste-ready.
>
> Don't include `OverrideBy`, `OverrideDate`, or `OverrideReason` in
> `set` — those are stamped automatically.

## Invariants Coach must respect

1. Never write to `Actual*` columns.
2. Never write to `Skipped` or `SkipReason`.
3. Never write to original `Planned*` / `Movement` / `SessionType`
   columns — only their `*Override` siblings.
4. Don't include `OverrideBy` / `OverrideDate` / `OverrideReason` in
   `set` — the app stamps them.
5. Group changes that share a reason into a single payload. Split
   payloads when reasons differ — preserves audit-log clarity.
6. The HRVLog, Dashboard, Benchmarks sheets are read-only to the
   Coach API.

## Versioning

This spec applies to **app v6.1.7b and above** — when the override
columns, the display layer, and the paste-and-apply panel are all
live. v6.1.7a has the display layer but no apply-from-UI, so writes
have to be typed into cells manually.
