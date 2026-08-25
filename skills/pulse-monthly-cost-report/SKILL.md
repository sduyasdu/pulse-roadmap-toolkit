---
name: pulse-monthly-cost-report
description: >
  Generate a 4-sheet Excel workbook (.xlsx) that lists a Pulse's tasks for a
  given month with resources assigned and hours worked, applies a monthly
  hour cap per resource (proportionally reducing hours on each of their
  tasks if they'd otherwise exceed it), and computes cost per task line, per
  resource, and per epic from a monthly-cost input the user fills in. Use
  this whenever the user asks for a monthly hours report, a resource cost
  breakdown, capacity/utilization numbers for a Pulse, or wants to know what
  a month of roadmap work costs — even if they don't name this skill
  directly. Always ask which Pulse and which month before pulling data or
  generating anything, unless the request already specifies both.
---

# Pulse Monthly Hours & Cost Report

Produces a Yasdu-branded workbook with:
1. **Detalle `<mes>`** — one row per task-resource pair: business days
   worked within the month, uncapped hours, the resource's uncapped monthly
   total, an adjustment factor, capped hours, hourly rate, and line cost.
2. **Costo por Recurso** — one row per resource with a yellow input cell for
   monthly cost; hourly rate is auto-computed as `monthly cost / cap hours`.
3. **Resumen por Recurso** — capped/uncapped hour totals, whether the cap
   was hit, and total cost per resource.
4. **Resumen por Epic y Tarea** — hours and cost rolled up by epic and by
   task.

Every number in the workbook is a live formula (NETWORKDAYS, SUMIFS,
INDEX/MATCH) — nothing is precomputed and pasted in as a static value, so
the whole sheet recalculates the moment someone fills in a monthly cost.
Visual style is fixed to the Yasdu brand (logo top-left on every sheet,
orange/navy palette) — there is no style question to ask.

## Step 1 — Ask what's needed

Ask before pulling any data unless the request already specifies these:

1. **Which Pulse.** Call `Pulse:list_pulses` and offer the names.
2. **Which month.** Default to the current month if the user doesn't say,
   but confirm rather than assume — state the month you're using in your
   reply either way.

The monthly hour cap defaults to 160 — only ask about it if the user's
request hints at a different figure (e.g. "cap at 150" or "part-time
resources"); otherwise just use 160 and mention the default in your reply.

## Step 2 — Pull task data and filter to the month

1. `Pulse:get_pulse` for the chosen Pulse — gives every task's `epic`,
   `status`, `startDate`, `endDate`, and `assignees` (with
   `allocationPercent`).
2. Keep only tasks whose `[startDate, endDate]` overlaps the target month
   (`start <= monthEnd AND end >= monthStart`) — a task doesn't need to fit
   entirely inside the month, just intersect it.
3. Tasks with no assignees still count — keep them with an empty
   `assignees` list so unstaffed work stays visible in the report rather
   than disappearing.

## Step 3 — Build the input JSON and generate

Read `references/input_schema.md` for the exact shape. Write the JSON to a
scratch file, then:

```bash
python3 scripts/generate_hours_report.py <input.json> <output.xlsx> <skill_dir>
```

(`<skill_dir>` is this skill's root folder — needed so the script can find
the `assets/yasdu/` logo; omit it and the script infers its own location,
which usually works too.)

## Step 4 — Recalculate and verify

This workbook uses formulas (NETWORKDAYS, SUMIFS, INDEX/MATCH) — always run
the xlsx skill's recalculation step before delivering:

```bash
python3 /mnt/skills/public/xlsx/scripts/recalc.py <output.xlsx> 90
```

Confirm `status: success` and `total_errors: 0`. If you want to sanity-check
the cost chain before delivering, temporarily fill in one or two monthly
costs in a scratch copy, recalc, and read back a few `Detalle` rows and the
resource/epic summaries with `data_only=True` to confirm the numbers flow
through correctly — then discard that scratch copy and ship the version
with the cost column left blank for the user to fill in themselves.

Also render to PDF and check visually (landscape, fits page width, logo
present on each sheet, yellow cost column visible, no clipped headers):

```bash
soffice --headless --convert-to pdf <output.xlsx>
pdftoppm -jpeg -r 110 <output>.pdf page
```

## Step 5 — Present the file

Copy to the outputs directory and use `present_files`. Mention in one line:
task/resource counts, how many resources hit the cap (if any figures are
already known), and that costs will read $0 until the yellow column is
filled in — no long postamble.
