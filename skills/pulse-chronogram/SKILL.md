---
name: pulse-chronogram
description: >
  Generate a Yasdu-branded weekly chronogram / Gantt-style spreadsheet
  (.xlsx) for tasks in a Pulse, with epics and tasks listed down the left
  and months-then-weeks across the top, shading each week a task is active.
  Use whenever the user asks for a "chronogram", "cronograma", a Gantt
  chart, a weekly timeline, or a visual schedule of Pulse tasks — even if
  they don't name this skill directly. If the user doesn't specify which
  tasks or epics to include, always ask before pulling data or generating
  anything.
---

# Pulse Chronogram

Produces a landscape spreadsheet: task rows grouped under bold epic-header
rows on the left, week columns (grouped under merged month headers) across
the top, with each cell shaded solid orange for every week the task is
active. Visual style is fixed to the Yasdu brand (see Step 4) — there is no
style question to ask; only scope is elicited.

## Step 1 — Scope: which tasks/epics to include

If the user's request already specifies the scope (an epic name, "only
stalled tasks", "all tasks", a specific task list, etc.), use that directly
and skip to Step 2.

**Otherwise, always ask before doing anything else.** Fetch the epic list
for the relevant Pulse first (`Pulse:get_pulse`, or reuse data already in
context), then use `ask_user_input_v0`:

- One question, options built from the actual epics present, plus an "All
  epics" option and an "Let me specify tasks" option if that seems likely to
  fit the request better than epic-level filtering.
- If the user picks specific epics, that's your task filter for Step 3.

Don't guess at scope and don't silently default to "all tasks" — a
full-roadmap chronogram and a single-epic one serve very different
purposes.

## Step 2 — Time window

Default to the next 8 weeks from today unless the user's request states a
different window (e.g. "next quarter", "this month", "next 4 weeks", a date
range). You don't need to ask about this separately — state the default
you're using in your reply so the user can correct it if it's wrong.

## Step 3 — Pull and filter task data

1. `Pulse:get_pulse` (or `Pulse:search_tasks` filtered by epic, if the scope
   from Step 1 narrows to specific epics) for the chosen Pulse — gives every
   task's `epic`, `status`, `startDate`, `endDate`.
2. Compute the window: Monday of the week containing "today" through the end
   of the Nth week out (see `references/input_schema.md` for exactly how the
   generator does this — mirror it here for your own filtering math).
3. Keep only tasks whose `[startDate, endDate]` overlaps that window
   (`start <= windowEnd AND end >= windowStart`) — a task doesn't need to
   fit entirely inside the window, just intersect it.
4. If Step 1's scope was epic-based, also drop tasks outside those epics
   (should already be filtered if you used `search_tasks` with an epic
   filter).

## Step 4 — Build the input JSON and generate

Read `references/input_schema.md` for the exact shape and
`references/styles.md` for the fixed Yasdu style. Group filtered tasks by
epic, preserve a sensible `epicOrder` (e.g. descending task count, or the
order epics appeared in `get_pulse`). Write the JSON to a scratch file, then:

```bash
python3 scripts/generate_chronogram.py <input.json> <output.xlsx> <skill_dir>
```

(`<skill_dir>` is this skill's root folder — needed so the script can find
the `assets/yasdu/` logo; omit it and the script infers its own location,
which usually works too.)

This always produces the Yasdu-branded spreadsheet: merged month headers in
navy, week headers with dd/mm date ranges in orange, epic section rows in
tan, alternating row shading, solid orange fill on every week a task is
active, a legend row, frozen panes below the header, and page setup already
set to landscape / fit-to-width so it prints or exports cleanly on one page
width. There is no other style to choose.

## Step 5 — Verify before sending

Convert to PDF and render to an image to check visually — column headers not
clipped at page edges, month header labels fully visible, shading lines up
with the dates you expect for at least two or three spot-checked tasks, and
the legend doesn't overlap the swatch:

```bash
soffice --headless --convert-to pdf <output.xlsx>
pdftoppm -jpeg -r 130 <output>.pdf page
```

View the resulting page images. If a month label clips at the page edge, the
fix used previously is shortening month names (e.g. "Ago" not "Agosto") — do
that via `monthNames` in the input JSON rather than editing the script.

## Step 6 — Present the file

Copy the .xlsx to the outputs directory and use `present_files`. Mention in
one line what scope and window it covers (e.g. "31 tasks across 8 epics,
next 8 weeks") — no long postamble.
