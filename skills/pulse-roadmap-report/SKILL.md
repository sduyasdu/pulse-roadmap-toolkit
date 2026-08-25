---
name: pulse-roadmap-report
description: >
  Generate a Yasdu-branded status report (Word or PDF) for a Pulse project,
  grouping tasks by status (Completed / Ongoing / Stalled / Planned) and then
  by epic, with leads, start/end dates, and bulleted subtasks per task. Use
  this skill whenever the user asks for a roadmap status report, a Pulse
  report for management, a project status document, or wants to "reportar el
  roadmap" / share progress on a Pulse — even if they don't name this skill
  directly. Always trigger it instead of building an ad-hoc report by hand
  when the request involves summarizing Pulse task data by status/epic into a
  shareable document. Before producing anything, this skill always asks the
  user three things: report language, output filetype (Word or PDF), and
  which Pulse/project to report on — do not skip or assume these.
---

# Pulse Roadmap Report

Produces a Yasdu-branded status report: tasks grouped by status, then by
epic, each row showing the task, its lead, start/end dates, and its
subtasks as a bulleted list — plus a closing "key risks" callout flagging
unowned or stalled work. Visual style is fixed to the Yasdu brand (see Step
5) — there is no style question to ask; only language, filetype, and Pulse
selection are elicited.

## Step 1 — Ask the three required questions

Always ask these up front, even if the user's request seems to imply an
answer. Use `ask_user_input_v0` (or your platform's equivalent elicitation
tool) with one call:

1. **Language** — options: common languages for the user's Pulse (e.g.
   English, Spanish), plus an "Other" the user can type into.
2. **Filetype** — options: "Word (.docx)", "PDF".
3. **Pulse / project** — call `Pulse:list_pulses` first, then offer each
   returned Pulse name as an option (skip archived ones unless the user asks
   for one). If there's only one Pulse available, you can still confirm it
   rather than silently assuming.

Wait for the user's answers before doing any data pulling or file
generation.

## Step 2 — Pull the Pulse data

1. `Pulse:get_pulse` with the chosen `pulseId` — gives you every task with
   epic, status, dates, and assignees (including each assignee's `isLead`
   flag — that's your Lead column, no separate lookup needed).
2. `Pulse:search_tasks` with just the `pulseId` (no filters, `limit: 50` or
   the pulse's task count if higher) — this is the call that returns each
   task's `subtasks` array (title + status). `get_pulse` does not include
   subtasks, so don't skip this call. If the pulse has more tasks than one
   call's limit allows, page through by status instead (call once per
   distinct status value).
3. Merge: for each task from `get_pulse`, attach the `subtasks` array found
   for the same `taskId`/title in the `search_tasks` results.

## Step 3 — Decide the status buckets and group by epic

Inspect the distinct `status` values actually present. Map them into four
buckets in this order — adjust bucket labels to match the language chosen in
Step 1 (see `references/labels.md`):

- **Completed** — status `Done` (or equivalent)
- **Ongoing** — everything actively being worked (e.g. `Dev`, `QA`,
  `Design`) — keep a Stage column in this section only
- **Stalled** — status `Stalled`/`Blocked` (render this whole section in
  warning red — see `colorOverride` in the schema)
- **Planned** — status `Planned`/`Backlog`/not yet started

If the pulse's status vocabulary doesn't map cleanly onto these four, ask the
user how to bucket them rather than guessing.

Within each bucket, group tasks by `epic` (tasks with no epic go in a "No
epic" group, kept as its own group — see `references/labels.md` for the
localized label). Order epic groups by descending task count, or reuse
whatever order the user's Pulse conventionally uses if they mention one.

For **Completed** tasks, show Start Date and Finished Date (use `finishedOn`
if present, otherwise fall back to `endDate`). For all other buckets show
Start Date and End Date, plus a Stage column for Ongoing.

For the **Subtasks** cell of each task, list subtask titles as plain
strings; append the localized "(done)" suffix (see `references/labels.md`)
to any subtask whose own status is Done. A task with no subtasks gets an
empty list — the renderer shows "—" automatically.

Determine the **Lead** for each task from its assignees' `isLead: true`
flag; if no assignee is marked lead, use "—".

## Step 4 — Build the "key risks" note

Scan the grouped data for:
- Any epic/group with zero assignees across all its tasks in the Ongoing or
  Planned buckets (an unowned epic).
- Any tasks in the Stalled bucket.

Summarize these in one or two sentences in the target language. Skip this
note if nothing stands out.

## Step 5 — Assemble the input JSON and generate the document

Read `references/input_schema.md` for the exact JSON shape, and
`references/labels.md` for label translations (English and Spanish given
directly; translate yourself for any other requested language — keep task
titles, epic names, and people's names as-is, never translate proper nouns).

Write the JSON to a scratch file, then run:

```bash
node scripts/generate_report.js <input.json> <output.docx> assets/yasdu/logo-horizontal-light.png
```

This always produces a Word document in the Yasdu brand style: orange
(#D85A28) headings and table headers, navy (#123359) epic sub-headings and
key-risks text, the Yasdu logo in the header, a tan key-risks callout,
red-flagged Stalled section, one-bullet-per-line subtasks, and no table rows
splitting across a page break. There is no other style to choose — always
pass the Yasdu logo path shown above.

## Step 6 — Convert to PDF if requested

If the user chose PDF in Step 1, convert the generated .docx:

```bash
soffice --headless --convert-to pdf <output.docx>
```

(Use the same LibreOffice invocation documented in the docx skill if
`soffice` isn't directly on PATH in this environment.) Present the PDF
instead of the docx in that case.

## Step 7 — Verify before sending

Render the docx/pdf to images (e.g. `pdftoppm`) and view at least the first
page and one page with a dense subtasks list, to confirm: logo present,
tables not splitting rows across pages, dates each fit on one line, Stalled
rows in red, key-risks callout rendered. Fix and rebuild before presenting
if anything looks off.

## Step 8 — Present the file

Copy the final file to the outputs directory and use `present_files`. Don't
add a long postamble — a one-line summary of what's in the report (task
counts per bucket, anything flagged in key risks) is enough.
