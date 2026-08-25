# Pulse Roadmap Toolkit

A Claude Code / Claude Cowork plugin that connects to **Pulse**
(`pulse.yasdu.com`) — a project roadmap and resource-planning tool — and
generates three deliverables from live Pulse data:

| Skill | What it produces |
|---|---|
| `pulse-roadmap-report` | A Yasdu-branded status report (Word or PDF) grouping tasks by status (Completed/Ongoing/Stalled/Planned) then by epic, with leads, dates, and subtasks. Asks for report language, output filetype, and which Pulse to report on. |
| `pulse-chronogram` | A Yasdu-branded weekly chronogram/Gantt-style spreadsheet (.xlsx): epics and tasks down the left, months → weeks across the top, shaded cells for active weeks. Asks for task/epic scope. |
| `pulse-monthly-cost-report` | A Yasdu-branded 4-sheet Excel workbook: task/resource hours for a chosen month, a monthly-hours cap applied proportionally per resource, and cost computed from a monthly-cost input you fill in per resource. Asks which Pulse and which month. |

## Who this is for

This plugin is **exclusively for Pulse users** — it only works if your
organization has an account on `pulse.yasdu.com` and access to the Pulse
MCP server. It has no functionality outside of a Pulse account: every skill
depends entirely on the `Pulse:*` tools this plugin bundles to read epics,
tasks, dates, assignees, and allocations. If you don't use Pulse, this
plugin will not be useful to you.

## What's bundled

- **MCP server**: `pulse` (`https://pulse.yasdu.com/mcp`) — gives Claude the
  `Pulse:*` tools (`list_pulses`, `get_pulse`, `search_tasks`, etc.) these
  skills are built on. This is the only data source the plugin uses; there
  is no fallback for non-Pulse project data.
- **Three skills**, each self-contained with its own `SKILL.md`, a
  generator script (Node or Python), and reference docs describing the
  input format the script expects. All three always render their output in
  a fixed Yasdu brand style (logo, orange/navy palette) — there is no style
  choice to make; this has no bearing on who can use the plugin.

## Requirements

- Node.js (for `pulse-roadmap-report`'s `docx` generation) and Python 3 with
  `openpyxl` (for the two spreadsheet skills) available in the environment
  Claude is running in.
- An active account on `pulse.yasdu.com` with access to the Pulse(s) you
  want to report on.
- The MCP server requires the user to authenticate (OAuth) to their Pulse
  account on first use — this happens automatically the first time Claude
  tries to call a `Pulse:*` tool after the plugin is installed.

## Installing locally (before publishing)

From a directory containing this plugin folder:

```bash
claude plugin validate ./pulse-roadmap-toolkit
claude --plugin-dir ./pulse-roadmap-toolkit
```

Or install into an existing Claude Code session pointed at a local path
marketplace (see publishing instructions).

## License

MIT — update to your preferred license before publishing.

