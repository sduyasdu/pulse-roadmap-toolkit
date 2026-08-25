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

## Connecting this plugin to Claude Cowork

Cowork installs plugins from a **marketplace** — a Git repository containing
a `.claude-plugin/marketplace.json` catalog file. If this plugin's own
`.claude-plugin` folder is the repo root, `marketplace.json`'s `source`
should be `"./"`; if the plugin lives inside a `plugins/` subfolder (the
layout used when a repo hosts more than one plugin), `source` should point
at that subfolder, e.g. `"./plugins/pulse-roadmap-toolkit"`. This repo needs
its own `marketplace.json` at the repo root either way — it isn't included
automatically.

### Add the marketplace

1. Open **Claude Cowork**.
2. Click **Customize** in the left sidebar.
3. Go to the **Plugins** tab.
4. Click **Browse plugins**, then **Add marketplace**.
5. Enter the repository URL — either form works:
   ```
   https://github.com/sduyasdu/pulse-roadmap-toolkit
   ```
   or the shorthand:
   ```
   sduyasdu/pulse-roadmap-toolkit
   ```
6. Cowork reads `.claude-plugin/marketplace.json` at the repo root and lists
   the plugin(s) it finds — **pulse-roadmap-toolkit** should appear.

If you get **"Marketplace sync failed. Check the repository URL and try
again,"** check these two known gotchas before anything else — both have
caused this exact error:

- **`marketplace.json`'s `"name"` field must be lowercase kebab-case**
  (letters, digits, hyphens only). A name like `"My-Marketplace"` fails
  sync; `"my-marketplace"` works.
- **`.mcp.json`'s `"type"` field must be `"http"`**, not `"url"`. Only
  `"http"` (or its alias `"streamable-http"`), `"sse"`, and `"ws"` are valid
  values — this repo's own `.mcp.json` is already correct, but if you're
  troubleshooting a fork or a hand-edited copy, this is worth checking.

After fixing either of those, push the change and click **Update** on the
marketplace in Cowork rather than waiting — it doesn't always re-check
automatically.

### Install the plugin

1. Click on **pulse-roadmap-toolkit** in the marketplace listing.
2. Click **Install**.
3. Since this plugin bundles the `pulse` MCP connector, you'll be prompted
   to sign in to `pulse.yasdu.com` right after install — complete that
   authentication.
4. Open the installed plugin under **Customize → Plugins** to confirm its
   three skills and the `pulse` connector appear as tabs.

### Where to find it again later

Added marketplaces and installed plugins both live under
**Customize → Plugins → Browse plugins** — your marketplace appears there
alongside Anthropic's default official catalog. Click **Update** on it any
time you've pushed new commits to the repo, so Cowork picks up the changes
without needing to remove and re-add the marketplace.

### For organization-wide rollout

On Team and Enterprise plans, an administrator can make this plugin
required for everyone in the org from **Organization settings → Plugins**,
so it installs automatically rather than per-person. See Anthropic's
[deployment guide](https://claude.com/docs/cowork/3p/extensions) for that
flow.

## License

MIT — update to your preferred license before publishing.

