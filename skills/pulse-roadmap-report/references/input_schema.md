# Input JSON schema for generate_report.js

Build this JSON entirely in the target language the user selected — every
label, heading, and cell string is rendered verbatim, nothing is translated
by the script.

```json
{
  "meta": {
    "title": "Product Roadmap",
    "subtitle": "Status Report for Management — <date>",
    "intro": "One or two sentence summary of scope (task count, epic count).",
    "footerText": "Product Roadmap — Status Report",
    "keyRisksLabel": "KEY RISKS"
  },
  "sections": [
    {
      "icon": "✔",
      "heading": "Completed",
      "introText": "N tasks finished or closed out.",
      "colorOverride": null,
      "columns": [
        { "label": "Task", "width": 0.19 },
        { "label": "Lead", "width": 0.15 },
        { "label": "Start Date", "width": 0.14 },
        { "label": "Finished Date", "width": 0.14 },
        { "label": "Subtasks", "width": 0.38 }
      ],
      "epics": [
        {
          "name": "Conectores",
          "rows": [
            {
              "cells": ["Macro", "—", "2026-06-21", "2026-07-19", { "subtasks": [] }]
            },
            {
              "cells": ["Proceso Conectores 3 dias", "Brian Walker", "2026-07-20", "2026-08-04", { "subtasks": ["Done", "Piloto"] }]
            }
          ]
        }
      ]
    }
  ],
  "keyRisks": "Free-text paragraph, e.g. flagging unowned tasks or stalled work."
}
```

## Field notes

- **`columns[].width`** — fractions of the usable page width; each section's
  widths should sum to ~1.0. Keep "Task" and "Subtasks" the widest columns.
- **`epics[].name`** — omit or set to the localized "No epic" label for tasks
  without an epic; keep as its own group, don't merge into another epic.
- **A subtasks cell** is always `{ "subtasks": [...] }`, an array of plain
  label strings (append the localized "(done)" suffix to a label yourself if
  that subtask's status is Done — the script doesn't compute this).
- **`colorOverride`** — set to a hex string (e.g. `"991B1B"`) for a section
  that should render in warning red (e.g. a Stalled section); leave `null`
  for normal sections. When set, it recolors both the section heading and
  every row's text in that section.
- Sections appear in the document in the array order given — order them
  Completed → Ongoing → Stalled → Planned (or the local-language equivalent
  status buckets actually present in the Pulse).
- `keyRisks` is optional — omit the top-level key entirely to skip the callout
  box.
