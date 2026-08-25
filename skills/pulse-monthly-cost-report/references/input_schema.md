# Input JSON schema for generate_hours_report.py

```json
{
  "meta": {
    "monthLabel": "Agosto 2026",
    "detailSheetName": "Detalle Agosto",
    "font": "Arial"
  },
  "monthStart": "2026-08-01",
  "monthEnd": "2026-08-31",
  "capHours": 160,
  "tasks": [
    {
      "epic": "Conectores",
      "title": "CR 3.0",
      "start": "2026-07-20",
      "end": "2026-08-30",
      "assignees": [
        { "name": "Marcelo Nieto", "allocationPercent": 100 },
        { "name": "Sdu (copy)", "allocationPercent": 100 }
      ]
    }
  ]
}
```

## Field notes

- **`monthStart` / `monthEnd`** — the first and last calendar day of the
  target month, ISO format. `monthEnd` must be the actual last day (28–31
  depending on the month) — the script does not compute it for you.
- **`capHours`** — the monthly hour cap per resource. Defaults to 160 if
  omitted. Every column header and the assumption note automatically pick up
  whatever value you pass, so you don't need to edit the script to change it.
- **`detailSheetName`** — Excel sheet names cap at 31 characters; the script
  truncates automatically but pick something short and clear
  (`"Detalle Agosto"`, `"Detalle Q3"`, etc).
- **`tasks[].start` / `.end`** — ISO `YYYY-MM-DD`. Only include tasks whose
  range overlaps `[monthStart, monthEnd]` — filter *before* writing this
  file; the script does not filter for you.
- **`tasks[].assignees`** — list of `{name, allocationPercent}`. An empty
  list (`[]`) is fine for an unassigned task — it renders as one row with
  `"— (sin asignar)"` and 0 hours/cost, so unstaffed work is visible rather
  than silently dropped.
- Every hours/cost figure in the workbook is a **live formula** — NETWORKDAYS
  for business-day overlap, SUMIFS for the resource's uncapped monthly total,
  an adjustment factor (`MIN(1, capHours / total)`) applied per line, then an
  INDEX/MATCH lookup against the Costo por Recurso sheet for cost. Nothing is
  precomputed in Python and pasted in as a static number.
