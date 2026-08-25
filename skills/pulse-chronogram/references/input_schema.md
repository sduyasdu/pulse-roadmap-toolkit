# Input JSON schema for generate_chronogram.py

```json
{
  "meta": {
    "sheetName": "Cronograma",
    "title": "Cronograma — Product Roadmap Pulse",
    "subtitle": "Próximas {n} semanas — {start} a {end}",
    "epicColumnLabel": "Epic / Tarea",
    "legendPrefix": "Leyenda:",
    "legendLabel": "Semana con tarea activa"
  },
  "startDate": "2026-08-18",
  "weeks": 8,
  "monthNames": { "8": "Ago", "9": "Sep", "10": "Oct" },
  "epicOrder": ["Conectores", "Arcos", "No epic"],
  "tasks": [
    { "epic": "Conectores", "title": "CR 3.0", "start": "2026-07-20", "end": "2026-08-30" }
  ]
}
```

## Field notes

- **`startDate`** — any date inside the first week you want shown; the script
  snaps back to that week's Monday automatically. Normally "today".
- **`weeks`** — how many week-columns to render. Default the skill uses is 8,
  but honor whatever the user asked for (e.g. "next quarter" → ~13).
- **`subtitle`** — `{n}`, `{start}`, `{end}` are filled in automatically if
  present in the string; write it already in the target language.
- **`monthNames`** — keys are month numbers as strings ("1"–"12"). Omit the
  whole field to fall back to Spanish abbreviations; pass your own dict for
  English ("Aug", "Sep", "Oct", …) or any other language.
- **`epicOrder`** — controls which epics appear and in what order; an epic
  with zero tasks in the filtered set is simply skipped, not shown empty.
  Include a "No epic" (or localized equivalent) entry for tasks with no
  epic — don't drop or merge them elsewhere.
- **`tasks[].start` / `.end`** — ISO `YYYY-MM-DD`. Only include tasks that
  overlap the visible week range (a task's `[start, end]` intersects
  `[weekStart, lastWeekEnd]`) — filter *before* writing this file, the
  script does not filter for you.
- A task whose date range extends beyond the visible window still renders —
  its shading simply runs through to the last column without implying it
  ends there. No need to clip start/end dates to the window.
