# Style

This skill always renders in the Yasdu brand style — there is no style
choice to elicit from the user. The generator script hardcodes it, so you
never pass style colors or a logo path yourself; just point it at the
bundled logo asset (see SKILL.md Step 4).

| Element | Value |
|---|---|
| Logo | `assets/yasdu/logo-horizontal-light.png`, shown top-left |
| Header / active-week fill | Yasdu orange `#D85A28` |
| Month header | Yasdu navy `#123359` |
| Epic row background | Yasdu accent tan `#F7E8DA` |
| Font | Calibri (Inter fallback) |

Layout stays the same regardless: merged month headers, week columns with
dd/mm ranges, epic section rows, alternating row shading, solid fill on
active weeks, legend row, landscape + fit-to-width page setup. The logo
reserves the first few rows and the title/subtitle shift down automatically
— you don't need to adjust row numbers in your input JSON for this.
