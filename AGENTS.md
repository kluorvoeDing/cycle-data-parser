# Cycle 數據可視化 — AGENTS.md

## Project
Offline battery charge/discharge data visualization app.
This is the lightweight offline version for local use, with a small footprint and fast startup.
Treat the offline build as the source of truth; keep edits simple and runtime-light.

## Key Files
- `index.html` — Main entrypoint (CDN libraries, lightweight ~30KB)
- `index_offline.html` — Fully offline version (embedded libraries ~5MB)
- `js/app.js` — Application logic (parsing, data processing, chart rendering)
- `css/style.css` — Styling with embedded gg sans fonts
- `run.bat` — Opens `index.html` in default browser
- `recreated_CH005_11260319010002_202605271.xlsx` — Sample input workbook
- `fonts/` — Embedded gg sans font files

## Editing Rule
- Primary editing target: `js/app.js` (application logic)
- Styling: `css/style.css`
- HTML structure: `index.html`
- For offline deployment, regenerate `index_offline.html` with embedded libraries
- Keep changes minimal and test after each edit

## ⚠️ MANDATORY: Verify After Every Edit
**Lesson learned**: Every code edit must be followed by a verification pass to catch upload-breaking syntax errors. Past failures were caused by:
- Residual `}` braces left after partial edits (e.g., old discharge block left outside `forEach`)
- Duplicate `const` declarations when replacing blocks incompletely
- Typo `Ahref` (missing underscore) breaking `r.PROPERTY` access → NaN values

```powershell
# If a verifier exists in this folder, run it AFTER every edit.
# Otherwise do a quick smoke check: open index_offline.html and confirm the charts load without console errors.
```

`verify.ps1` (when present) checks:
- `Ahref` typo count (must be 0)
- Brace balance `{` = `}` in main inline script
- `fileInput.addEventListener('change', ...)` handler present
- `function handleFile` present
- No residual old discharge reversal formula
- All 3 render functions present (`renderCycleLife`, `renderCDCurves`, `renderDQDV`)

**Rule**: If `verify.ps1` fails, do NOT claim the fix is done. Fix the issues and re-run until it passes.

## Charts
| Tab | X-axis | Y-axis | Details |
|-----|--------|--------|---------|
| Cycle Life | Cycle Number (0–500) | Retention% / CC Ratio% / CE% (0–100) | Scatter only; dtick=100/20, minor dtick=50/10 |
| C/D Curves | SOC% (0–100) | Voltage (V, 0–5) | 3 modes (Charge/Discharge/Overlay); gradient coloring |
| dQ/dV | Voltage (V, 2.5–5) | Auto | Central diff + SG smoothing; Charge/Discharge/Both |

## Design Conventions
- Font: `gg sans` (embedded woff: Regular 400, Medium 500, Semibold 600, Bold 700)
- Chart bg: `#FAF9F5`, body bg: `#F5F0E6`
- All text/lines: `#000000`
- Axis lines: `showline: true, linecolor: #000000, mirror: true, ticks: outside`
- Modebar: bottom-right, scrollZoom enabled, toImage & resetScale2d kept
- Chart title: Plotly `layout.title`, centered, bold, 2x global font size, `y: 0.98`
- Layout margins: dynamic based on font size (`padL`, `padB`, `margin.t`)

## Data
- Input: Excel with columns: Action, Model, Loop1, Cycle, Vol_Avg, Cur_Avg, Cap_Ah, Tim, etc.
- Units: V (not mV), A (not mA), Ah (not mAh), s (not ms)
- First-cycle CHARGE auto-skipped if it precedes any DISCHARGE
- CC ratio: current-drop threshold at 95% of max current

## Key JS Functions
- `parseExcel(ab)` / `parseFud(text)` — parse workbook inputs
- `processData()` — builds cycle summaries and derived metrics
- `baseLayout()` — returns Plotly layout with shared config
- `renderCycleLife()`, `renderCDCurves()`, `renderDQDV()` — chart renderers
- `setupResizeObserver()` / `resizeActiveChart()` — auto-resize on summary toggle

## Axis Input ID Convention
- Prefix `cl` (Cycle Life), `cd` (C/D), `dq` (dQ/dV)
- Suffix `Xmin`, `Xmax`, `Ymin`, `Ymax`

## GitHub
- **Repo:** https://github.com/kluorvoeDing/cycle-data-parser
- **Owner:** kluorvoeDing
