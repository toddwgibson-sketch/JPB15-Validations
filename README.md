
## Notes

- All processing happens in memory / temporary files
- Nothing is stored on the server after you download the output
- Multiple cutsheets are supported in all tools (highly recommended)

## Central Error Logging & Dashboard

All validation tools (QFAB, GFAB, HOPS, T0-to-Host, etc.) can log error counts to a shared file:

`data/validation_error_log.xlsx`

- This file lives inside the repo so the **10_Dashboard** page can aggregate errors by hall, building, and category.
- The `data/` folder contains a `.gitkeep` so it survives cloning.
- On first run (or after a fresh clone) the folder may appear empty — use the **"📁 Force-create the GitHub data/ folder now"** button on the QFAB page to create it.
- The actual log file is git-ignored (we don't want to commit every run's data).
- The Dashboard auto-refreshes the view from this file.
