# HSE Compliance Checklist website

Upload the Master Training Register (Excel, sheet `KENT TRAINING MATRIX`); the HSE Compliance Checklist is filled from it.

## Run
1. Install Python 3.10+.  2. Windows: double-click `run_windows.bat` (Mac/Linux: `./run_mac_linux.sh`).  3. Open http://127.0.0.1:5000

## Pages
- **Upload** - choose the latest register.
- **Checklist** - shows which register column each checklist block used, totals, warnings.
- **HSE Compliance Checklist** (download) - your template (`audit_formats/HSE_Compliance_Checklist_TEMPLATE.xlsx`) filled; charts/logos kept.

## Rules
- Employee Category number (1-56) gives the headcount in the master table.
- ATTENDED counts only valid (non-expired) records. The 18 Yes/No questions are not touched.
- H2S L1 counted first; L2 "To Attend" = headcount - L1 attended; L2 attended counts only people without a valid L1.
- Red rows in the register (red NAME cell or 10+ red cells) = inactive, never counted.
- Columns are found by header text (merged Issue/Expiry headers, missing sub-header row, "Valid till" etc. are handled).

## Settings - `config/settings.json`
`project_short`, `contractor`, `expiry_soon_days`, `as_of_date` (null = today), `company_parents`, `h2s_pool_rule`,
`exclude_red_rows`, `red_row_min_cells`, `training_type_overrides`, `checklist.block_overrides`.

## Hosting
Needs a Python server (Render / Railway / PythonAnywhere): start command `gunicorn app:app --timeout 300 --workers 1` (see `Procfile`).
Set environment variable `APP_PASSWORD` - the site then asks for it (the register contains employee data).

### Render.com in 5 steps (free)
1. Create a free GitHub account; New repository -> **Private**; upload all files of this folder (drag & drop in the browser; `data/` stays empty).
2. On render.com sign up with GitHub -> **New +** -> **Blueprint** -> pick the repository (it reads `render.yaml`).
3. When asked for `APP_PASSWORD`, type a strong password -> **Apply**.
4. After the build (2-3 minutes) Render shows the link `https://hse-checklist-xxxx.onrender.com`. Open it; user name = anything, password = APP_PASSWORD.
5. Free plan sleeps after 15 min idle (first open takes ~1 min) and forgets the uploaded register on restart - just upload it again.
