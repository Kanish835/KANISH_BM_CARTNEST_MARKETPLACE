# CartNest settlement dashboard

This project turns the four source CSV files into a normalized, auditable settlement dashboard and a client-ready Excel report. It includes:

- mixed-format date and identifier cleaning;
- rate-card commission, fixed fee, shipping, GST, TCS and TDS calculations;
- item-level matching between sales and settlement records;
- T+7 timing separation from unpaid shipped items;
- bank cash reconciliation;
- category and daily performance charts;
- a filterable exception queue.
- `CartNest_Settlement_Report_Sep_2026.xlsx` with a client summary, exception report, category summary and cleaning log.

## Run locally

```powershell
python -m pip install -r requirements.txt
python clean_data.py
python serve.py
```

Open http://127.0.0.1:8000. The generated `dashboard_data.js` is the only data file used by the static page, and the Excel report is available from the dashboard download button. Both are refreshed when you run `python clean_data.py`.

## Create a new GitHub repository and push

Create an empty repository in GitHub first, then run from this folder:

```powershell
git init
git add .
git commit -m "Build settlement reconciliation dashboard"
git branch -M main
git remote add origin https://github.com/<your-username>/<new-repository>.git
git push -u origin main
```

Keep the generated `dashboard_data.js` in the repository if using static hosting. The CSV files are included as the reproducible source data; remove them before publishing if they contain confidential financial information.
