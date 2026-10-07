# my-job-search-agent
Job search agent looks into job portals and delivers an email daily basis

## Agents

| Agent | Workflow | What it does |
|---|---|---|
| Email agent | `daily_agent.yml` → `portal_scraper_agent.py` | Emails the last 24h of jobs as an Excel attachment. |
| Tracker agent | `job_tracker.yml` → `excel_tracker_agent.py` | Appends new jobs to `data/job_history.xlsx`, keeps only the last 30 days, and publishes a web UI. |

The tracker reuses the email agent's search and filters (it imports them), so changing a filter in `portal_scraper_agent.py` changes both.

### Tracker details
- Runs daily at 08:30 UTC (or manually from the Actions tab).
- Each row has a **Date Found**. Jobs already in the file (same link, or same company + title) are not added again.
- Rows older than 30 days are deleted on every run.
- `data/jobs.json` holds the same rows for the UI.

### Web UI (`web/`)
React + Vite app, deployed to GitHub Pages: https://sudheernookala.github.io/my-job-search-agent/

Filter by date range (or quick ranges: today / 3 / 7 / 14 / 30 days), search by company or location, filter by source, sort columns, download the Excel.

**One-time setup:** repo **Settings → Pages → Build and deployment → Source: GitHub Actions**.

Run locally:
```bash
python excel_tracker_agent.py   # creates data/jobs.json
cd web && npm install && npm run dev
```
