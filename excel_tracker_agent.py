"""
Job Tracker Agent - rolling 30-day Excel history (no email)
===========================================================

Runs the same search as `portal_scraper_agent.py` (it imports the fetchers,
so filters stay in one place), then:

  1. Appends new jobs to `data/job_history.xlsx` with the date they were found.
  2. Skips jobs already in the file (same link, or same company + title).
  3. Deletes rows older than RETENTION_DAYS (30).
  4. Writes `data/jobs.json` - the same rows, read by the React UI in `web/`.

The email agent is not touched and keeps running on its own.
"""

import json
import os
from datetime import datetime, timedelta, timezone

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from portal_scraper_agent import fetch_all_jobs

RETENTION_DAYS = 30
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
EXCEL_PATH = os.path.join(DATA_DIR, "job_history.xlsx")
JSON_PATH = os.path.join(DATA_DIR, "jobs.json")

COLUMNS = [
    ("date_found", "Date Found", 14),
    ("source", "Source", 16),
    ("company", "Company Name", 26),
    ("location", "Location", 22),
    ("title", "Role Title", 40),
    ("tech_stack", "Key Tech Stack", 30),
    ("contact_info", "Contact / Application Link", 60),
]


def job_key(job):
    link = (job.get("contact_info") or "").strip().lower()
    if link and link != "n/a":
        return ("link", link)
    return ("name", (job.get("company") or "").strip().lower(), (job.get("title") or "").strip().lower())


def load_history():
    """Read existing rows from the Excel file (header is row 1)."""
    if not os.path.exists(EXCEL_PATH):
        return []
    ws = openpyxl.load_workbook(EXCEL_PATH).active
    keys = [c[0] for c in COLUMNS]
    rows = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if not values or not values[0]:
            continue
        row = dict(zip(keys, values))
        if isinstance(row["date_found"], datetime):
            row["date_found"] = row["date_found"].date().isoformat()
        row = {k: ("" if v is None else str(v)) for k, v in row.items()}
        rows.append(row)
    return rows


def save_history(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Job History (30 days)"

    ws.append([c[1] for c in COLUMNS])
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in rows:
        ws.append([row.get(c[0], "") for c in COLUMNS])

    for idx, (_, _, width) in enumerate(COLUMNS):
        ws.column_dimensions[openpyxl.utils.get_column_letter(idx + 1)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(EXCEL_PATH)


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    today = datetime.now(timezone.utc).date()
    cutoff = (today - timedelta(days=RETENTION_DAYS - 1)).isoformat()  # today + 29 previous days

    history = load_history()
    before = len(history)
    history = [r for r in history if r["date_found"] >= cutoff]
    removed = before - len(history)

    seen = {job_key(r) for r in history}
    added = 0
    for job in fetch_all_jobs():
        key = job_key(job)
        if key in seen:
            continue
        seen.add(key)
        history.append({"date_found": today.isoformat(), **{k: str(job.get(k, "")) for k, _, _ in COLUMNS[1:]}})
        added += 1

    history.sort(key=lambda r: (r["date_found"], r["company"].lower()), reverse=True)
    save_history(history)

    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {"generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
             "retention_days": RETENTION_DAYS,
             "jobs": history},
            f, ensure_ascii=False, indent=1,
        )

    print(f"Added {added} new jobs, removed {removed} older than {RETENTION_DAYS} days, total {len(history)}.")


if __name__ == "__main__":
    main()
