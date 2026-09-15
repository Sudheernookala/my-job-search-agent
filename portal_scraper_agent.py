"""
Daily Senior Java Fullstack Job Aggregator - Germany / Netherlands / EU Remote
================================================================================

SOURCES USED (all public, no-ToS-violation access):
  - Arbeitsagentur (Bundesagentur fur Arbeit) - Germany's official job database.
    Uses a widely-used, community-documented public endpoint (client key
    "jobboerse-jobsuche"). Not an officially supported partner API, but it's
    the same endpoint the arbeitsagentur.de website itself calls, so it's
    stable and not a ToS violation the way scraping a rendered page is.
  - Remotive       - public API, no key required
  - Arbeitnow      - public API, no key required
  - Jobicy         - public API, no key required

SOURCES DELIBERATELY *NOT* INCLUDED (and why):
  - Indeed.de   -> Public search API deprecated since 2023. Current ToS
                   explicitly forbids scraping/automated access.
  - StepStone   -> No public API for job seekers. Heavy anti-bot protection.
  - WeAreDevelopers -> No public API or RSS feed found for their job board.
  If you need those specifically, the only legitimate routes are their
  paid/approved partner programs, or a third-party scraping service that
  assumes the ToS risk on your behalf (e.g. Apify) - not a DIY scraper.

FRESHNESS:
  Every job is checked against `MAX_AGE_HOURS` (default 24h) using each
  source's own posting-date field. Arbeitsagentur additionally uses its
  native `veroeffentlichtseit` (days-since-published) filter server-side,
  but note that parameter's granularity is in whole days, not hours - so a
  small number of jobs posted 24-36h ago may still slip through from that
  source. The others are filtered precisely using timestamps.
"""

import os
import re
import base64
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime, timedelta, timezone
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

RECIPIENT_EMAIL = "sudheernookala@gmail.com"
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "your_email@gmail.com")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "your_app_password")

MAX_AGE_HOURS = 24
NOW = datetime.now(timezone.utc)


# ----------------------------------------------------------------------
# Shared filtering logic
# ----------------------------------------------------------------------

def is_target_role(title, location, desc):
    """Region + seniority + stack filter, applied across all sources."""
    text = f"{title} {location} {desc}".lower()

    valid_region = any(r in text for r in [
        "germany", "deutschland", "netherlands", "amsterdam", "berlin",
        "munich", "munchen", "hamburg", "frankfurt", "cologne", "koln",
        "stuttgart", "europe", "eu remote", "worldwide"
    ])

    has_java = "java" in text or "spring" in text
    has_senior = any(s in text for s in ["senior", "lead", "sr.", "sr ", "principal"])
    has_fullstack = (
        "fullstack" in text or "full stack" in text or "full-stack" in text
        or "react" in text or "angular" in text or "vue" in text
    )

    return valid_region and has_java and has_senior and has_fullstack


def extract_stack(text):
    text = text.lower()
    found = []
    for tech in ["Java", "Spring Boot", "Angular", "React", "TypeScript",
                 "Kafka", "Docker", "Kubernetes", "AWS"]:
        if re.search(r'\b' + re.escape(tech.lower()) + r'\b', text):
            found.append(tech)
    return ", ".join(found) if found else "Java, Spring Boot"


def within_freshness_window(posted_dt):
    """posted_dt must be a timezone-aware datetime."""
    if posted_dt is None:
        return True  # unknown date -> don't silently drop it, let it through
    return (NOW - posted_dt) <= timedelta(hours=MAX_AGE_HOURS)


# ----------------------------------------------------------------------
# Source: Remotive
# ----------------------------------------------------------------------

def fetch_remotive():
    jobs = []
    url = "https://remotive.com/api/remote-jobs?category=software-dev&search=Java"
    try:
        res = requests.get(url, timeout=15)
        res.raise_for_status()
        for item in res.json().get("jobs", []):
            title = item.get("title", "")
            loc = item.get("candidate_required_location", "")
            desc = item.get("description", "")

            posted_dt = None
            pub_date = item.get("publication_date")
            if pub_date:
                try:
                    posted_dt = datetime.strptime(pub_date, "%Y-%m-%dT%H:%M:%S").replace(tzinfo=timezone.utc)
                except ValueError:
                    pass

            if is_target_role(title, loc, desc) and within_freshness_window(posted_dt):
                jobs.append({
                    "source": "Remotive",
                    "company": item.get("company_name", "N/A"),
                    "location": loc or "EU Remote",
                    "title": title,
                    "tech_stack": extract_stack(f"{title} {desc}"),
                    "contact_info": item.get("url", "N/A"),
                })
    except Exception as e:
        print(f"[Remotive] fetch error: {e}")
    return jobs


# ----------------------------------------------------------------------
# Source: Arbeitsagentur (Bundesagentur fur Arbeit) - official German data
# ----------------------------------------------------------------------

def fetch_arbeitsagentur():
    jobs = []
    base_url = "https://rest.arbeitsagentur.de/jobboerse/jobsuche-service/pc/v4/jobs"
    headers = {"X-API-Key": "jobboerse-jobsuche"}

    # Run separate searches per city/region - the API's "wo" free-text field
    # does not reliably match "Germany" as a whole.
    search_locations = ["Berlin", "Munchen", "Hamburg", "Frankfurt", "Stuttgart", "Koln"]

    for wo in search_locations:
        params = {
            "was": "Java Entwickler",
            "wo": wo,
            "umkreis": 50,
            "veroeffentlichtseit": 1,  # jobs published within the last 1 day
            "angebotsart": 1,          # 1 = regular employment (ARBEIT)
            "page": 1,
            "size": 50,
        }
        try:
            res = requests.get(base_url, headers=headers, params=params, timeout=15)
            res.raise_for_status()
            data = res.json()
            for item in data.get("stellenangebote", []):
                title = item.get("titel", "") or item.get("beruf", "")
                arbeitgeber = item.get("arbeitgeber", "N/A")
                ort = (item.get("arbeitsort") or {}).get("ort", wo)
                refnr = item.get("refnr", "")

                if not is_target_role(title, ort, title):
                    continue

                # Build the public job-detail link from the reference number
                link = f"https://www.arbeitsagentur.de/jobsuche/jobdetail/{refnr}" if refnr else "N/A"

                jobs.append({
                    "source": "Arbeitsagentur",
                    "company": arbeitgeber,
                    "location": ort,
                    "title": title,
                    "tech_stack": extract_stack(title),
                    "contact_info": link,
                })
        except Exception as e:
            print(f"[Arbeitsagentur:{wo}] fetch error: {e}")

    return jobs


# ----------------------------------------------------------------------
# Source: Arbeitnow (good German/EU tech coverage, no key required)
# ----------------------------------------------------------------------

def fetch_arbeitnow():
    jobs = []
    url = "https://www.arbeitnow.com/api/job-board-api"
    try:
        res = requests.get(url, timeout=15)
        res.raise_for_status()
        for item in res.json().get("data", []):
            title = item.get("title", "")
            location = item.get("location", "") or ("Remote" if item.get("remote") else "")
            desc = item.get("description", "")
            tags = " ".join(item.get("tags", []) or [])

            posted_dt = None
            created_at = item.get("created_at")
            if created_at:
                try:
                    posted_dt = datetime.fromtimestamp(int(created_at), tz=timezone.utc)
                except (ValueError, TypeError):
                    pass

            if is_target_role(title, location, f"{desc} {tags}") and within_freshness_window(posted_dt):
                jobs.append({
                    "source": "Arbeitnow",
                    "company": item.get("company_name", "N/A"),
                    "location": location or "EU Remote",
                    "title": title,
                    "tech_stack": extract_stack(f"{title} {desc} {tags}"),
                    "contact_info": item.get("url", "N/A"),
                })
    except Exception as e:
        print(f"[Arbeitnow] fetch error: {e}")
    return jobs


# ----------------------------------------------------------------------
# Source: Jobicy (remote-focused, has an EU/Europe geo filter)
# ----------------------------------------------------------------------

def fetch_jobicy():
    jobs = []
    url = "https://jobicy.com/api/v2/remote-jobs"
    params = {"count": 50, "geo": "europe", "tag": "java"}
    try:
        res = requests.get(url, params=params, timeout=15)
        res.raise_for_status()
        for item in res.json().get("jobs", []):
            title = item.get("jobTitle", "")
            location = item.get("jobGeo", "")
            desc = item.get("jobDescription", "") or item.get("jobExcerpt", "")

            posted_dt = None
            pub_date = item.get("pubDate")
            if pub_date:
                try:
                    posted_dt = datetime.strptime(pub_date, "%Y-%m-%dT%H:%M:%S").replace(tzinfo=timezone.utc)
                except ValueError:
                    pass

            if is_target_role(title, location, desc) and within_freshness_window(posted_dt):
                jobs.append({
                    "source": "Jobicy",
                    "company": item.get("companyName", "N/A"),
                    "location": location or "EU Remote",
                    "title": title,
                    "tech_stack": extract_stack(f"{title} {desc}"),
                    "contact_info": item.get("url", "N/A"),
                })
    except Exception as e:
        print(f"[Jobicy] fetch error: {e}")
    return jobs


# ----------------------------------------------------------------------
# Aggregation
# ----------------------------------------------------------------------

def fetch_all_jobs():
    all_jobs = []
    all_jobs += fetch_remotive()
    all_jobs += fetch_arbeitsagentur()
    all_jobs += fetch_arbeitnow()
    all_jobs += fetch_jobicy()

    # De-dupe on (company, title) - different sources sometimes list the same posting
    seen = set()
    deduped = []
    for job in all_jobs:
        key = (job["company"].strip().lower(), job["title"].strip().lower())
        if key not in seen:
            seen.add(key)
            deduped.append(job)

    return deduped


# ----------------------------------------------------------------------
# Excel output
# ----------------------------------------------------------------------

def generate_excel(jobs, filename="Daily_EU_Java_Roles.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daily Job Matches"

    ws['A1'] = "Senior Java Fullstack Roles - Germany, Netherlands & EU Remote (last 24h)"
    ws['A1'].font = Font(name='Calibri', size=14, bold=True, color="1F4E78")

    headers = ["Source", "Company Name", "Location", "Role Title", "Key Tech Stack", "Contact / Application Link"]
    ws.append([])
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name='Calibri', size=11, bold=True, color="FFFFFF")

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for row_idx, job in enumerate(jobs, 4):
        ws.append([
            job['source'],
            job['company'],
            job['location'],
            job['title'],
            job['tech_stack'],
            job['contact_info'],
        ])
        fill_color = "F9FAFB" if row_idx % 2 == 0 else "FFFFFF"
        for col_idx in range(1, 7):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            cell.font = Font(name='Calibri', size=10)

    widths = {'A': 16, 'B': 22, 'C': 22, 'D': 35, 'E': 30, 'F': 55}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    wb.save(filename)
    return filename


# ----------------------------------------------------------------------
# Email
# ----------------------------------------------------------------------

def send_daily_email(recipient, excel_path, job_count):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = recipient
    msg['Subject'] = f"Daily Senior Java Fullstack Job Matches - {job_count} new roles (last 24h)"

    body = (
        f"Hi,\n\n"
        f"Attached: {job_count} Senior Java Fullstack roles posted in the last 24 hours "
        f"across Germany, Netherlands, and EU Remote, aggregated from Arbeitsagentur, "
        f"Remotive, Arbeitnow, and Jobicy.\n\n"
        f"Note: StepStone, Indeed.de, and WeAreDevelopers are not included - none of "
        f"them offer a public job-search API, and scraping them violates their Terms "
        f"of Service.\n\n"
        f"Best regards,\nJob Search Agent"
    )
    msg.attach(MIMEText(body, 'plain'))

    with open(excel_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(excel_path))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(excel_path)}"'
        msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(SENDER_EMAIL, GMAIL_APP_PASSWORD)
            server.sendmail(SENDER_EMAIL, recipient, msg.as_string())
        print("Daily email successfully sent!")
    except Exception as e:
        print(f"Email send failed: {e}")
        raise


if __name__ == "__main__":
    jobs = fetch_all_jobs()
    if jobs:
        file_path = generate_excel(jobs)
        send_daily_email(RECIPIENT_EMAIL, file_path, len(jobs))
    else:
        print("No matching roles found in the last 24 hours.")
