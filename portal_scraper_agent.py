import os
import re
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

RECIPIENT_EMAIL = "sudheernookala@gmail.com"
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "your_email@gmail.com")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "your_app_password")

USER_AGENT = {"User-Agent": "Mozilla/5.0"}

# ============================================================
# COMMON FILTERS
# ============================================================

def is_target_role(title, location, desc):
    text = f"{title} {location} {desc}".lower()
    valid_region = any(r in text for r in [
        "germany", "deutschland", "netherlands", "amsterdam",
        "berlin", "europe", "eu remote", "worldwide"
    ])
    has_java = "java" in text or "spring" in text
    has_senior = any(s in text for s in ["senior", "lead", "sr", "principal"])
    has_fullstack = (
        "fullstack" in text or "full stack" in text or
        "react" in text or "angular" in text or "vue" in text
    )
    return valid_region and has_java and has_senior and has_fullstack


def extract_stack(text):
    text = text.lower()
    found = []
    for tech in ["Java", "Spring Boot", "Angular", "TypeScript", "Kafka", "Docker", "Kubernetes", "AWS"]:
        if re.search(r'\b' + re.escape(tech.lower()) + r'\b', text):
            found.append(tech)
    return ", ".join(found) if found else "Java, Spring Boot, Microservices"


# ============================================================
# COMMON SCRAPER ENGINE
# ============================================================

def scrape_portal(url, parser_fn, source_name, is_json=False):
    jobs = []
    try:
        res = requests.get(url, headers=USER_AGENT, timeout=10)
        if res.status_code == 200:
            data = res.json() if is_json else res.text
            parsed = parser_fn(data)

            for item in parsed:
                title = item.get("title", "")
                location = item.get("location", "")
                desc = item.get("desc", "")

                if is_target_role(title, location, desc):
                    jobs.append({
                        "company": item.get("company", "N/A"),
                        "location": location,
                        "title": title,
                        "tech_stack": extract_stack(f"{title} {desc}"),
                        "contact_info": item.get("link", ""),
                        "source": source_name
                    })
    except Exception as e:
        print(f"[{source_name}] Error: {e}")

    return jobs


# ============================================================
# PORTAL PARSERS
# ============================================================

def parse_remotive(json_data):
    jobs = []
    for item in json_data.get("jobs", []):
        jobs.append({
            "company": item.get("company_name", ""),
            "location": item.get("candidate_required_location", "Remote"),
            "title": item.get("title", ""),
            "desc": item.get("description", ""),
            "link": item.get("url", "")
        })
    return jobs


def parse_stepstone(html):
    pattern = re.findall(
        r'"title":"(.*?)".*?"company":"(.*?)".*?"location":"(.*?)".*?"url":"(.*?)"',
        html
    )
    return [
        {"title": t, "company": c, "location": l, "desc": t, "link": u}
        for t, c, l, u in pattern
    ]


def parse_indeed(html):
    titles = re.findall(r'jobTitle">(.*?)<', html)
    companies = re.findall(r'companyName">(.*?)<', html)
    locations = re.findall(r'companyLocation">(.*?)<', html)
    links = re.findall(r'href="(/rc/clk.*?)"', html)

    jobs = []
    for i in range(min(len(titles), len(companies), len(locations), len(links))):
        jobs.append({
            "title": titles[i],
            "company": companies[i],
            "location": locations[i],
            "desc": titles[i],
            "link": "https://de.indeed.com" + links[i]
        })
    return jobs


def parse_wearedevelopers(html):
    pattern = re.findall(
        r'"title":"(.*?)".*?"companyName":"(.*?)".*?"location":"(.*?)".*?"slug":"(.*?)"',
        html
    )
    return [
        {
            "title": t,
            "company": c,
            "location": l,
            "desc": t,
            "link": f"https://www.wearedevelopers.com/jobs/{slug}"
        }
        for t, c, l, slug in pattern
    ]


def parse_remoteok(json_data):
    jobs = []
    for item in json_data:
        if isinstance(item, dict) and "position" in item:
            jobs.append({
                "title": item.get("position", ""),
                "company": item.get("company", ""),
                "location": item.get("location", "Remote"),
                "desc": item.get("description", ""),
                "link": item.get("url", "")
            })
    return jobs


def parse_arbeitnow(json_data):
    jobs = []
    for item in json_data.get("data", []):
        jobs.append({
            "title": item.get("title", ""),
            "company": item.get("company", ""),
            "location": item.get("location", "Germany"),
            "desc": item.get("description", ""),
            "link": item.get("url", "")
        })
    return jobs


# ============================================================
# AGGREGATOR
# ============================================================

def fetch_live_portal_jobs():
    portals = [
        ("https://remotive.com/api/remote-jobs?category=software-dev&search=Java", parse_remotive, "Remotive", True),
        ("https://www.stepstone.de/jobs/java-developer/in-deutschland", parse_stepstone, "StepStone", False),
        ("https://de.indeed.com/jobs?q=senior+java+developer&l=Germany", parse_indeed, "Indeed", False),
        ("https://www.wearedevelopers.com/jobs?query=java&location=germany", parse_wearedevelopers, "WeAreDevelopers", False),
        ("https://remoteok.com/api", parse_remoteok, "RemoteOK", True),
        ("https://api.arbeitnow.com/api/job-board-api", parse_arbeitnow, "Arbeitnow", True)
    ]

    all_jobs = []
    for url, parser, source, is_json in portals:
        all_jobs += scrape_portal(url, parser, source, is_json)

    # Deduplicate
    unique = {}
    for job in all_jobs:
        key = (job["company"], job["title"], job["location"])
        unique[key] = job

    return list(unique.values())


# ============================================================
# EXCEL + EMAIL (unchanged)
# ============================================================

# (Your existing Excel + email code stays exactly the same)
