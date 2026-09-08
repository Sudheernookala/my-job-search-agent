import os
import re
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

RECIPIENT_EMAIL = "sudheernookala@gmail.com"
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "your_email@gmail.com")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "your_app_password")

def fetch_live_portal_jobs():
    """
    Fetches live postings from European tech career portals (e.g. WeAreDevelopers, Remotive, API feeds).
    """
    target_jobs = []
    
    # Query Remotive API for EU Remote / DE / NL Java Fullstack roles
    url = "https://remotive.com/api/remote-jobs?category=software-dev&search=Java"
    try:
        res = requests.get(url, timeout=10)
        if res.status_code == 200:
            data = res.json().get("jobs", [])
            for item in data:
                title = item.get("title", "")
                candidate_loc = item.get("candidate_required_location", "")
                desc = item.get("description", "")
                
                # Check Seniority + Fullstack + Target Regions
                if is_target_role(title, candidate_loc, desc):
                    target_jobs.append({
                        "company": item.get("company_name", "N/A"),
                        "location": candidate_loc if candidate_loc else "EU Remote",
                        "title": title,
                        "tech_stack": extract_stack(f"{title} {desc}"),
                        "contact_info": item.get("url", "N/A")
                    })
    except Exception as e:
        print(f"Error fetching API jobs: {e}")

    return target_jobs

def is_target_role(title, location, desc):
    text = f"{title} {location} {desc}".lower()
    
    # Check regions
    valid_region = any(r in text for r in ["germany", "deutschland", "netherlands", "amsterdam", "berlin", "europe", "eu remote", "worldwide"])
    
    # Check tech stack & seniority
    has_java = "java" in text or "spring" in text
    has_senior = any(s in text for s in ["senior", "lead", "sr", "principal"])
    has_fullstack = "fullstack" in text or "full stack" in text or ("react" in text or "angular" in text or "vue" in text)
    
    return valid_region and has_java and has_senior and has_fullstack

def extract_stack(text):
    text = text.lower()
    found = []
    for tech in ["Java", "Spring Boot", "Angular", "React", "Vue", "TypeScript", "Kafka", "Docker", "Kubernetes", "AWS"]:
        if re.search(r'\b' + re.escape(tech.lower()) + r'\b', text):
            found.append(tech)
    return ", ".join(found) if found else "Java, Spring Boot, Microservices"

def generate_excel(jobs, filename="Daily_EU_Java_Roles.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daily Job Matches"
    ws.views.sheetView[0].showGridLines = True

    # Title
    ws['A1'] = "Senior Java Fullstack Roles - Germany, Netherlands & EU Remote"
    ws['A1'].font = Font(name='Calibri', size=14, bold=True, color="1F4E78")

    # Table Header
    headers = ["Company Name", "Location", "Role Title", "Key Tech Stack", "Contact / Application Link"]
    ws.append([])
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name='Calibri', size=11, bold=True, color="FFFFFF")

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Rows
    for row_idx, job in enumerate(jobs, 4):
        ws.append([
            job['company'],
            job['location'],
            job['title'],
            job['tech_stack'],
            job['contact_info']
        ])
        fill_color = "F9FAFB" if row_idx % 2 == 0 else "FFFFFF"
        for col_idx in range(1, 6):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            cell.font = Font(name='Calibri', size=10)

    widths = {'A': 22, 'B': 28, 'C': 35, 'D': 40, 'E': 55}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    wb.save(filename)
    return filename

def send_daily_email(recipient, excel_path):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = recipient
    msg['Subject'] = "Daily Senior Java Fullstack Job Matches (Portal Search)"

    body = "Hi Sudheer,\n\nPlease find attached today's freshly scraped Excel report containing Senior Java Fullstack Developer roles across Germany, Netherlands, and EU Remote.\n\nBest regards,\nJob Search Agent"
    msg.attach(MIMEText(body, 'plain'))

    with open(excel_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=excel_path)
        part['Content-Disposition'] = f'attachment; filename="{excel_path}"'
        msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(SENDER_EMAIL, GMAIL_APP_PASSWORD)
        server.sendmail(SENDER_EMAIL, recipient, msg.as_string())
    print("Daily email successfully sent!")

if __name__ == "__main__":
    jobs = fetch_live_portal_jobs()
    if jobs:
        file_path = generate_excel(jobs)
        send_daily_email(RECIPIENT_EMAIL, file_path)
    else:
        print("No matching roles found today.")
