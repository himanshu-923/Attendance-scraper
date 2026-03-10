from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import date, timedelta, datetime
import os
import re

# === CONFIGURATION ===
URL = "https://student.cgc.ac.in/Attendance.aspx"
OUT_FILE = "attendance_tracker.xlsx"
START_DATE = date(2026, 1, 9)          # adjust as needed
END_DATE = date.today()

RAW_SHEET = "daily_log"
SUBJECT_SHEET = "subject_summary"
DAILY_SHEET = "daily_summary"

# Credentials from environment variables (set in GitHub Secrets)
ERP_USER = os.environ.get("ERP_USERNAME", "")
ERP_PASS = os.environ.get("ERP_PASSWORD", "")

# === HELPER FUNCTIONS (same as yours) ===
def daterange(start, end):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)

def parse_rows(panel_text, current_date):
    # ... (your exact parse_rows function here) ...
    rows = []
    lines = [x.strip() for x in panel_text.splitlines() if x.strip()]

    idxs = [i for i, v in enumerate(lines) if re.fullmatch(r"\d+", v)]
    for k, start in enumerate(idxs):
        end = idxs[k + 1] if k + 1 < len(idxs) else len(lines)
        chunk = lines[start:end]
        if not chunk:
            continue

        lecture_no = int(chunk[0])
        subject = chunk[1] if len(chunk) > 1 else ""

        status = ""
        time_slot = ""
        teacher = ""

        status_tokens = []
        teacher_tokens = []

        for b in chunk[2:]:
            lb = b.lower()
            if re.search(r"\bpresent\b", lb):
                status_tokens.append("Present")
            if re.search(r"\babsent\b", lb):
                status_tokens.append("Absent")
            if not time_slot and re.search(r"\d{1,2}:\d{2}\s*[AP]M\s*to\s*\d{1,2}:\d{2}\s*[AP]M", b, re.I):
                time_slot = b
            if re.fullmatch(r"[A-Z]{2,5}", b):
                teacher_tokens.append(b)

        if status_tokens:
            status = status_tokens[-1]
        if teacher_tokens:
            teacher = teacher_tokens[-1]

        rows.append({
            "date": current_date.strftime("%Y-%m-%d"),
            "lecture_no": lecture_no,
            "subject": subject,
            "status": status,
            "time_slot": time_slot,
            "teacher": teacher,
        })
    return rows

def find_left_date_panel(page):
    # ... (your exact find_left_date_panel function here) ...
    panel = page.locator("div.col-md-4, div[class*='col-md-4']").filter(
        has_text=re.compile(r"Lecture\s*Date", re.I)
    ).first
    if panel.count() > 0:
        return panel

    cols = page.locator("div.col-md-4, div[class*='col-md-4']")
    for i in range(cols.count()):
        c = cols.nth(i)
        txt = (c.inner_text() or "").lower()
        if "today" in txt or "yesterday" in txt:
            continue
        has_input = c.locator("input[id*='txt'], input[name*='txt'], input[type='text']").count() > 0
        has_submit = (
            c.get_by_role("button", name=re.compile("submit", re.I)).count() > 0
            or c.locator("input[value*='SUBMIT'], button:has-text('SUBMIT')").count() > 0
        )
        if has_input and has_submit:
            return c
    return None

def login_if_needed(page):
    """If the current page is a login page, fill credentials and submit."""
    page.wait_for_load_state("networkidle")
    current_url = page.url.lower()
    if "login" in current_url or "signin" in current_url:
        print("Login page detected. Attempting automatic login.")
        # Try common field names – you may need to adjust these selectors!
        page.locator("input[name*='user'], input[name*='User'], input[type='text']").first.fill(ERP_USER)
        page.locator("input[type='password']").first.fill(ERP_PASS)
        page.locator("button[type='submit'], input[type='submit']").first.click()
        page.wait_for_load_state("networkidle")
        # After login, we should be at the attendance page
        print("Login submitted.")
    else:
        print("Already on attendance page, no login needed.")

def make_summary():
    # ... (your exact make_summary function here) ...
    df = pd.read_excel(OUT_FILE, sheet_name=RAW_SHEET)
    df["status"] = df["status"].astype(str).str.strip().str.title()
    df["subject"] = df["subject"].astype(str).str.strip()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df[df["status"].isin(["Present", "Absent"])]
    df = df[df["subject"] != ""]
    df = df.dropna(subset=["date"])

    subject_summary = df.groupby("subject").agg(
        total_lectures=("status", "count"),
        present=("status", lambda s: (s == "Present").sum()),
        absent=("status", lambda s: (s == "Absent").sum()),
    ).reset_index()
    subject_summary["percentage"] = (subject_summary["present"] / subject_summary["total_lectures"] * 100).round(2)
    subject_summary = subject_summary.sort_values("percentage").reset_index(drop=True)

    daily_summary = df.groupby("date").agg(
        lectures=("status", "count"),
        attended=("status", lambda s: (s == "Present").sum()),
        absent=("status", lambda s: (s == "Absent").sum()),
    ).reset_index()
    daily_summary["percentage"] = (daily_summary["attended"] / daily_summary["lectures"] * 100).round(2)
    daily_summary = daily_summary.sort_values("date").reset_index(drop=True)
    daily_summary["cum_lectures"] = daily_summary["lectures"].cumsum()
    daily_summary["cum_attended"] = daily_summary["attended"].cumsum()
    daily_summary["cumulative_percentage"] = (daily_summary["cum_attended"] / daily_summary["cum_lectures"] * 100).round(2)
    daily_summary["date"] = daily_summary["date"].dt.strftime("%Y-%m-%d")

    with pd.ExcelWriter(OUT_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        subject_summary.to_excel(writer, sheet_name=SUBJECT_SHEET, index=False)
        daily_summary.to_excel(writer, sheet_name=DAILY_SHEET, index=False)

    print(f"Summary sheets updated.")

def scrape_attendance():
    all_rows = []
    with sync_playwright() as p:
        # Use headless mode on CI, but we need to see if login works headlessly.
        # We'll keep headless=False temporarily for debugging, but on GitHub Actions we must use headless=True.
        browser = p.chromium.launch(headless=True)   # <<< headless mode for CI
        page = browser.new_page()
        page.goto(URL, wait_until="domcontentloaded")
        
        # If login is required, do it now
        login_if_needed(page)

        # Wait for the Lecture Date element to be sure we are on the attendance page
        try:
            page.wait_for_selector("text=Lecture Date", timeout=15000)
        except Exception:
            print("Warning: 'Lecture Date' not found. Page may not be ready.")
            # Optionally take a screenshot for debugging
            page.screenshot(path="debug_page.png")

        for d in daterange(START_DATE, END_DATE):
            # Format date like "Jan 9, 2026" (without leading zero)
            date_str = d.strftime("%b %-d, %Y") if os.name != "nt" else d.strftime("%b %d, %Y").replace(" 0", " ")

            panel = find_left_date_panel(page)
            if not panel:
                print(f"{d}: left panel not found")
                continue

            try:
                date_box = panel.locator("input[id*='txt'], input[name*='txt'], input[type='text']").first
                date_box.fill("")
                date_box.type(date_str)
            except Exception as e:
                print(f"{d}: date fill failed -> {e}")
                continue

            try:
                panel.get_by_role("button", name=re.compile("submit", re.I)).first.click()
            except Exception:
                try:
                    panel.locator("input[value*='SUBMIT'], button:has-text('SUBMIT')").first.click()
                except Exception as e:
                    print(f"{d}: submit click failed -> {e}")
                    continue

            page.wait_for_timeout(2000)

            try:
                panel_text = panel.inner_text()
                day_rows = parse_rows(panel_text, d)
                day_rows = [r for r in day_rows if r["subject"] and r["status"] in ["Present", "Absent"]]
                all_rows.extend(day_rows)
                print(f"{d}: {len(day_rows)} rows")
            except Exception as e:
                print(f"{d}: parse failed -> {e}")
                page.screenshot(path=f"error_{d}.png")

        browser.close()

    if not all_rows:
        raise Exception("No attendance rows parsed.")

    df = pd.DataFrame(all_rows).drop_duplicates(subset=["date", "lecture_no", "subject"])
    df["scrape_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with pd.ExcelWriter(OUT_FILE, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=RAW_SHEET, index=False)

    print(f"Saved raw data to {OUT_FILE} ({len(df)} rows)")

def main():
    scrape_attendance()
    make_summary()
    print("All done.")

if __name__ == "__main__":
    if not ERP_USER or not ERP_PASS:
        raise ValueError("ERP_USERNAME and ERP_PASSWORD environment variables must be set.")
    main()
