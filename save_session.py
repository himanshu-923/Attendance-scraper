from playwright.sync_api import sync_playwright

ATTENDANCE_URL = "https://student.cgc.ac.in/Attendance.aspx"

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(
        user_data_dir="erp_profile",
        headless=False
    )
    page = context.new_page()
    page.goto(ATTENDANCE_URL, wait_until="domcontentloaded")
    print("Login manually and wait until Attendance page is visible, then press Enter.")
    input()
    context.close()
