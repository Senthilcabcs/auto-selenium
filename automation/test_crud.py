import os
import re
import time
from urllib.parse import urlparse

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.drawing.image import Image as XLImage



BASE_URL = "http://127.0.0.1:8000"
INDEX_URL = f"{BASE_URL}/posts"
REPORT_DIR = Path(__file__).resolve().parent / "reports"
SCREENSHOT_DIR = REPORT_DIR / "screenshots"
REPORT_PATH = REPORT_DIR / "crud_test_report.xlsx"

options = webdriver.ChromeOptions()
# Uncomment for headless mode:
# options.add_argument("--headless")

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)


def fill_field(name, value):
    field = wait.until(
        EC.visibility_of_element_located((By.NAME, name))
    )
    field.clear()
    field.send_keys(value)


def click_form_submit(form):
    # Matches buttons without an explicit type, as well as submit inputs.
    buttons = form.find_elements(
        By.CSS_SELECTOR, "button, input[type='submit']"
    )

    if buttons:
        buttons[0].click()
    else:
        # Fallback if the form has no detectable submit button.
        driver.execute_script("arguments[0].requestSubmit()", form)


def post_ids_on_index():
    driver.get(INDEX_URL)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    ids = set()

    for link in driver.find_elements(By.CSS_SELECTOR, "a[href]"):
        path = urlparse(link.get_attribute("href")).path
        match = re.search(r"/posts/(\d+)/edit/?$", path)

        if match:
            ids.add(match.group(1))

    return ids



report_rows = []
current_step = "Startup"


TEST_CASES = {
    "Login":  ("TC-001", "Log in with valid credentials"),
    "Create": ("TC-002", "Create a post"),
    "Edit":   ("TC-003", "Edit a post"),
    "Delete": ("TC-004", "Delete a post"),
}


def record_result(step, status, details=""):
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    test_case_id, test_case = TEST_CASES.get(
        step, ("TC-000", step)
    )

    safe_step = re.sub(r"[^A-Za-z0-9_-]+", "_", step)
    filename = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    screenshot_path = SCREENSHOT_DIR / f"{filename}_{safe_step}.png"

    try:
        driver.save_screenshot(str(screenshot_path))
        screenshot_file = str(screenshot_path)
    except Exception:
        screenshot_file = ""

    report_rows.append([
        timestamp,
        test_case_id,
        test_case,
        step,
        status,
        details,
        screenshot_file,
    ])


def save_excel_report():
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "CRUD Test Report"

    sheet.append([
        "Timestamp",
        "Test Case ID",
        "Test Case",
        "Step",
        "Status",
        "Details",
        "Screenshot",
    ])

    for row_number, result in enumerate(report_rows, start=2):
        timestamp, case_id, case, step, status, details, screenshot_file = result

        sheet.append([
            timestamp, case_id, case, step, status, details, ""
        ])

        if screenshot_file and Path(screenshot_file).exists():
            image = XLImage(screenshot_file)

            scale = min(320 / image.width, 180 / image.height, 1)
            image.width = int(image.width * scale)
            image.height = int(image.height * scale)

            sheet.add_image(image, f"G{row_number}")
            sheet.row_dimensions[row_number].height = image.height * 0.75 + 10

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for column, width in {
        "A": 22, "B": 16, "C": 36, "D": 14,
        "E": 12, "F": 40, "G": 45,
    }.items():
        sheet.column_dimensions[column].width = width

    workbook.save(REPORT_PATH)
    print(f"Excel report saved: {REPORT_PATH}")
    
try:
    # Log in using credentials set in PowerShell environment variables.
    email = os.environ["TEST_EMAIL"]
    password = os.environ["TEST_PASSWORD"]
    current_step = "Login"
    driver.get(f"{BASE_URL}/login")

    fill_field("email", email)
    fill_field("password", password)

    login_form = wait.until(
        EC.presence_of_element_located((By.TAG_NAME, "form"))
    )
    click_form_submit(login_form)

    try:
        wait.until(
            lambda browser:
            urlparse(browser.current_url).path != "/login"
        )
    except TimeoutException:
        page_text = driver.find_element(By.TAG_NAME, "body").text
        raise RuntimeError(
            f"Login did not complete. Check the test account credentials. "
            f"Page says: {page_text}"
        )

    print("Logged in")
    record_result("Login", "PASS", "Logged in successfully")

    # CREATE
    existing_ids = post_ids_on_index()
    title = f"Selenium test {int(time.time())}"
    current_step = "Create"
    driver.get(f"{BASE_URL}/posts/create")
    fill_field("title", title)
    fill_field("body", "Created by Selenium")

    create_form = wait.until(
        EC.presence_of_element_located((By.TAG_NAME, "form"))
    )
    click_form_submit(create_form)

    # Find the new post using its edit link on the posts index.
    new_ids = post_ids_on_index() - existing_ids

    if not new_ids:
        raise RuntimeError(
            "Could not find the new post's edit link. Check that the post "
            "was created and that the index page has edit links."
        )

    post_id = new_ids.pop()
    print(f"Created post {post_id}")
    record_result("Create", "PASS", f"Created post {post_id}")

    # EDIT
    edited_title = title + " edited"
    current_step = "Edit"
    driver.get(f"{BASE_URL}/posts/{post_id}/edit")
    fill_field("title", edited_title)
    fill_field("body", "Edited by Selenium")

    edit_form = wait.until(
        EC.presence_of_element_located((By.TAG_NAME, "form"))
    )
    click_form_submit(edit_form)

    driver.get(INDEX_URL)
    wait.until(
        lambda browser:
        edited_title in browser.find_element(By.TAG_NAME, "body").text
    )
    print("Edited post successfully")
    record_result("Edit", "PASS", f"Edited post {post_id}")

    # DELETE
    current_step = "Delete"
    delete_button = wait.until(
        EC.element_to_be_clickable((By.ID, str(post_id)))
    )
    delete_button.click()

    # Accept a JavaScript confirmation dialog, if destroy() shows one.
    try:
        wait.until(EC.alert_is_present()).accept()
    except TimeoutException:
        pass

    # Wait for React to remove the post from the page.
    wait.until(
        lambda browser:
        edited_title not in browser.find_element(By.TAG_NAME, "body").text
    )

    print("Deleted post successfully")
    print("CRUD test passed")
    record_result("Delete", "PASS", f"Deleted post {post_id}")

except Exception as error:
    record_result(current_step, "FAIL", str(error))
    raise
finally:
    save_excel_report()
    driver.quit()