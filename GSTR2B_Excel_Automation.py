import re
import time
import shutil
import traceback
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import messagebox

from openpyxl import Workbook, load_workbook

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
    WebDriverException,
)


# ============================================================
# 1. CONFIGURATION
# ============================================================

LOGIN_URL = "https://services.gst.gov.in/services/login"
WELCOME_URL_PART = "/services/auth/fowelcome"

MAX_LOGIN_ATTEMPTS = 3
DOWNLOAD_TIMEOUT = 120
PAGE_TIMEOUT = 45

# Financial-year month numbering used by the selection prompt:
# 1=April, 2=May, 3=June, ... 12=March.
MONTH_NUMBER_TO_NAME = {
    1: "April",
    2: "May",
    3: "June",
    4: "July",
    5: "August",
    6: "September",
    7: "October",
    8: "November",
    9: "December",
    10: "January",
    11: "February",
    12: "March",
}
MONTHS = [MONTH_NUMBER_TO_NAME[i] for i in range(1, 13)]

QUARTERS = {
    "April": "Quarter 1 (Apr - Jun)",
    "May": "Quarter 1 (Apr - Jun)",
    "June": "Quarter 1 (Apr - Jun)",
    "July": "Quarter 2 (Jul - Sep)",
    "August": "Quarter 2 (Jul - Sep)",
    "September": "Quarter 2 (Jul - Sep)",
    "October": "Quarter 3 (Oct - Dec)",
    "November": "Quarter 3 (Oct - Dec)",
    "December": "Quarter 3 (Oct - Dec)",
    "January": "Quarter 4 (Jan - Mar)",
    "February": "Quarter 4 (Jan - Mar)",
    "March": "Quarter 4 (Jan - Mar)",
}

MAIN_OUTPUT_FOLDER = Path.home() / "Downloads" / "GSTR2B_DOWNLOADS"
MAIN_OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

LOG_FILE = MAIN_OUTPUT_FOLDER / "GSTR2B_Download_Log.xlsx"
BACKUP_LOG_FILE = MAIN_OUTPUT_FOLDER / "GSTR2B_Download_Log_Backup.xlsx"

LOG_HEADERS = [
    "GSTIN", "Trade Name", "Financial Year", "Month", "Status",
    "File", "File Path", "Remarks", "Date and Time",
]


# ============================================================
# 2. TKINTER POPUPS
# ============================================================

root = tk.Tk()
root.withdraw()


def show_message(title, message):
    messagebox.showinfo(title, message, parent=root)


def captcha_popup(gstin, attempt):
    result = {"value": None}

    popup = tk.Toplevel(root)
    popup.title("GST CAPTCHA")
    popup.geometry("390x185")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)

    tk.Label(
        popup,
        text=f"Enter CAPTCHA for GSTIN:\n{gstin}",
        font=("Arial", 11, "bold"),
        pady=10,
    ).pack()

    tk.Label(
        popup,
        text=f"Attempt {attempt} of {MAX_LOGIN_ATTEMPTS}",
    ).pack()

    entry = tk.Entry(popup, font=("Arial", 13), width=26)
    entry.pack(pady=10)
    entry.focus_set()

    def submit():
        result["value"] = entry.get().strip()
        popup.destroy()

    tk.Button(popup, text="OK", width=12, command=submit).pack()
    popup.bind("<Return>", lambda event: submit())
    popup.protocol("WM_DELETE_WINDOW", popup.destroy)
    popup.grab_set()
    root.wait_window(popup)

    return result["value"]


def credential_popup(gstin, username, password):
    result = {
        "username": username,
        "password": password,
        "cancelled": False,
    }

    popup = tk.Toplevel(root)
    popup.title("Correct GST Credentials")
    popup.geometry("430x245")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)

    tk.Label(
        popup,
        text=f"Check credentials for {gstin}",
        font=("Arial", 11, "bold"),
        pady=10,
    ).pack()

    tk.Label(popup, text="Username").pack()
    username_entry = tk.Entry(popup, width=44)
    username_entry.pack()
    username_entry.insert(0, username)

    tk.Label(popup, text="Password").pack()
    password_entry = tk.Entry(popup, width=44, show="*")
    password_entry.pack()
    password_entry.insert(0, password)

    def update():
        result["username"] = username_entry.get().strip()
        result["password"] = password_entry.get()
        popup.destroy()

    def cancel():
        result["cancelled"] = True
        popup.destroy()

    buttons = tk.Frame(popup)
    buttons.pack(pady=15)

    tk.Button(buttons, text="Update", width=12, command=update).pack(
        side="left", padx=5
    )
    tk.Button(buttons, text="Cancel", width=12, command=cancel).pack(
        side="left", padx=5
    )

    popup.protocol("WM_DELETE_WINDOW", cancel)
    popup.grab_set()
    root.wait_window(popup)
    return result


# ============================================================
# 3. EXCEL CREDENTIALS
# ============================================================

def normalize_header(value):
    return re.sub(r"[^a-z]", "", str(value).lower())


def read_credentials(file_path):
    workbook = load_workbook(file_path, data_only=True)
    try:
        sheet = workbook.active
        headers = {
            normalize_header(cell.value): cell.column
            for cell in sheet[1]
            if cell.value is not None
        }

        def get_column(aliases):
            for alias in aliases:
                key = normalize_header(alias)
                if key in headers:
                    return headers[key]
            raise ValueError(
                f"Missing Excel column. Expected one of: {aliases}"
            )

        gstin_col = get_column(["GSTIN", "GST Number", "GST No"])
        username_col = get_column(["Username", "User ID", "Userid"])
        password_col = get_column(["Password", "Pass"])

        trade_name_col = None
        for alias in ["Trade Name", "TradeName", "Business Name", "Legal Name"]:
            key = normalize_header(alias)
            if key in headers:
                trade_name_col = headers[key]
                break

        records = []
        seen_gstins = set()
        for row in range(2, sheet.max_row + 1):
            gstin = sheet.cell(row, gstin_col).value
            username = sheet.cell(row, username_col).value
            password = sheet.cell(row, password_col).value
            trade_name = (
                sheet.cell(row, trade_name_col).value
                if trade_name_col is not None else None
            )

            if gstin and username and password:
                gstin_text = str(gstin).strip()
                if gstin_text in seen_gstins:
                    raise ValueError(
                        f"Duplicate GSTIN found in Excel row {row}: {gstin_text}"
                    )
                seen_gstins.add(gstin_text)
                records.append({
                    "gstin": gstin_text,
                    "trade_name": str(trade_name).strip() if trade_name else gstin_text,
                    "username": str(username).strip(),
                    "password": str(password).strip(),
                })

        if not records:
            raise ValueError("No complete GSTIN credentials found.")
        return records
    finally:
        workbook.close()


# ============================================================
# 4. EXCEL LOGGING (MAIN + BACKUP)
# ============================================================

def create_log_file(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Download Log"
    ws.append(LOG_HEADERS)
    wb.save(path)
    wb.close()


def append_log_row(path, row):
    if not path.exists():
        create_log_file(path)

    # If an older PDF-version log exists with different columns, preserve it
    # and start a fresh log with the current Excel-download structure.
    try:
        check_wb = load_workbook(path, read_only=True, data_only=True)
        check_ws = check_wb.active
        existing_headers = [cell.value for cell in check_ws[1]]
        check_wb.close()
    except Exception:
        existing_headers = []

    if existing_headers != LOG_HEADERS:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        legacy_path = path.with_name(
            f"{path.stem}_Previous_{stamp}{path.suffix}"
        )
        shutil.copy2(path, legacy_path)
        create_log_file(path)
        print("Previous log format preserved as:", legacy_path.name)

    wb = load_workbook(path)
    try:
        wb.active.append(row)
        wb.save(path)
    finally:
        wb.close()


def write_log(gstin, financial_year, month, status, filename="", remarks="",
              trade_name="", file_path=""):
    row = [
        gstin,
        trade_name or gstin,
        financial_year,
        month,
        status,
        filename,
        str(file_path) if file_path else "",
        remarks,
        datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
    ]

    for attempt in range(1, 4):
        try:
            append_log_row(LOG_FILE, row)
            print("Main Excel log updated.")
            return True
        except PermissionError:
            print(f"Main log is open/locked. Retry {attempt}/3.")
            time.sleep(2)
        except Exception as exc:
            print("Main log error:", exc)
            break

    try:
        append_log_row(BACKUP_LOG_FILE, row)
        print("Entry saved in backup log.")
        return True
    except Exception as exc:
        print("Backup log error:", exc)
        return False


# ============================================================
# 5. CHROME
# ============================================================

def create_driver(download_folder):
    download_folder.mkdir(parents=True, exist_ok=True)

    options = webdriver.ChromeOptions()
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": str(download_folder.resolve()),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "safebrowsing.enabled": True,
        },
    )
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(60)
    return driver


# ============================================================
# 6. SELENIUM HELPERS
# ============================================================

def wait_for_page(driver):
    WebDriverWait(driver, PAGE_TIMEOUT).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )


def page_text(driver):
    return driver.find_element(By.TAG_NAME, "body").text.lower()


def check_access_denied(driver):
    if "accessdenied" in driver.current_url.lower():
        raise RuntimeError(
            "GST portal displayed Access Denied. "
            "The portal session or navigation was rejected."
        )


def safe_click(driver, element):
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});", element
    )
    time.sleep(0.4)
    try:
        element.click()
    except (ElementClickInterceptedException, StaleElementReferenceException):
        driver.execute_script("arguments[0].click();", element)


# ============================================================
# 7. CAPTCHA FIELD
# ============================================================

def find_captcha_field(driver):
    selectors = [
        "input#captcha",
        "input[name='captcha']",
        "input[id*='captcha']",
        "input[name*='captcha']",
        "input[placeholder*='Captcha']",
        "input[placeholder*='captcha']",
    ]

    for selector in selectors:
        for field in driver.find_elements(By.CSS_SELECTOR, selector):
            if field.is_displayed() and field.is_enabled():
                return field

    for field in driver.find_elements(By.TAG_NAME, "input"):
        attributes = " ".join([
            field.get_attribute("id") or "",
            field.get_attribute("name") or "",
            field.get_attribute("placeholder") or "",
        ]).lower()

        if "captcha" in attributes and field.is_displayed() and field.is_enabled():
            return field

    raise RuntimeError(
        "Could not identify CAPTCHA input. The GST login page selector may have changed."
    )


# ============================================================
# 8. LOGIN AND WELCOME PAGE
# ============================================================

def is_welcome_page(driver):
    return WELCOME_URL_PART in driver.current_url.lower()


def wait_for_login_success(driver, gstin):
    print("Waiting for GST welcome page...")

    try:
        WebDriverWait(driver, 30).until(
            lambda d: is_welcome_page(d)
        )
        print("GST login successful.")
        return True
    except TimeoutException:
        pass

    show_message(
        "GST Verification Required",
        f"The GST welcome page has not opened for {gstin}.\n\n"
        "Check Chrome. If the portal asks for OTP or additional "
        "verification, complete it manually, then click OK.",
    )

    WebDriverWait(driver, 180).until(
        lambda d: is_welcome_page(d)
    )
    print("GST login successful after manual verification.")
    return True


def handle_metadata_popup(driver):
    """
    Clicks only the portal's explicit 'NO-REMIND ME LATER' option.
    If it is not found, the function continues without clicking another option.
    """
    print("Checking for Principal Place of Business popup...")

    xpath = (
        "//*[self::button or self::a or self::input]"
        "[contains(translate("
        "concat(normalize-space(.), ' ', @value),"
        "'abcdefghijklmnopqrstuvwxyz',"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
        "'NO-REMIND ME LATER')]"
    )

    try:
        button = WebDriverWait(driver, 12).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
    except TimeoutException:
        print("Metadata popup not found.")
        return False

    print("Metadata popup found. Clicking NO-REMIND ME LATER...")
    safe_click(driver, button)

    # Wait for the popup text/button to disappear or detect server redirect.
    end_time = time.time() + 15
    while time.time() < end_time:
        check_access_denied(driver)

        try:
            matches = driver.find_elements(By.XPATH, xpath)
            visible = any(item.is_displayed() for item in matches)
        except WebDriverException:
            visible = False

        if not visible:
            print("NO-REMIND ME LATER popup closed.")
            return True

        time.sleep(0.5)

    print(
        "The metadata prompt did not disappear within the wait period. "
        "Leaving Chrome open for inspection."
    )
    return False


def login_to_gst(driver, credential):
    gstin = credential["gstin"]
    username = credential["username"]
    password = credential["password"]

    for attempt in range(1, MAX_LOGIN_ATTEMPTS + 1):
        print(f"\nLogin attempt {attempt}/{MAX_LOGIN_ATTEMPTS}: {gstin}")

        driver.get(LOGIN_URL)
        wait_for_page(driver)

        username_field = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "input#username"))
        )
        password_field = driver.find_element(By.CSS_SELECTOR, "input#user_pass")

        username_field.clear()
        username_field.send_keys(username)
        password_field.clear()
        password_field.send_keys(password)

        captcha_field = find_captcha_field(driver)
        captcha = captcha_popup(gstin, attempt)

        if not captcha:
            raise RuntimeError("CAPTCHA entry was cancelled or left blank.")

        captcha_field.clear()
        captcha_field.send_keys(captcha)

        login_buttons = driver.find_elements(
            By.XPATH,
            "//button[contains(translate(normalize-space(.),"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'login')]"
            " | //input[@type='submit']",
        )
        login_buttons = [
            button for button in login_buttons
            if button.is_displayed() and button.is_enabled()
        ]

        if not login_buttons:
            raise RuntimeError("Login button not found.")

        safe_click(driver, login_buttons[0])
        time.sleep(4)

        current_text = page_text(driver)

        if any(term in current_text for term in [
            "invalid captcha",
            "incorrect captcha",
            "captcha is invalid",
        ]):
            print("Invalid CAPTCHA.")
            if attempt == MAX_LOGIN_ATTEMPTS:
                raise RuntimeError("Maximum CAPTCHA attempts reached.")
            continue

        if any(term in current_text for term in [
            "invalid username",
            "invalid password",
            "invalid credentials",
        ]):
            print("Invalid username/password.")
            updated = credential_popup(gstin, username, password)

            if updated["cancelled"]:
                raise RuntimeError("Credential correction cancelled.")

            username = updated["username"]
            password = updated["password"]
            credential["username"] = username
            credential["password"] = password

            if attempt == MAX_LOGIN_ATTEMPTS:
                raise RuntimeError("Maximum login attempts reached.")
            continue

        wait_for_login_success(driver, gstin)

        # Handle this portal prompt after successful login.
        handle_metadata_popup(driver)

        check_access_denied(driver)
        return True

    raise RuntimeError("Unable to log in after maximum attempts.")


# ============================================================
# 9. RETURN DASHBOARD NAVIGATION
# ============================================================

def open_return_dashboard(driver):
    """Open the Returns Dashboard using the portal's visible navigation."""
    check_access_denied(driver)

    # If already on the Returns Dashboard, do not navigate away.
    if "return.gst.gov.in" in driver.current_url.lower():
        if any(
            item.is_displayed()
            for item in driver.find_elements(By.TAG_NAME, "select")
        ):
            print("Already on Returns Dashboard.")
            return True

    # If the browser is on a GSTR-2B summary, use its own BACK TO DASHBOARD button.
    if "gstr2b.gst.gov.in" in driver.current_url.lower():
        return_to_dashboard(driver)
        return True

    print("Locating Return Dashboard link on GST welcome page...")
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)

    locators = [
        (By.XPATH, "//span[@title='Return Dashboard']"),
        (By.XPATH,
         "//*[self::a or self::span or self::button]"
         "[normalize-space()='Return Dashboard']"),
    ]

    link = None
    for locator in locators:
        for element in driver.find_elements(*locator):
            try:
                if element.is_displayed() and element.is_enabled():
                    link = element
                    break
            except StaleElementReferenceException:
                continue
        if link:
            break

    if link is None:
        save_debug_files(driver, MAIN_OUTPUT_FOLDER, "Return_Dashboard_Link_Debug")
        raise RuntimeError(
            "Could not find the Return Dashboard link on the current page. "
            "A screenshot and HTML debug file were saved."
        )

    safe_click(driver, link)
    print("Return Dashboard link clicked.")

    WebDriverWait(driver, 45).until(
        lambda d: (
            "accessdenied" in d.current_url.lower()
            or "return.gst.gov.in" in d.current_url.lower()
        )
    )
    check_access_denied(driver)

    WebDriverWait(driver, 35).until(
        lambda d: any(
            element.is_displayed()
            for element in d.find_elements(By.TAG_NAME, "select")
        )
    )
    print("Return Dashboard loaded.")
    return True

def select_native_option(driver, target):
    target_clean = re.sub(r"\s+", " ", target.strip().lower())

    for dropdown in driver.find_elements(By.TAG_NAME, "select"):
        try:
            if not dropdown.is_displayed() or not dropdown.is_enabled():
                continue

            select = Select(dropdown)

            for option in select.options:
                option_clean = re.sub(
                    r"\s+", " ", option.text.strip().lower()
                )

                if option_clean == target_clean:
                    select.select_by_visible_text(option.text)
                    time.sleep(1)
                    print("Selected:", option.text)
                    return True
        except Exception:
            continue

    raise RuntimeError(f"Dropdown option not found: {target}")


def select_period_details(driver, financial_year, month):
    if month not in MONTHS:
        raise ValueError(f"Invalid financial-year month: {month}")

    expected_number = next(
        number for number, name in MONTH_NUMBER_TO_NAME.items()
        if name == month
    )
    print(
        f"\nSelecting financial-year month {expected_number}: {month}"
    )

    print("Selecting Financial Year:", financial_year)
    select_native_option(driver, financial_year)

    print("Selecting Quarter:", QUARTERS[month])
    select_native_option(driver, QUARTERS[month])

    # Wait for the period dropdown to refresh after the quarter selection.
    WebDriverWait(driver, 20).until(
        lambda d: any(
            element.is_displayed()
            and any(
                re.sub(r"\\s+", " ", option.text.strip().lower()) == month.lower()
                for option in Select(element).options
            )
            for element in d.find_elements(By.TAG_NAME, "select")
            if element.is_enabled()
        )
    )

    print(f"Selecting Period: {month}")
    select_native_option(driver, month)

    # Verify that the requested month, rather than the dropdown's default
    # (which may be May), is actually selected before pressing SEARCH.
    selected_month_visible = any(
        element.is_displayed()
        and Select(element).first_selected_option.text.strip().lower() == month.lower()
        for element in driver.find_elements(By.TAG_NAME, "select")
        if element.is_enabled()
    )
    if not selected_month_visible:
        raise RuntimeError(
            f"Month selection verification failed: expected {month}. "
            "SEARCH was not clicked."
        )

    print(f"Verified selected month: {month}")


# ============================================================
# 11. SEARCH
# ============================================================

def click_search(driver):
    print("Clicking SEARCH...")

    search_button = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//button[normalize-space()='SEARCH' "
            "or normalize-space()='Search']",
        ))
    )

    safe_click(driver, search_button)
    print("SEARCH clicked.")

    WebDriverWait(driver, 45).until(
        lambda d: "gstr-2b" in d.find_element(By.TAG_NAME, "body").text.lower()
    )

    check_access_denied(driver)
    time.sleep(2)
    print("Search results loaded.")


# ============================================================
# 11A. DETECT NO GSTR-2B AND SAVE SCREENSHOT
# ============================================================

def capture_no_gstr2b_if_present(driver, gstin_folder, gstin, financial_year, month,
                                  trade_name=None):
    """Detect the no-generation notice on the GSTR-2B summary page and save a screenshot."""
    check_access_denied(driver)
    body = driver.find_element(By.TAG_NAME, "body").text
    normalized = re.sub(r"\s+", " ", body).lower()

    no_2b_phrases = [
        "gstr-2b could not be generated by the system",
        "gstr 2b could not be generated by the system",
        "there are no records to generate gstr 2b",
        "system will not generate gstr 2b",
        "gstr-2b could not be generated",
    ]

    if not any(phrase in normalized for phrase in no_2b_phrases):
        return None

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    screenshot_name = (
        f"NO_GSTR2B_{safe_windows_name(trade_name or gstin)}_"
        f"{financial_year}_{month}_{timestamp}.png"
    )
    screenshot_path = gstin_folder / screenshot_name
    driver.save_screenshot(str(screenshot_path))

    print(f"No GSTR-2B was generated for {month}.")
    print("Evidence screenshot saved:", screenshot_path)
    return screenshot_path

def save_debug_files(driver, gstin_folder, prefix="GST_Debug"):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    screenshot_path = gstin_folder / f"{prefix}_{timestamp}.png"
    html_path = gstin_folder / f"{prefix}_{timestamp}.html"

    try:
        driver.save_screenshot(str(screenshot_path))
        print("Screenshot saved:", screenshot_path)
    except Exception as exc:
        print("Could not save screenshot:", exc)

    try:
        html_path.write_text(driver.page_source, encoding="utf-8")
        print("Page HTML saved:", html_path)
    except Exception as exc:
        print("Could not save HTML:", exc)


# ============================================================
# 13. OPEN GSTR-2B VIEW
# ============================================================

def open_gstr2b(driver, gstin_folder):
    print("Locating GSTR-2B card and its VIEW button...")
    check_access_denied(driver)

    view_button = driver.execute_script("""
        function clean(value) {
            return (value || '').replace(/\\\\s+/g, ' ').trim().toLowerCase();
        }

        const leafHeadings = Array.from(document.querySelectorAll('body *'))
            .filter(el => el.children.length === 0 &&
                          clean(el.textContent) === 'gstr-2b');

        for (const heading of leafHeadings) {
            let parent = heading;

            for (let level = 0; level < 10 && parent; level++) {
                const controls = Array.from(
                    parent.querySelectorAll('button, a, input[type="button"]')
                ).filter(el => el.offsetParent !== null);

                const view = controls.find(el =>
                    clean(el.innerText || el.value) === 'view'
                );

                const download = controls.find(el =>
                    clean(el.innerText || el.value).includes('download')
                );

                if (view && download) return view;
                parent = parent.parentElement;
            }
        }
        return null;
    """)

    if view_button is None:
        save_debug_files(driver, gstin_folder, "GSTR2B_View_Debug")
        raise RuntimeError(
            "Could not identify the GSTR-2B VIEW button. "
            "Debug screenshot and HTML were saved."
        )

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        view_button,
    )
    time.sleep(0.5)
    driver.execute_script("arguments[0].click();", view_button)
    time.sleep(4)
    check_access_denied(driver)
    print("GSTR-2B VIEW clicked.")


# ============================================================
# 14. DOWNLOAD GSTR-2B EXCEL
# ============================================================

def safe_windows_name(value):
    """Return a Windows-safe folder/file base name."""
    name = re.sub(r'[<>:"/\\\\|?*]', "_", str(value or "")).strip()
    name = re.sub(r"\\s+", " ", name).rstrip(". ")
    return name[:100] if name else "Unnamed_Trade"


def click_first_visible(driver, xpath, description, timeout=20):
    candidates = WebDriverWait(driver, timeout).until(
        EC.presence_of_all_elements_located((By.XPATH, xpath))
    )
    for element in candidates:
        try:
            if element.is_displayed() and element.is_enabled():
                safe_click(driver, element)
                print(description)
                return True
        except (StaleElementReferenceException, WebDriverException):
            continue
    raise RuntimeError(f"Could not find a visible, enabled control: {description}")


def open_gstr2b_download_page(driver, gstin_folder):
    """
    Open the GSTR-2B VIEW/summary page (the same page used by the
    earlier PDF downloader). Do NOT click the tile-level Download action,
    because that opens the separate offline/generation workflow.
    """
    check_access_denied(driver)
    print("Opening GSTR-2B summary through the VIEW button...")
    open_gstr2b(driver, gstin_folder)

    details_excel_xpath = (
        "//*[self::button or self::a or self::input or self::span]"
        "[contains(translate(normalize-space(concat(., ' ', @value)),"
        "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
        "'DOWNLOAD GSTR-2B DETAILS (EXCEL)')]"
    )

    try:
        WebDriverWait(driver, 25).until(
            lambda d: any(
                element.is_displayed() and element.is_enabled()
                for element in d.find_elements(By.XPATH, details_excel_xpath)
            )
        )
    except TimeoutException:
        save_debug_files(driver, gstin_folder, "GSTR2B_Details_Excel_Button_Not_Found")
        raise RuntimeError(
            "The GSTR-2B summary page opened, but "
            "'DOWNLOAD GSTR-2B DETAILS (EXCEL)' was not found. "
            "Debug screenshot and HTML were saved. The script will not "
            "open the separate Excel generation page."
        )

    print("Found DOWNLOAD GSTR-2B DETAILS (EXCEL) on the summary page.")
    return "details"

def wait_for_new_excel(download_folder, previous_files):
    start_time = time.time()
    excel_extensions = {".xlsx", ".xls"}
    while time.time() - start_time < DOWNLOAD_TIMEOUT:
        files = list(download_folder.iterdir())
        in_progress = [
            f for f in files
            if f.name.lower().endswith((".crdownload", ".tmp", ".part"))
        ]
        new_excels = [
            f for f in files
            if f.is_file()
            and f.suffix.lower() in excel_extensions
            and f.resolve() not in previous_files
        ]
        if new_excels and not in_progress:
            latest = max(new_excels, key=lambda f: f.stat().st_mtime)
            if latest.stat().st_size > 0:
                return latest
        time.sleep(1)
    raise TimeoutException("GSTR-2B Excel download timed out.")


def download_month_excel(driver, gstin, financial_year, month, output_folder,
                         trade_name, single_month):
    previous_files = {
        f.resolve() for f in output_folder.iterdir()
        if f.is_file() and f.suffix.lower() in {".xlsx", ".xls"}
    }

    route = open_gstr2b_download_page(driver, output_folder)

    # The portal may show the no-generation message only after opening
    # the GSTR-2B tile/download page. Capture it here before looking for Excel.
    page_text_now = re.sub(
        r"\\s+", " ",
        driver.find_element(By.TAG_NAME, "body").text
    ).lower()
    no_2b_phrases = [
        "gstr-2b could not be generated by the system",
        "gstr 2b could not be generated by the system",
        "gstr-2b could not be generated",
    ]
    if any(phrase in page_text_now for phrase in no_2b_phrases):
        screenshot_path = output_folder / (
            f"NO_GSTR2B_{safe_windows_name(trade_name or gstin)}_"
            f"{financial_year}_{month}_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        )
        driver.save_screenshot(str(screenshot_path))
        print("No GSTR-2B notice found. Screenshot saved:", screenshot_path)
        return None

    button_xpath = (
        "//*[self::button or self::a or self::input or self::span]"
        "[contains(translate(normalize-space(concat(., ' ', @value)),"
        "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
        "'DOWNLOAD GSTR-2B DETAILS (EXCEL)')]"
    )
    click_first_visible(
        driver, button_xpath,
        "DOWNLOAD GSTR-2B DETAILS (EXCEL) clicked."
    )

    downloaded = wait_for_new_excel(output_folder, previous_files)

    if single_month:
        final_name = f"{safe_windows_name(trade_name)}{downloaded.suffix.lower()}"
    else:
        final_name = (
            f"GSTR2B_{safe_windows_name(trade_name)}_"
            f"{financial_year}_{month}{downloaded.suffix.lower()}"
        )

    final_path = output_folder / final_name
    if final_path.exists():
        # Preserve existing data; add a timestamp rather than overwriting.
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        final_path = output_folder / (
            f"{final_path.stem}_{stamp}{final_path.suffix}"
        )

    shutil.move(str(downloaded), str(final_path))
    print("GSTR-2B Excel saved:", final_path)
    return final_path


# ============================================================
# 14. DOWNLOAD SUMMARY PDF
# ============================================================

def click_download_summary(driver):
    xpath = (
        "//*[self::button or self::a or self::span]"
        "[contains(translate(normalize-space(.),"
        "'abcdefghijklmnopqrstuvwxyz',"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
        "'DOWNLOAD GSTR-2B SUMMARY')]"
    )

    elements = driver.find_elements(By.XPATH, xpath)

    for element in elements:
        try:
            if element.is_displayed() and element.is_enabled():
                safe_click(driver, element)
                print("DOWNLOAD GSTR-2B SUMMARY clicked.")
                return True
        except Exception:
            continue

    raise RuntimeError(
        "DOWNLOAD GSTR-2B SUMMARY (PDF) button not found."
    )


def wait_for_new_pdf(download_folder, previous_files):
    start_time = time.time()

    while time.time() - start_time < DOWNLOAD_TIMEOUT:
        all_files = list(download_folder.iterdir())

        in_progress = [
            item for item in all_files
            if item.name.lower().endswith(".crdownload")
        ]

        new_pdfs = [
            item for item in all_files
            if item.suffix.lower() == ".pdf"
            and item.resolve() not in previous_files
        ]

        if new_pdfs and not in_progress:
            latest = max(new_pdfs, key=lambda item: item.stat().st_mtime)
            if latest.stat().st_size > 0:
                return latest

        time.sleep(1)

    raise TimeoutException("PDF download timed out.")


def download_month_pdf(driver, gstin, financial_year, month, gstin_folder):
    previous_files = {
        item.resolve() for item in gstin_folder.glob("*.pdf")
    }

    click_download_summary(driver)

    downloaded = wait_for_new_pdf(gstin_folder, previous_files)

    final_name = f"GSTR2B_{gstin}_{financial_year}_{month}.pdf"
    final_path = gstin_folder / final_name

    if final_path.exists():
        final_path.unlink()

    shutil.move(str(downloaded), str(final_path))
    print("PDF saved:", final_path)
    return final_path


# ============================================================
# 15. RETURN FROM VIEW PAGE
# ============================================================

def return_to_dashboard(driver):
    """Return from GSTR-2B summary using the portal's BACK TO DASHBOARD button."""
    check_access_denied(driver)

    # If already on the returns search page, no action is required.
    if "return.gst.gov.in" in driver.current_url.lower():
        has_selects = any(
            element.is_displayed()
            for element in driver.find_elements(By.TAG_NAME, "select")
        )
        if has_selects:
            print("Already on Returns Dashboard.")
            return True

    print("Looking for BACK TO DASHBOARD button...")
    xpath = (
        "//*[self::button or self::a or self::input]"
        "[contains(translate(concat(normalize-space(.), ' ', @value),"
        "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
        "'BACK TO DASHBOARD')]"
    )

    button = None
    try:
        candidates = WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((By.XPATH, xpath))
        )
        for candidate in candidates:
            try:
                if candidate.is_displayed() and candidate.is_enabled():
                    button = candidate
                    break
            except StaleElementReferenceException:
                continue
    except TimeoutException:
        button = None

    if button is not None:
        safe_click(driver, button)
        print("BACK TO DASHBOARD clicked.")
    else:
        # Fallback only if the portal button is not present.
        print("BACK TO DASHBOARD button not found; trying browser Back once.")
        driver.back()

    WebDriverWait(driver, 40).until(
        lambda d: "accessdenied" in d.current_url.lower()
        or (
            "return.gst.gov.in" in d.current_url.lower()
            and any(
                element.is_displayed()
                for element in d.find_elements(By.TAG_NAME, "select")
            )
        )
    )
    check_access_denied(driver)
    print("Returned to Returns Dashboard.")
    return True

def process_gstin(credential, financial_year, selected_months):
    gstin = credential["gstin"]
    trade_name = credential.get("trade_name") or gstin
    safe_trade = safe_windows_name(trade_name)
    multiple_months = len(selected_months) > 1

    # Multiple months -> one Trade Name folder.
    # Single month -> Excel file directly in the main output folder.
    output_folder = (
        MAIN_OUTPUT_FOLDER / safe_trade if multiple_months else MAIN_OUTPUT_FOLDER
    )
    output_folder.mkdir(parents=True, exist_ok=True)

    driver = None
    processing_completed = False
    keep_browser_open = False

    print("\\n" + "=" * 65)
    print("PROCESSING GSTIN:", gstin)
    print("TRADE NAME:", trade_name)
    print("Selected months:", ", ".join(selected_months))
    print("Saving location:", output_folder)
    print("=" * 65)

    try:
        driver = create_driver(output_folder)
        login_to_gst(driver, credential)
        open_return_dashboard(driver)

        for index, month in enumerate(selected_months):
            print("\\n" + "-" * 55)
            print(f"PROCESSING {trade_name} | {financial_year} | {month}")
            print("-" * 55)

            try:
                expected_extension_candidates = [".xlsx", ".xls"]
                if multiple_months:
                    expected_stem = (
                        f"GSTR2B_{safe_trade}_{financial_year}_{month}"
                    )
                else:
                    expected_stem = safe_trade

                existing_download = next(
                    (
                        output_folder / f"{expected_stem}{ext}"
                        for ext in expected_extension_candidates
                        if (output_folder / f"{expected_stem}{ext}").exists()
                    ),
                    None,
                )
                if existing_download:
                    print("Already downloaded; skipping:", existing_download.name)
                    write_log(
                        gstin, financial_year, month, "SKIPPED",
                        existing_download.name,
                        "Existing Excel file found; skipped to avoid duplicate download.",
                        trade_name=trade_name, file_path=existing_download,
                    )
                    continue

                select_period_details(driver, financial_year, month)
                click_search(driver)

                # Check for the no-generation notice before opening any tile.
                body_now = re.sub(
                    r"\\s+", " ",
                    driver.find_element(By.TAG_NAME, "body").text
                ).lower()
                no_notice = (
                    "gstr-2b could not be generated by the system" in body_now
                    or "gstr 2b could not be generated by the system" in body_now
                )

                if no_notice:
                    screenshot_path = output_folder / (
                        f"NO_GSTR2B_{safe_trade}_{financial_year}_{month}_"
                        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                    )
                    driver.save_screenshot(str(screenshot_path))
                    write_log(
                        gstin, financial_year, month, "NO GSTR-2B",
                        screenshot_path.name,
                        "Portal reported that GSTR-2B could not be generated.",
                        trade_name=trade_name, file_path=screenshot_path,
                    )
                    print("No GSTR-2B. Screenshot saved:", screenshot_path)
                else:
                    excel_path = download_month_excel(
                        driver, gstin, financial_year, month, output_folder,
                        trade_name=trade_name, single_month=not multiple_months,
                    )
                    if excel_path is None:
                        screenshot_path = max(
                            output_folder.glob(f"NO_GSTR2B_{safe_trade}_{financial_year}_{month}_*.png"),
                            key=lambda item: item.stat().st_mtime,
                        )
                        write_log(
                            gstin, financial_year, month, "NO GSTR-2B",
                            screenshot_path.name,
                            "Portal reported that GSTR-2B could not be generated.",
                            trade_name=trade_name, file_path=screenshot_path,
                        )
                    else:
                        write_log(
                            gstin, financial_year, month, "SUCCESS",
                            excel_path.name,
                            "GSTR-2B Excel downloaded successfully.",
                            trade_name=trade_name, file_path=excel_path,
                        )
                        print("Completed:", excel_path)

                if index < len(selected_months) - 1:
                    # Download route may be on offline page or summary page.
                    # Use portal's Back control where available, otherwise the
                    # browser Back action to return to the GST returns dashboard.
                    return_to_dashboard(driver)

            except Exception as month_error:
                print(f"{month} FAILED: {month_error}")
                traceback.print_exc()
                write_log(
                    gstin, financial_year, month, "FAILED", "",
                    str(month_error), trade_name=trade_name,
                )

                if driver is None or "accessdenied" in driver.current_url.lower():
                    raise

                if index < len(selected_months) - 1:
                    return_to_dashboard(driver)

        processing_completed = True
        print("\\nCompleted Trade Name:", trade_name)

    except Exception as error:
        keep_browser_open = True
        print("\\nGSTIN PROCESS FAILED:", gstin, str(error))
        traceback.print_exc()
        write_log(
            gstin, financial_year, "ALL", "FAILED", "", str(error),
            trade_name=trade_name,
        )
        show_message(
            "GST Process Error",
            f"GSTIN: {gstin}\\nTrade Name: {trade_name}\\n\\n"
            f"Error:\\n{error}\\n\\n"
            "Chrome will remain open for inspection. Check the browser "
            "and Python terminal.",
        )

    finally:
        if driver:
            if keep_browser_open or not processing_completed:
                print("\\nChrome is being kept open for inspection.")
                input(
                    "\\nAfter inspecting Chrome, press ENTER here "
                    "to close this browser..."
                )
            try:
                driver.quit()
            except Exception:
                pass


def ask_month_selection():
    """
    Let the user choose individual months, a month range, multiple ranges,
    or all 12 months.

    Examples:
      1       = April
      1-3     = April to June
      4-6     = July to September
      1-3,7-9 = April to June and October to December
      1,3,5   = April, June and August
      13      = All months (April to March)
    """
    print("\nChoose the GSTR-2B download month(s):")

    for number, month in enumerate(MONTHS, start=1):
        print(f"{number:2}. {month}")

    print("13. All months (April to March)")
    print("\nYou can enter a single month, a range, or combine selections.")
    print("Examples: 1-3 | 4-6 | 1-3,7-9 | 1,3,5 | 13")

    choice = input("\nEnter your month selection: ").strip().replace(" ", "")

    if not choice:
        raise ValueError("No month selection entered.")

    if choice == "13":
        print("\nSelected: All months (April to March)")
        return MONTHS.copy()

    # The all-months option must be selected on its own.
    if "13" in choice.split(",") or re.search(r"(^|[, -])13($|[, -])", choice):
        raise ValueError("Use 13 by itself to select all 12 months.")

    selected_numbers = set()

    try:
        for part in choice.split(","):
            if not part:
                raise ValueError

            if "-" in part:
                bounds = part.split("-")
                if len(bounds) != 2 or not all(bound.isdigit() for bound in bounds):
                    raise ValueError

                start_num, end_num = map(int, bounds)

                if not (1 <= start_num <= 12 and 1 <= end_num <= 12):
                    raise ValueError
                if start_num > end_num:
                    raise ValueError

                selected_numbers.update(range(start_num, end_num + 1))
            else:
                if not part.isdigit():
                    raise ValueError

                month_num = int(part)
                if not 1 <= month_num <= 12:
                    raise ValueError

                selected_numbers.add(month_num)

    except ValueError:
        raise ValueError(
            "Invalid selection. Use 1-12, a range such as 1-3, "
            "or combined selections such as 1-3,7-9. Enter 13 for all months."
        )

    if not selected_numbers:
        raise ValueError("Please select at least one month.")

    # Preserve financial-year order (April to March), regardless of input order.
    selected_months = [
        MONTH_NUMBER_TO_NAME[number]
        for number in sorted(selected_numbers)
    ]

    if len(selected_months) == 1:
        print(f"\nSelected month: {selected_months[0]}")
    else:
        print("\nSelected months:")
        for month in selected_months:
            print(f" - {month}")

    return selected_months


# ============================================================
# 17. MAIN PROGRAM
# ============================================================

def main():
    print("=" * 65)
    print("             GSTR-2B EXCEL DOWNLOADER")
    print("=" * 65)

    print("\nExcel output folder:")
    print(MAIN_OUTPUT_FOLDER)

    excel_path = input(
        "\nEnter full path of Excel credentials file: "
    ).strip().strip('"')

    excel_path = Path(excel_path)

    if not excel_path.exists():
        print("Excel credentials file not found.")
        return

    financial_year = input(
        "\nEnter Financial Year (example: 2025-26): "
    ).strip()

    if not re.fullmatch(r"\d{4}-\d{2}", financial_year):
        print("Invalid financial year. Example: 2025-26")
        return

    try:
        credentials = read_credentials(excel_path)

        # Ask the user to choose one named month or all 12 months.
        selected_months = ask_month_selection()

        if not LOG_FILE.exists():
            create_log_file(LOG_FILE)

        print(f"\nTotal GSTINs found: {len(credentials)}")
        print("Months selected:", ", ".join(selected_months))

        for credential in credentials:
            process_gstin(credential, financial_year, selected_months)

        print("\n" + "=" * 65)
        print("PROCESSING FINISHED")
        print("=" * 65)
        print("Output folder:", MAIN_OUTPUT_FOLDER)
        print("Main log:", LOG_FILE)
        print("Backup log:", BACKUP_LOG_FILE)

        show_message(
            "GSTR-2B Processing Finished",
            f"Output folder:\n{MAIN_OUTPUT_FOLDER}\n\n"
            f"Main log:\n{LOG_FILE}\n\n"
            f"Backup log:\n{BACKUP_LOG_FILE}",
        )

    except Exception as error:
        print("\nPROCESS ERROR:", error)
        traceback.print_exc()
        show_message("Process Error", str(error))

    finally:
        try:
            root.destroy()
        except Exception:
            pass


if __name__ == "__main__":
    main()
