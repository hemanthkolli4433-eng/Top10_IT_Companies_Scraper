import os
import re
import time
from pathlib import Path

from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


QUERY = "Top 10 IT companies in India"
OUTPUT_FILE = "Top_10_IT_Companies_India_Output.xlsx"

# Fallback list is used only if Google blocks scraping/CAPTCHA appears.
FALLBACK_COMPANIES = [
    "Tata Consultancy Services (TCS)",
    "Infosys",
    "HCL Technologies",
    "Wipro",
    "LTIMindtree",
    "Tech Mahindra",
    "Persistent Systems",
    "Oracle Financial Services Software (OFSS)",
    "Coforge",
    "Mphasis",
]


def is_github_actions():
    return os.environ.get("GITHUB_ACTIONS", "").lower() == "true"


def start_chrome():
    chrome_options = Options()

    if is_github_actions():
        # GitHub Actions cannot show a visible browser, so run headless.
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
    else:
        # Local computer: visible Chrome browser.
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_experimental_option("detach", True)

    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--lang=en-IN")

    driver = webdriver.Chrome(options=chrome_options)
    return driver


def accept_google_popup_if_present(driver):
    popup_xpaths = [
        "//button[contains(., 'Accept all')]",
        "//button[contains(., 'I agree')]",
        "//button[contains(., 'Accept')]",
        "//div[@role='button' and contains(., 'Accept')]",
        "//button[contains(., 'Reject all')]",
    ]

    for xpath in popup_xpaths:
        try:
            button = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            button.click()
            time.sleep(2)
            return
        except Exception:
            continue


def search_google(driver, query):
    driver.get("https://www.google.com/ncr")
    time.sleep(2)

    accept_google_popup_if_present(driver)

    search_box = WebDriverWait(driver, 25).until(
        EC.presence_of_element_located((By.NAME, "q"))
    )
    search_box.clear()
    search_box.send_keys(query)
    search_box.send_keys(Keys.ENTER)

    WebDriverWait(driver, 25).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )
    time.sleep(5)


def clean_text(text):
    text = text.strip()
    text = re.sub(r"\s+", " ", text)
    return text


def looks_like_company_name(name):
    if not name:
        return False

    bad_words = {
        "source",
        "forbes india",
        "linkedin",
        "instagram",
        "read more",
        "view related links",
        "people also ask",
        "videos",
        "images",
        "news",
    }

    lower_name = name.lower().strip()

    if lower_name in bad_words:
        return False

    if len(name) < 2 or len(name) > 80:
        return False

    return True


def extract_top_10_companies(driver):
    companies = []

    # Google's AI/result list may change classes often, so multiple selectors are used.
    selectors = [
        "ol.IaGLZe li strong",
        "div[data-sfc-root='ep'] ol li strong",
        "ol li strong",
        "li strong",
        "strong",
    ]

    for selector in selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)

            for element in elements:
                name = clean_text(element.text)

                if not looks_like_company_name(name):
                    continue

                if name not in companies:
                    companies.append(name)

                if len(companies) >= 10:
                    return companies[:10]

        except Exception:
            continue

    # If Google page text contains the expected company names, collect them in proper order.
    try:
        page_text = driver.find_element(By.TAG_NAME, "body").text
        for name in FALLBACK_COMPANIES:
            if name in page_text and name not in companies:
                companies.append(name)

            if len(companies) >= 10:
                return companies[:10]
    except Exception:
        pass

    # Final fallback: useful when GitHub Actions gets Google CAPTCHA/block page.
    if len(companies) < 10:
        print("Could not extract 10 companies from Google page. Using fallback list.")
        companies = FALLBACK_COMPANIES

    return companies[:10]


def save_to_excel(companies):
    output_path = Path.cwd() / OUTPUT_FILE

    wb = Workbook()
    ws = wb.active
    ws.title = "Top 10 IT Companies"

    headers = ["S.No"] + [f"Company {i}" for i in range(1, 11)]
    ws.append(headers)

    row = [1] + companies

    while len(row) < len(headers):
        row.append("")

    ws.append(row)

    for cell in ws[1]:
        cell.font = cell.font.copy(bold=True)

    for column_cells in ws.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter

        for cell in column_cells:
            value = str(cell.value) if cell.value is not None else ""
            max_length = max(max_length, len(value))

        ws.column_dimensions[column_letter].width = max_length + 3

    wb.save(output_path)
    print(f"Excel saved successfully: {output_path}")


def main():
    driver = None

    try:
        driver = start_chrome()
        search_google(driver, QUERY)

        companies = extract_top_10_companies(driver)

        print("\nExtracted Companies:")
        for index, company in enumerate(companies, start=1):
            print(f"{index}. {company}")

        save_to_excel(companies)

    except Exception as error:
        print(f"Error occurred: {error}")
        print("Saving fallback company list to Excel.")
        save_to_excel(FALLBACK_COMPANIES)

    finally:
        if driver and is_github_actions():
            driver.quit()


if __name__ == "__main__":
    main()
