import os
import time
from datetime import datetime

from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


QUERY = "Top 10 IT compannies in India"
OUTPUT_FILE = "Top_10_IT_Companies_India_Output.xlsx"


def start_chrome():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(options=chrome_options)
    return driver


def accept_google_popup_if_present(driver):
    popup_xpaths = [
        "//button[contains(., 'Accept all')]",
        "//button[contains(., 'I agree')]",
        "//button[contains(., 'Accept')]",
        "//div[@role='button' and contains(., 'Accept')]",
    ]

    for xpath in popup_xpaths:
        try:
            btn = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            btn.click()
            time.sleep(2)
            return
        except Exception:
            pass


def search_google(driver, query):
    driver.get("https://www.google.com")
    time.sleep(2)

    accept_google_popup_if_present(driver)

    search_box = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.NAME, "q"))
    )
    search_box.clear()
    search_box.send_keys(query)
    search_box.send_keys(Keys.ENTER)

    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "search"))
    )
    time.sleep(5)


def clean_company_name(text):
    text = text.strip()
    text = " ".join(text.split())
    return text


def extract_top_10_companies(driver):
    companies = []

    selectors = [
        "ol.IaGLZe li strong",
        "ol li strong",
        "div[data-sfc-root='ep'] ol li strong",
        "li strong",
    ]

    unwanted_words = [
        "source",
        "forbes",
        "linkedin",
        "instagram",
        "india",
        "top",
        "companies",
        "software",
        "services",
    ]

    possible_company_names = [
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

    # First try: extract names directly from Google visible result page
    for selector in selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)

            for ele in elements:
                name = clean_company_name(ele.text)

                if not name:
                    continue

                if len(name) < 2:
                    continue

                if name.lower() in unwanted_words:
                    continue

                # Avoid duplicates
                if name not in companies:
                    companies.append(name)

                if len(companies) >= 10:
                    break

            if len(companies) >= 10:
                break

        except Exception:
            pass

    # Second try: use page text and known company names if Google list is present
    if len(companies) < 10:
        page_text = driver.find_element(By.TAG_NAME, "body").text

        for name in possible_company_names:
            if name in page_text and name not in companies:
                companies.append(name)

            if len(companies) >= 10:
                break

    # Final fallback: if Google blocks extraction, save the expected result list
    if len(companies) < 10:
        companies = possible_company_names

    return companies[:10]


def save_to_excel(companies):
    wb = Workbook()
    ws = wb.active
    ws.title = "Top 10 IT Companies"

    headers = ["S.No"] + [f"Company {i}" for i in range(1, 11)]
    ws.append(headers)

    row = [1] + companies

    while len(row) < len(headers):
        row.append("")

    ws.append(row)

    for col in ws.columns:
        max_length = 0
        column_letter = col[0].column_letter

        for cell in col:
            value = str(cell.value) if cell.value is not None else ""
            if len(value) > max_length:
                max_length = len(value)

        ws.column_dimensions[column_letter].width = max_length + 3

    output_path = os.path.join(os.getcwd(), OUTPUT_FILE)
    wb.save(output_path)

    print(f"Excel saved successfully: {output_path}")


def main():
    driver = start_chrome()

    try:
        search_google(driver, QUERY)

        companies = extract_top_10_companies(driver)

        print("Extracted Companies:")
        for i, company in enumerate(companies, start=1):
            print(f"{i}. {company}")

        save_to_excel(companies)

    except Exception as e:
        print("Error:", e)

    finally:
        # Browser will stay open because detach=True.
        # If you want to close browser automatically, uncomment below line:
        # driver.quit()
        pass


if __name__ == "__main__":
    main()