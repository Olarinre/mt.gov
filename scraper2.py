import time
import pandas as pd
from selenium.webdriver import Remote, ChromeOptions
from selenium.webdriver.chromium.remote_connection import ChromiumRemoteConnection
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

AUTH = 'brd-customer-hl_fc61121b-zone-scraping_browser1:x6w2qsffsm5h'
SBR_WEBDRIVER = f'https://{AUTH}@brd.superproxy.io:9515'

def main():
    print("Connecting to Browser API...")
    sbr_connection = ChromiumRemoteConnection(SBR_WEBDRIVER, "goog", "chrome")

    with Remote(sbr_connection, options=ChromeOptions()) as driver:
        driver.get("https://mt.gov.sa/e-services/forms/licensed-activities-inquiry")

        wait = WebDriverWait(driver, 30)

        # Select "Special Accommodation Facilities"
        select_from_ngselect(driver, wait, "activity", "Special Accommodation Facilities")

        # Select "JEDDAH"
        select_from_ngselect(driver, wait, "city", "JEDDAH")

        handle_cookie_popup(driver)

        # Click Search
        search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Search']]")))
        search_btn.click()
        print("Search clicked")

        # Wait for results
        time.sleep(5)

        all_records = []
        page = 1

        while True:
            print(f"Processing page {page}...")

            # Wait for at least one record
            # Wait until at least one card is present
            # Wait until at least one card is loaded
            wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, "//div[contains(@class,'col-start') and contains(@class,'flex-grow-1')]")
))

            # Get all card elements on the page
            records = driver.find_elements(
            By.XPATH, "//div[contains(@class,'col-start') and contains(@class,'flex-grow-1')]"
)



            for rec in records:
                try:
                    data = extract_record(rec)
                    all_records.append(data)
                except Exception as e:
                    print("Error extracting record:", e)

            print(f"Extracted {len(records)} records from page {page}")

            # Next button
            try:
                next_btn = driver.find_element(By.XPATH, "//*[@id='inner-content']/div/div/div[2]/app-licensed-inquiry-results/app-special-accommodation-facilities-results/div/div[3]/app-mt-paginator/nav/ul/li[3]/a")
                if not next_btn.is_enabled():
                    break
                driver.execute_script("arguments[0].click();", next_btn)
                time.sleep(3)
                page += 1
            except:
                break

        # Save
        if all_records:
            df = pd.DataFrame(all_records)
            df.to_excel("tourism_licenses_jeddah.xlsx", index=False)
            print(f"Saved {len(all_records)} records to tourism_licenses_jeddah.xlsx")
        else:
            print("No records found")


def select_from_ngselect(driver, wait, formcontrol, value):
    for attempt in range(3):
        try:
            ng_input = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//ng-select[@formcontrolname='{formcontrol}']//div[@role='combobox']/input")
                )
            )
            ng_input.click()
            ng_input.clear()
            ng_input.send_keys(value)
            time.sleep(1)

            option = wait.until(
                EC.element_to_be_clickable((By.XPATH, f"//div[@role='option' and contains(., '{value}')]"))
            )
            driver.execute_script("arguments[0].click();", option)
            print(f"Selected {value} for {formcontrol}")
            return
        except Exception as e:
            print(f"Retry {attempt+1}/3 for {formcontrol}: {e}")
            time.sleep(1)
    raise Exception(f"Failed to select {value} for {formcontrol}")


def handle_cookie_popup(driver):
    try:
        cookie_btn = driver.find_element(By.XPATH, "//app-mt-cookie//button")
        driver.execute_script("arguments[0].click();", cookie_btn)
        print("Closed cookie popup")
        time.sleep(1)
    except:
        print("No cookie popup found")




def extract_record(record):
    """Extract fields from a record card"""
    def safe_find(xpath):
        try:
            return record.find_element(By.XPATH, xpath).text.strip()
        except:
            return "Not specified"
        
    def safe_text(xpath):
        try:
            return record.find_element(By.XPATH, xpath).text.strip()
        except:
            return "Not specified"

    def safe_attr(xpath, attr):
        try:
            return record.find_element(By.XPATH, xpath).get_attribute(attr).strip()
        except:
            return "Not specified"
    
    
    return {
        "Company Name": safe_find(".//h2 | .//h3 | .//h4 | .//strong"),
        "Classification": safe_find(".//*[contains(text(),'Classification')]/following-sibling::*"),
        "License Status": safe_find(".//*[contains(text(),'License Status')]/following-sibling::*"),
        "Facility Type": safe_find(".//*[contains(text(),'Facility Type')]/following-sibling::*"),
        "License Number": safe_text(".//label[contains(.,'License Number')]/following-sibling::strong"),
        "Website": safe_find(".//*[contains(text(),'Website')]/following-sibling::*"),
        "Facility Location": safe_attr(".//label[contains(.,'Facility Location')]/following-sibling::a", "href"),
    }


if __name__ == "__main__":
    main()
