import re
import time
import random
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# User-Agent pool for randomization
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/89.0",
    "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:89.0) Gecko/20100101 Firefox/89.0",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 14_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15E148 Safari/604.1",
    "Mozilla/5.0 (Linux; Android 10; SM-A505FN) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Mobile Safari/537.36"
]

# Helper function to randomize delays
def random_sleep(min_time=1, max_time=3):
    time.sleep(random.uniform(min_time, max_time))

# Helper function to extract emails and phone numbers from a page
def extract_contact_info(driver):
    emails = []
    phones = []
    try:
        page_text = driver.find_element(By.TAG_NAME, "body").text

        # Email regex
        email_pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
        
        # Refined phone number regex
        phone_pattern = r"""
        \b                 # Start of word boundary
        (?:\+?\d{1,3}[\s.-]?)?  # Optional country code, e.g., +1, +44
        (?:\(?\d{2,4}\)?[\s.-]?)  # Area code: (123), 123-
        (?:\d{3,4}[\s.-]?\d{3,4})  # Central office and station codes
        (?:\s?(?:x|ext|extension)\s?\d{1,5})?  # Optional extension
        \b                 # End of word boundary
        """

        # Use re.findall to extract emails and phones
        emails = re.findall(email_pattern, page_text)
        phones = re.findall(phone_pattern, page_text, re.VERBOSE)  # Enable VERBOSE mode for phone regex

    except Exception as e:
        print(f"Error extracting contact info: {e}")

    # Remove duplicates using set() and return the results
    return list(set(emails)), list(set(phones))

# Manual CAPTCHA handling
def handle_captcha(driver):
    try:
        # Check for Google's reCAPTCHA iframe
        captcha_frame = driver.find_element(By.XPATH, "//iframe[contains(@src, 'recaptcha')]")
        if captcha_frame:
            print("CAPTCHA detected. Please solve it manually in the browser.")
            while True:
                time.sleep(5)
                try:
                    # If the CAPTCHA iframe disappears, assume it's solved
                    driver.find_element(By.XPATH, "//iframe[contains(@src, 'recaptcha')]")
                except:
                    print("CAPTCHA solved. Resuming scraping...")
                    break
    except Exception:
        # No CAPTCHA detected
        print("No CAPTCHA detected.")

# Function to handle tab switching
def switch_to_new_tab(driver):
    current_tab = driver.current_window_handle
    for tab in driver.window_handles:
        if tab != current_tab:
            driver.switch_to.window(tab)
            print(f"Switched to new tab: {driver.current_url}")
            return
    print("No new tab detected.")

# Function to close all tabs except the first one
def close_extra_tabs(driver):
    original_tab = driver.window_handles[0]
    for tab in driver.window_handles[1:]:
        driver.switch_to.window(tab)
        driver.close()
    driver.switch_to.window(original_tab)

# Function to check if a URL is a social media link
def is_social_media_link(url):
    social_domains = ["instagram.com", "linkedin.com", "facebook.com", "twitter.com"]
    return any(domain in url for domain in social_domains)

# Configure undetected-chromedriver options
selected_user_agent = random.choice(USER_AGENTS)
options = uc.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument(f"user-agent={selected_user_agent}")

# Start undetected-chromedriver
driver = uc.Chrome(options=options)

# Load the Excel file containing company names
input_file_path = '621.xlsx'  # Your input Excel file
output_file_path = 'output_contacts.xlsx'  # Output Excel file

data = pd.read_excel(input_file_path)
company_names = data[data.columns[0]].dropna()

# Output data list
output_data = []

for idx, company_name in enumerate(company_names, start=1):
    print(f"Processing {idx}/{len(company_names)}: {company_name}")
    try:
        driver.get("https://www.google.com")
        random_sleep()

        ActionChains(driver).move_by_offset(5, 5).perform()
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "q"))
        )
        search_box.send_keys(f"{company_name} website")
        search_box.submit()
        random_sleep()

        handle_captcha(driver)

        search_results = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "h3"))
        )

        selected_result = None
        for i in range(min(5, len(search_results))):
            parent_element = search_results[i].find_element(By.XPATH, "ancestor::a")
            url = parent_element.get_attribute("href")
            if not is_social_media_link(url):
                selected_result = parent_element
                break

        if selected_result:
            ActionChains(driver).move_to_element(selected_result).click().perform()
            random_sleep(2, 5)

            # Handle potential new tab
            switch_to_new_tab(driver)

            try:
                cookie_banner = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'accept') or contains(text(), 'Accept') button[contains(text(), 'Close') or contains(text(), 'close')]")
                ))
                cookie_banner.click()
                random_sleep()
            except:
                pass

            possible_contact_links = driver.find_elements(By.PARTIAL_LINK_TEXT, "Contact") + \
                                     driver.find_elements(By.PARTIAL_LINK_TEXT, "Kontakt") + \
                                     driver.find_elements(By.PARTIAL_LINK_TEXT, "Contacto") + \
                                     driver.find_elements(By.PARTIAL_LINK_TEXT, "联系我们") + \
                                     driver.find_elements(By.PARTIAL_LINK_TEXT, "Contatti") + \
                                     driver.find_elements(By.PARTIAL_LINK_TEXT, "اتصال")

            if possible_contact_links:
                ActionChains(driver).move_to_element(possible_contact_links[0]).click().perform()
                random_sleep(2, 5)

                # Handle potential new tab
                switch_to_new_tab(driver)

            emails, phones = extract_contact_info(driver)

            output_data.append({
                "Company Name": company_name,
                "Company Website": driver.current_url,
                "Company E-mail": ", ".join(emails) if emails else "Not Found",
                "Company Telephone": ", ".join(phones) if phones else "Not Found"
            })
        else:
            print(f"No suitable result found for {company_name}")
            output_data.append({
                "Company Name": company_name,
                "Company Website": "Not Found",
                "Company E-mail": "Not Found",
                "Company Telephone": "Not Found"
            })

    except Exception as e:
        print(f"Error processing {company_name}: {e}")
        output_data.append({
            "Company Name": company_name,
            "Company Website": "Error",
            "Company E-mail": "Error",
            "Company Telephone": "Error"
        })

    # Close extra tabs after processing
    close_extra_tabs(driver)

output_df = pd.DataFrame(output_data)
output_df.to_excel(output_file_path, index=False)
print(f"Scraping completed. Data saved to {output_file_path}")

driver.quit()
