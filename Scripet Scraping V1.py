import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Read the Excel file containing the URLs
excel_file = "Pages.xlsx"
urls_df = pd.read_excel(excel_file)

# Get the URLs as a list
urls = urls_df['URL'].tolist()

# Launch the browser
driver = webdriver.Chrome()  # Or specify the path to your WebDriver executable

# Initialize lists to store data
data = {
    "Company Name": [],
    "Company Address": [],
    "Category": [],
    "Description": [],
    "Website": [],
    "Phone": [],
    "Keywords": [],
    "About Us": [],
    "WhatsApp": [],
    "Location": []
}

# Iterate over each URL
for url in urls:
    try:
        # Navigate to the webpage
        driver.get(url)

        # Wait for the page to load
        wait = WebDriverWait(driver, 20)  # Increased timeout duration to 20 seconds

        # Find all company cards
        company_cards = wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, ".item-details")))

        # Iterate over company cards and extract information
        for card in company_cards:
            # Extract information from each card
            try:
                company_name = card.find_element(By.CSS_SELECTOR, ".item-title").text.strip()
            except NoSuchElementException:
                company_name = "N/A"

            try:
                company_address = card.find_element(By.CSS_SELECTOR, ".address-text").text.strip()
            except NoSuchElementException:
                company_address = "N/A"

            try:
                category = card.find_element(By.CSS_SELECTOR, ".category").text.strip()
            except NoSuchElementException:
                category = "N/A"

            try:
                description = card.find_element(By.CSS_SELECTOR, ".item-aboutUs a").text.strip()
            except NoSuchElementException:
                description = "N/A"

            try:
                website_elem = card.find_element(By.CSS_SELECTOR, ".website")
                website = website_elem.get_attribute('href') if website_elem else "N/A"
            except NoSuchElementException:
                website = "N/A"

     #       try:
     #          phone_elem = card.find_element(By.CSS_SELECTOR, ".rtl .fa-location-arrow, .rtl i.fa-phone")
     #           phone_elem.click()  # Click on the element to display the phone numbers
     #           wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".internal-breadcrumb, .popover-content")))
     #           time.sleep(.5)  # Wait for 5 seconds after phone numbers are displayed
     #           phone_elem = card.find_element(By.CSS_SELECTOR, ".rtl .fa-location-arrow, .rtl i.fa-phone")
     #           phone_elem.click()  # Click on the element to display the phone numbers
     #           phone = driver.find_element(By.CSS_SELECTOR, ".internal-breadcrumb, .popover-phones a").text.strip()
     #       except NoSuchElementException:
     #           phone = "N/A"

            try:
                phone_elem = card.find_element(By.CSS_SELECTOR, ".rtl .fa-location-arrow, .rtl i.fa-phone")
                phone_elem.click()  # Click on the element to display the phone numbers
                wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".internal-breadcrumb, .popover-content")))
                time.sleep(.5)  # Wait for 0.5 seconds after phone numbers are displayed

                # Find all phone number elements
                phone_elems = driver.find_elements(By.CSS_SELECTOR, ".internal-breadcrumb, .popover-phones a")
                
                # Extract phone numbers and store them in a list
                phone_numbers = [phone_elem.text.strip() for phone_elem in phone_elems]

                # Concatenate phone numbers with special character between them
                phone = " & ".join(phone_numbers)
                
            except NoSuchElementException:
                phone = "N/A"



            try:
                keywords_elem = card.find_element(By.CSS_SELECTOR, ".two-words")
                keywords = keywords_elem.text.strip() if keywords_elem else "N/A"
            except NoSuchElementException:
                keywords = "N/A"

            try:
                about_us_elem = card.find_element(By.CSS_SELECTOR, ".item-aboutUs")
                about_us = about_us_elem.text.strip() if about_us_elem else "N/A"
            except NoSuchElementException:
                about_us = "N/A"

            try:
                whatsapp_elem = card.find_element(By.CSS_SELECTOR, ".whatsAppLink")
                whatsapp = whatsapp_elem.get_attribute('href') if whatsapp_elem else "N/A"
            except NoSuchElementException:
                whatsapp = "N/A"

            try:
                location_elem = card.find_element(By.CSS_SELECTOR, ".showMapSearch")
                location = location_elem.get_attribute('href') if location_elem else "N/A"
            except NoSuchElementException:
                location = "N/A"

            # Append data to lists
            data["Company Name"].append(company_name)
            data["Company Address"].append(company_address)
            data["Category"].append(category)
            data["Description"].append(description)
            data["Website"].append(website)
            data["Phone"].append(phone)
            data["Keywords"].append(keywords)
            data["About Us"].append(about_us)
            data["WhatsApp"].append(whatsapp)
            data["Location"].append(location)
    except Exception as e:
        print(f"Error occurred while processing URL {url}: {str(e)}")

# Create a DataFrame from the data
df = pd.DataFrame(data)

# Export DataFrame to Excel
df.to_excel("مصاعد.xlsx", index=False)

# Close the browser
driver.quit()


