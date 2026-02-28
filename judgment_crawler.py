import time
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

def crawl_via_selenium(keyword):
    """
    Example of using Selenium to crawl judgment.judicial.gov.tw.
    This handles the ASP.NET PostBack and dynamic content.
    """
    print(f"Starting crawl for keyword: {keyword}")
    
    # Setup Chrome options
    chrome_options = Options()
    # chrome_options.add_argument("--headless")  # Run defined headless if needed
    
    # Initialize driver (ensure you have chromedriver installed)
    driver = webdriver.Chrome(options=chrome_options)
    
    try:
        # 1. Navigate to the search page
        url = "https://judgment.judicial.gov.tw/FJUD/default.aspx"
        driver.get(url)
        
        # 2. Find the search box and input keyword
        # Note: ID might change, so looking by name or surrounding label is safer, 
        # but 'txtKW' is a common ID for keyword inputs in ASP.NET apps.
        # Alternatively, use XPath to find input near "檢索字詞"
        try:
            search_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "txtKW"))
            )
            search_input.clear()
            search_input.send_keys(keyword)
        except:
            print("Could not find input by ID 'txtKW', trying generic XPath...")
            search_input = driver.find_element(By.XPATH, "//input[@type='text']")
            search_input.send_keys(keyword)

        # 3. Click the Search button
        # Verified ID from page source: 'btnSimpleQry'
        try:
            search_btn = driver.find_element(By.ID, "btnSimpleQry")
            search_btn.click()
        except:
            print("Could not find button by ID 'btnSimpleQry', submitting form...")
            search_input.submit()

        # 4. Wait for results to load
        # The result page typically has a list of cases in an iframe or main table
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "table"))
        )
        
        print("Search results loaded. Extracting links...")
        
        # 5. Extract links to judgments
        # This is strictly an example selector - inspect the page for exact classes
        links = driver.find_elements(By.XPATH, "//a[contains(@href, 'data.aspx')]")
        
        for link in links[:5]:  # Print first 5
            print(f"Found Case: {link.text}")
            print(f"Link: {link.get_attribute('href')}")
            
            # Implementation to download PDF or content would go here
            # Usually involves clicking the link, switching to the new window/frame, 
            # and finding the 'PDF' button or extracting text.

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        driver.quit()

def check_open_data_api():
    """
    Check the Judicial Yuan Open Data Platform.
    The user is encouraged to use https://opendata.judicial.gov.tw/ for bulk data.
    """
    print("\n--- API Information ---")
    print("The Judicial Yuan 'Open Data Platform' (https://opendata.judicial.gov.tw/)")
    print("offers datasets and APIs which are more stable than crawling.")
    print("Check for 'Judgment' (裁判書) datasets there.")

if __name__ == "__main__":
    # You need to install selenium: pip install selenium webdriver-manager
    print("Select mode:")
    print("1. Selenium Crawler (Demo)")
    print("2. API Info")
    
    choice = input("Enter choice (1/2): ")
    if choice == "1":
        kw = input("Enter search keyword (default: test): ") or "test"
        crawl_via_selenium(kw)
    else:
        check_open_data_api()
