
#Please Use this code as main code for download data from watson's shop
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException, TimeoutException, 
    ElementClickInterceptedException, StaleElementReferenceException
)
import pandas as pd
from openpyxl import Workbook
from datetime import datetime
import logging
from bs4 import BeautifulSoup

# Define driver path
chrome_driver_path = r"C:\Users\puriwat.s.COMETS\Desktop\Pongpang_2024\chromedriver-win64\chromedriver.exe"

# Set up the service and options for Chrome
service = Service(executable_path=chrome_driver_path)
options = webdriver.ChromeOptions()
# options.add_argument('--headless')  # Uncomment to run in headless mode
options.add_argument('--disable-gpu')  # Necessary for headless mode

# Open the browser
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()

# Target URL : THAI
base_url = "https://www.watsons.co.th/th/c/030000"

# Target URL : ENGLISH
# base_url = "https://www.watsons.co.th/en/c/030000"

def close_overlay():
    try:
        overlay_close_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[aria-label="close"]'))
        )
        overlay_close_button.click()
    except TimeoutException:
        pass  # No overlay found

    try:
        privacy_policy_overlay = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '.onetrust-pc-dark-filter'))
        )
        driver.execute_script("arguments[0].style.visibility='hidden'", privacy_policy_overlay)
    except TimeoutException:
        pass  # No privacy policy overlay found

def get_product_links():
    retries = 3
    while retries > 0:
        try:
            product_links = WebDriverWait(driver, 60).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'h2.productName a'))
            )
            return [link.get_attribute('href') for link in product_links]
        except StaleElementReferenceException:
            retries -= 1
            logging.warning("StaleElementReferenceException encountered. Retrying...")
            time.sleep(2)
        except TimeoutException:
            retries -= 1
            logging.error("Timeout while waiting for product links. Retrying...")
            if retries == 0:
                return []
            driver.refresh()
            time.sleep(5)
    return []

def extract_element_text(selector, by=By.CSS_SELECTOR, attribute='text'):
    try:
        element = driver.find_element(by, selector)
        return element.get_attribute(attribute) if attribute != 'text' else element.text
    except NoSuchElementException:
        return None

def extract_category_data():
    try:
        categories = driver.find_elements(By.CSS_SELECTOR, 'span[itemprop="name"].ng-star-inserted')
        category = categories[1].text if len(categories) > 1 else None
        sub_category_tier1 = categories[2].text if len(categories) > 2 else None
        sub_category_tier2 = categories[3].text if len(categories) > 3 else None
        return category, sub_category_tier1, sub_category_tier2
    except Exception as e:
        logging.error(f"Error extracting category data: {e}")
        return "เครื่องสำอาง และน้ำหอม", None, None

def extract_product_image(product_name):
    retries = 3
    while retries > 0:
        try:
            images = driver.find_elements(By.CSS_SELECTOR, 'e2-media img')
            for img in images:
                if product_name in img.get_attribute('alt'):
                    return img.get_attribute('src')
            # Alternative method if primary selector fails
            images_alt = driver.find_elements(By.CSS_SELECTOR, 'img.product-main-image')
            for img in images_alt:
                if product_name in img.get_attribute('alt'):
                    return img.get_attribute('src')
            return None
        except (NoSuchElementException, StaleElementReferenceException) as e:
            retries -= 1
            logging.warning(f"Error extracting product image: {e}. Retrying {retries} more times.")
            time.sleep(2)
    return None

def extract_product_data():
    category, sub_category_tier1, sub_category_tier2 = extract_category_data()
    product_name = extract_element_text('.product-name')
    
    # Corrected method to extract Number_of_Product_Sold
    try:
        social_proof = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.social-proof.ng-star-inserted'))
        )
        number_of_product_sold_text = social_proof.text.strip() if social_proof else None
    except (NoSuchElementException, TimeoutException):
        number_of_product_sold_text = None

    data = {}
    try:
        data['Brand'] = extract_element_text('h2.product-brand a')
    except NoSuchElementException:
        data['Brand'] = None

    data['Product_Name'] = product_name
    data['Number_of_Product_Sold'] = number_of_product_sold_text
    
    try:
        data['Current_Price'] = extract_element_text('span.price')
    except NoSuchElementException:
        data['Current_Price'] = None
    
    try:
        data['Original_Price'] = extract_element_text('span.retail-price')
    except NoSuchElementException:
        data['Original_Price'] = None
    
    try:
        data['Product_Desc'] = extract_element_text('.content')
    except NoSuchElementException:
        data['Product_Desc'] = None
    
    try:
        data['ประเทศผู้ผลิต'] = extract_element_text('//h4[text()="ผลิตที่"]/following-sibling::p', By.XPATH)
    except NoSuchElementException:
        data['ประเทศผู้ผลิต'] = None

    try:
        data['เลขที่ใบจดแจ้ง'] = driver.find_element(By.XPATH, '//h4[text()=" เลขที่ใบจดแจ้ง "]/following-sibling::p').text
    except NoSuchElementException:
        data['เลขที่ใบจดแจ้ง'] = None

    try:
        data['วิธีการใช้งาน'] = extract_element_text('//h4[text()="วิธีการใช้งาน"]/following-sibling::p', By.XPATH)
    except NoSuchElementException:
        data['วิธีการใช้งาน'] = None

    try:
        data['ส่วนประกอบ'] = driver.find_element(By.XPATH, '//h4[text()=" ส่วนประกอบ "]/following-sibling::p').text
    except NoSuchElementException:
        data['ส่วนประกอบ'] = None

    try:
        data['ประโยชน์'] = extract_element_text('//td[text()="ประโยชน์"]/following-sibling::td', By.XPATH)
    except NoSuchElementException:
        data['ประโยชน์'] = None

    try:
        data['ประเภทของผิว'] = extract_element_text('//td[text()="ประเภทของผิว"]/following-sibling::td', By.XPATH)
    except NoSuchElementException:
        data['ประเภทของผิว'] = None

    try:
        data['ขนาด'] = extract_element_text('//td[text()="ขนาด"]/following-sibling::td', By.XPATH)
    except NoSuchElementException:
        data['ขนาด'] = None

    try:
        data['รหัสสินค้า'] = driver.find_element(By.XPATH, '//div[@class="info"]/div[@class="name" and text()="รหัสสินค้า"]/following-sibling::div[@class="value"]').text
    except NoSuchElementException:
        data['รหัสสินค้า'] = None

    try:
        data['ความกว้าง'] = extract_element_text('//h4[text()="ความกว้าง"]/following-sibling::p', By.XPATH)
    except NoSuchElementException:
        data['ความกว้าง'] = None

    try:
        data['ความสูง'] = extract_element_text('//h4[text()="ความสูง"]/following-sibling::p', By.XPATH)
    except NoSuchElementException:
        data['ความสูง'] = None

    try:
        data['ความลึก'] = extract_element_text('//h4[text()="ความลึก"]/following-sibling::p', By.XPATH)
    except NoSuchElementException:
        data['ความลึก'] = None

    try:
        data['Latest_Promotion'] = ' '.join([promo.text for promo in driver.find_elements(By.CSS_SELECTOR, '.promotion-group .promotion .description')])
    except NoSuchElementException:
        data['Latest_Promotion'] = None

    try:
        data['คำเตือน'] = extract_element_text('//h4[text()="คำเตือน"]/following-sibling::p', By.XPATH)
    except NoSuchElementException:
        data['คำเตือน'] = None

    data['Product_Image'] = extract_product_image(product_name)

    try:
        data['Hot_Promotion'] = extract_element_text('.tabContainer .tab')
    except NoSuchElementException:
        data['Hot_Promotion'] = None

    data['Category'] = category
    data['Sub_Category_Tier1'] = sub_category_tier1
    data['Sub_Category_Tier2'] = sub_category_tier2

    current_datetime = datetime.now()
    data['Revision_Datetime'] = current_datetime.strftime('%Y-%m-%d %H:%M:%S')
    data['Revision_Date'] = current_datetime.strftime('%d-%b-%Y')

    if not data['Product_Image']:
        logging.warning(f"No image found for product: {data['Product_Name']}")

    if not data['Number_of_Product_Sold']:
        logging.warning(f"No sales data found for product: {data['Product_Name']}")

    return data

def save_to_excel(products_data, file_name):
    df = pd.DataFrame(products_data)
    df.to_excel(file_name, index=False)
    logging.info(f"Data saved to Excel successfully as {file_name}.")

def click_next_page():
    retries = 3
    while retries > 0:
        try:
            next_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'li.page-item a.page-link i.icon-arrow-right'))
            )
            driver.execute_script("arguments[0].click();", next_button)
            WebDriverWait(driver, 20).until(
                EC.staleness_of(next_button)
            )
            time.sleep(5)
            return True
        except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
            retries -= 1
            logging.error(f"Error clicking next page: {e}. Retrying {retries} more times.")
            driver.refresh()
            time.sleep(5)
    return False

def load_page_with_number(current_page):
    try:
        url_with_page = f"{base_url}?currentPage={current_page - 1}"
        driver.get(url_with_page)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'h2.productName a'))
        )
        return True
    except TimeoutException as e:
        logging.error(f"Timeout while loading page {current_page}: {e}")
        return False

# Part 5: Main Logic and Execution
try:
    user_input = input("Please specify file name: ")
    current_page = int(input("Enter the starting page number: "))
    last_page = int(input("Enter the last page number: "))

    total_pages = last_page - current_page + 1
    items_per_page = 32
    total_items = total_pages * items_per_page

    print(f"Total Pages to be downloaded: {total_pages}")
    print(f"Items to be downloaded: {total_items}")

    # Generate file name and log file name
    file_name = f"WATSONS_BestSeller_MakeupProduct_Update_070824_{user_input}_Page{current_page}to{last_page}.xlsx"
    log_file_name = f"scraping_log_070824_{user_input}_from_Page{current_page}to{last_page}.log"

    # Set up logging to file and console with dynamic log file name
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', handlers=[
        logging.FileHandler(log_file_name),
        logging.StreamHandler()
    ])

    all_products_data = []
    unique_product_names = set()

    driver.execute_script("window.open('');")
    product_tab = driver.window_handles[1]

    while current_page <= last_page:
        if not load_page_with_number(current_page):
            break
        
        close_overlay()
        product_links = get_product_links()
        if not product_links:
            logging.error(f"Failed to get product links on page {current_page}.")
            break
        logging.info(f"Number of products found on current page {current_page}: {len(product_links)}")

        for i in range(len(product_links)):
            retries = 3
            while retries > 0:
                try:
                    logging.info(f"Processing product {i + 1}/{len(product_links)}")

                    product_url = product_links[i]
                    driver.switch_to.window(product_tab)
                    driver.get(product_url)

                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'h2.product-brand a'))
                    )
                    product_data = extract_product_data()
                    if product_data['Product_Name'] not in unique_product_names:
                        all_products_data.append(product_data)
                        unique_product_names.add(product_data['Product_Name'])

                    driver.switch_to.window(driver.window_handles[0])
                    time.sleep(2)
                    break
                except StaleElementReferenceException as e:
                    retries -= 1
                    logging.warning(f"Stale element reference exception for product {i + 1}: {e}. Retrying {retries} more times.")
                    driver.refresh()
                    time.sleep(2)
                except Exception as e:
                    logging.error(f"Error processing product {i + 1}: {e}")
                    driver.switch_to.window(driver.window_handles[0])
                    time.sleep(2)
                    break

        if len(product_links) == items_per_page and current_page < last_page:
            if not click_next_page():
                break

        current_page += 1

    save_to_excel(all_products_data, file_name)

finally:
    driver.quit()

