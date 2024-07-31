from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time
# import re

# กำหนดพาธของ chromedriver
chrome_driver_path = r"C:\Users\puriwat.s.COMETS\Desktop\Pongpang_2024\chromedriver-win64\chromedriver.exe"
# เรียกใช้ Service และ ChromeOptions
service = Service(executable_path=chrome_driver_path)
options = webdriver.ChromeOptions()


# Update chorme driver to include headless mode
# options.add_argument("--headless")

# เปิดเบราว์เซอร์
driver = webdriver.Chrome(service=service, options=options)

# เปิด URL ของเว็บไซต์
url = "https://pertento.fda.moph.go.th/FDA_SEARCH_CENTER/PRODUCT/FRM_SEARCH_CMT.aspx"
driver.get(url)

# ค้นหาและเติมคำค้นหา
#search_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$txt_trade")))


# สร้าง list ของคำค้นหา
 # "SIVANNA" "

search_keywords = [ 
      "ASHLEY"
      
   
      

]



data_list = []

for keyword in search_keywords:

    # ค้นหาชื่อผู้ประกอบการ

    #search_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$txt_oper")))

    ## In case of slow internet speed (e.g. <= 4 MB/s) or other very little retardation which timeout error can be found
    #search_input = WebDriverWait(driver, 1000000).until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$txt_oper")))

    # ค้นหาจากชื่อแบรนด์
    search_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$txt_trade")))
    ## In case of slow internet speed (e.g. <= 4 MB/s) or other very little retardation which timeout error can be found

    # search_input = WebDriverWait(driver, 100000).until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$txt_trade")))





    driver.execute_script("arguments[0].scrollIntoView();", search_input)
    search_input.click()
    # ล้างช่องค้นหา
    search_input.clear()
    
    search_input.click()
    # ใส่คำค้นหาใหม่
    search_input.send_keys(keyword)

    # คลิกปุ่มค้นหา
    search_button = driver.find_element(By.ID, "ContentPlaceHolder1_btn_sea_cmt")
    search_button.click()
    time.sleep(20)
    span = driver.find_element(By.ID, "ContentPlaceHolder1_lb_num").text
    while len(data_list) < int(span):
        try:
            # รอให้หน้าเว็บโหลดเสร็จสมบูรณ์
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//table[@id='ContentPlaceHolder1_RadGrid1_ctl00']/tbody/tr")))
            
             # In case of slow internet speed (e.g. <= 4 MB/s) or other very little retardation which timeout error can be found 
            #WebDriverWait(driver, 200000).until(EC.presence_of_element_located((By.XPATH, "//table[@id='ContentPlaceHolder1_RadGrid1_ctl00']/tbody/tr")))


            # หาตารางทั้งหมด
            tables = driver.find_elements(By.XPATH, "//table[@id='ContentPlaceHolder1_RadGrid1_ctl00']/tbody/tr")
            span_element = len(tables)
            
            # วนลูปในตารางและคลิกทุกลิงค์ "ดูข้อมูล"
            for n in range(span_element):
                xpath = f"//*[@id='ContentPlaceHolder1_RadGrid1_ctl00__{n}']/td[20]/a"
                # คลิกลิงค์ "ดูข้อมูล"
                link = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath)))

                #In case of slow internet speed (e.g. <= 4 MB/s) or other very little retardation which timeout error can be found 
                #link = WebDriverWait(driver, 200000).until(EC.element_to_be_clickable((By.XPATH, xpath)))

                link.click()

                # รอให้มีการเปิดแท็บใหม่
                WebDriverWait(driver, 30).until(lambda driver: len(driver.window_handles) > 1)

                # In case of slow internet speed (e.g. <= 4 MB/s) or other very little retardation which timeout error can be found
                #WebDriverWait(driver, 300000).until(lambda driver: len(driver.window_handles) > 1)

                
                # เปลี่ยนไปที่แท็บใหม่
                driver.switch_to.window(driver.window_handles[1])
                
                try :
                    # ดึงข้อมูลจากแต่ละ element ในหน้าใหม่
                    status = driver.find_element(By.ID, "ContentPlaceHolder1_lb_status").text
                    status_location = driver.find_element(By.ID, "ContentPlaceHolder1_lb_status_lct").text
                    registration_number = driver.find_element(By.ID, "ContentPlaceHolder1_lb_no_regnos").text
                    registration_type = driver.find_element(By.ID, "ContentPlaceHolder1_lb_type").text
                    ประเภทการจดแจ้ง = driver.find_element(By.ID, "ContentPlaceHolder1_lb_format_regnos").text
                    รูปแบบการจดแจ้ง = driver.find_element(By.ID, "ContentPlaceHolder1_lb_type").text
                    ชื่อการค้า = driver.find_element(By.ID, "ContentPlaceHolder1_lb_trade_Tpop").text
                    ชื่อเครื่องสำอาง = driver.find_element(By.ID, "ContentPlaceHolder1_lb_cosnm_Tpop").text
                    วันที่อนุญาต = driver.find_element(By.ID, "ContentPlaceHolder1_lb_appdate").text
                    วันที่หมดอายุ = driver.find_element(By.ID, "ContentPlaceHolder1_lb_expdate").text
                    รูปแบบการใช้ผลิตภัณฑ์ = driver.find_element(By.ID, "ContentPlaceHolder1_lb_mode").text
                    บริเวณที่ใช้ผลิตภัณฑ์ = driver.find_element(By.ID, "ContentPlaceHolder1_lb_applicability_name").text
                    วัตถุประสงค์ในการใช้งานผลิตภัณฑ์ = driver.find_element(By.ID, "ContentPlaceHolder1_lb_applicability_name").text
                    เงื่อนไขการใช้ผลิตภัณฑ์ = driver.find_element(By.ID, "ContentPlaceHolder1_lb_condition").text
                    ชื่อผู้ประกอบการ = driver.find_element(By.ID, "ContentPlaceHolder1_lb_usernm_pop").text
                    ที่ตั้งผู้ผลิต = driver.find_element(By.ID, "ContentPlaceHolder1_lb_locat_pop").text
                    ชื่อและที่อยู่ผู้ผลิตต่างประเทศ = driver.find_element(By.ID, "ContentPlaceHolder1_lb_fac_pop").text
                    เลขอ้างอิงสำหรับ = driver.find_element(By.ID, "ContentPlaceHolder1_lb_NO_pop").text
                    SKU = driver.find_element(By.ID, "ContentPlaceHolder1_lb_fileattach").text

                # เพิ่มข้อมูลลงใน new_data
                    new_data = [status, status_location, registration_number, registration_type,ประเภทการจดแจ้ง,รูปแบบการจดแจ้ง,ชื่อการค้า,ชื่อเครื่องสำอาง,วันที่อนุญาต,วันที่หมดอายุ,รูปแบบการใช้ผลิตภัณฑ์,บริเวณที่ใช้ผลิตภัณฑ์,วัตถุประสงค์ในการใช้งานผลิตภัณฑ์,
                            เงื่อนไขการใช้ผลิตภัณฑ์,ชื่อผู้ประกอบการ,ที่ตั้งผู้ผลิต,ชื่อและที่อยู่ผู้ผลิตต่างประเทศ,เลขอ้างอิงสำหรับ,SKU]
                    
                # นำ new_data ไปเพิ่มลงใน data_list
                    data_list.append(new_data)
                
                # ปิดแท็บใหม่
                    driver.close()        
                except NoSuchElementException:
                    driver.close() 
                    
                # เปลี่ยนกลับไปที่แท็บเดิม
                driver.switch_to.window(driver.window_handles[0])
                time.sleep(5)
                print(n)
            print("จบหน้า")  

            try:
                # รอให้ปุ่ม Next โหลดและกลายเป็นสถานะที่เป็นไปได้
                next_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "rgPageNext")))

                # In case of slow internet speed (e.g. <= 4 MB/s) or other very little retardation which timeout error can be found
                #next_button = WebDriverWait(driver, 100000).until(EC.element_to_be_clickable((By.CLASS_NAME, "rgPageNext")))

                # Scroll into view
                driver.execute_script("arguments[0].scrollIntoView();", next_button)

                # Click using JavaScript if regular click doesn't work
                driver.execute_script("arguments[0].click();", next_button)

                # Wait for new data to load
                WebDriverWait(driver, 20).until(EC.staleness_of(tables[-1]))  # Wait for last table to become stale

                # In case of slow internet speed (e.g. <= 4 MB/s) or other very little retardation which timeout error can be found
                #WebDriverWait(driver, 200000).until(EC.staleness_of(tables[-1]))  # Wait for last table to become stale


            except TimeoutException:
                print("ไม่พบปุ่ม next")
                break
        except TimeoutException:
            print("Timed out waiting for page to load")
            break

# ปิดเบราว์เซอร์
driver.quit()

# แสดงผลลัพธ์ในรูปแบบของ DataFrame
df = pd.DataFrame(data_list, columns=["Status", "Status Location", "Registration Number", "Registration Type","ประเภทการจดแจ้ง","รูปแบบการจดแจ้ง","ชื่อการค้า","ชื่อเครื่องสำอาง","วันที่อนุญาต","วันที่หมดอายุ","รูปแบบการใช้ผลิตภัณฑ์","บริเวณที่ใช้ผลิตภัณฑ์","วัตถุประสงค์ในการใช้งานผลิตภัณฑ์",
                            "เงื่อนไขการใช้ผลิตภัณฑ์","ชื่อผู้ประกอบการ","ที่ตั้งผู้ผลิต","ชื่อและที่อยู่ผู้ผลิตต่างประเทศ","เลขอ้างอิงสำหรับ","SKU"])



# เซฟข้อมูลเป็น Dataframe และ export ออกไปเป็น Excel file (อย่าลืมแก้ไขชื่อไฟล์ก่อน Automate)
print(df)
df.to_excel('scraped_data_FDA_030724_ASHLEY.xlsx', index=False)
