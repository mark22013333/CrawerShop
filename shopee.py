from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

# 使用 undetected_chromedriver
import undetected_chromedriver as uc

import pyodbc
import urllib.parse
import random

def connect_to_mssql():
    conn_str = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=localhost;"
        "DATABASE=WCS;"
        "UID=root;"
        "PWD=root;"
    )
    try:
        conn = pyodbc.connect(conn_str)
        return conn
    except Exception as e:
        return None

def fetch_row_by_seller(connection, seller):
    cursor = connection.cursor()
    cursor.execute("SELECT * FROM WCST198 WHERE Auction = ?", (seller,))
    
    columns = [column[0] for column in cursor.description]
    
    row = cursor.fetchone()
    if row:
        row_data = dict(zip(columns, row))
        return row_data
    else:
        return None

def update_data(connection, auction, keyword, category, district, fans, fans_Add, fans_Last, commodity, commodity_Add, commodity_Last, point, point_Add, point_Last, companyName, last_DATE):
    cursor = connection.cursor()
    query = " UPDATE WCST198 SET Keyword = ?, Category = ?, District = ?, Fans_NOW = ?, Fans_ADD = ?, Fans_LAST = ?, Commodity_NOW = ?, Commodity_ADD = ?, Commodity_LAST = ?, Point_NOW = ?, Point_ADD = ?, Point_LAST = ?, CompanyName = ?, EXTRACT_DATE = CONVERT(varchar(256), GETDATE(), 120), Last_DATE = ? WHERE Auction = ?"
    cursor.execute(query, (keyword, category, district, fans, fans_Add, fans_Last, commodity, commodity_Add, commodity_Last, point, point_Add, point_Last, companyName, last_DATE, auction))
    connection.commit()

def insert_data(connection, auction, keyword, category, district, fans, commodity, point, companyName):
    cursor = connection.cursor()
    query = "insert into WCST198 (Auction, Keyword, Category, District, Fans_NOW, Fans_ADD, Fans_LAST, Commodity_NOW, Commodity_ADD, Commodity_LAST, Point_NOW, Point_ADD, Point_LAST, CompanyName, EXTRACT_DATE, First_DATE, Last_DATE) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CONVERT(varchar(256), GETDATE(), 120), CONVERT(varchar(256), GETDATE(), 120), CONVERT(varchar(256), GETDATE(), 120))"
    cursor.execute(query, (auction, keyword, category, district, fans, 0, fans, commodity, 0, commodity, point, 0, point, companyName))
    connection.commit();

def convert_wan_to_number(wan_string):
    """將以「萬」為單位的字符串轉換為數字"""
    if "萬" in wan_string:
        # 去掉「萬」並轉換為浮點數
        number = float(wan_string.replace("萬", "").strip())
        return int(number * 10000)  # 轉換為整數
    else:
        return int(wan_string)  # 如果沒有「萬」，直接轉換為整數

keyword = "現貨"
page = 0

connection = connect_to_mssql()
keyword = urllib.parse.quote(keyword)

options = uc.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
#options.add_experimental_option("excludeSwitches", ["enable-automation"])
#options.add_experimental_option('useAutomationExtension', False)

# 使用 undetected_chromedriver
driver = uc.Chrome(options=options)
driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

# 將你的頁面加載過程稍微延遲以避免檢測過快
driver.get(f"https://shopee.tw/search_user/?keyword={keyword}")
time.sleep(120)  # 等待頁面載入
actions = ActionChains(driver)

#print("[Page1]")
#print("########################")

#page_source = driver.page_source
#soup = BeautifulSoup(page_source, "html.parser")
#stores = soup.find_all('div', class_='shopee-search-user-item__nickname')
#sellers = soup.find_all('div', class_='shopee-search-user-item__username')
#fans = soup.find_all('span', class_='shopee-search-user-item__follow-count-number')
#following = soup.find_all('span', class_='shopee-search-user-item__follow-count-number')

#for i in range(len(sellers)):
#    store_name = stores[i].text.strip()
#    seller_name = sellers[i].text.strip()
#    fans_count = fans[i * 2].text.strip()  # Assuming fans and following are alternating, adjust if not
#    following_count = following[i * 2 + 1].text.strip()
#    result = f"[{store_name}] 賣家：{seller_name} ,粉絲數 : {fans_count} ,關注中數量：{following_count}"
#    print(result)

while True:

    print(f"[Page{page}]")
    print("########################")
    while True:
        driver.get(f"https://shopee.tw/search_user/?keyword={keyword}&page={page}")
        time.sleep(random.randint(5, 10))
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, "html.parser")
        stores = soup.find_all('div', class_='shopee-search-user-item__nickname')
        sellers = soup.find_all('div', class_='shopee-search-user-item__username')
        fans = soup.find_all('span', class_='shopee-search-user-item__follow-count-number')
        following = soup.find_all('span', class_='shopee-search-user-item__follow-count-number')
        sellers_count = len(sellers)
        empty_result = soup.find_all('div', class_='search-user-page__empty-result')
        if len(empty_result) > 0:
            break
        elif sellers_count > 0:
            break
        else:
            time.sleep(120)

    if sellers_count == 0:
        break
    
    for j in range(random.randint(1, 3)):
        actions.send_keys(Keys.PAGE_DOWN).perform()
    time.sleep(random.randint(5, 10))
    for j in range(random.randint(1, 3)):
        actions.send_keys(Keys.PAGE_UP).perform()
    time.sleep(random.randint(5, 10))
    page += 1

    for i in range(sellers_count):
        while True:
            try:
                store_name = stores[i].text.strip()
                seller_name = sellers[i].text.strip()

                # if 'biligitte' == seller_name:
                #     break
                # if 'okokokmike' == seller_name:
                #     break
                # if 'shutingpan464' == seller_name:
                #     break
                # if 'xxx123zzz' == seller_name:
                #     break
                # if 'cafele2' == seller_name:
                #     break
                # if 'piggy36422' == seller_name:
                #     break
                
                fans_count = convert_wan_to_number(fans[i * 2].text.strip().replace(",", ""))  # Assuming fans and following are alternating, adjust if not
                following_count = following[i * 2 + 1].text.strip()
                district = ""

                driver.get(f"https://shopee.tw/{seller_name}")
                time.sleep(random.randint(5, 10))
                if len(BeautifulSoup(driver.page_source, "html.parser").find_all('div', id='NEW_CAPTCHA')) > 0:
                    time.sleep(120)
                    continue
                if len(BeautifulSoup(driver.page_source, "html.parser").find_all('div', class_='ePKaWw')) > 0:
                    break
                for j in range(random.randint(1, 3)):
                    actions.send_keys(Keys.PAGE_DOWN).perform()
                time.sleep(random.randint(5, 10))
                actions.send_keys(Keys.END).perform()
                time.sleep(random.randint(5, 10))
                if len(BeautifulSoup(driver.page_source, "html.parser").find_all('div', id='NEW_CAPTCHA')) > 0:
                    time.sleep(120)
                    continue
                if len(driver.find_elements(By.CSS_SELECTOR, '.shop-page__no-product-placeholder')) > 0:
                    time.sleep(random.randint(5, 10))
                elif len(driver.find_elements(By.CSS_SELECTOR, '.shop-search-result-view__item:nth-child(1)')) > 0:
                    driver.find_element(By.CSS_SELECTOR, '.shop-search-result-view__item:nth-child(1)').click()
                    time.sleep(random.randint(5, 10))
                    for j in range(random.randint(1, 3)):
                        actions.send_keys(Keys.PAGE_DOWN).perform()
                    time.sleep(random.randint(5, 10))
                    actions.send_keys(Keys.END).perform()
                    time.sleep(random.randint(5, 10))
                    for ybxj32 in driver.find_elements(By.CSS_SELECTOR, '.ybxj32'):
                        if ybxj32.find_element(By.CSS_SELECTOR, 'h3').text == '出貨地':
                            district = ybxj32.find_element(By.CSS_SELECTOR, 'div').text
                elif len(driver.find_elements(By.CSS_SELECTOR, '.shop-collection-view__item:nth-child(1)')) > 0:
                    driver.find_element(By.CSS_SELECTOR, '.shop-collection-view__item:nth-child(1)').click()
                    time.sleep(random.randint(5, 10))
                    for j in range(random.randint(1, 3)):
                        actions.send_keys(Keys.PAGE_DOWN).perform()
                    time.sleep(random.randint(5, 10))
                    actions.send_keys(Keys.END).perform()
                    time.sleep(random.randint(5, 10))
                    for ybxj32 in driver.find_elements(By.CSS_SELECTOR, '.ybxj32'):
                        if ybxj32.find_element(By.CSS_SELECTOR, 'h3').text == '出貨地':
                            district = ybxj32.find_element(By.CSS_SELECTOR, 'div').text
                elif len(driver.find_elements(By.CSS_SELECTOR, '.navbar-with-more-menu__items')) > 0:
                    time.sleep(random.randint(5, 10))
                else:
                    time.sleep(120)
                    continue
                
                driver.get(f"https://shopee.tw/{seller_name}")
                time.sleep(random.randint(5, 10))
                if len(BeautifulSoup(driver.page_source, "html.parser").find_all('div', id='NEW_CAPTCHA')) > 0:
                    time.sleep(120)
                    continue
                if len(BeautifulSoup(driver.page_source, "html.parser").find_all('div', class_='ePKaWw')) > 0:
                    break
                for j in range(random.randint(1, 3)):
                    actions.send_keys(Keys.PAGE_DOWN).perform()
                # time.sleep(random.randint(5, 10))
                # for j in range(random.randint(1, 3)):
                #     actions.send_keys(Keys.PAGE_UP).perform()
                driver.execute_script("window.scrollTo(0, 0);")
                time.sleep(random.randint(5, 10))
                commodity_now = convert_wan_to_number(driver.find_element(By.CSS_SELECTOR, '.section-seller-overview__item:nth-child(1) .section-seller-overview__item-text-value').text.strip().replace(",", ""))
                driver.find_element(By.CSS_SELECTOR, '.section-seller-overview__item--clickable:nth-child(4)').click()
                time.sleep(random.randint(5, 10))
                for j in range(random.randint(1, 3)):
                    actions.send_keys(Keys.PAGE_DOWN).perform()
                time.sleep(random.randint(5, 10))
                rating_good = convert_wan_to_number(driver.find_element(By.CSS_SELECTOR, '.product-rating-overview__filter:nth-child(2)').text.strip().replace("5 星 (", "").replace(")", "").replace(",", ""))
                
                result = f"[{store_name}] 賣家：{seller_name} ,粉絲數 : {fans_count} ,關注中數量：{following_count}, 5星 : {rating_good}, {commodity_now}, 出貨地:{district}"
                print(result)
                kw = " "
                category = "一般賣家"
                row_data = fetch_row_by_seller(connection, seller_name)
                if row_data:
                    last_DATE = row_data['EXTRACT_DATE']
                    fans_Last = row_data['Fans_NOW']
                    commodity_Last = row_data['Commodity_NOW']
                    point_Last = row_data['Point_NOW']
                    if district == "":
                        district = row_data['District']
                    update_data(connection, seller_name, kw, category, district, fans_count, fans_count - fans_Last, fans_Last, commodity_now, commodity_now - commodity_Last, commodity_Last, rating_good, rating_good - point_Last, point_Last, seller_name, last_DATE)
                else:
                    insert_data(connection, seller_name, kw, category, district, fans_count, commodity_now, rating_good, seller_name)

            except Exception as e:
                print(f"{e}")
                time.sleep(120)
            else:
                break


connection.close()

driver.quit()
