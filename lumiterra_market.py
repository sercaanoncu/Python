#!/usr/bin/env python
# coding: utf-8

# In[1]:


from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import threading
import os

service = Service(r'C:\Users\Sercan\Downloads\geckodriver-v0.35.0-win32\geckodriver.exe')
options = Options()
options.add_argument("--headless")
options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
driver = webdriver.Firefox(service=service, options=options)

base_url = "https://marketplace.skymavis.com/collections/lumiterra?auction=Sale&page={}&type=consumable&type=material&type=taskticket"
file_path = r'C:\Users\Sercan\dist\nft_market_material_prices_sorted.xlsx'

def update_data():
    while not stop_flag.is_set():
        new_data = pd.DataFrame(columns=['Item Name', 'Price'])
        
        for page in range(1, 7): 
            url = base_url.format(page)
            driver.get(url)
            
            try:
                element_present = EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'Card_Footer__vcM3_')]"))
                WebDriverWait(driver, 5).until(element_present)
                
                items = driver.find_elements(By.XPATH, "//div[contains(@class, 'Card_Footer__vcM3_')]")
                
                if not items:
                    print(f"Sayfada öğe bulunamadı: {url}")
                    continue
                
                for item in items:
                    try:
                        item_name_elements = item.find_elements(By.XPATH, ".//div[contains(@class, 'Erc1155Card_name__Z9fOp')]")
                        item_name = item_name_elements[-1].text.strip() if item_name_elements else 'N/A'
                        
                        price_element = item.find_element(By.XPATH, ".//div[contains(@class, 'Erc1155Card_prices__kLn2L')]//h5[contains(@class, 'TokenPrice_price__ZrJBZ')]")
                        price = price_element.text.strip() if price_element else 'N/A'
                        
                        print(f"Item: {item_name}, Price: {price}")
                        
                        new_data = pd.concat([new_data, pd.DataFrame([{'Item Name': item_name, 'Price': price}])], ignore_index=True)
                    except Exception as e:
                        print(f"Öğe verileri alınırken bir hata oluştu: {e}")
                        continue
            
            except Exception as e:
                print(f"Bir hata oluştu: {e}")
                continue
        
        new_data['Price'] = pd.to_numeric(new_data['Price'], errors='coerce')
        new_data.sort_values(by='Item Name', inplace=True)
        
        try:
            if os.path.exists(file_path):
                existing_data = pd.read_excel(file_path)
                updated_data = pd.concat([existing_data, new_data])
                updated_data = updated_data.groupby('Item Name').agg({'Price': 'last'}).reset_index()
            else:
                updated_data = new_data.groupby('Item Name').agg({'Price': 'last'}).reset_index()
            
            updated_data.to_excel(file_path, index=False)
            print("Excel dosyası başarıyla güncellendi.")
        except Exception as e:
            print(f"Excel dosyasına yazarken bir hata oluştu: {e}")
        
        time.sleep(5)

stop_flag = threading.Event()

def wait_for_exit():
    input("Verileri güncellemeye başladık. Durdurmak için Enter'a basın...")
    stop_flag.set()
    print("İşlem durduruluyor...")

threading.Thread(target=update_data, daemon=True).start()
wait_for_exit()

driver.quit()


# In[ ]:




