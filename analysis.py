import time
import pandas as pd
import matplotlib.pyplot as plt
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
chrome_driver_path = r'C:\\Users\\forty\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe'
chrome_options = Options()
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.5938.62 Safari/537.36")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--headless")  

service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.set_page_load_timeout(60)
url = "https://www.barchart.com/futures"
driver.get(url)
try:
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, "//div[contains(@class, 'bc-datatable-header-tooltip')]")))


    headers = driver.find_elements(By.XPATH, "//div[contains(@class, 'bc-datatable-header-tooltip')]")
    header_names = [header.text.strip() for header in headers]


    rows = driver.find_elements(By.XPATH, "//div[@role='row']")

    print(f"Number of rows fetched: {len(rows)}")

    data = []


    for row in rows[1:]: 
        cells = row.find_elements(By.XPATH, ".//div[@role='cell']") 
        if len(cells) == len(header_names):  
            row_data = [cell.text.strip() for cell in cells]
            data.append(row_data)

   
    df = pd.DataFrame(data, columns=header_names)
    print(df.head())  

    df['High'] = pd.to_numeric(df['High'], errors='coerce') 
    df['Low'] = pd.to_numeric(df['Low'], errors='coerce') 
    df['Mean'] = (df['High'] + df['Low']) / 2  

    plt.figure(figsize=(10,6))
    plt.plot(df['Contract Name'], df['High'], label='High', marker='o')
    plt.plot(df['Contract Name'], df['Low'], label='Low', marker='o')
    plt.plot(df['Contract Name'], df['Mean'], label='Mean', marker='x')
    plt.xlabel('Contract Name')
    plt.ylabel('Price')
    plt.title('High, Low, and Mean Prices')
    plt.xticks(rotation=90)  
    plt.legend()
    plt.tight_layout()
    plt.show()


    df['Change'] = pd.to_numeric(df['Change'], errors='coerce')  
    largest_change_row = df.loc[df['Change'].idxmax()]
    print(f"Contract Name with largest change: {largest_change_row['Contract Name']}")
    print(f"Last Price: {largest_change_row['Last']}")

  
    output_path = r"C:\\Users\\forty\\Downloads\\futures_data.xlsx"
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Raw Data', index=False)

    print(f"Data has been saved to {output_path}")

except Exception as e:
    print(f"Error: {e}")

finally:
    driver.quit()  