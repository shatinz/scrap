import pandas as pd
import requests
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
# NOTE: Make sure you have installed selenium and geckodriver.(im using firefox and geckodriver)
# pip install selenium
# Download geckodriver from https://github.com/mozilla/geckodriver/releases

# Read the table from SampleSites.xlsx
print('Reading data from SampleSites.xlsx...')
df = pd.read_excel('SampleSites.xlsx')

# Function to get price using Selenium
def get_price_selenium(product_url):
    options = Options()
    # options.add_argument('--headless')  # Comment this out for now
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    driver = webdriver.Firefox(options=options)
    try:
        driver.set_page_load_timeout(60)  # 60 seconds
        print("Opening URL...")
        driver.get(product_url)
        print("URL opened.")
        print("Finding element...")
        price_elem = driver.find_element(By.CSS_SELECTOR, '.priceVal')
        print("Element found.")
        price = price_elem.text.strip()
        print(f"[SELENIUM DEBUG] Price found: {price}")
        return price
    except Exception as e:
        print(f"[SELENIUM DEBUG] Error: {e}")
        return ''
    finally:
        driver.quit()

# Function to search and get price from surfaceiran.com
def get_surfaceiran_price(product_name, features):
    search_url = f'https://surfaceiran.com/products/getShortList?search={product_name}'
    print(f"  [DEBUG] Search URL: {search_url}")
    try:
        resp = requests.get(search_url, timeout=10)
        print(f"  [DEBUG] Search response status: {resp.status_code}")
        data = resp.json()
        print("[DEBUG] Products returned from API:")
        for p in data.get('rows', []):
            print(p)
        if not data or not data.get('rows'):
            print(f"  [DEBUG] No products found for product name: {product_name}")
            return ''
        # Remove color from features for matching (assume color is last in the list)
        match_features = features[:-1]
        def feature_match(product):
            name = product.get('productname', '').lower()
            return all(str(f).lower() in name for f in match_features if f)
        filtered = [p for p in data['rows'] if feature_match(p)]
        if not filtered:
            print(f"  [DEBUG] No products matched features: {features}")
            return ''
        product = filtered[0]
        product_id = product['_id']
        product_url = f'https://surfaceiran.com/p/{product_id}'
        print(f"  [DEBUG] Product URL: {product_url}")
        # Use Selenium to get the price
        price = get_price_selenium(product_url)
        return price
    except Exception as e:
        print(f"Error for {product_name} {features}: {e}")
        return ''

# For each row, fill in the price for surfaceiran.com
df['surfaceiran.com'] = df['surfaceiran.com'].astype('object')
for idx, row in df.iterrows():
    product_name = row['Product name']
    features = [row['Cpu'], row['Ram'], row['SSD'], row['Color']]
    print(f"Updating row {idx+1}: {product_name}, features: {features}")
    price = get_surfaceiran_price(product_name, features)
    print(f"  -> Price found: {price}")
    df.at[idx, 'surfaceiran.com'] = price
    time.sleep(1)  # Be polite to the server

# Save the updated table back to SampleSites.xlsx
print('Saving updated data to SampleSites.xlsx...')
df.to_excel('SampleSites.xlsx', index=False)

print('Done. Prices updated in SampleSites.xlsx.') 