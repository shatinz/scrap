import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re

CATEGORY_URLS = [
    "https://mysurface.ir/surface-pro/",
    "https://mysurface.ir/surface-laptop/",
    "https://mysurface.ir/surface-go/",
]

def get_all_products_from_category(category_url):
    from selenium.webdriver.common.by import By
    import random
    products = []
    page = 1
    options = Options()
    options.add_argument('--headless')  # Uncomment for headless mode if desired
    profile = FirefoxProfile()
    profile.set_preference("general.useragent.override", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
    options.profile = profile
    driver = webdriver.Firefox(options=options)
    try:
        max_pages = 30  # Optional: set a reasonable upper limit

        while page <= max_pages:
            url = category_url if page == 1 else f"{category_url.rstrip('/')}/page/{page}/"
            print(f"[DEBUG] Scraping (Selenium): {url}")
            tries = 0
            success = False
            while tries < 3:
                try:
                    driver.set_page_load_timeout(60)
                    driver.get(url)
                    try:
                        WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, 'a.woocommerce-LoopProduct-link, a.woocommerce-loop-product__link'))
                        )
                    except Exception as e:
                        print(f"[WARNING] Page did not load products in time: {e}")
                    success = True
                    break
                except Exception as e:
                    tries += 1
                    print(f"[WARNING] Timeout or error on {url}, retrying ({tries}/3)... {e}")
                    time.sleep(2 + random.random() * 2)
            if not success:
                print(f"[ERROR] Failed to fetch {url} after 3 retries, assuming end of pagination.")
                break
            # Find all product links
            product_links = driver.find_elements(By.CSS_SELECTOR, 'a.woocommerce-LoopProduct-link, a.woocommerce-loop-product__link')
            if not product_links:
                print(f"[DEBUG] No products found on {url}, assuming end of pagination.")
                break
            for link in product_links:
                try:
                    title = link.get_attribute('title') or link.text
                    title = title.strip()
                    href = link.get_attribute('href').strip()
                    # Find parent div with class containing 'box-text'
                    parent = link.find_element(By.XPATH, "ancestor::div[contains(@class, 'box-text')]")
                    price = ''
                    try:
                        price_wrapper = parent.find_element(By.CLASS_NAME, 'price-wrapper')
                        price_span = price_wrapper.find_element(By.CLASS_NAME, 'woocommerce-Price-amount')
                        price = price_span.text.strip()
                    except Exception:
                        pass
                    if title and href:
                        products.append({'title': title, 'url': href, 'cat_price': price})
                except Exception as e:
                    print(f"[WARNING] Error extracting product info: {e}")
            # Pagination: check for next page
            try:
                next_page = driver.find_element(By.CSS_SELECTOR, 'a.next')
                if not next_page:
                    break
                else:
                    page += 1
            except Exception:
                break
            time.sleep(1)
    finally:
        driver.quit()
    return products

def get_all_products():
    all_products = []
    for cat_url in CATEGORY_URLS:
        all_products.extend(get_all_products_from_category(cat_url))
    print(f"[DEBUG] Total products scraped: {len(all_products)}")
    return all_products

def normalize(text):
    return re.sub(r'[^a-zA-Z0-9آ-ی]', '', str(text).lower())

def best_match(product_name, features, products, min_match=2):
    search_terms = [normalize(product_name)] + [normalize(f) for f in features if f]
    best = None
    best_score = 0
    for prod in products:
        title = normalize(prod['title'])
        score = sum(term in title for term in search_terms)
        if score > best_score:
            best = prod
            best_score = score
    if best_score < min_match:
        print(f"    [DEBUG] No strong match for search terms: {search_terms}")
        print("    [DEBUG] Candidate product titles:")
        for prod in products:
            print(f"      - {prod['title']}")
        return None
    return best

def get_price_mysurface(product_url):
    options = Options()
    options.add_argument('--headless')  # Uncomment for headless mode if desired
    driver = webdriver.Firefox(options=options)
    try:
        driver.set_page_load_timeout(60)
        print(f"Opening mysurface.ir URL: {product_url}")
        driver.get(product_url)
        print("URL opened.")
        print("Finding price element...")
        price_texts = []
        try:
            price_p = driver.find_element(By.CSS_SELECTOR, 'p.price')
            price_spans = price_p.find_elements(By.CSS_SELECTOR, 'span.woocommerce-Price-amount.amount')
            price_texts = [span.text.strip() for span in price_spans if span.text.strip()]
        except Exception:
            pass
        if not price_texts:
            try:
                price_span = driver.find_element(By.CSS_SELECTOR, 'span.price')
                price_spans = price_span.find_elements(By.CSS_SELECTOR, 'span.woocommerce-Price-amount.amount')
                price_texts = [span.text.strip() for span in price_spans if span.text.strip()]
            except Exception:
                pass
        if not price_texts:
            try:
                price_spans = driver.find_elements(By.CSS_SELECTOR, 'span.woocommerce-Price-amount.amount')
                price_texts = [span.text.strip() for span in price_spans if span.text.strip()]
            except Exception:
                pass
        if price_texts:
            price = ' - '.join(price_texts)
            print(f"[SELENIUM DEBUG] Price found: {price}")
            return price
        else:
            print("[WARNING] Price element not found or empty! Printing page title for debug:")
            print(driver.title)
            return ''
    except Exception as e:
        print(f"[SELENIUM DEBUG] Error: {e}")
        return ''
    finally:
        driver.quit()

# --- Main script ---

df = pd.read_excel('SampleSites.xlsx')
if 'mysurface.ir' not in df.columns:
    df['mysurface.ir'] = ''
df['mysurface.ir'] = df['mysurface.ir'].astype('object')
if 'mysurfaceproducturl' not in df.columns:
    idx = list(df.columns).index('mysurface.ir') + 1
    df.insert(idx, 'mysurfaceproducturl', '')
df['mysurfaceproducturl'] = df['mysurfaceproducturl'].astype('object')

all_products = get_all_products()

for idx, row in df.iterrows():
    product_name = row['Product name']
    features = [row['Cpu'], row['Ram'], row['SSD'], row['Color']]
    print(f"Processing row {idx+1} for mysurface.ir: {product_name}, features: {features}")
    match = best_match(product_name, features, all_products)
    if match:
        print(f"  [DEBUG] Matched product: {match['title']} ({match['url']})")
        if match.get('cat_price'):
            price = match['cat_price']
            print(f"  [DEBUG] Category page price used: {price}")
        else:
            price = get_price_mysurface(match['url'])
        print(f"  [DEBUG] Product URL: {match['url']}")
        df.at[idx, 'mysurfaceproducturl'] = match['url']
    else:
        print("  [DEBUG] No matching product found.")
        price = ''
        df.at[idx, 'mysurfaceproducturl'] = ''
    print(f"  -> Price found: {price}")
    df.at[idx, 'mysurface.ir'] = price
    time.sleep(1)

df.to_excel('SampleSites.xlsx', index=False)
print('Done. Prices and product URLs updated in SampleSites.xlsx.') 