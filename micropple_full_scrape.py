import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
import re

CATEGORY_URLS = [
    "https://micropple.ir/product-category/microsoft/laptop-microsoft/",
    "https://micropple.ir/product-category/microsoft/tablet-microsoft/",
    "https://micropple.ir/product-category/microsoft/surface-laptop-studio/",
    "https://micropple.ir/product-category/microsoft/surface-go/",
    "https://micropple.ir/product-category/microsoft/surface-studio/",
    "https://micropple.ir/product-category/microsoft/surface-laptop-go/",
    "https://micropple.ir/product-category/microsoft/surface-book/",
]

def get_all_products_from_category(category_url):
    import random
    products = []
    page = 1
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    }
    while True:
        url = category_url if page == 1 else f"{category_url.rstrip('/')}/page/{page}/"
        print(f"[DEBUG] Scraping: {url}")
        tries = 0
        while tries < 3:
            try:
                resp = requests.get(url, timeout=40, headers=headers)
                break
            except requests.exceptions.ReadTimeout:
                tries += 1
                print(f"[WARNING] Timeout on {url}, retrying ({tries}/3)...")
                time.sleep(2 + random.random() * 2)
        else:
            print(f"[ERROR] Failed to fetch {url} after 3 retries, skipping.")
            break
        soup = BeautifulSoup(resp.text, 'html.parser')
        product_links = soup.find_all('a', class_='product-image-link')
        if not product_links:
            break
        for link in product_links:
            parent = link.find_parent(class_='product-grid-item')
            title = ''
            if parent:
                title_tag = parent.find('h3', class_='wd-entities-title')
                if title_tag and title_tag.a:
                    title = title_tag.a.get_text(strip=True)
            href = link.get('href', '').strip()
            if title and href:
                products.append({'title': title, 'url': href})
        next_page = soup.find('a', class_='next')
        if not next_page:
            if len(product_links) < 10:
                break
            else:
                page += 1
        else:
            page += 1
        time.sleep(1)
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

def get_price_micropple(product_url):
    options = Options()
    # options.add_argument('--headless')  # Uncomment for headless mode if desired
    driver = webdriver.Firefox(options=options)
    try:
        driver.set_page_load_timeout(60)
        print(f"Opening micropple.ir URL: {product_url}")
        driver.get(product_url)
        print("URL opened.")
        print("Finding price element...")

        price_texts = []
        # Try <p class='price'>
        try:
            price_p = driver.find_element(By.CSS_SELECTOR, 'p.price')
            price_spans = price_p.find_elements(By.CSS_SELECTOR, 'span.woocommerce-Price-amount.amount')
            price_texts = [span.text.strip() for span in price_spans if span.text.strip()]
        except Exception:
            pass

        # Try <span class='price'>
        if not price_texts:
            try:
                price_span = driver.find_element(By.CSS_SELECTOR, 'span.price')
                price_spans = price_span.find_elements(By.CSS_SELECTOR, 'span.woocommerce-Price-amount.amount')
                price_texts = [span.text.strip() for span in price_spans if span.text.strip()]
            except Exception:
                pass

        # Try any span.woocommerce-Price-amount.amount on the page
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
if 'micropple.ir' not in df.columns:
    df['micropple.ir'] = ''
df['micropple.ir'] = df['micropple.ir'].astype('object')
# Add the new column for product URLs if not present
if 'micropleproducturl' not in df.columns:
    # Insert it right after 'micropple.ir' if possible
    col_list = list(df.columns)
    idx = col_list.index('micropple.ir') + 1
    df.insert(idx, 'micropleproducturl', '')
df['micropleproducturl'] = df['micropleproducturl'].astype('object')

all_products = get_all_products()

for idx, row in df.iterrows():
    product_name = row['Product name']
    features = [row['Cpu'], row['Ram'], row['SSD'], row['Color']]
    print(f"Processing row {idx+1}: {product_name}, features: {features}")
    match = best_match(product_name, features, all_products)
    if match:
        print(f"  [DEBUG] Matched product: {match['title']} ({match['url']})")
        price = get_price_micropple(match['url'])
        print(f"  [DEBUG] Product URL: {match['url']}")
        df.at[idx, 'micropleproducturl'] = match['url']
    else:
        print("  [DEBUG] No matching product found.")
        price = ''
        df.at[idx, 'micropleproducturl'] = ''
    print(f"  -> Price found: {price}")
    df.at[idx, 'micropple.ir'] = price
    time.sleep(1)

df.to_excel('SampleSites.xlsx', index=False)
print('Done. Prices and product URLs updated in SampleSites.xlsx.') 