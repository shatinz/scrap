import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import re

def persian_to_english_digits(text):
    persian_digits = '۰۱۲۳۴۵۶۷۸۹'
    english_digits = '0123456789'
    for p, e in zip(persian_digits, english_digits):
        text = text.replace(p, e)
    return text

def normalize(text):
    text = str(text).lower()
    text = persian_to_english_digits(text)
    return re.sub(r'[^a-zA-Z0-9آ-ی]', '', text)

def best_match(product_name, features, products, min_match=2):
    # Combine product name and features for more robust matching
    search_terms = [normalize(product_name)] + [normalize(f) for f in features if f]
    print(f"    [DEBUG] Normalized search terms: {search_terms}")
    print(f"    [DEBUG] Original features: {features}")
    best = None
    best_score = 0
    for prod in products:
        title = normalize(prod['name'])
        print(f"    [DEBUG] Normalized product title: {title}")
        print(f"    [DEBUG] Original product name: {prod['name']}")
        score = sum(term in title for term in search_terms)
        print(f"    [DEBUG] Match score: {score}")
        if score > best_score:
            best = prod
            best_score = score
    if best and best_score >= min_match:
        print(f"    [DEBUG] MATCH FOUND (score={best_score}): {best['name']}")
        return best
    print(f"    [DEBUG] No strong match for search terms: {search_terms}")
    print("    [DEBUG] Candidate product titles:")
    for prod in products:
        print(f"      - {prod['name']}")
    return None

def scrape_all_products_from_micropple(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; MicroppleScraper/1.0)"
    }
    products = []
    page = 1
    while True:
        page_url = url if page == 1 else f"{url.rstrip('/')}/page/{page}/"
        print(f"[DEBUG] Scraping: {page_url}")
        try:
            resp = requests.get(page_url, headers=headers, timeout=20)
            resp.raise_for_status()
        except Exception as e:
            print(f"[DEBUG] Error scraping {page_url}: {e}")
            break
        soup = BeautifulSoup(resp.text, 'html.parser')
        product_divs = soup.find_all('div', class_='product-grid-item')
        if not product_divs:
            break
        for prod_div in product_divs:
            # Name
            name_tag = prod_div.find('h3', class_='wd-entities-title')
            name = name_tag.get_text(strip=True) if name_tag else ''
            # URL
            a_tag = prod_div.find('a', class_='product-image-link', href=True)
            url_full = a_tag['href'] if a_tag else ''
            # Price
            price_tag = prod_div.find('span', class_='woocommerce-Price-amount')
            price = price_tag.get_text(strip=True) if price_tag else ''
            products.append({'name': name, 'price': price, 'url': url_full})
        # Pagination: look for next page
        next_page = soup.find('a', class_='next')
        if not next_page:
            break
        page += 1
        time.sleep(1)
    return products

# --- Main script ---

CATEGORY_URL = 'https://micropple.ir/product-category/microsoft/tablet-microsoft/'

all_products = scrape_all_products_from_micropple(CATEGORY_URL)

print("[ALL SCRAPED PRODUCTS]")
for prod in all_products:
    print(f"  - Name: {prod['name']}")
    print(f"    URL: {prod['url']}")
    print(f"    Price: {prod['price']}")

# Load Excel and match

df = pd.read_excel('SampleSites.xlsx')
if 'micropple.ir' not in df.columns:
    df['micropple.ir'] = ''
df['micropple.ir'] = df['micropple.ir'].astype('object')
if 'microppleproducturl' not in df.columns:
    idx = list(df.columns).index('micropple.ir') + 1
    df.insert(idx, 'microppleproducturl', '')
df['microppleproducturl'] = df['microppleproducturl'].astype('object')

for idx, row in df.iterrows():
    product_name = row['Product name']
    features = [row['Cpu'], row['Ram'], row['SSD'], row['Color']]
    print(f"Processing row {idx+1} for micropple.ir: {product_name}, features: {features}")
    match = best_match(product_name, features, all_products)
    if match:
        print(f"  [DEBUG] Matched product: {match['name']}")
        url = match['url']
        price = match['price']
        df.at[idx, 'microppleproducturl'] = url
    else:
        print("  [DEBUG] No matching product found.")
        price = ''
        df.at[idx, 'microppleproducturl'] = ''
    print(f"  -> Price found: {price}")
    df.at[idx, 'micropple.ir'] = price
    time.sleep(1)

df.to_excel('SampleSites.xlsx', index=False)
print('Done. Prices and product URLs updated in SampleSites.xlsx.') 