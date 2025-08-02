import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import re
#solved
#surfaceiran.com does not have surface 11 available
def normalize(text):
    # Replace Persian 'گیگابایت' with 'gb'
    text = str(text).replace('گیگابایت', 'gb')
    return re.sub(r'[^a-zA-Z0-9آ-ی]', '', text.lower())

def best_match(product_name, features, products):
    search_terms = [normalize(f) for f in features if f]
    # print(f"    [DEBUG] Normalized search terms: {search_terms}")
    # print(f"    [DEBUG] Original features: {features}")
    for prod in products:
        title = normalize(prod['name'])
        # print(f"    [DEBUG] Normalized product title: {title}")
        # print(f"    [DEBUG] Original product name: {prod['name']}")
        if all(term in title for term in search_terms):
            # print(f"    [DEBUG] MATCH FOUND: {prod['name']}")
            return prod
    # print(f"    [DEBUG] No full feature match for search terms: {search_terms}")
    # print("    [DEBUG] Candidate product titles:")
    # for prod in products:
        # print(f"      - {prod['name']}")
    #return None

def scrape_all_products_from_url(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; SurfaceIranScraper/1.0)"
    }
    try:
        resp = requests.get(url, headers=headers, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'html.parser')
        products = []
        for prod_div in soup.find_all('div', class_='productItem'):
            name_tag = prod_div.find('div', class_='productname')
            name = name_tag.get_text(strip=True) if name_tag else ''
            price_tag = prod_div.find('span', class_='price')
            price_text = price_tag.get_text(strip=True) if price_tag else ''
            price = re.sub(r'[^\d]', '', price_text) if price_text else ''
            a_tag = prod_div.find('a', href=True)
            url_path = a_tag['href'] if a_tag else ''
            if url_path and not url_path.startswith('http'):
                url_full = 'https://surfaceiran.com' + url_path
            else:
                url_full = url_path
            products.append({'name': name, 'price': price, 'url': url_full})
        return products
    except Exception as e:
        # print(f"[DEBUG] Error scraping products from url: {e}")
        return []

# --- Main script ---

CATEGORY_URL = 'https://surfaceiran.com/products/65e24e454b49f2d824666a29/%D8%B3%D8%B1%D9%81%DB%8C%D8%B3-%D9%BE%D8%B1%D9%88'

all_products = scrape_all_products_from_url(CATEGORY_URL)

# print("[ALL SCRAPED PRODUCTS]")
# for prod in all_products:
    # print(f"  - Name: {prod['name']}")
    # print(f"    URL: {prod['url']}")
    # print(f"    Price: {prod['price']}")

# Load Excel and match

df = pd.read_excel('SampleSites.xlsx')
if 'surfaceiran.com' not in df.columns:
    df['surfaceiran.com'] = ''
df['surfaceiran.com'] = df['surfaceiran.com'].astype('object')
if 'surfaceiranproducturl' not in df.columns:
    idx = list(df.columns).index('surfaceiran.com') + 1
    df.insert(idx, 'surfaceiranproducturl', '')
df['surfaceiranproducturl'] = df['surfaceiranproducturl'].astype('object')

for idx, row in df.iterrows():
    product_name = row['Product name']
    features = [row['Cpu'], row['Ram'], row['SSD']]
    # print(f"Processing row {idx+1} for surfaceiran.com: {product_name}, features: {features}")
    match = best_match(product_name, features, all_products)
    if match:
        # print(f"  [DEBUG] Matched product: {match['name']}")
        url = match['url']
        price = match['price']
        df.at[idx, 'surfaceiranproducturl'] = url
    else:
        # print("  [DEBUG] No matching product found.")
        price = ''
        df.at[idx, 'surfaceiranproducturl'] = ''
    # print(f"  -> Price found: {price}")
    df.at[idx, 'surfaceiran.com'] = price
    time.sleep(1)

df.to_excel('SampleSites.xlsx', index=False)
# print('Done. Prices and product URLs updated in SampleSites.xlsx.')
