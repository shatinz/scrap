import pandas as pd
import requests
from urllib.parse import quote
import time
import re

# Color mapping from English to Persian
COLOR_MAP = {
    'platinum': 'پلاتینی',
    'graphite': 'نوک مدادی',
    'black': 'مشکی',
    'silver': 'نقره ای',
    'sandstone': 'سنگ شنی',
    'ice blue': 'آبی یخی',
    'poppy red': 'قرمز',
    'sapphire': 'یاقوتی',
    'forest': 'جنگلی',
    'dune': 'طلایی'
}

#solved

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
    return None

def scrape_all_products_from_api(base_api_url, pid):
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; SurfaceIranScraper/2.0)"
    }
    all_products = []
    
    # Scrape from offset 0 to 3
    for offset in range(0, 4):
        paginated_url = f"{base_api_url}?limit=12&offset={offset}&pid={pid}"
        #print(f"Scraping API page: {paginated_url}")
        try:
            resp = requests.get(paginated_url, headers=headers, timeout=20)
            resp.raise_for_status()
            data = resp.json()
            
            products_on_page = data.get('rows', [])
            if not products_on_page:
                print(f"No more products found at offset {offset}. Stopping.")
                break

            for prod_data in products_on_page:
                base_name = prod_data.get('productname', '')
                prod_id = prod_data.get('_id')
                price_list = prod_data.get('price', [])

                if not price_list:
                    continue

                for price_variant in price_list:
                    variant_info = price_variant.get('priceinfo', '')
                    price = price_variant.get('price', '')
                    
                    # Combine base name with variant info for a unique, searchable name
                    full_name = f"{base_name} {variant_info}"
                    
                    # Construct URL
                    url_full = f"https://surfaceiran.com/p/{prod_id}/{quote(base_name)}" if prod_id and base_name else ''
                    
                    all_products.append({'name': full_name, 'price': str(price), 'url': url_full})

            time.sleep(1) # Be polite to the server

        except requests.exceptions.RequestException as e:
            print(f"[DEBUG] Error scraping API at offset {offset}: {e}")
            continue # Continue to the next offset
            
    return all_products

# --- Main script ---

API_BASE_URL = 'https://surfaceiran.com/products/getList'
PRODUCT_GROUP_ID = '65e24e454b49f2d824666a29'
all_products = scrape_all_products_from_api(API_BASE_URL, PRODUCT_GROUP_ID)

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
    color_en = row.get('Color')
    color_fa = None
    if isinstance(color_en, str):
        color_fa = COLOR_MAP.get(color_en.lower())

    features = [row['Cpu'], row['Ram'], row['SSD'], color_fa]
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
