import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import re

def normalize(text):
    # Replace Persian 'گیگابایت' with 'gb'
    text = str(text).replace('گیگابایت', 'gb')
    return re.sub(r'[^a-zA-Z0-9آ-ی]', '', text.lower())

def best_match(product_name, features, products):
    search_terms = [normalize(f) for f in features if f]
    for prod in products:
        title = normalize(prod['title'])
        if all(term in title for term in search_terms):
            return prod
    print(f"    [DEBUG] No full feature match for search terms: {search_terms}")
    print("    [DEBUG] Candidate product titles:")
    for prod in products:
        print(f"      - {prod['title']}")
    return None

def search_raayaatech(product_name):
    search_query = product_name
    search_url = f'https://raayaatech.com/search?q={requests.utils.quote(search_query)}'
    print(f"[DEBUG] raayaatech.com search URL: {search_url}")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    }
    try:
        resp = requests.get(search_url, timeout=20, headers=headers)
        soup = BeautifulSoup(resp.text, 'html.parser')
        products = []
        for prod_div in soup.select('div.col-xl-3.price_on, div.col-lg-4.price_on, div.col-md-4.price_on'):
            a = prod_div.select_one('a.title.overflow-hidden')
            if not a:
                continue
            title = a.get('title', '').strip() or a.text.strip()
            href = a.get('href', '').strip()
            if href and not href.startswith('http'):
                href = 'https://raayaatech.com' + href
            price = ''
            price_tag = prod_div.select_one('div.price-area span.price')
            if price_tag:
                price = price_tag.get_text(strip=True)
            if title and href:
                products.append({'title': title, 'url': href, 'cat_price': price})
        return products
    except Exception as e:
        print(f"[DEBUG] Error searching raayaatech.com: {e}")
        return []

# --- Main script ---

df = pd.read_excel('SampleSites.xlsx')
if 'raayaatech.com' not in df.columns:
    df['raayaatech.com'] = ''
df['raayaatech.com'] = df['raayaatech.com'].astype('object')
if 'raayaatechproducturl' not in df.columns:
    idx = list(df.columns).index('raayaatech.com') + 1
    df.insert(idx, 'raayaatechproducturl', '')
df['raayaatechproducturl'] = df['raayaatechproducturl'].astype('object')

for idx, row in df.iterrows():
    product_name = row['Product name']
    features = [row['Cpu'], row['Ram'], row['SSD']]
    print(f"Processing row {idx+1} for raayaatech.com: {product_name}, features: {features}")
    products = search_raayaatech(product_name)
    match = best_match(product_name, features, products)
    if match:
        print(f"  [DEBUG] Matched product: {match['title']} ({match['url']})")
        price = match['cat_price']
        print(f"  [DEBUG] Category page price used: {price}")
        df.at[idx, 'raayaatechproducturl'] = match['url']
    else:
        print("  [DEBUG] No matching product found.")
        price = ''
        df.at[idx, 'raayaatechproducturl'] = ''
    print(f"  -> Price found: {price}")
    df.at[idx, 'raayaatech.com'] = price
    time.sleep(1)

df.to_excel('SampleSites.xlsx', index=False)
print('Done. Prices and product URLs updated in SampleSites.xlsx.') 