import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import urllib.parse
import re

#color is not a feature for yasinrayan.com
def search_yasinrayan_url(product_name, features):
    search_query = product_name + ' ' + ' '.join(str(f) for f in features if f)
    search_url = f'https://www.yasinrayan.com/?s={urllib.parse.quote(search_query)}&post_type=product'
    print(f"[DEBUG] yasinrayan.com search URL: {search_url}")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    }
    try:
        resp = requests.get(search_url, timeout=20, headers=headers)
        soup = BeautifulSoup(resp.text, 'html.parser')
        products = []
        for h3 in soup.select('h3.wd-entities-title'):
            a = h3.find('a')
            if not a:
                continue
            title = a.get('title', '').strip() or a.text.strip()
            href = a.get('href', '').strip()
            # Find the price in the next siblings
            price = ''
            price_span = None
            for sib in h3.find_all_next(['span', 'p'], limit=5):
                if 'price' in sib.get('class', []):
                    price_span = sib.find('span', class_='woocommerce-Price-amount')
                    if price_span:
                        price = price_span.get_text(strip=True)
                        break
            if title and href:
                products.append({'title': title, 'url': href, 'cat_price': price})
        return products
    except Exception as e:
        print(f"[DEBUG] Error searching yasinrayan.com: {e}")
        return []

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

def get_available_colors(product_url):
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; YasinrayanScraper/1.0)"
    }
    try:
        resp = requests.get(product_url, headers=headers, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'html.parser')
        color_divs = soup.select('div.wd-swatches-product .wd-swatch-text')
        colors = [div.get_text(strip=True) for div in color_divs]
        return colors
    except Exception as e:
        print(f"[DEBUG] Error fetching colors from {product_url}: {e}")
        return []

def map_color_name(color):
    color_map = {
        'platinum': 'پلاتینی',
        'graphite': 'مشکی',
        'black': 'مشکی',
        'sapphire': 'آبی (Sapphire)',
        'gold': 'شنی طلایی',
        # Add more mappings as needed
    }
    color_norm = normalize(color)
    for en, fa in color_map.items():
        if color_norm == normalize(en) or color_norm == normalize(fa):
            return fa
    return color  # fallback to original

# --- Main script ---

df = pd.read_excel('SampleSites.xlsx')
if 'yasinrayan.com' not in df.columns:
    df['yasinrayan.com'] = ''
df['yasinrayan.com'] = df['yasinrayan.com'].astype('object')
if 'yasinrayanproducturl' not in df.columns:
    idx = list(df.columns).index('yasinrayan.com') + 1
    df.insert(idx, 'yasinrayanproducturl', '')
df['yasinrayanproducturl'] = df['yasinrayanproducturl'].astype('object')

for idx, row in df.iterrows():
    product_name = row['Product name']
    features = [row['Cpu'], row['Ram'], row['SSD']]
    desired_color = str(row['Color']).strip()
    desired_color_mapped = map_color_name(desired_color)
    print(f"Processing row {idx+1} for yasinrayan.com: {product_name}, features: {features}, color: {desired_color} (mapped: {desired_color_mapped})")
    products = search_yasinrayan_url(product_name, features)
    match = best_match(product_name, features, products)
    if match:
        available_colors = get_available_colors(match['url'])
        print(f"  [DEBUG] Available colors: {available_colors}")
        if any(normalize(desired_color_mapped) == normalize(c) for c in available_colors):
            print(f"  [DEBUG] Color match found: {desired_color_mapped}")
            price = match['cat_price']
            df.at[idx, 'yasinrayanproducturl'] = match['url']
        else:
            print(f"  [DEBUG] Desired color '{desired_color}' (mapped: '{desired_color_mapped}') not found in available colors.")
            price = ''
            df.at[idx, 'yasinrayanproducturl'] = ''
    else:
        print("  [DEBUG] No matching product found.")
        price = ''
        df.at[idx, 'yasinrayanproducturl'] = ''
    print(f"  -> Price found: {price}")
    df.at[idx, 'yasinrayan.com'] = price
    time.sleep(1)

df.to_excel('SampleSites.xlsx', index=False)
print('Done. Prices and product URLs updated in SampleSites.xlsx.') 