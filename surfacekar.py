
import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import urllib.parse
import re
#solves
# can not switch between colors

def search_surfacekar_url(product_name, features, color):
    search_query = product_name + ' ' + ' '.join(str(f) for f in features if f)
    search_url = f'https://surfacekar.com/?s={urllib.parse.quote(search_query)}&post_type=product'
    print(f"[DEBUG] surfacekar.com search URL: {search_url}")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    }
    try:
        resp = requests.get(search_url, timeout=20, headers=headers)
        soup = BeautifulSoup(resp.text, 'html.parser')
        products = []
        for product_div in soup.select('.product-grid-item'):
            h3 = product_div.select_one('h3.wd-entities-title')
            if not h3:
                continue
            a = h3.find('a')
            if not a:
                continue
            title = a.get('title', '').strip() or a.text.strip()
            href = a.get('href', '').strip()

            # Extract available colors
            available_colors = []
            for swatch in product_div.select('.wd-swatch-text'):
                available_colors.append(normalize(swatch.get_text(strip=True)))
            
            # Find the price
            price = ''
            price_span = product_div.select_one('.price .woocommerce-Price-amount')
            if price_span:
                price = price_span.get_text(strip=True)

            if title and href:
                # Remove 'تومان' and convert Persian digits to English
                if price:
                    price = price.replace('تومان', '').strip()
                    price = persian_to_english_digits(price)
                products.append({'title': title, 'url': href, 'cat_price': price, 'colors': available_colors})
        return products
    except Exception as e:
        print(f"[DEBUG] Error searching surfacekar.com: {e}")
        return []

def normalize(text):
    text = str(text).lower()
    text = text.replace('پلاتینیوم', 'platinum')
    text = text.replace('مشکی', 'black')
    text = text.replace('graphite', 'black')
    return re.sub(r'[^a-zA-Z0-9]', '', text)

def best_match(product_name, features, color, products):
    search_terms = [normalize(product_name)] + [normalize(f) for f in features if f]
    normalized_color = normalize(color) if color else None
    print(f"    [DEBUG] Normalized search terms for full match: {search_terms}")
    if normalized_color:
        print(f"    [DEBUG] Normalized color for match: {normalized_color}")

    for prod in products:
        title = normalize(prod['title'])
        print(f"    [DEBUG] Normalized product title: {title}")
        
        # Check for feature match
        feature_match = all(term in title for term in search_terms)
        
        # Check for color match
        color_match = (normalized_color is None) or (normalized_color in prod['colors'])
        
        if feature_match and color_match:
            print(f"    [DEBUG] FULL MATCH FOUND (Features & Color): {prod['title']}")
            return prod
            
    print(f"    [DEBUG] No full match found for search terms and color.")
    print("    [DEBUG] Candidate product titles:")
    for prod in products:
        print(f"      - {prod['title']} (Colors: {prod['colors']})")
    return None

def persian_to_english_digits(text):
    persian_digits = '۰۱۲۳۴۵۶۷۸۹'
    english_digits = '0123456789'
    for p, e in zip(persian_digits, english_digits):
        text = text.replace(p, e)
    return text

# --- Main script ---

df = pd.read_excel('SampleSites.xlsx')
if 'surfacekar.com' not in df.columns:
    df['surfacekar.com'] = ''
df['surfacekar.com'] = df['surfacekar.com'].astype('object')
if 'surfacekarproducturl' not in df.columns:
    idx = list(df.columns).index('surfacekar.com') + 1
    df.insert(idx, 'surfacekarproducturl', '')
df['surfacekarproducturl'] = df['surfacekarproducturl'].astype('object')

for idx, row in df.iterrows():
    product_name = row['Product name']
    color = row.get('Color') 
    features = [row['Cpu'], row['Ram'], 'SSD']
    print(f"Processing row {idx+1} for surfacekar.com: {product_name}, features: {features}, color: {color}")
    products = search_surfacekar_url(product_name, features, color)
    match = best_match(product_name, features, color, products)
    if match:
        print(f"  [DEBUG] Matched product: {match['title']} ({match['url']})")
        price = match['cat_price']
        print(f"  [DEBUG] Category page price used: {price}")
        df.at[idx, 'surfacekarproducturl'] = match['url']
    else:
        print("  [DEBUG] No matching product found.")
        price = ''
        df.at[idx, 'surfacekarproducturl'] = ''
    print(f"  -> Price found: {price}")
    df.at[idx, 'surfacekar.com'] = price
    time.sleep(1)

df.to_excel('SampleSites.xlsx', index=False)
print('Done. Prices and product URLs updated in SampleSites.xlsx.')
