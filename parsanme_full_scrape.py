import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import re

CATEGORY_URLS = [
    "https://parsanme.com/store/microsoft-surface",
    "https://parsanme.com/store/surface-pro"
]

# color is not a feature for parsanme.com

def get_all_products_from_category(category_url):
    products = []
    page = 1
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    }
    while True:
        url = category_url if page == 1 else f"{category_url}?page={page}"
        print(f"[DEBUG] Scraping: {url}")
        try:
            resp = requests.get(url, timeout=20, headers=headers)
            soup = BeautifulSoup(resp.text, 'html.parser')
            product_links = soup.select('a.title.ellipsis-2')
            if not product_links:
                print(f"[DEBUG] No product links found on {url}")
                break
            for a in product_links:
                title = a.get('title', '').strip() or a.text.strip()
                href = a.get('href', '').strip()
                if href and not href.startswith('http'):
                    href = 'https://parsanme.com' + href
                # Find the closest parent with class 'container'
                container = a.find_parent('div', class_='container')
                price = ''
                if container:
                    price_tag = container.select_one('div.price-container strong.price')
                    if price_tag:
                        price = price_tag.get_text(strip=True)
                if title and href:
                    products.append({'title': title, 'url': href, 'cat_price': price})
        except Exception as e:
            print(f"[DEBUG] Error scraping {url}: {e}")
            break
        # Pagination: check for next page (adjust if needed)
        next_page = soup.find('a', {'aria-label': 'Next'})
        if not next_page or len(product_links) < 10:
            break
        page += 1
        time.sleep(1)
    return products

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

# --- Main script ---

df = pd.read_excel('SampleSites.xlsx')
if 'parsanme.com' not in df.columns:
    df['parsanme.com'] = ''
df['parsanme.com'] = df['parsanme.com'].astype('object')
if 'parsanmeproducturl' not in df.columns:
    idx = list(df.columns).index('parsanme.com') + 1
    df.insert(idx, 'parsanmeproducturl', '')
df['parsanmeproducturl'] = df['parsanmeproducturl'].astype('object')

all_products = []
for category_url in CATEGORY_URLS:
    all_products.extend(get_all_products_from_category(category_url))

# Print all scraped products
print("\n[ALL SCRAPED PRODUCTS]")
for prod in all_products:
    print(f"- {prod['title']} | {prod['url']}")

for idx, row in df.iterrows():
    product_name = row['Product name']
    # Remove 'Color' from features for parsanme.com
    features = [row['Cpu'], row['Ram'], row['SSD']]
    print(f"Processing row {idx+1} for parsanme.com: {product_name}, features: {features}")
    match = best_match(product_name, features, all_products)
    if match:
        print(f"  [DEBUG] Matched product: {match['title']} ({match['url']})")
        price = match['cat_price']
        print(f"  [DEBUG] Category page price used: {price}")
        df.at[idx, 'parsanmeproducturl'] = match['url']
    else:
        print("  [DEBUG] No matching product found.")
        price = ''
        df.at[idx, 'parsanmeproducturl'] = ''
    print(f"  -> Price found: {price}")
    df.at[idx, 'parsanme.com'] = price
    time.sleep(1)

df.to_excel('SampleSites.xlsx', index=False)
print('Done. Prices and product URLs updated in SampleSites.xlsx.') 