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


# Synonym dictionaries for robust matching
CPU_SYNONYMS = {
    'ultra7': ['ultra7', 'coreultra7', 'intelultra7', 'ultra 7'],
    'ultra5': ['ultra5', 'coreultra5', 'intelultra5', 'ultra 5'],
    'xplus': ['xplus', 'snapdragonxplus', 'x plus', 'snapdragon x plus'],
    'xelite': ['xelite', 'snapdragonxelite', 'x elite', 'snapdragon x elite'],
    'n200': ['n200', 'intel n200'],
    'i7': ['i7', 'intel i7'],
    'i5': ['i5', 'intel i5'],
    'sq1': ['sq1', 'sq 1'],
}
RAM_SYNONYMS = {
    '8gb': ['8gb', '8', '8 گیگ', '8g'],
    '16gb': ['16gb', '16', '16 گیگ', '16g'],
    '32gb': ['32gb', '32', '32 گیگ', '32g'],
    '64gb': ['64gb', '64', '64 گیگ', '64g'],
}
SSD_SYNONYMS = {
    '128gb': ['128gb', '128', '128 گیگ', '128ssd', '128g'],
    '256gb': ['256gb', '256', '256 گیگ', '256ssd', '256g'],
    '512gb': ['512gb', '512', '512 گیگ', '512ssd', '512g'],
    '1tb': ['1tb', '1t', '1000gb', '1 ترابایت', '1tbssd', '1tb ssd', '1tb', '1000g'],
}
COLOR_SYNONYMS = {
    'platinum': ['platinum', 'پلاتینیوم', 'platinum'],
    'black': ['black', 'مشکی', 'black'],
}

def normalize(text):
    text = str(text).lower()
    text = persian_to_english_digits(text)
    text = re.sub(r'[^a-zA-Z0-9آ-ی]', '', text)
    return text

def synonym_match(value, synonyms_dict):
    norm = normalize(value)
    for key, syns in synonyms_dict.items():
        for s in syns:
            if norm == normalize(s):
                return key
    return norm


def best_match(product_name, features, products):
    # features: [Cpu, Ram, SSD, Color]
    cpu = synonym_match(features[0], CPU_SYNONYMS) if features[0] else ''
    ram = synonym_match(features[1], RAM_SYNONYMS) if features[1] else ''
    ssd = synonym_match(features[2], SSD_SYNONYMS) if features[2] else ''
    color = synonym_match(features[3], COLOR_SYNONYMS) if len(features) > 3 and features[3] else ''
    prod_name_norm = normalize(product_name)
    print(f"    [DEBUG] Normalized search terms: product={prod_name_norm}, cpu={cpu}, ram={ram}, ssd={ssd}, color={color}")
    def tokenize(text):
        # Split by non-alphanumeric, keep Persian letters
        return re.findall(r'[a-zA-Zآ-ی0-9]+', text)

    for prod in products:
        title = normalize(prod['name'])
        url = normalize(prod['url'])
        title_tokens = tokenize(title)
        failed = []
        # Check if each feature is present as a token or substring in any token
        if not any(prod_name_norm in t for t in title_tokens):
            failed.append('product_name')
        if not any(cpu in t for t in title_tokens):
            failed.append('cpu')
        if not any(ram in t for t in title_tokens):
            failed.append('ram')
        if not any(ssd in t for t in title_tokens):
            failed.append('ssd')
        # Color is optional if all other features match
        if not failed:
            if color and (not any(color in t for t in title_tokens) and color not in url):
                print(f"    [DEBUG] MATCH FOUND (all features except color): {prod['name']}")
                return prod
            else:
                print(f"    [DEBUG] MATCH FOUND: {prod['name']}")
                return prod
        else:
            print(f"    [DEBUG] No match for {prod['name']}: missing {failed}")
    print(f"    [DEBUG] No strong match for search terms: product={prod_name_norm}, cpu={cpu}, ram={ram}, ssd={ssd}, color={color}")
    return None

def scrape_all_products_from_micropple(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; MicroppleScraper/1.0)"
    }
    products = []
    for page in range(1, 3):  # Only first two pages
        page_url = url if page == 1 else f"{url.rstrip('/')}/page/{page}/"
        print(f"[DEBUG] Scraping page {page}: {page_url}")
        try:
            resp = requests.get(page_url, headers=headers, timeout=20)
            resp.raise_for_status()
        except Exception as e:
            print(f"[DEBUG] Error scraping {page_url}: {e}")
            break
        soup = BeautifulSoup(resp.text, 'html.parser')
        product_divs = soup.find_all('div', class_='product-grid-item')
        if not product_divs:
            print(f"[DEBUG] No products found on page {page}.")
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