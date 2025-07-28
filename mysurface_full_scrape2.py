import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import re

#still cant find matches . previoslly could .

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

# English to Persian color names
EN_TO_FA_COLOR = {
    "platinum": "پلاتینیوم",
    "black": "مشکی",
    "blue": "آبی",
    "gold": "طلایی",
    "silver": "نقره ای",
    # Add more as needed
}
FA_TO_EN_COLOR = {v: k for k, v in EN_TO_FA_COLOR.items()}

def get_persian_color_name(color):
    return EN_TO_FA_COLOR.get(color.strip().lower(), color)

def get_english_color_name(color):
    return FA_TO_EN_COLOR.get(color.strip().lower(), color)

def normalize_color(text):
    text = get_persian_color_name(text)
    persian_digits = '۰۱۲۳۴۵۶۷۸۹'
    english_digits = '0123456789'
    for p, e in zip(persian_digits, english_digits):
        text = text.replace(p, e)
    return re.sub(r'[^a-zA-Z0-9آ-ی]', '', str(text).strip().lower())

# English to Persian product names
EN_TO_FA_PRODUCT = {
    "surface pro 11": "سرفیس پرو 11",
    "surface pro 10": "سرفیس پرو 10",
    "surface pro 9": "سرفیس پرو 9",
    "surface pro 8": "سرفیس پرو 8",
    "surface laptop 7": "سرفیس لپ تاپ 7",
    "surface laptop 6": "سرفیس لپ تاپ 6",
    # Add more mappings as needed
}

def get_persian_product_name(english_name):
    key = english_name.strip().lower()
    return EN_TO_FA_PRODUCT.get(key, english_name)

EN_TO_FA_CPU = {
    "x elite": "X Elite",
    "x plus": "X Plus",
    # Add more mappings as needed
}

def get_persian_cpu(cpu):
    return EN_TO_FA_CPU.get(cpu.strip().lower(), cpu)

def best_match(product_name, features, products):
    persian_name = get_persian_product_name(product_name)
    cpu_en = features[0] if len(features) > 0 else ""
    cpu_fa = get_persian_cpu(cpu_en)
    ram_num = re.sub(r'[^0-9]', '', persian_to_english_digits(str(features[1]))) if len(features) > 1 else ''
    ssd_num = re.sub(r'[^0-9]', '', persian_to_english_digits(str(features[2]))) if len(features) > 2 else ''
    search_terms = [normalize(product_name), normalize(persian_name), normalize(cpu_en), normalize(cpu_fa)]
    print(f"    [DEBUG] Normalized search terms: {search_terms}, RAM: {ram_num}, SSD: {ssd_num}")
    best = None
    for prod in products:
        title = normalize(prod['name'])
        # Check if all search terms are present
        terms_match = all(term in title for term in search_terms if term)
        # Check if RAM and SSD numbers are present as standalone numbers
        ram_match = ram_num and re.search(rf'(?<!\d){ram_num}(?!\d)', title)
        ssd_match = ssd_num and re.search(rf'(?<!\d){ssd_num}(?!\d)', title)
        if terms_match and ram_match and ssd_match:
            best = prod
            print(f"    [DEBUG] FULL MATCH FOUND: {prod['name']}")
            break
    if not best:
        print(f"    [DEBUG] No full feature match for search terms: {search_terms}, RAM: {ram_num}, SSD: {ssd_num}")
        print("    [DEBUG] Candidate product titles:")
        for prod in products:
            print(f"      - {prod['name']}")
    return best

def scrape_all_products_from_mysurface(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; MysurfaceScraper/1.0)"
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
        product_divs = soup.find_all('div', class_='product-small')
        if not product_divs:
            break
        for prod_div in product_divs:
            # Name
            name_tag = prod_div.find('p', class_='name')
            name = name_tag.get_text(strip=True) if name_tag else ''
            # URL
            a_tag = prod_div.find('a', class_='woocommerce-LoopProduct-link', href=True)
            url_full = a_tag['href'] if a_tag else ''
            # Price (prefer <ins> if present, else <del> or just <span class='woocommerce-Price-amount'>)
            price = ''
            price_tag = prod_div.find('ins')
            if price_tag:
                price_amt = price_tag.find('span', class_='woocommerce-Price-amount')
                price = price_amt.get_text(strip=True) if price_amt else ''
            if not price:
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

CATEGORY_URL = 'https://mysurface.ir/surface-pro/'

all_products = scrape_all_products_from_mysurface(CATEGORY_URL)

print("[ALL SCRAPED PRODUCTS]")
for prod in all_products:
    print(f"  - Name: {prod['name']}")
    print(f"    URL: {prod['url']}")
    print(f"    Price: {prod['price']}")



# Load Excel and match

df = pd.read_excel('SampleSites.xlsx')
if 'mysurface.ir' not in df.columns:
    df['mysurface.ir'] = ''
df['mysurface.ir'] = df['mysurface.ir'].astype('object')
if 'mysurfaceproducturl' not in df.columns:
    idx = list(df.columns).index('mysurface.ir') + 1
    df.insert(idx, 'mysurfaceproducturl', '')
df['mysurfaceproducturl'] = df['mysurfaceproducturl'].astype('object')

for idx, row in df.iterrows():
    product_name = row['Product name']
    # Only use Cpu, Ram, SSD for matching (remove Color)
    features = [row['Cpu'], row['Ram'], row['SSD']]
    print(f"Processing row {idx+1} for mysurface.ir: {product_name}, features: {features}")
    match = best_match(product_name, features, all_products)
    if match:
        print(f"  [DEBUG] Matched product: {match['name']}")
        url = match['url']
        price = match['price']
        df.at[idx, 'mysurfaceproducturl'] = url
    else:
        print("  [DEBUG] No matching product found.")
        price = ''
        df.at[idx, 'mysurfaceproducturl'] = ''
    print(f"  -> Price found: {price}")
    df.at[idx, 'mysurface.ir'] = price
    time.sleep(1)


df.to_excel('SampleSites.xlsx', index=False)
print('Done. Prices and product URLs updated in SampleSites.xlsx.')