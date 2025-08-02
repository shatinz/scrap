import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import re
#solved. debug avainlable.

CATEGORY_URLS = [
    "https://parsanme.com/store/microsoft-surface",
    "https://parsanme.com/store/surface-pro"
]

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

# English to Persian color names
EN_TO_FA_COLOR = {
    "platinum": "پلاتینیوم",
    "black": "مشکی",
    "purple": "بنفش",
    "ocean": "اقیانوسی",
    "silver": "نقره ای",
    # Add more mappings as needed
}

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
            pass
            break
        # Pagination: check for next page (adjust if needed)
        next_page = soup.find('a', {'aria-label': 'Next'})
        if not next_page or len(product_links) < 10:
            break
        page += 1
        time.sleep(1)
    return products

def get_product_colors_and_variants(url):
    """Extract all color titles and their English keys from the product page."""
    try:
        resp = requests.get(url, timeout=20)
        soup = BeautifulSoup(resp.text, 'html.parser')
        color_list = []
        color_block = soup.select_one('div.block.group:-soup-contains("انتخاب رنگ")')
        if color_block:
            ul = color_block.find('ul')
            if ul:
                for li in ul.find_all('li'):
                    container_span = li.find('span', class_='container')
                    persian_color = container_span.get_text(strip=True) if container_span else ''
                    input_tag = li.find('input', class_='variant_item')
                    color_key = input_tag.get('value') if input_tag else ''
                    if persian_color:
                        color_list.append((persian_color, color_key))
        return color_list
    except Exception as e:
        print(f"[DEBUG] Error fetching colors from {url}: {e}")
        pass
    return []

def normalize(text):
    return re.sub(r'[^a-zA-Z0-9آ-ی]', '', str(text).lower())

def get_persian_product_name(english_name):
    key = english_name.strip().lower()
    return EN_TO_FA_PRODUCT.get(key, english_name)

def get_persian_color_name(english_color):
    key = english_color.strip().lower()
    return EN_TO_FA_COLOR.get(key, english_color)

def persian_to_english_digits(text):
    persian_digits = '۰۱۲۳۴۵۶۷۸۹'
    english_digits = '0123456789'
    for p, e in zip(persian_digits, english_digits):
        text = text.replace(p, e)
    return text

def normalize_color(text):
    # Try to convert English color to Persian if possible
    text = get_persian_color_name(text)
    text = persian_to_english_digits(str(text).strip().lower())
    text = re.sub(r'[^a-zA-Z0-9آ-ی]', '', text)
    return text

def best_match(product_name, features, products):
    # Use Persian name if available
    persian_name = get_persian_product_name(product_name)
    search_terms = [normalize(persian_name)] + [normalize(f) for f in features if f]
    for prod in products:
        title = normalize(prod['title'])
        if all(term in title for term in search_terms):
            return prod
    print(f"    [DEBUG] No strong match for search terms: {search_terms}")
    print("    [DEBUG] Candidate product titles:")
    for prod in products:
        print(f"      - {prod['title']}")
    return None

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

print("\n[ALL SCRAPED PRODUCTS]")
for prod in all_products:
    print(f"- {prod['title']} | {prod['url']}")

for idx, row in df.iterrows():
    product_name = row['Product name']
    features = [row['Cpu'], row['Ram'], row['SSD']]
    color = str(row['Color']).strip() if 'Color' in row and pd.notna(row['Color']) else ''
    print(f"Processing row {idx+1} for parsanme.com: {product_name}, features: {features}, color: {color}")
    match = best_match(product_name, features, all_products)
    if match and color:
        product_colors = get_product_colors_and_variants(match['url'])
        print(f"  [DEBUG] Product page colors: {product_colors}")
        color_norm = normalize_color(color)
        found = False
        for persian_color, color_key in product_colors:
            persian_norm = normalize_color(persian_color)
            if color_norm == persian_norm:
                # Build variant URL
                variant_url = match['url']
                if '?' in variant_url:
                    variant_url += f"&variant[color]={color_key}"
                else:
                    variant_url += f"?variant[color]={color_key}"
                print(f"  [DEBUG] Matched color: {persian_color} ({color_key}), URL: {variant_url}")
                price = match['cat_price']
                df.at[idx, 'parsanmeproducturl'] = variant_url
                df.at[idx, 'parsanme.com'] = price
                found = True
                break
        if not found:
            print("  [DEBUG] Color does not match any product color. Not inserting product.")
            df.at[idx, 'parsanmeproducturl'] = ''
            df.at[idx, 'parsanme.com'] = ''
    else:
        print("  [DEBUG] No matching product found or no color specified.")
        df.at[idx, 'parsanmeproducturl'] = ''
        df.at[idx, 'parsanme.com'] = ''
    time.sleep(1)

df.to_excel('SampleSites.xlsx', index=False)
print('Done. Prices and product URLs updated in SampleSites.xlsx.')
