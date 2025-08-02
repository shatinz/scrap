#solved . debug activated
EN_TO_FA_PRODUCT = {
    "surface pro 11": "سرفیس پرو 11",
    "surface pro 10": "سرفیس پرو 10",
    "surface pro 9": "سرفیس پرو 9",
    "surface pro 8": "سرفیس پرو 8",
    "surface laptop 7": "سرفیس لپ تاپ 7",
    "surface laptop 6": "سرفیس لپ تاپ 6",
    # Add more mappings as needed
}

EN_TO_FA_COLOR = {
    "platinum": ["پلاتینیوم", "پلاتینی"],
    "black": ["مشکی"],
    "blue": ["آبی"],
    "gold": ["طلایی"],
    "silver": ["نقره ای"],
    # Add more mappings as needed
}
FA_TO_EN_COLOR = {v: k for k, v_list in EN_TO_FA_COLOR.items() for v in v_list}

def get_persian_product_name(english_name):
    key = english_name.strip().lower()
    return EN_TO_FA_PRODUCT.get(key, english_name)

def get_persian_color_name(color):
    return EN_TO_FA_COLOR.get(color.strip().lower(), [color])

def get_english_color_name(color):
    return FA_TO_EN_COLOR.get(color.strip().lower(), color)
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import time

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

def scrape_all_products_mysurface(base_url):
    products = []
    url = f"{base_url.rstrip('/')}/page/1/"
    print(f"[DEBUG] Scraping: {url}")
    try:
        resp = requests.get(url, timeout=20)
        resp.raise_for_status()
    except Exception as e:
        print(f"[DEBUG] Error scraping {url}: {e}")
        return products
    soup = BeautifulSoup(resp.text, 'html.parser')
    product_divs = soup.find_all('div', class_='product-small')
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
        # Clean price
        price = price.replace('تومان', '').replace(',', '').strip()
        price = persian_to_english_digits(price)
        products.append({'name': name, 'price': price, 'url': url_full})
    return products

def get_color_from_product_page(url):
    try:
        resp = requests.get(url, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'html.parser')
        description_div = soup.find('div', class_='product-short-description')
        if description_div:
            list_items = description_div.find_all('li')
            for item in list_items:
                text = item.get_text()
                if 'رنگ:' in text:
                    color = text.replace('رنگ:', '').strip()
                    return normalize(color)
    except Exception as e:
        print(f"[DEBUG] Error scraping product page {url}: {e}")
    return None

def best_match(product_name, features, products):
    # Only match if all normalized search terms (including numbers for RAM/SSD) are present
    cpu = features[0] if len(features) > 0 else ''
    ram_num = re.sub(r'[^0-9]', '', persian_to_english_digits(str(features[1]))) if len(features) > 1 else ''
    ssd_num = re.sub(r'[^0-9]', '', persian_to_english_digits(str(features[2]))) if len(features) > 2 else ''
    color = features[3] if len(features) > 3 else ''

    # English search terms
    english_product_name = product_name
    normalized_cpu = normalize(cpu).replace('ultra', 'coreultra')
    search_terms_en = [normalize(english_product_name), normalized_cpu]
    if ram_num:
        search_terms_en.append(ram_num)
    if ssd_num:
        search_terms_en.append(ssd_num)
    search_terms_en = [term for term in search_terms_en if term]

    # Persian search terms
    persian_product_name = get_persian_product_name(product_name)
    search_terms_fa = [normalize(persian_product_name), normalized_cpu]
    if ram_num:
        search_terms_fa.append(ram_num)
    if ssd_num:
        search_terms_fa.append(ssd_num)
    search_terms_fa = [term for term in search_terms_fa if term]

    print(f"\n[DEBUG] Searching for: {product_name} | CPU: {cpu} | RAM: {ram_num} | SSD: {ssd_num} | Color: {color}")
    print(f"[DEBUG] English search terms: {search_terms_en}")
    print(f"[DEBUG] Persian search terms: {search_terms_fa}")

    for prod in products:
        title = normalize(prod['name'])
        print(f"[DEBUG]   Comparing with: {prod['name']} -> {title}")
        
        # Try matching with English terms first, then Persian
        if all(term in title for term in search_terms_en) or all(term in title for term in search_terms_fa):
            print(f"[DEBUG]   Potential match found: {prod['name']}. Checking color...")
            page_color = get_color_from_product_page(prod['url'])
            print(f"[DEBUG]     Page color: {page_color}")
            persian_colors = get_persian_color_name(color)
            if page_color and (normalize(color) in page_color or page_color in normalize(color) or any(normalize(c) in page_color for c in persian_colors) or any(page_color in normalize(c) for c in persian_colors)):
                print(f"[DEBUG]   COLOR MATCH FOUND: {prod['name']}")
                return prod
            else:
                print(f"[DEBUG]     Color mismatch.")

    print("[DEBUG]   NO MATCH FOUND")
    return None

if __name__ == "__main__":
    BASE_URL = "https://mysurface.ir/surface-pro/"
    products = scrape_all_products_mysurface(BASE_URL)
    print("\n[ALL SCRAPED PRODUCTS]")
    for prod in products:
        print(f"- {prod['name']} | {prod['url']} | {prod['price']}")
    print("-" * 20)

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
        features = [row['Cpu'], row['Ram'], row['SSD'], row['Color']]
        
        print(f"\n{'='*20}\n[INFO] Matching product from Excel row {idx+2}: {product_name} {features}")
        
        match = best_match(product_name, features, products)
        
        if match:
            print(f"[INFO]   -> Matched to: {match['name']} | Price: {match['price']}")
            df.at[idx, 'mysurfaceproducturl'] = match['url']
            df.at[idx, 'mysurface.ir'] = match['price']
        else:
            print("[INFO]   -> No match found in scraped products.")
            df.at[idx, 'mysurfaceproducturl'] = ''
            df.at[idx, 'mysurface.ir'] = ''
            
    df.to_excel('SampleSites.xlsx', index=False)
    print('\nDone. Prices and product URLs updated in SampleSites.xlsx.')
