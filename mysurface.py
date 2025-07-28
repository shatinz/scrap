#working
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
    "platinum": "پلاتینیوم",
    "black": "مشکی",
    "blue": "آبی",
    "gold": "طلایی",
    "silver": "نقره ای",
    # Add more mappings as needed
}
FA_TO_EN_COLOR = {v: k for k, v in EN_TO_FA_COLOR.items()}

def get_persian_product_name(english_name):
    key = english_name.strip().lower()
    return EN_TO_FA_PRODUCT.get(key, english_name)

def get_persian_color_name(color):
    return EN_TO_FA_COLOR.get(color.strip().lower(), color)

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

def best_match(product_name, features, products):
    # Only match if all normalized search terms (including numbers for RAM/SSD) are present
    cpu = features[0] if len(features) > 0 else ''
    ram_num = re.sub(r'[^0-9]', '', persian_to_english_digits(str(features[1]))) if len(features) > 1 else ''
    ssd_num = re.sub(r'[^0-9]', '', persian_to_english_digits(str(features[2]))) if len(features) > 2 else ''
    search_terms = [normalize(product_name), normalize(cpu)]
    if ram_num:
        search_terms.append(ram_num)
    if ssd_num:
        search_terms.append(ssd_num)
    for prod in products:
        title = normalize(prod['name'])
        if all(term in title for term in search_terms if term):
            return prod
    return None

if __name__ == "__main__":
    BASE_URL = "https://mysurface.ir/surface-pro/"
    products = scrape_all_products_mysurface(BASE_URL)
    print("[ALL SCRAPED PRODUCTS]")
    for prod in products:
        print(f"- {prod['name']} | {prod['url']} | {prod['price']}")

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
        features = [row['Cpu'], row['Ram'], row['SSD']]
        match = best_match(product_name, features, products)
        if match:
            df.at[idx, 'mysurfaceproducturl'] = match['url']
            df.at[idx, 'mysurface.ir'] = match['price']
        else:
            df.at[idx, 'mysurfaceproducturl'] = ''
            df.at[idx, 'mysurface.ir'] = ''
    df.to_excel('SampleSites.xlsx', index=False)
    print('Done. Prices and product URLs updated in SampleSites.xlsx.')
