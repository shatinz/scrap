import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import re
from selenium import webdriver
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.common.by import By

#working but cant switch btw colors

# Dictionaries for feature translation
EN_TO_FA_CPU = {
    "ultra7": "Ultra 7",
    "ultra5": "Ultra 5",
    "i7": "i7",
    "i5": "i5",
    # Add more as needed
}

EN_TO_FA_RAM = {
    "32gb": "32 گیگابایت",
    "16gb": "16 گیگابایت",
    "8gb": "8 گیگابایت",
    # Add more as needed
}

EN_TO_FA_SSD = {
    "512gb": "512 گیگابایت",
    "256gb": "256 گیگابایت",
    "1tb": "1 ترابایت",
    # Add more as needed
}

def get_persian_feature(feature, feature_type):
    if feature_type == "cpu":
        return EN_TO_FA_CPU.get(feature.lower(), feature)
    if feature_type == "ram":
        return EN_TO_FA_RAM.get(feature.lower(), feature)
    if feature_type == "ssd":
        return EN_TO_FA_SSD.get(feature.lower(), feature)
    return feature

def normalize(text):
    return re.sub(r'[^a-zA-Z0-9آ-ی]', '', str(text).lower())

def best_match(product_name, features, products):
    # features: [Cpu, Ram, SSD]
    cpu_en = features[0] if len(features) > 0 else ""
    cpu_fa = get_persian_feature(cpu_en, "cpu")
    ram_en = features[1] if len(features) > 1 else ""
    ram_fa = get_persian_feature(ram_en, "ram")
    ssd_en = features[2] if len(features) > 2 else ""
    ssd_fa = get_persian_feature(ssd_en, "ssd")

    # Normalize all forms
    search_terms = [
        [normalize(product_name)],  # Product name (usually English)
        [normalize(cpu_en), normalize(cpu_fa)],
        [normalize(ram_en), normalize(ram_fa)],
        [normalize(ssd_en), normalize(ssd_fa)],
    ]

    for prod in products:
        title = normalize(prod['title'])
        # For each feature, at least one form must be in the title
        if all(any(term and term in title for term in term_group) for term_group in search_terms):
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

def get_product_colors_raayaatech_selenium(url):
    """Extract all color options from the product page, robust to different structures."""
    try:
        resp = requests.get(url, timeout=20)
        soup = BeautifulSoup(resp.text, 'html.parser')
        color_list = []

        # Find all divs with class 'selector-variant form-group'
        for div in soup.find_all('div', class_='selector-variant'):
            label = div.find('label')
            if label and ('رنگ' in label.get_text() or 'color' in label.get_text().lower()):
                select = div.find('select')
                if select:
                    for option in select.find_all('option'):
                        color_list.append(option.get_text(strip=True))
        return color_list
    except Exception as e:
        print(f"[DEBUG] Error fetching colors from {url}: {e}")
    return []

# English to Persian color names
EN_TO_FA_COLOR = {
    "platinum": "پلاتینیوم",
    "black": "مشکی",
    "purple": "بنفش",
    "ocean": "آبی",
    "gold": "طلایی",
    "silver": "نقره ای",
    "blue": "آبی",
    # Add more mappings as needed
}
# Persian to English color names (reverse mapping)
FA_TO_EN_COLOR = {v: k for k, v in EN_TO_FA_COLOR.items()}

def get_persian_color_name(english_color):
    key = english_color.strip().lower()
    return EN_TO_FA_COLOR.get(key, english_color)

def get_english_color_name(persian_color):
    key = persian_color.strip().lower()
    return FA_TO_EN_COLOR.get(key, persian_color)

def normalize_color(text):
    # Try to convert English color to Persian if possible
    text = get_persian_color_name(text)
    persian_digits = '۰۱۲۳۴۵۶۷۸۹'
    english_digits = '0123456789'
    for p, e in zip(persian_digits, english_digits):
        text = text.replace(p, e)
    return re.sub(r'[^a-zA-Z0-9آ-ی]', '', str(text).strip().lower())

def get_product_colors_raayaatech_selenium(url):
    """Extract all color options from the product page using Selenium and Firefox."""
    try:
        options = FirefoxOptions()
        options.add_argument('--headless')
        driver = webdriver.Firefox(options=options)
        driver.get(url)
        color_list = []

        # Find all select elements
        selects = driver.find_elements(By.TAG_NAME, 'select')
        for select in selects:
            # Try to find the label for this select
            label_text = ""
            select_id = select.get_attribute('id')
            if select_id:
                labels = driver.find_elements(By.XPATH, f"//label[@for='{select_id}']")
                if labels:
                    label_text = labels[0].text
            # Or check parent label
            parent = select.find_element(By.XPATH, '..')
            if parent.tag_name == 'label':
                label_text = parent.text
            if 'رنگ' in label_text or 'color' in label_text.lower():
                options_elems = select.find_elements(By.TAG_NAME, 'option')
                for option in options_elems:
                    color_list.append(option.text.strip())
        driver.quit()
        return color_list
    except Exception as e:
        print(f"[DEBUG] Error fetching colors from {url} with Selenium/Firefox: {e}")
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
    color = str(row['Color']).strip() if 'Color' in row and pd.notna(row['Color']) else ''
    print(f"Processing row {idx+1} for raayaatech.com: {product_name}, features: {features}, color: {color}")
    products = search_raayaatech(product_name)
    match = best_match(product_name, features, products)
    if match and color:
        product_colors = get_product_colors_raayaatech_selenium(match['url'])
        print(f"  [DEBUG] Product page colors: {product_colors}")
        color_norm = normalize_color(color)
        # Normalize both Persian and English for each product color
        product_colors_norm = []
        for c in product_colors:
            product_colors_norm.append(normalize_color(c))
            # Also add the English version if available
            en_c = get_english_color_name(c)
            if en_c != c:
                product_colors_norm.append(normalize_color(en_c))
        if any(color_norm == pc for pc in product_colors_norm):
            print(f"  [DEBUG] Matched product: {match['title']} ({match['url']})")
            price = match['cat_price']
            print(f"  [DEBUG] Category page price used: {price}")
            df.at[idx, 'raayaatechproducturl'] = match['url']
            df.at[idx, 'raayaatech.com'] = price
        else:
            print("  [DEBUG] Color does not match any product color. Not inserting product.")
            df.at[idx, 'raayaatechproducturl'] = ''
            df.at[idx, 'raayaatech.com'] = ''
    elif match:
        # No color specified, insert as before
        print(f"  [DEBUG] Matched product: {match['title']} ({match['url']})")
        price = match['cat_price']
        print(f"  [DEBUG] Category page price used: {price}")
        df.at[idx, 'raayaatechproducturl'] = match['url']
        df.at[idx, 'raayaatech.com'] = price
    else:
        print("  [DEBUG] No matching product found.")
        df.at[idx, 'raayaatechproducturl'] = ''
        df.at[idx, 'raayaatech.com'] = ''
    time.sleep(1)

df.to_excel('SampleSites.xlsx', index=False)
print('Done. Prices and product URLs updated in SampleSites.xlsx.')