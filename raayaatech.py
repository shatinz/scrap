import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import re
from selenium import webdriver
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.common.by import By

#

# Dictionaries for feature translation
EN_TO_FA_CPU = {
    "ultra7": "Ultra 7",
    "ultra5": "Ultra 5",
    "i7": "i7",
    "i5": "i5",
    "xplus": "X Plus",
    "xelite": "X Elite",
    # Add more as needed
}

EN_TO_FA_RAM = {
    "64gb": "64 گیگابایت",
    "32gb": "32 گیگابایت",
    "16gb": "16 گیگابایت",
    "8gb": "8 گیگابایت",
    # Add more as needed
}

EN_TO_FA_SSD = {
    "1t": "1 ترابایت",
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
        (normalize(product_name), 2),  # Product name gets a higher weight
        (normalize(cpu_en), 1),
        (normalize(cpu_fa), 1),
        (normalize(ram_en), 1),
        (normalize(ram_fa), 1),
        (normalize(ssd_en), 1),
        (normalize(ssd_fa), 1),
    ]

    best_score = 0
    best_prod = None

    for prod in products:
        title = normalize(prod['title'])
        
        # Exclude accessories like keyboards and pens
        if 'keyboard' in title or 'pen' in title or 'کیبورد' in title or 'قلم' in title:
            continue

        current_score = 0
        for term, weight in search_terms:
            if term and term in title:
                current_score += weight
        
        if current_score > best_score:
            best_score = current_score
            best_prod = prod

    if best_prod:
        print(f"    [DEBUG] Best match found with score {best_score}: {best_prod['title']}")
    else:
        print(f"    [DEBUG] No suitable match found for search terms.")
        print("    [DEBUG] Candidate product titles (after filtering):")
        for prod in products:
            # Re-check filter for logging purposes
            title = normalize(prod['title'])
            if 'keyboard' not in title and 'pen' not in title and 'کیبورد' not in title and 'قلم' not in title:
                 print(f"      - {prod['title']}")

    return best_prod

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

def get_product_details_raayaatech_selenium(url, target_color=None):
    """
    Extracts color options and price from the product page using Selenium.
    If a target_color is provided, it selects the color and gets the updated price.
    Returns a tuple: (list_of_available_colors, price).
    """
    try:
        options = FirefoxOptions()
        options.add_argument('--headless')
        driver = webdriver.Firefox(options=options)
        driver.get(url)
        time.sleep(3) # Initial wait for page load

        color_list = []
        target_color_selected = False

        # Find all select elements and check if they are for color
        selects = driver.find_elements(By.TAG_NAME, 'select')
        for select in selects:
            label_text = ""
            try:
                # Logic to find the label for the select element
                select_id = select.get_attribute('id')
                if select_id:
                    labels = driver.find_elements(By.XPATH, f"//label[@for='{select_id}']")
                    if labels:
                        label_text = labels[0].text
                if not label_text:
                    # Try to find label as a parent or in a parent div
                    parent = select.find_element(By.XPATH, '..')
                    if parent.tag_name == 'label':
                        label_text = parent.text
                    else:
                        # Common pattern: <div class="..."><label>...</label><select>...</select></div>
                        grandparent = parent.find_element(By.XPATH, '..')
                        if 'selector-variant' in grandparent.get_attribute('class'):
                             label_tag = grandparent.find_element(By.TAG_NAME, 'label')
                             if label_tag:
                                 label_text = label_tag.text

            except Exception:
                pass # Ignore if label not found

            if 'رنگ' in label_text or 'color' in label_text.lower():
                options_elems = select.find_elements(By.TAG_NAME, 'option')
                for option in options_elems:
                    color_name = option.text.strip()
                    if color_name:
                        color_list.append(color_name)
                    
                    if target_color and not target_color_selected:
                        # Normalize and compare colors
                        norm_option_color = normalize_color(color_name)
                        norm_target_color = normalize_color(target_color)
                        en_option_color = normalize_color(get_english_color_name(color_name))

                        if norm_option_color == norm_target_color or en_option_color == norm_target_color:
                            print(f"  [DEBUG] Selecting color: {color_name}")
                            option.click()
                            target_color_selected = True
                            time.sleep(2) # Wait for price to update
                # Found color select, no need to check other selects
                break
        
        price = ''
        try:
            # More robust price finding
            price_selectors = [
                'div.price-area span.price',
                'span.price',
                '.product-info-price .price-new',
                '.price-box .price',
                '.price-container .price',
                '#price-old',
                '#price-new'
            ]
            for selector in price_selectors:
                try:
                    price_tag = driver.find_element(By.CSS_SELECTOR, selector)
                    if price_tag and price_tag.is_displayed():
                        price = price_tag.text.strip()
                        if price:
                            break
                except Exception:
                    continue
        except Exception as e:
            print(f"  [DEBUG] Could not find price on page {url}. Error: {e}")

        driver.quit()
        return color_list, price

    except Exception as e:
        print(f"[DEBUG] Error in get_product_details_raayaatech_selenium for {url}: {e}")
        if 'driver' in locals() and driver:
            driver.quit()
        return [], None

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
    
    if match:
        product_colors, price = get_product_details_raayaatech_selenium(match['url'], color if color else None)
        
        if color:
            print(f"  [DEBUG] Product page colors: {product_colors}")
            color_norm = normalize_color(color)
            
            product_colors_norm = []
            for c in product_colors:
                product_colors_norm.append(normalize_color(c))
                en_c = get_english_color_name(c)
                if en_c != c:
                    product_colors_norm.append(normalize_color(en_c))

            if any(color_norm == pc for pc in product_colors_norm):
                print(f"  [DEBUG] Matched product with correct color: {match['title']} ({match['url']})")
                if price:
                    print(f"  [DEBUG] Price for color '{color}' from product page: {price}")
                    df.at[idx, 'raayaatechproducturl'] = match['url']
                    df.at[idx, 'raayaatech.com'] = price
                else:
                    print(f"  [DEBUG] Color matched, but no price found on product page. Using category page price.")
                    df.at[idx, 'raayaatechproducturl'] = match['url']
                    df.at[idx, 'raayaatech.com'] = match['cat_price']
            else:
                print(f"  [DEBUG] Color '{color}' not found in available colors: {product_colors}. Not inserting product.")
                df.at[idx, 'raayaatechproducturl'] = ''
                df.at[idx, 'raayaatech.com'] = ''
        else: # No color specified
            print(f"  [DEBUG] Matched product (no color specified): {match['title']} ({match['url']})")
            if price:
                 print(f"  [DEBUG] Default price from product page: {price}")
                 df.at[idx, 'raayaatechproducturl'] = match['url']
                 df.at[idx, 'raayaatech.com'] = price
            else:
                print(f"  [DEBUG] No price on product page, using category page price: {match['cat_price']}")
                df.at[idx, 'raayaatechproducturl'] = match['url']
                df.at[idx, 'raayaatech.com'] = match['cat_price']
    else:
        # No match found
        print("  [DEBUG] No matching product found.")
        df.at[idx, 'raayaatechproducturl'] = ''
        df.at[idx, 'raayaatech.com'] = ''
    time.sleep(1)

df.to_excel('SampleSites.xlsx', index=False)
print('Done. Prices and product URLs updated in SampleSites.xlsx.')
