import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import re
import json

#solved and can find matched color with matched price

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

    for prod in products:
        title = prod['title']
        normalized_title = normalize(title)

        # Exclude accessories like keyboards and pens
        if 'keyboard' in normalized_title or 'pen' in normalized_title or 'کیبورد' in normalized_title or 'قلم' in normalized_title:
            continue

        # Check for product name
        if normalize(product_name) not in normalized_title:
            continue

        # Check for all specified features
        all_features_found = True
        if cpu_en and not (normalize(cpu_en) in normalized_title or normalize(cpu_fa) in normalized_title):
            all_features_found = False
        if ram_en and not (normalize(ram_en) in normalized_title or normalize(ram_fa) in normalized_title):
            all_features_found = False
        if ssd_en and not (normalize(ssd_en) in normalized_title or normalize(ssd_fa) in normalized_title):
            all_features_found = False

        if all_features_found:
            print(f"    [DEBUG] Full match found: {title}")
            return prod

    print(f"    [DEBUG] No full match found for {product_name} with features {features}.")
    return None

def make_request_with_retries(url, headers, retries=3, delay=5):
    for i in range(retries):
        try:
            resp = requests.get(url, timeout=30, headers=headers)
            resp.raise_for_status()
            return resp
        except (requests.exceptions.RequestException, requests.exceptions.SSLError) as e:
            print(f"  [DEBUG] Request to {url} failed (attempt {i+1}/{retries}): {e}")
            if i < retries - 1:
                time.sleep(delay)
            else:
                print(f"[DEBUG] All retries failed for {url}.")
                return None

def search_raayaatech(product_name):
    search_query = product_name
    search_url = f'https://raayaatech.com/search?q={requests.utils.quote(search_query)}'
    print(f"[DEBUG] raayaatech.com search URL: {search_url}")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    }
    resp = make_request_with_retries(search_url, headers)
    if not resp:
        return []
    
    try:
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
        print(f"[DEBUG] Error parsing search results for raayaatech.com: {e}")
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

def get_product_details_raayaatech(url, target_color=None):
    """
    Extracts color options and prices from the product page using BeautifulSoup.
    It assumes variant information is stored in a JavaScript object in the page source.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    }
    resp = make_request_with_retries(url, headers)
    if not resp:
        return [], None

    try:
        soup = BeautifulSoup(resp.text, 'html.parser')
        variants = []
        variants_data = None

        # Method 1: Look for a specific script tag with a JSON data attribute (common in Shopify)
        script_tag = soup.find('script', {'type': 'application/json', 'data-product-json': True})
        if script_tag and script_tag.string:
            try:
                product_json = json.loads(script_tag.string)
                variants = product_json.get('variants', [])
            except json.JSONDecodeError:
                print(f"  [DEBUG] Failed to decode product JSON from data-product-json tag on {url}")
                variants = []

        # Method 2: If method 1 fails, look for a script tag containing variant information by regex
        if not variants:
            for script in soup.find_all('script'):
                if not script.string:
                    continue
                # More flexible regex to find a variants array
                match = re.search(r'variants\s*[:=]\s*(\[.+?\])\s*[;,]', script.string, re.DOTALL)
                if match:
                    variants_data = match.group(1)
                    try:
                        variants = json.loads(variants_data)
                        break
                    except json.JSONDecodeError:
                        print(f"  [DEBUG] Failed to decode variants JSON from regex match on {url}")
                        variants = []
        
        # If we found variant information, process it
        if variants:
            color_price_map = {}
            available_colors = []
            for variant in variants:
                color = variant.get('option1')
                price_val = variant.get('price')
                if color and price_val:
                    # Format price with commas, assuming it's in the main currency unit (Toman)
                    formatted_price = f"{int(price_val):,d}"
                    normalized_c = normalize_color(color)
                    if normalized_c not in color_price_map:
                        color_price_map[normalized_c] = formatted_price
                    if color not in available_colors:
                        available_colors.append(color)
            
            if not color_price_map:
                print(f"  [DEBUG] No color-price mapping found in variants JSON on {url}")
                return [], None

            if target_color:
                norm_target_color = normalize_color(target_color)
                if norm_target_color in color_price_map:
                    return available_colors, color_price_map[norm_target_color]
                
                en_target_color = normalize_color(get_english_color_name(target_color))
                if en_target_color in color_price_map:
                    return available_colors, color_price_map[en_target_color]

                return available_colors, None # Target color not found
            
            if available_colors:
                first_color_norm = normalize_color(available_colors[0])
                return available_colors, color_price_map.get(first_color_norm)
            
            return [], None

        # Fallback to scraping visible elements if no JSON found
        print(f"  [DEBUG] Could not find variants JSON on page {url}. Scraping visible elements as fallback.")
        options = soup.select('select#variant_id option')
        if not options:
            price_tag = soup.select_one('span.price#ProductPrice')
            price = price_tag.get_text(strip=True).split(" ")[0] if price_tag else None
            return [], price

        available_colors = [opt.text.split('/')[0].strip() for opt in options if opt.text.split('/')[0].strip()]
        price_tag = soup.select_one('span.price#ProductPrice')
        price = price_tag.get_text(strip=True).split(" ")[0] if price_tag else None

        if not available_colors:
            return [], price

        if target_color:
            norm_target_color = normalize_color(target_color)
            for color in available_colors:
                if normalize_color(color) == norm_target_color or normalize_color(get_english_color_name(color)) == norm_target_color:
                    return available_colors, price
            return available_colors, None
        
        return available_colors, price

    except Exception as e:
        print(f"[DEBUG] Error in get_product_details_raayaatech for {url}: {e}")
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
        product_colors, price = get_product_details_raayaatech(match['url'], color if color else None)
        
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
