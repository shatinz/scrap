import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import urllib.parse
import re
from selenium import webdriver
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.common.by import By

#solved
def search_yasinrayan_url(product_name, features):
    search_query = product_name + ' ' + ' '.join(str(f) for f in features if f)
    search_url = f'https://www.yasinrayan.com/?s={urllib.parse.quote(search_query)}&post_type=product'
    # print(f"[DEBUG] yasinrayan.com search URL: {search_url}")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    }
    try:
        resp = requests.get(search_url, timeout=20, headers=headers)
        soup = BeautifulSoup(resp.text, 'html.parser')
        products = []
        for h3 in soup.select('h3.wd-entities-title'):
            a = h3.find('a')
            if not a:
                continue
            title = a.get('title', '').strip() or a.text.strip()
            href = a.get('href', '').strip()
            # Find the price in the next siblings
            price = ''
            price_span = None
            for sib in h3.find_all_next(['span', 'p'], limit=5):
                if 'price' in sib.get('class', []):
                    price_span = sib.find('span', class_='woocommerce-Price-amount')
                    if price_span:
                        price = price_span.get_text(strip=True)
                        break
            if title and href:
                products.append({'title': title, 'url': href, 'cat_price': price})
        return products
    except Exception as e:
        # print(f"[DEBUG] Error searching yasinrayan.com: {e}")
        return []

def normalize(text):
    return re.sub(r'[^a-zA-Z0-9آ-ی]', '', str(text).lower())

def persian_to_english_digits(persian_string):
    """Converts a string with Persian digits to English digits."""
    persian_to_english_map = {
        '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
        '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9'
    }
    # Also handle Arabic digits if they appear
    persian_to_english_map.update({
        '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4',
        '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9'
    })
    # Remove non-digit characters like commas
    persian_string = re.sub(r'[,]', '', persian_string)
    translation_table = str.maketrans(persian_to_english_map)
    return persian_string.translate(translation_table)

def get_strong_matches(product_name, features, products):
    """Finds all products that fully match the search terms (product name + features)."""
    search_terms = [normalize(product_name)] + [normalize(f) for f in features if f]
    full_matches = []
    for prod in products:
        title = normalize(prod['title'])
        # Check if all search terms are in the title
        if all(term in title for term in search_terms):
            full_matches.append(prod)

    if not full_matches:
        # print(f"    [DEBUG] No full matches found for search terms: {search_terms}")
        # print("    [DEBUG] Candidate product titles:")
        # for prod in products:
            # print(f"      - {prod['title']}")
        return []

    return full_matches

def get_available_colors(product_url):
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; YasinrayanScraper/1.0)"
    }
    try:
        resp = requests.get(product_url, headers=headers, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'html.parser')
        color_divs = soup.select('div.wd-swatches-product .wd-swatch')
        colors = []
        for div in color_divs:
            color_name = div.get('data-title') or div.get('title') or ''
            if not color_name:
                text_span = div.find('span', class_='wd-swatch-text')
                if text_span:
                    color_name = text_span.get_text(strip=True)
            classes = div.get('class', [])
            enabled = 'wd-enabled' in classes and 'wd-disabled' not in classes
            colors.append((color_name.strip(), enabled))
        return colors
    except Exception as e:
        # print(f"[DEBUG] Error fetching colors from {product_url}: {e}")
        return []

def get_available_colors_selenium(product_url):
    """Extract all available colors and their enabled/disabled status using Selenium."""
    try:
        options = FirefoxOptions()
        options.add_argument('--headless')
        driver = webdriver.Firefox(options=options)
        driver.get(product_url)
        colors = []
        # Wait for swatches to load (optional: add WebDriverWait for robustness)
        swatches = driver.find_elements(By.CSS_SELECTOR, 'div.wd-swatches-product .wd-swatch')
        for swatch in swatches:
            color_name = swatch.get_attribute('data-title') or swatch.get_attribute('title') or ''
            if not color_name:
                text_spans = swatch.find_elements(By.CSS_SELECTOR, 'span.wd-swatch-text')
                if text_spans:
                    color_name = text_spans[0].text.strip()
            classes = swatch.get_attribute('class').split()
            enabled = 'wd-enabled' in classes and 'wd-disabled' not in classes
            colors.append((color_name.strip(), enabled))
        driver.quit()
        return colors
    except Exception as e:
        # print(f"[DEBUG] Error fetching colors from {product_url} with Selenium: {e}")
        return []

def map_color_name(color):
    color_map = {
        'platinum': 'پلاتینی',
        'graphite': 'مشکی',
        'black': 'مشکی',
        'sapphire': 'آبی (Sapphire)',
        'gold': 'شنی طلایی',
        # Add more mappings as needed
    }
    color_norm = normalize(color)
    for en, fa in color_map.items():
        if color_norm == normalize(en) or color_norm == normalize(fa):
            return fa
    return color  # fallback to original

def map_product_name(name):
    """Translates English product names to Persian for better matching."""
    name_map = {
        'surface pro 10': 'سرفیس پرو 10',
        'surface pro 11': 'سرفیس پرو 11',
        'surface laptop 6': 'سرفیس لپ تاپ 6',
    }
    # Normalize the input name for a case-insensitive lookup
    name_lower = str(name).lower().strip()
    return name_map.get(name_lower, name)

# --- Main script ---

df = pd.read_excel('SampleSites.xlsx')
if 'yasinrayan.com' not in df.columns:
    df['yasinrayan.com'] = ''
df['yasinrayan.com'] = df['yasinrayan.com'].astype('object')
if 'yasinrayanproducturl' not in df.columns:
    idx = list(df.columns).index('yasinrayan.com') + 1
    df.insert(idx, 'yasinrayanproducturl', '')
df['yasinrayanproducturl'] = df['yasinrayanproducturl'].astype('object')

for idx, row in df.iterrows():
    product_name = row['Product name']
    features = [row['Cpu'], row['Ram'], row['SSD']]
    desired_color = str(row['Color']).strip()
    
    # Map product name and color to Persian
    product_name_mapped = map_product_name(product_name)
    desired_color_mapped = map_color_name(desired_color)

    # print(f"Processing row {idx+1} for yasinrayan.com: {product_name} (mapped: {product_name_mapped}), features: {features}, color: {desired_color} (mapped: {desired_color_mapped})")
    
    # Use the original English name for the URL search query, as it's more likely to work with the site's search engine
    products = search_yasinrayan_url(product_name, features)
    # But use the mapped Persian name for title matching
    strong_matches = get_strong_matches(product_name_mapped, features, products)

    price = ''
    product_url = ''
    found_match_with_color = False

    if strong_matches:
        # print(f"  [DEBUG] Found {len(strong_matches)} strong match(es). Checking for color availability...")
        for i, match in enumerate(strong_matches):
            # print(f"    [DEBUG] Checking match {i+1}/{len(strong_matches)}: {match['title']}")
            available_colors = get_available_colors_selenium(match['url'])
            # print(f"      [DEBUG] Available colors: {available_colors}")
            
            # Check if the desired color is present and enabled
            if any(normalize(desired_color_mapped) == normalize(c[0]) and c[1] for c in available_colors):
                # print(f"      [DEBUG] Color match found and enabled: {desired_color_mapped}")
                price = match['cat_price']
                price = price.replace('تومان', '').strip()
                price = persian_to_english_digits(price)
                product_url = match['url']
                found_match_with_color = True
                break  # Stop after finding the first product with the right color
            #else:
                # print(f"      [DEBUG] Desired color '{desired_color_mapped}' not found or is disabled for this product.")
        
        #if not found_match_with_color:
             #print(f"  [DEBUG] Checked all {len(strong_matches)} strong matches, but none had the desired color '{desired_color_mapped}' available.")
            #return
    #else:
         #print("  [DEBUG] No matching product found.")
        
        #return

    df.at[idx, 'yasinrayanproducturl'] = product_url
    # print(f"  -> Price found: {price}")
    df.at[idx, 'yasinrayan.com'] = price
    time.sleep(1)

df.to_excel('SampleSites.xlsx', index=False)
# print('Done. Prices and product URLs updated in SampleSites.xlsx.')
