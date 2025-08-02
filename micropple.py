#!git clone https://github.com/shatinz/scrap.git
#!git switch main
import pandas as pd
import requests
from bs4 import BeautifulSoup
import urllib.parse
import re
import json
# solved . scrap all colors but not matching color with new price needed.
print("starting")
df_sites = pd.read_excel('SampleSites.xlsx')

base_url = "https://micropple.ir/product-category/microsoft/tablet-microsoft/page/{}/"
scraped_products = []

for page_num in range(1, 3):
    url = base_url.format(page_num)
    print(f"Scraping page: {url}")
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    products = soup.find_all('div', class_='wd-product')

    for product in products:
        product_link = product.find('a', class_='product-image-link')
        if product_link and product_link.has_attr('href'):
            product_name = product_link.get('href').split('/')[-2]
            product_url = product_link.get('href')
            price_element = product.find('span', class_='price')
            price = price_element.text.strip() if price_element else 'N/A'
            scraped_products.append({'product_name': product_name, 'price': price, 'product_url': product_url})

scraped_df = pd.DataFrame(scraped_products)

decoded_products = []
for _, row in scraped_df.iterrows():
    decoded_name = urllib.parse.unquote(row['product_name'])
    decoded_products.append({'product_name': row['product_name'], 'decoded_name': decoded_name, 'price': row['price'], 'product_url': row['product_url']})

decoded_df = pd.DataFrame(decoded_products)

def fetch_product_page(url):
    if url != 'Not Found':
        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            return response.text
        except requests.exceptions.RequestException as e:
            print(f"Error fetching {url}: {e}")
            return None
    return None

def extract_product_variations(html_content):
    if html_content:
        soup = BeautifulSoup(html_content, 'html.parser')
        variations_form = soup.find('form', class_='variations_form')
        if variations_form and 'data-product_variations' in variations_form.attrs:
            return variations_form['data-product_variations']
    return None

def extract_colors_from_variations(variations_json):
    if variations_json:
        try:
            variations = json.loads(variations_json)
            colors = []
            for variation in variations:
                color = variation.get('attributes', {}).get('attribute_pa_color')
                if color:
                    colors.append(color)
            return colors
        except (json.JSONDecodeError, IndexError, KeyError) as e:
            print(f"Error parsing variations: {e}")
            return []
    return []

def decode_and_normalize_color(color):
    if color is None:
        return None
    
    decoded_color = urllib.parse.unquote(color)
    
    color_map = {
        'مشکی': 'black',
        'پلاتینیوم': 'platinum',
        'پلاتینی': 'platinum',
        'گرافیت': 'graphite',
        'نقره-ای': 'silver',
        'دون-گلد': 'platinum'
    }
    
    return color_map.get(decoded_color, decoded_color)

def find_price_and_url_v3(row):
    product_name_xlsx = str(row['Product name']).lower()
    ram_xlsx = str(row['Ram']).lower()
    ssd_xlsx = str(row['SSD']).lower()
    cpu_xlsx = str(row['Cpu']).lower().replace(' ', '').replace('ultra', 'coreultra')
    color_xlsx = str(row['Color']).lower()

    excel_numbers = set(re.findall(r'\d+', product_name_xlsx + " " + ram_xlsx + " " + ssd_xlsx))

    print("\n" + "="*50)
    print(f"Processing Excel row: {row['Product name']} | RAM: {row['Ram']} | SSD: {row['SSD']} | CPU: {row['Cpu']} | Color: {row['Color']}")
    print(f"Normalized Excel data: name='{product_name_xlsx}', ram='{ram_xlsx}', ssd='{ssd_xlsx}', cpu='{cpu_xlsx}', color='{color_xlsx}'")
    print(f"Extracted numbers from Excel row: {excel_numbers}")
    print("="*50)

    for _, scraped_row in decoded_df.iterrows():
        decoded_name = scraped_row['decoded_name'].lower()
        
        print(f"\n--- Comparing with scraped product ---")
        print(f"Name: {decoded_name}")
        print(f"URL: {scraped_row['product_url']}")

        scraped_numbers = set(re.findall(r'\d+', decoded_name))
        
        scraped_cpu = ""
        if 'coreultra' in decoded_name:
            match = re.search(r'coreultra(\d+)', decoded_name)
            if match:
                scraped_cpu = 'coreultra' + match.group(1)
        elif 'x-plus' in decoded_name:
            scraped_cpu = 'xplus'
        elif 'x-elite' in decoded_name:
            scraped_cpu = 'xelite'

        page_content = fetch_product_page(scraped_row['product_url'])
        variations_json = extract_product_variations(page_content)
        scraped_colors = extract_colors_from_variations(variations_json)
        normalized_scraped_colors = [decode_and_normalize_color(c) for c in scraped_colors]

        print(f"Scraped Details: numbers={scraped_numbers}, cpu='{scraped_cpu}', colors='{normalized_scraped_colors}' (raw: '{scraped_colors}')")

        numbers_match = excel_numbers.issubset(scraped_numbers)
        cpu_match = cpu_xlsx in scraped_cpu
        color_match = color_xlsx in normalized_scraped_colors

        print(f"Checking match: Numbers ({numbers_match}), CPU ({cpu_match}), Color ({color_match})")

        if numbers_match and cpu_match and color_match:
            print(">>> Match Found! <<<")
            return scraped_row['price'], scraped_row['product_url']
        else:
            print("--- No Match ---")

    print("\n" + "="*50)
    print(f"Finished processing Excel row: {row['Product name']}. No match found in all scraped products.")
    print("="*50)
    return 'Not Found', 'Not Found'

df_sites[['micropple.ir', 'microppleproducturl']] = df_sites.apply(find_price_and_url_v3, axis=1, result_type='expand')

def clean_price(price):
    if isinstance(price, str):
        price = price.replace('تومان', '').replace(',', '').replace('.', '').strip()
        price = re.sub(r'–.*', '', price)
        if price.isdigit():
            return int(price)
    return None

df_sites['micropple.ir'] = df_sites['micropple.ir'].apply(clean_price)

df_sites.to_excel('SampleSites.xlsx', index=False)

print(df_sites[['Product name', 'Cpu', 'Ram', 'SSD', 'Color', 'micropple.ir', 'microppleproducturl']])
