#!git clone https://github.com/shatinz/scrap.git
import pandas as pd
import requests
from bs4 import BeautifulSoup
import urllib.parse
import re
import json
# completely working . but cant switch btw colors .

df_sites = pd.read_excel('SampleSites.xlsx')

url = "https://micropple.ir/product-category/microsoft/tablet-microsoft/page/1/"
response = requests.get(url)
soup = BeautifulSoup(response.content, 'html.parser')

products = soup.find_all('div', class_='wd-product')

scraped_products = []
for product in products:
    product_link = product.find('a', class_='product-image-link')
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
            response = requests.get(url)
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

def extract_color_from_variations(variations_json):
    if variations_json:
        try:
            variations = json.loads(variations_json)
            if variations:
                return variations[0]['attributes'].get('attribute_pa_color')
        except (json.JSONDecodeError, IndexError, KeyError) as e:
            print(f"Error parsing variations: {e}")
            return None
    return None

def decode_and_normalize_color(color):
    if color is None:
        return None
    
    decoded_color = urllib.parse.unquote(color)
    
    color_map = {
        'مشکی': 'black',
        'پلاتینیوم': 'platinum',
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

    for _, scraped_row in decoded_df.iterrows():
        decoded_name = scraped_row['decoded_name'].lower()
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
        scraped_color = extract_color_from_variations(variations_json)
        normalized_scraped_color = decode_and_normalize_color(scraped_color)


        if excel_numbers.issubset(scraped_numbers) and cpu_xlsx in scraped_cpu and color_xlsx == normalized_scraped_color:
            return scraped_row['price'], scraped_row['product_url']

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