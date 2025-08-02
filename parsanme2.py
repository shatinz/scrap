import requests
from bs4 import BeautifulSoup


def get_all_products_from_category(category_url):
    products = []
    page = 1
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    }
    while True:
        url = category_url if page == 1 else f"{category_url}?page={page}"
        try:
            resp = requests.get(url, timeout=20, headers=headers)
            soup = BeautifulSoup(resp.text, 'html.parser')
            # Find all product links
            product_links = soup.select('a.title.ellipsis-2')
            # print(product_links)
            if not product_links:
                break
        except requests.exceptions.RequestException as e:
            # print(f"Error fetching {url}: {e}")
            break
