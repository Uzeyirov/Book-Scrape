import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

#Define URL (Book)

url="https://books.toscrape.com/"

# Request page
headers = {"User-Agent": "Mozilla/5.0"}
response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.text, "html.parser")

# Prepare excel
wb = Workbook()
wb = wb.active
ws.title = "Book Scrape"
ws.append(["Title", "Price", "Attributes"])


# Scrape listing
bookscrape=soup.find_all("div", class_="product_pod")

for product in bookscrape:
    title_tag = product.find("div", class_="product_pod__name")
    price_tag = product.find("div", class_="product_pod-price")
    attribute_tag = product.find("div", class_="product_pod__attributes")

    title=title_tag.get_text(strip=True) if title_tag else ""
    price=price_tag.get_text(strip=True) if price_tag else ""
    attribute=attribute_tag.get_text(strip=True) if attribute_tag else ""

    if title and price and attribute:
        ws.append([title, price, attribute])

    # Save Excel
wb.save("Bookscrape.xlsx")
print("Done: book_scrape.xlsx")

    


