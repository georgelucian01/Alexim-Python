import os
import subprocess
import platform
import re
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
import openpyxl

# Function to calculate final selling price
def calculate_final_price(item_price):

    # Used excel formulas as a model, that why the c2 d2 e2

    # Removing VAT
    c2 = item_price / 1.19
    
    # Calculate D2 using the nested IF conditions
    if c2 <= 4.99:
        d2 = c2 + 1.0 * c2 + 0.19 * (c2 + c2 * 1.0)
    elif c2 <= 9.99:
        d2 = c2 + 1.0 * c2 + 0.19 * (c2 + c2 * 1.0)
    elif c2 <= 14.99:
        d2 = c2 + 0.9 * c2 + 0.19 * (c2 + c2 * 0.9)
    elif c2 <= 24.99:
        d2 = c2 + 0.9 * c2 + 0.19 * (c2 + c2 * 0.9)
    elif c2 <= 29.99:
        d2 = c2 + 0.9 * c2 + 0.19 * (c2 + c2 * 0.9)
    elif c2 <= 34.99:
        d2 = c2 + 0.8 * c2 + 0.19 * (c2 + c2 * 0.8)
    elif c2 <= 39.99:
        d2 = c2 + 0.65 * c2 + 0.19 * (c2 + c2 * 0.65)
    elif c2 <= 49.99:
        d2 = c2 + 0.6 * c2 + 0.19 * (c2 + c2 * 0.6)
    elif c2 <= 54.99:
        d2 = c2 + 0.58 * c2 + 0.19 * (c2 + c2 * 0.58)
    elif c2 <= 64.99:
        d2 = c2 + 0.5 * c2 + 0.19 * (c2 + c2 * 0.5)
    elif c2 <= 74.99:
        d2 = c2 + 0.45 * c2 + 0.19 * (c2 + c2 * 0.45)
    elif c2 <= 89.99:
        d2 = c2 + 0.43 * c2 + 0.19 * (c2 + c2 * 0.43)
    elif c2 <= 99.99:
        d2 = c2 + 0.37 * c2 + 0.19 * (c2 + c2 * 0.37)
    elif c2 <= 124.99:
        d2 = c2 + 0.32 * c2 + 0.19 * (c2 + c2 * 0.32)
    elif c2 <= 144.99:
        d2 = c2 + 0.29 * c2 + 0.19 * (c2 + c2 * 0.29)
    elif c2 <= 174.99:
        d2 = c2 + 0.28 * c2 + 0.19 * (c2 + c2 * 0.28)
    elif c2 <= 199.99:
        d2 = c2 + 0.27 * c2 + 0.19 * (c2 + c2 * 0.27)
    elif c2 <= 229.99:
        d2 = c2 + 0.26 * c2 + 0.19 * (c2 + c2 * 0.26)
    elif c2 <= 299.99:
        d2 = c2 + 0.25 * c2 + 0.19 * (c2 + c2 * 0.25)
    elif c2 <= 399.99:
        d2 = c2 + 0.24 * c2 + 0.19 * (c2 + c2 * 0.24)
    elif c2 <= 499.99:
        d2 = c2 + 0.23 * c2 + 0.19 * (c2 + c2 * 0.23)
    elif c2 <= 649.99:
        d2 = c2 + 0.22 * c2 + 0.19 * (c2 + c2 * 0.22)
    elif c2 <= 749.99:
        d2 = c2 + 0.21 * c2 + 0.19 * (c2 + c2 * 0.21)
    elif c2 <= 799.99:
        d2 = c2 + 0.20 * c2 + 0.19 * (c2 + c2 * 0.20)
    elif c2 <= 899.99:
        d2 = c2 + 0.19 * c2 + 0.19 * (c2 + c2 * 0.19)
    elif c2 <= 999.99:
        d2 = c2 + 0.18 * c2 + 0.19 * (c2 + c2 * 0.18)
    elif c2 <= 1999.99:
        d2 = c2 + 0.17 * c2 + 0.19 * (c2 + c2 * 0.17)
    elif c2 <= 4999.99:
        d2 = c2 + 0.12 * c2 + 0.19 * (c2 + c2 * 0.12)
    else:
        d2 = c2 + 0.10 * c2 + 0.19 * (c2 + c2 * 0.10)
    
    # Calculate E2 (final price)
    rounded_d2 = round(d2)
    if rounded_d2 - d2 <= 0.5:
        e2 = rounded_d2 - 0.01
    else:
        e2 = d2 + (rounded_d2 - d2 - 0.5) - 0.01
        
    return e2

# Function to scrape product information
def scrape_product_info(site, parent_category, category):
    global item_sku # to open folders
    word_url = site # website
    
    # Set up Chrome options to run headless (without opening a browser window)
    chrome_options = Options()
    chrome_options.add_argument("--headless")

    # Specify the path to the ChromeDriver executable
    chromedriver_path = "./chromedriver.exe"

    # Add the ChromeDriver path to the PATH environment variable
    os.environ["PATH"] += os.pathsep + os.path.dirname(chromedriver_path)

    # Create a new Chrome browser instance with the specified options
    driver = webdriver.Chrome(options=chrome_options)

    ### ADDING DELAY, SO THE WEBSITE DOESNT SLOW DOWN MY CONNECTION
    time.sleep(1) # 1 SECONDs

    driver.get(word_url)

    # Wait for the dynamic content to load (you may need to adjust the waiting time depending on the page)
    driver.implicitly_wait(10)

    # Get the page source
    page_source = driver.page_source

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(page_source, 'html.parser')

    
    # Getting product information
    item_name = soup.find(class_="page-heading").text.strip() if soup.find(class_="page-heading") else "N/A"
    item_sku = soup.find(itemprop="sku").text.strip() if soup.find(itemprop="sku") else "N/A"
    item_price_str = soup.find(itemprop="price").text.strip() if soup.find(itemprop="price") else "N/A"
    # Remove "RON", non-numeric characters (except decimal points), and any extra spaces from item_price_str
    item_price_str = re.sub(r'[^0-9.,]', '', item_price_str.replace('RON', ''))
    # Replace comma (,) with period (.) as the decimal separator
    item_price_str = item_price_str.replace(',', '.')
    # Convert item_price_str to float
    try:
        item_price = float(item_price_str)
    except ValueError:
        item_price = None
    item_price = calculate_final_price(item_price)
    product_description = soup.find(class_="product-description typo").text.strip() if soup.find(class_="product-description typo") else "N/A"

    ### Add extra information
    Descriere_meta = f"Cumpara {item_name} cu {item_price} RON de la AccesMag!"
    Taxa = "TVA - 19%"
    Categorie = parent_category + ">" + category
    Producator = "AleximTOP"

    # buc or set
    if "set" in item_name:
        Unitate_de_masura = "set"
    else:
        Unitate_de_masura = "buc"

    Disponibilitate = "Disponibil in 2-3 zile de la comanda"
    Stoc = 1000
    Vizibilitate = "Produsul este vizibil"

    Imagine1 = ""
    Imagine2 = ""
    Imagine3 = ""
    Imagine4 = ""

    # Find the image elements within the "product-images" class and extract the big image URLs
    image_elements = soup.select('.product-images li a.thumb')
    image_urls = [element['data-zoom-image'] for element in image_elements]

    # 4 images (maximum)
    images = ["","","",""]

    # 
    i = 0
    for image_url in image_urls:
        print(f"Image {i+1} saved")
        images[i] = image_url
        i += 1
        if i == 4: break

    Imagine1 = images[0]
    Imagine2 = images[1]
    Imagine3 = images[2]
    Imagine4 = images[3]

    ###


    ### EXCEL INFORMATION ADDING ###

    # Check if the Excel file exists
    file_path = "feed_alexim.xlsx"
    is_file_exists = os.path.exists(file_path)

    # Create or open the Excel workbook
    workbook = openpyxl.Workbook() if not is_file_exists else openpyxl.load_workbook(file_path)

    # Select the active worksheet
    worksheet = workbook.active

    # If the file is newly created, add header row with column names
    if not is_file_exists:
        worksheet.append(["Nume_Produs", "Descriere", "Descriere_meta",	"Taxa_(%)",	"Pret_Produs", "Cod_produs_(SKU)", "Categorie",	"Producator",	"Unitate_de_masura"	,"Disponibilitate",	"Stoc",	"Vizibilitate",	"Imagine1",	"Imagine2",	"Imagine3",	"Imagine4"]) 


    # Append the scraped product information to the Excel file
    worksheet.append([item_name, product_description, Descriere_meta, Taxa, item_price, item_sku, Categorie, Producator, Unitate_de_masura, Disponibilitate, Stoc, Vizibilitate, Imagine1, Imagine2, Imagine3, Imagine4])

    # Save the changes to the Excel file
    workbook.save(file_path)

    print("Product information saved in feed")


    # Close the browser
    driver.quit()

# Function to scrape information from all products in a category page
def scrape_category_page(site, chrome_options, parent_category):
    word_url = site  # category page URL

    # Create a new Chrome browser instance with the specified options
    driver = webdriver.Chrome(options=chrome_options)

    ### ADDING DELAY, SO THE WEBSITE DOESNT SLOW DOWN MY CONNECTION
    time.sleep(1)  # 3 SECONDs

    driver.get(word_url)

    # Wait for the dynamic content to load (you may need to adjust the waiting time depending on the page)
    driver.implicitly_wait(10)

    # Get the page source
    page_source = driver.page_source

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(page_source, 'html.parser')

    # Getting category

    category = soup.find(class_="page-heading").text.strip() if soup.find(class_="page-heading") else "N/A"

    # Find all article blocks with product information
    article_blocks = soup.find_all('article', class_='product-miniature')

    # Scrape information from each product article block
    i=0
    for article in article_blocks:
        i+=1
        print(f"Produsul numarul {i}")
        # Extract the product page link from the product name (h5 tag)
        product_name_tag = article.find('h5', class_='product-name')
        product_page_link = product_name_tag.a['href'] if product_name_tag and product_name_tag.a else None

        if product_page_link:
            scrape_product_info(product_page_link, parent_category, category)
            print(f"Produsul numarul {i}")

    print("Finished the whole page!")
    # Close the browser
    driver.quit()

# Main code
def main():
    # Scrape product information and save files
    # Continue until I type exit

    print("Type exit to stop the script")

    chrome_options = Options()
    chrome_options.add_argument("--headless")

    # Scrape product information and save files
    # Continue until I type exit

    parent_category = input("Parent category: ")


    while True:
        
        print(f"You are in {parent_category}, type change to choose another one!")
        site = input("Category page URL to be scraped: ")
        if site == "exit": 
            print("You stopped the script")
            break
        if site == "change":
            parent_category = input("Parent category: ")
            site = input("Category page URL to be scraped: ")
        
        scrape_category_page(site, chrome_options, parent_category)

# Run code
if __name__ == "__main__":
    main()




