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

# Function to open the images folder
def open_images_folder(item_sku):
    folder_path = os.path.join("images", item_sku)

    if os.path.exists(folder_path):
        try:
            if platform.system() == "Darwin":  # macOS
                subprocess.run(["open", folder_path], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            elif platform.system() == "Windows":  # Windows
                subprocess.run(["explorer", folder_path], check=True, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            elif platform.system() == "Linux":  # Linux
                subprocess.run(["xdg-open", folder_path], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            else:
                print("Unsupported operating system.")
        except Exception:
            pass  # Suppress any exceptions
    else:
        print(f"{folder_path} does not exist.")

# Function to scrape product information
def scrape_product_info(site):
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

    item_price_str = soup.find(itemprop="price").text.strip() if soup.find(itemprop="price") else "0"
    
    # Remove "RON", non-numeric characters (except decimal points), and any extra spaces from item_price_str
    item_price_str = item_price_str.replace('RON', '')  # Remove 'RON' from the string

    # Convert comma (,) to period (.) as the decimal separator
    item_price_str = item_price_str.replace('.', '').replace(',', '.')

    # Convert item_price_str to float
    try:
        item_price = float(item_price_str)
    except ValueError:
        item_price = 0

    item_price = calculate_final_price(item_price)

    product_description = soup.find(class_="product-description typo").text.strip() if soup.find(class_="product-description typo") else "N/A"


    ### EXCEL INFORMATION ADDING ###

    # Check if the Excel file exists
    file_path = "product_info.xlsx"
    is_file_exists = os.path.exists(file_path)

    # Create or open the Excel workbook
    workbook = openpyxl.Workbook() if not is_file_exists else openpyxl.load_workbook(file_path)

    # Select the active worksheet
    worksheet = workbook.active

    # If the file is newly created, add header row with column names
    if not is_file_exists:
        worksheet.append(["Nume_Produs", "Descriere", "Descriere_meta",	"Taxa_(%)",	"Pret_Produs", "Cod_produs_(SKU)", "Categorie",	"Producator",	"Unitate_de_masura"	,"Disponibilitate",	"Stoc",	"Vizibilitate",	"Imagine1",	"Imagine2",	"Imagine3",	"Imagine4"])

    # Append the scraped product information to the Excel file
    worksheet.append([item_name, item_sku, item_price, product_description])

    # Save the changes to the Excel file
    workbook.save(file_path)

    print("Product information saved in product_info.xlsx")

    #################################

    # Find the image elements within the "product-images" class and extract the big image URLs
    image_elements = soup.select('.product-images li a.thumb')
    image_urls = [element['data-zoom-image'] for element in image_elements]

    # Create a folder to store the downloaded images
    item_sku_folder = os.path.join("images", item_sku)
    if not os.path.exists(item_sku_folder):
        os.makedirs(item_sku_folder)
    else:
        # Delete existing files in the item_sku folder
        existing_files = os.listdir(item_sku_folder)
        for file_name in existing_files:
            file_path = os.path.join(item_sku_folder, file_name)
            os.unlink(file_path)

    # Download the images
    for idx, url in enumerate(image_urls):
        try:
            # Download the image using requests
            response = requests.get(url)

            rewritten_name = re.sub(r'[^a-zA-Z0-9]', ' ', item_name)

            with open(f'images/{item_sku}/{idx+1}_{rewritten_name}.jpg', 'wb') as f:
                f.write(response.content)

            print(f"Image {idx+1} downloaded.")
        except Exception as e:
            print(f"An error occurred while downloading image {idx+1}: {e}")

    # Close the browser
    driver.quit()

# Scrape product information and save files
# Continue until I type exit

print("Type exit to stop the script")

while True:
    site = input("Page URL to be scraped: ")
    if site == "exit": 
        print("You stopped the script")
        break
    scrape_product_info(site)

# open_images_folder(item_sku)