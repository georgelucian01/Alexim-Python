### THIS ONE DOESNT WORK PROPERLY


import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import openpyxl
import time
from multiprocessing import Process, cpu_count, Lock


# Multiprocessing
# Function to scrape information from all products in a subcategory page
def scrape_subcategory_page(subcategory_url, parent_category, products):
    # Set the environment variable for the ChromeDriver executable path
    chromedriver_path = "./chromedriver.exe"
    os.environ["webdriver.chrome.driver"] = chromedriver_path

    # Create a new ChromeDriver instance for each process
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(options=chrome_options)
    
    # Call the main scraping function
    scrape_category_page(subcategory_url, parent_category, driver, products)
    
    # Quit the driver after scraping
    driver.quit()

# Function to calculate final selling price
def calculate_final_price(item_price):

    # Used excel formulas as a model, that why the c2 d2 e2

    # Removing VAT
    c2 = float(item_price / 1.19)
    
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

# Function to scrape product information and saves to excel
def scrape_product_info(product_page_link, parent_category, category, driver, products):
    
    word_url = product_page_link # website
    
    driver.get(word_url)

    # Wait for the dynamic content to load (you may need to adjust the waiting time depending on the page)
    driver.implicitly_wait(1)

    # Get the page source
    page_source = driver.page_source

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(page_source, 'lxml')

    
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

    print(f"You are at url: {product_page_link}")
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


    # Append the scraped product information to the products list
    print(products)
    products.append([item_name, product_description, Descriere_meta, Taxa, item_price, item_sku, Categorie, Producator, Unitate_de_masura, Disponibilitate, Stoc, Vizibilitate, Imagine1, Imagine2, Imagine3, Imagine4])
    
# Function to scrape information from all products in a category page
def scrape_category_page(site, parent_category, driver, products):
    
    word_url = site  # category page URL

    # Use the 'driver' to get the page source and interact with the web page
    driver.get(word_url)
    driver.implicitly_wait(1)  ## waiting 2 seconds
    page_source = driver.page_source
    soup = BeautifulSoup(page_source, 'lxml')

    # Getting category
    category = soup.find(class_="page-heading").text.strip() if soup.find(class_="page-heading") else "N/A"

    # Find all article blocks with product information
    article_blocks = soup.find_all('article', class_='product-miniature')

    print(f"You are in {parent_category}")

    # Scrape information from each product article block
    for article in article_blocks:
        # Extract the product page link from the product name (h5 tag)
        product_name_tag = article.find('h5', class_='product-name')
        product_page_link = product_name_tag.a['href'] if product_name_tag and product_name_tag.a else None

        if product_page_link:
            # Call the 'scrape_product_info' function to scrape information from the product page
            scrape_product_info(product_page_link, parent_category, category, driver, products)

    print("Finished scraping all products on the page!")

# Scape categories
def scrape_categories(html):

    soup = BeautifulSoup(html, 'lxml')
    categories = []

    for li in soup.find_all('li', {'data-depth': '0'}):
        parent_category = li.find('a').text.strip()
        subcategories = []

        for sub_li in li.find_all('li', {'data-depth': '1'}):
            subcategory_name = sub_li.find('a').text.strip()
            subcategory_url = sub_li.find('a')['href']
            subcategories.append({
                'name': subcategory_name,
                'url': subcategory_url
            })

        categories.append({
            'parent_category': parent_category,
            'subcategories': subcategories
        })

    return categories

# Main code
def main():

    ## HTML BLOCK PENTRU MENIU cu categorii
    html = """
    <div class="category-tree js-category-tree">
        
    <ul><li data-depth="0"><a href="https://aleximtop.ro/accesorii-auto" title="Accesorii Auto" data-category-id="561">Accesorii Auto</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar561"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar561">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/conectori-auto" title="Conectori Auto" data-category-id="554">Conectori Auto</a></li><li data-depth="1"><a href="https://aleximtop.ro/difuzoare-auto" title="Difuzoare Auto" data-category-id="557">Difuzoare Auto</a></li><li data-depth="1"><a href="https://aleximtop.ro/mufe-auto" title="Mufe Auto" data-category-id="556">Mufe Auto</a></li><li data-depth="1"><a href="https://aleximtop.ro/modulator-fm-mp3-player-auto" title="Modulator Fm-Mp3-Player Auto" data-category-id="559">Modulator Fm-Mp3-Player Auto</a></li><li data-depth="1"><a href="https://aleximtop.ro/diverse-auto" title="Diverse Auto" data-category-id="555">Diverse Auto</a></li><li data-depth="1"><a href="https://aleximtop.ro/lumini-auto" title="Lumini Auto " data-category-id="590">Lumini Auto </a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/accesorii-pc-laptop-telefon" title="Accesorii PC-Laptop-Telefon" data-category-id="383">Accesorii PC-Laptop-Telefon</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar383"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar383">
    </div>
    """



   # Start measuring the runtime
    start_time = time.time()

    # Create a list to store the processes
    processes = []

    # Create a list to store the product information
    products = []

    result = scrape_categories(html)

    for category in result:
        print("Parent Category:", category['parent_category'])
        for subcategory in category['subcategories']:
            print("  - Subcategory Name:", subcategory['name'])
            print("    Subcategory URL:", subcategory['url'])
            # Start a new process for each subcategory
            process = Process(target=scrape_subcategory_page, args=(subcategory['url'], category['parent_category'], products))
            processes.append(process)
            process.start()

    # Wait for all processes to complete
    for process in processes:
        process.join()

    # Write the product information to the Excel file

    print(products)
    write_to_excel(products)

    # End measuring the runtime
    end_time = time.time()
    runtime = end_time - start_time
    print(f"Main function finished in {runtime:.2f} seconds")


# Create a lock to handle concurrent writes to the Excel file
lock = Lock()

# Function to write product information to the Excel file
def write_to_excel(products):
    # Check if the Excel file exists
    file_path = "./feed_alexim.xlsx"
    is_file_exists = os.path.exists(file_path)

    # Create or open the Excel workbook
    workbook = openpyxl.Workbook() if not is_file_exists else openpyxl.load_workbook(file_path)

    # Select the active worksheet
    worksheet = workbook.active

    # If the file is newly created, add header row with column names
    if not is_file_exists:
        worksheet.append(["Nume_Produs", "Descriere", "Descriere_meta", "Taxa_(%)", "Pret_Produs", "Cod_produs_(SKU)", "Categorie", "Producator", "Unitate_de_masura", "Disponibilitate", "Stoc", "Vizibilitate", "Imagine1", "Imagine2", "Imagine3", "Imagine4"])

    # Append the scraped product information to the Excel file
    with lock:
        for product in products:
            worksheet.append(product)

    # Save the changes to the Excel file
    workbook.save(file_path)

    print("Product information saved in feed")


# Run code
if __name__ == "__main__":
    main()