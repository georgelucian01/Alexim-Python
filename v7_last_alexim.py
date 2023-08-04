
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import openpyxl
import time


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

# Function to scrape product information and save products list
def scrape_product_info(product_page_link, parent_category, category, driver, products, Stoc):
    
    word_url = product_page_link # website
    
    driver.get(word_url)

    # Wait for the dynamic content to load (you may need to adjust the waiting time depending on the page)
    driver.implicitly_wait(2)

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
    
    # stoc is checked in category
    
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
    products.append([item_name, product_description, Descriere_meta, Taxa, item_price, item_sku, Categorie, Producator, Unitate_de_masura, Disponibilitate, Stoc, Vizibilitate, Imagine1, Imagine2, Imagine3, Imagine4])
    
# Function to scrape all the products from a category and scrapes other pages aswell
def scrape_category_page(site, subcategory, parent_category, driver, products, extra_pages_list, nr_pages):

    word_url = site  # category page URL

    # Use the 'driver' to get the page source and interact with the web page
    driver.get(word_url)
    driver.implicitly_wait(2)  ## waiting 2 seconds
    page_source = driver.page_source
    soup = BeautifulSoup(page_source, 'lxml')

    # Getting category
    category = subcategory

    # Find all article blocks with product information
    article_blocks = soup.find_all('article', class_='product-miniature')

    print(f"You are in {parent_category}")

    # Scrape information from each product article block
    for article in article_blocks:
        # Extract the product page link from the product name (h5 tag)
        product_name_tag = article.find('h5', class_='product-name')
        product_page_link = product_name_tag.a['href'] if product_name_tag and product_name_tag.a else None

        if product_page_link:
            button_element = article.find(class_="grid-buy-button")
            button_text = button_element.find("span").text.strip()
            if "Cumpăra" in button_text:
                Stoc = 1000
            else:
                Stoc = 0
            # Call the 'scrape_product_info' function to scrape information from the product page
            scrape_product_info(product_page_link, parent_category, category, driver, products, Stoc)
            print(f"stoc {Stoc}", button_text)

    print("Finished scraping all products on the page!")

    if nr_pages > 0:
        last_in_list = extra_pages_list[nr_pages-1]
        print(f" -----------Extra page {last_in_list}.")
        scrape_category_page(last_in_list, category, parent_category, driver, products, extra_pages_list, nr_pages-1)
        
# Check if category has extra pages
def extra_pages(site, driver):

    word_url = site  # category page URL

    # Use the 'driver' to get the page source and interact with the web page
    driver.get(word_url)
    driver.implicitly_wait(1)  ## waiting 2 seconds
    page_source = driver.page_source
    soup = BeautifulSoup(page_source, 'lxml')


    # Check if there are extra pages and call the function recursively for each extra page
    pagination_block = soup.find('div', class_='pagination-wrapper')
    if pagination_block and pagination_block.find('ul', class_='page-list'):
        pagination_links = pagination_block.find_all('a', class_='js-search-link')
        del pagination_links[-1]
        links = [link['href'] for link in pagination_links]
        links.pop(0) # remove first page
        return links # return the links
    return [] # return empty

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
    for product in products:
        worksheet.append(product)

    # Save the changes to the Excel file
    workbook.save(file_path)

    print("Product information saved in feed")

# Main code
def main():

    ## HTML BLOCK PENTRU MENIU cu categorii
    html = """
    <div class="category-tree js-category-tree">
    <ul><li data-depth="0"><a href="https://aleximtop.ro/accesorii-auto" title="Accesorii Auto" data-category-id="561">Accesorii Auto</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar561"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar561">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/conectori-auto" title="Conectori Auto" data-category-id="554">Conectori Auto</a></li><li data-depth="1"><a href="https://aleximtop.ro/difuzoare-auto" title="Difuzoare Auto" data-category-id="557">Difuzoare Auto</a></li><li data-depth="1"><a href="https://aleximtop.ro/mufe-auto" title="Mufe Auto" data-category-id="556">Mufe Auto</a></li><li data-depth="1"><a href="https://aleximtop.ro/modulator-fm-mp3-player-auto" title="Modulator Fm-Mp3-Player Auto" data-category-id="559">Modulator Fm-Mp3-Player Auto</a></li><li data-depth="1"><a href="https://aleximtop.ro/diverse-auto" title="Diverse Auto" data-category-id="555">Diverse Auto</a></li><li data-depth="1"><a href="https://aleximtop.ro/lumini-auto" title="Lumini Auto " data-category-id="590">Lumini Auto </a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/accesorii-pc-laptop-telefon" title="Accesorii PC-Laptop-Telefon" data-category-id="383">Accesorii PC-Laptop-Telefon</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar383"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar383">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/incarcatoare-telefon" title="Incarcatoare Telefon" data-category-id="345">Incarcatoare Telefon</a></li><li data-depth="1"><a href="https://aleximtop.ro/cabluri-de-date" title="Cabluri de Date" data-category-id="563">Cabluri de Date</a></li><li data-depth="1"><a href="https://aleximtop.ro/convertoare-smart-video" title="Convertoare Smart-Video" data-category-id="564">Convertoare Smart-Video</a></li><li data-depth="1"><a href="https://aleximtop.ro/memorii-card-stik-usb" title="Memorii Card &amp; Stik Usb" data-category-id="525">Memorii Card &amp; Stik Usb</a></li><li data-depth="1"><a href="https://aleximtop.ro/sisteme-audio-pc-laptop-telefon" title="Sisteme Audio PC" data-category-id="562">Sisteme Audio PC</a></li><li data-depth="1"><a href="https://aleximtop.ro/mouse-tastaturi" title="Mouse-Tastaturi" data-category-id="566">Mouse-Tastaturi</a></li><li data-depth="1"><a href="https://aleximtop.ro/surse-alimentare-pc" title="Surse Alimentare PC" data-category-id="565">Surse Alimentare PC</a></li><li data-depth="1"><a href="https://aleximtop.ro/accesorii-pc-pc-diverse" title="Pc Diverse" data-category-id="384">Pc Diverse</a></li><li data-depth="1"><a href="https://aleximtop.ro/semnal-wireles" title="Semnal &amp; Wireles" data-category-id="570">Semnal &amp; Wireles</a></li><li data-depth="1"><a href="https://aleximtop.ro/casti-audio" title="Casti Audio" data-category-id="583">Casti Audio</a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/accesorii-motoare-pompe" title="Accesorii Motoare-Pompe" data-category-id="544">Accesorii Motoare-Pompe</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar544"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar544">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/condensatori-motor" title="Condensatori Motor" data-category-id="291">Condensatori Motor</a></li><li data-depth="1"><a href="https://aleximtop.ro/presetupe-pompe-hidrofoare" title="Presetupe Pompe - Hidrofoare" data-category-id="289">Presetupe Pompe - Hidrofoare</a></li><li data-depth="1"><a href="https://aleximtop.ro/carbuni-perii-colectoare" title="Carbuni Perii Colectoare" data-category-id="290">Carbuni Perii Colectoare</a></li><li data-depth="1"><a href="https://aleximtop.ro/piese-masini-de-spalat" title="Piese Masini de Spalat " data-category-id="313">Piese Masini de Spalat </a></li><li data-depth="1"><a href="https://aleximtop.ro/rulmenti-zz-2rs-skf" title="Rulmenti ZZ,2RS,SKF" data-category-id="292">Rulmenti ZZ,2RS,SKF</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar292"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar292">
    <ul><li data-depth="2"><a href="https://aleximtop.ro/skf" title="SKF" data-category-id="546">SKF</a></li><li data-depth="2"><a href="https://aleximtop.ro/zz" title="ZZ" data-category-id="547">ZZ</a></li><li data-depth="2"><a href="https://aleximtop.ro/2rs" title="2RS" data-category-id="592">2RS</a></li></ul></div></li><li data-depth="1"><a href="https://aleximtop.ro/simeringuri-granituri" title="Simeringuri &amp; Granituri" data-category-id="342">Simeringuri &amp; Granituri</a></li><li data-depth="1"><a href="https://aleximtop.ro/piese-electro-pompe-apa" title="Piese &amp; Electro-Pompe Apa" data-category-id="568">Piese &amp; Electro-Pompe Apa</a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/accesorii-de-lipit" title="Accesorii de Lipit" data-category-id="301">Accesorii de Lipit</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar301"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar301">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/fludor-pasta-decapanta" title="Fludor &amp; Pasta Decapanta" data-category-id="302">Fludor &amp; Pasta Decapanta</a></li><li data-depth="1"><a href="https://aleximtop.ro/letcon-pistol-de-lipit" title="Letcon &amp; Pistol de Lipit" data-category-id="303">Letcon &amp; Pistol de Lipit</a></li><li data-depth="1"><a href="https://aleximtop.ro/pistol-bara-de-silicon" title="Pistol &amp; Bara de Silicon " data-category-id="333">Pistol &amp; Bara de Silicon </a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/automatizari-relee" title="Automatizari - Relee" data-category-id="304">Automatizari - Relee</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar304"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar304">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/intrerupatoare-comutatoare-butoane-limitatoare-intrerupatoare-motoare-pompe-bs-nf-lei-ac3-monofazice-trifazice" title="Contactori-NF-AC3" data-category-id="305">Contactori-NF-AC3</a></li><li data-depth="1"><a href="https://aleximtop.ro/punti-redresoare" title="Punti Redresoare" data-category-id="527">Punti Redresoare</a></li><li data-depth="1"><a href="https://aleximtop.ro/martori-indicatori-tensiune" title="Martori - Indicatori Tensiune" data-category-id="569">Martori - Indicatori Tensiune</a></li><li data-depth="1"><a href="https://aleximtop.ro/intrerupatoare-comutatoare-butoane-limitatoare-accesorii-pentru-automatizari-si-control" title="Timere - Control" data-category-id="372">Timere - Control</a></li><li data-depth="1"><a href="https://aleximtop.ro/intrerupatoare-comutatoare-butoane-limitatoare-relee" title="Relee" data-category-id="393">Relee</a></li><li data-depth="1"><a href="https://aleximtop.ro/papucielectriciputereautocontactori" title="Papuci de Putere" data-category-id="516">Papuci de Putere</a></li><li data-depth="1"><a href="https://aleximtop.ro/papuci-electrici" title="Papuci Electrici" data-category-id="332">Papuci Electrici</a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/intrerupatoare-butoane" title="Intrerupatoare - Butoane" data-category-id="361">Intrerupatoare - Butoane</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar361"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar361">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/butoane" title="Butoane " data-category-id="323">Butoane </a></li><li data-depth="1"><a href="https://aleximtop.ro/intrerupatoare-irs" title="Intrerupatoare IRS" data-category-id="306">Intrerupatoare IRS</a></li><li data-depth="1"><a href="https://aleximtop.ro/limitatoare" title="Limitatoare" data-category-id="307">Limitatoare</a></li><li data-depth="1"><a href="https://aleximtop.ro/push-butoane" title="Push Butoane" data-category-id="371">Push Butoane</a></li><li data-depth="1"><a href="https://aleximtop.ro/intrerupatoare-motoare" title="Intrerupatoare Motoare" data-category-id="422">Intrerupatoare Motoare</a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/alimentatoare-transformatoare" title="Alimentatoare-Transformatoare" data-category-id="318">Alimentatoare-Transformatoare</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar318"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar318">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/alimentatoare-stabilizate" title="Alimentatoare Stabilizate " data-category-id="322">Alimentatoare Stabilizate </a></li><li data-depth="1"><a href="https://aleximtop.ro/alimentatoare-reglabile" title="Alimentatoare Reglabile " data-category-id="350">Alimentatoare Reglabile </a></li><li data-depth="1"><a href="https://aleximtop.ro/invertoare-de-tensiune" title="Invertoare De Tensiune " data-category-id="573">Invertoare De Tensiune </a></li><li data-depth="1"><a href="https://aleximtop.ro/drivere-leduri" title="Drivere Leduri" data-category-id="387">Drivere Leduri</a></li><li data-depth="1"><a href="https://aleximtop.ro/surse-alimentare-metalice" title="Surse Alimentare Metalice" data-category-id="549">Surse Alimentare Metalice</a></li><li data-depth="1"><a href="https://aleximtop.ro/convertoare-de-tensiune" title="Convertoare De Tensiune " data-category-id="571">Convertoare De Tensiune </a></li><li data-depth="1"><a href="https://aleximtop.ro/ups-uri" title="UPS-uri" data-category-id="572">UPS-uri</a></li><li data-depth="1"><a href="https://aleximtop.ro/incarcatoare-de-acumulatorii" title="Incarcatoare de Acumulatorii " data-category-id="373">Incarcatoare de Acumulatorii </a></li><li data-depth="1"><a href="https://aleximtop.ro/stabilizatoare-tensiune" title="Stabilizatoare Tensiune " data-category-id="591">Stabilizatoare Tensiune </a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/acumulatori-reincarcabili-plumb-acid" title="Acumulatori Reincarcabili" data-category-id="505">Acumulatori Reincarcabili</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar505"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar505">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/acumulatori-reincarcabili-acumulatori-r6-aa" title="R6 AA" data-category-id="363">R6 AA</a></li><li data-depth="1"><a href="https://aleximtop.ro/acumulatori-r3-aaa" title="R3 AAA" data-category-id="359">R3 AAA</a></li><li data-depth="1"><a href="https://aleximtop.ro/speciali-li-ion" title="Speciali &amp; Li-ion" data-category-id="380">Speciali &amp; Li-ion</a></li><li data-depth="1"><a href="https://aleximtop.ro/plumb-acid" title="Plumb-Acid" data-category-id="360">Plumb-Acid</a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/baterii-acumulatori" title="Baterii" data-category-id="508">Baterii</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar508"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar508">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/bateri-r6-aa" title="Bateri R6 AA" data-category-id="379">Bateri R6 AA</a></li><li data-depth="1"><a href="https://aleximtop.ro/bateri-r3-aaa" title="Bateri R3 AAA" data-category-id="378">Bateri R3 AAA</a></li><li data-depth="1"><a href="https://aleximtop.ro/bateri-alkalinenealkalineli-ion-bateri-9vr14r20" title="Bateri 9V,R14,R20" data-category-id="377">Bateri 9V,R14,R20</a></li><li data-depth="1"><a href="https://aleximtop.ro/bateri-alkalinenealkalineli-ion-bateri-speciale" title="Bateri Speciale" data-category-id="348">Bateri Speciale</a></li><li data-depth="1"><a href="https://aleximtop.ro/bateri-alkaline-ag" title="Bateri Alkaline AG " data-category-id="339">Bateri Alkaline AG </a></li><li data-depth="1"><a href="https://aleximtop.ro/bateri-li-ion-tip-cr" title="Bateri Li-ion tip CR" data-category-id="347">Bateri Li-ion tip CR</a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/cabluri-la-rola" title="Cabluri la rola" data-category-id="354">Cabluri la rola</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar354"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar354">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/cablu-de-boxe-difuzoare" title="Cablu de Boxe - Difuzoare" data-category-id="355">Cablu de Boxe - Difuzoare</a></li><li data-depth="1"><a href="https://aleximtop.ro/cablu-coaxial-tv" title="Cablu Coaxial Tv" data-category-id="357">Cablu Coaxial Tv</a></li><li data-depth="1"><a href="https://aleximtop.ro/cabluri-electrice" title="Cabluri Electrice" data-category-id="530">Cabluri Electrice</a></li><li data-depth="1"><a href="https://aleximtop.ro/cablu-ftp-utp" title="Cablu Ftp &amp; Utp" data-category-id="356">Cablu Ftp &amp; Utp</a></li><li data-depth="1"><a href="https://aleximtop.ro/cabluri-alarma" title="Cabluri Alarma" data-category-id="574">Cabluri Alarma</a></li><li data-depth="1"><a href="https://aleximtop.ro/cablu-auto-de-putere" title="Cablu Auto de Putere" data-category-id="575">Cablu Auto de Putere</a></li><li data-depth="1"><a href="https://aleximtop.ro/cablu-microfon" title="Cablu Microfon " data-category-id="576">Cablu Microfon </a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/cabluri-audio-video-ac-dc" title="Cabluri Audio – Video &amp; Ac-Dc" data-category-id="293">Cabluri Audio – Video &amp; Ac-Dc</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar293"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar293">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/jack" title="Jack" data-category-id="294">Jack</a></li><li data-depth="1"><a href="https://aleximtop.ro/xlr-jack-spikon" title="Xlr-Jack-Spikon" data-category-id="295">Xlr-Jack-Spikon</a></li><li data-depth="1"><a href="https://aleximtop.ro/cabluri-alimentare-audio-video-pc-ac-dc-cabluri-audio-video-cu-mufe-hdmi" title="Hdmi" data-category-id="296">Hdmi</a></li><li data-depth="1"><a href="https://aleximtop.ro/cabluri-alimentare-audio-video-pc-ac-dc-cabluri-audio-video-cu-mufe-scartvga" title="Vga" data-category-id="297">Vga</a></li><li data-depth="1"><a href="https://aleximtop.ro/cabluri-alimentare-audio-video-pc-ac-dc-cabluri-audio-cu-mufe-rca" title="Rca" data-category-id="315">Rca</a></li><li data-depth="1"><a href="https://aleximtop.ro/cabluri-alimentare-audio-video-pc-ac-dc-cabluri-usb-otg-mini-usb-micro-usb" title="Usb" data-category-id="316">Usb</a></li><li data-depth="1"><a href="https://aleximtop.ro/cabluri-alimentare-audio-video-pc-ac-dc-diverse-cabluri-utile" title="Utile" data-category-id="317">Utile</a></li><li data-depth="1"><a href="https://aleximtop.ro/cabluri-alimentare-audio-video-pc-ac-dc-cablu-utp-tel" title="Ftp-Utp-Tel" data-category-id="381">Ftp-Utp-Tel</a></li><li data-depth="1"><a href="https://aleximtop.ro/cabu-alimentare-ac-dc-rc" title="Ac-Dc" data-category-id="504">Ac-Dc</a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/adaptori-conectori-mufe" title="Adaptori Conectori &amp; Mufe " data-category-id="352">Adaptori Conectori &amp; Mufe </a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar352"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar352">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/mufe-tv-f" title="Mufe TV &amp; F" data-category-id="353">Mufe TV &amp; F</a></li><li data-depth="1"><a href="https://aleximtop.ro/xlr-speak-on" title="XLR &amp; Speak-on" data-category-id="368">XLR &amp; Speak-on</a></li><li data-depth="1"><a href="https://aleximtop.ro/conectori-ac-dc" title="Conectori AC-DC" data-category-id="382">Conectori AC-DC</a></li><li data-depth="1"><a href="https://aleximtop.ro/conectori-mufe-mufa-conector-bnc" title="BNC &amp; Diverse" data-category-id="388">BNC &amp; Diverse</a></li><li data-depth="1"><a href="https://aleximtop.ro/conectori-jac-rca" title="Conectori JAC , RCA" data-category-id="389">Conectori JAC , RCA</a></li><li data-depth="1"><a href="https://aleximtop.ro/amplificatoare-semnal-tv-splitere-cablu-tv" title="Spliter-Amplificator" data-category-id="328">Spliter-Amplificator</a></li><li data-depth="1"><a href="https://aleximtop.ro/conectori-hdmi-dvi-vga" title="Conectori Hdmi , Dvi ,Vga" data-category-id="364">Conectori Hdmi , Dvi ,Vga</a></li><li data-depth="1"><a href="https://aleximtop.ro/conectori-tel-utp" title=" Conectori Tel , Utp" data-category-id="365"> Conectori Tel , Utp</a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/electronice" title="Electronice" data-category-id="513">Electronice</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar513"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar513">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/boxe-portabileradio" title="Boxe Portabile&amp;Radio" data-category-id="515">Boxe Portabile&amp;Radio</a></li><li data-depth="1"><a href="https://aleximtop.ro/ceasuri-de-camera" title="Ceasuri de Camera" data-category-id="567">Ceasuri de Camera</a></li><li data-depth="1"><a href="https://aleximtop.ro/termometre-higrometre" title="Termometre-Higrometre" data-category-id="320">Termometre-Higrometre</a></li><li data-depth="1"><a href="https://aleximtop.ro/telecomenzi-tv-lcd-led-air" title="Telecomenzi TV-LCD-LED-AIR" data-category-id="584">Telecomenzi TV-LCD-LED-AIR</a></li><li data-depth="1"><a href="https://aleximtop.ro/suporturi-tv" title="Suporturi Tv " data-category-id="589">Suporturi Tv </a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/electrice-aparataj-electric" title="Electrice &amp; Aparataj Electric" data-category-id="298">Electrice &amp; Aparataj Electric</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar298"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar298">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/adaptoare-ac-dc" title="Adaptoare Ac-Dc" data-category-id="300">Adaptoare Ac-Dc</a></li><li data-depth="1"><a href="https://aleximtop.ro/banda-izolatoare-adeziva" title="Banda Izolatoare &amp; Adeziva" data-category-id="299">Banda Izolatoare &amp; Adeziva</a></li><li data-depth="1"><a href="https://aleximtop.ro/duli-stechere-cuple" title="Duli - Stechere - Cuple" data-category-id="519">Duli - Stechere - Cuple</a></li><li data-depth="1"><a href="https://aleximtop.ro/intrerupatoare-prize" title="Intrerupatoare &amp; Prize " data-category-id="520">Intrerupatoare &amp; Prize </a></li><li data-depth="1"><a href="https://aleximtop.ro/prelungitoare-electrice" title="Prelungitoare Electrice" data-category-id="385">Prelungitoare Electrice</a></li><li data-depth="1"><a href="https://aleximtop.ro/triplu-stecher" title="Triplu Stecher " data-category-id="522">Triplu Stecher </a></li><li data-depth="1"><a href="https://aleximtop.ro/sonerii-alarme" title="Sonerii &amp; Alarme" data-category-id="331">Sonerii &amp; Alarme</a></li><li data-depth="1"><a href="https://aleximtop.ro/ventilatoare" title="Ventilatoare" data-category-id="335">Ventilatoare</a></li><li data-depth="1"><a href="https://aleximtop.ro/tub-termo" title="Tub Termo" data-category-id="394">Tub Termo</a></li><li data-depth="1"><a href="https://aleximtop.ro/reglete-conectori-electrici" title="Reglete &amp; Conectori Electrici" data-category-id="517">Reglete &amp; Conectori Electrici</a></li><li data-depth="1"><a href="https://aleximtop.ro/bride-coliere-de-prindere" title="Bride &amp; Coliere de Prindere" data-category-id="518">Bride &amp; Coliere de Prindere</a></li><li data-depth="1"><a href="https://aleximtop.ro/senzori" title="Senzori" data-category-id="523">Senzori</a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/rezistente-electrice" title="Rezistente Electrice" data-category-id="308">Rezistente Electrice</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar308"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar308">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/rezistente-din-nichelina" title="Rezistente din Nichelina" data-category-id="309">Rezistente din Nichelina</a></li><li data-depth="1"><a href="https://aleximtop.ro/rezistente-de-boilere" title="Rezistente de Boilere" data-category-id="310">Rezistente de Boilere</a></li><li data-depth="1"><a href="https://aleximtop.ro/flanse-garnituri-boilere" title="Flanse &amp; Garnituri Boilere" data-category-id="321">Flanse &amp; Garnituri Boilere</a></li><li data-depth="1"><a href="https://aleximtop.ro/rezistente-de-halogen-si-quart" title="Rezistente de Halogen si Quart" data-category-id="367">Rezistente de Halogen si Quart</a></li><li data-depth="1"><a href="https://aleximtop.ro/rezistente-pentru-masini-de-spalat" title="Rezistente pentru Masini de Spalat " data-category-id="374">Rezistente pentru Masini de Spalat </a></li><li data-depth="1"><a href="https://aleximtop.ro/piese-aragaze-cuptoare" title="Piese Aragaze &amp; Cuptoare" data-category-id="521">Piese Aragaze &amp; Cuptoare</a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/termostate-protecti" title="Termostate &amp; Protecti" data-category-id="311">Termostate &amp; Protecti</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar311"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar311">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/termostate-de-boilere-si-calorifere" title="Termostate de Boilere si Calorifere" data-category-id="312">Termostate de Boilere si Calorifere</a></li><li data-depth="1"><a href="https://aleximtop.ro/termostate-de-cuptoare-friptoze-frigidere" title="Termostate de Cuptoare , Friptoze , Frigidere" data-category-id="325">Termostate de Cuptoare , Friptoze , Frigidere</a></li><li data-depth="1"><a href="https://aleximtop.ro/protecti-termice" title="Protecti Termice" data-category-id="324">Protecti Termice</a></li><li data-depth="1"><a href="https://aleximtop.ro/termostate-bimetal" title="Termostate Bimetal" data-category-id="366">Termostate Bimetal</a></li><li data-depth="1"><a href="https://aleximtop.ro/sigurante-termice" title="Sigurante Termice" data-category-id="369">Sigurante Termice</a></li><li data-depth="1"><a href="https://aleximtop.ro/sigurante-fuzibile" title="Sigurante Fuzibile" data-category-id="588">Sigurante Fuzibile</a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/scule-unelte-utile-clesti-surubelnite" title="Scule - Unelte" data-category-id="507">Scule - Unelte</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar507"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar507">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/multimetre-testere" title="Multimetre &amp; Testere" data-category-id="541">Multimetre &amp; Testere</a></li><li data-depth="1"><a href="https://aleximtop.ro/clesti-patenti" title="Clesti &amp; Patenti" data-category-id="509">Clesti &amp; Patenti</a></li><li data-depth="1"><a href="https://aleximtop.ro/lupa-de-marit" title="Lupa de Marit " data-category-id="512">Lupa de Marit </a></li><li data-depth="1"><a href="https://aleximtop.ro/rulete-cuttere" title="Rulete &amp; Cuttere" data-category-id="376">Rulete &amp; Cuttere</a></li><li data-depth="1"><a href="https://aleximtop.ro/clesti-sertizare-compresie" title="Clesti Sertizare &amp; Compresie" data-category-id="343">Clesti Sertizare &amp; Compresie</a></li><li data-depth="1"><a href="https://aleximtop.ro/surubelnite-creioane-tensiune" title="Surubelnite &amp; Creioane Tensiune" data-category-id="375">Surubelnite &amp; Creioane Tensiune</a></li><li data-depth="1"><a href="https://aleximtop.ro/truse-scule" title="Truse Scule " data-category-id="334">Truse Scule </a></li><li data-depth="1"><a href="https://aleximtop.ro/diverse" title="Diverse" data-category-id="526">Diverse</a></li><li data-depth="1"><a href="https://aleximtop.ro/chei-reglabile-scule" title="Chei Reglabile &amp; Scule" data-category-id="577">Chei Reglabile &amp; Scule</a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/iluminat-led" title="Iluminat Led" data-category-id="511">Iluminat Led</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar511"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar511">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/banda-led" title="Banda Led" data-category-id="543">Banda Led</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar543"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar543">
    <ul><li data-depth="2"><a href="https://aleximtop.ro/controlere-led" title="Controlere Led " data-category-id="593">Controlere Led </a></li></ul></div></li><li data-depth="1"><a href="https://aleximtop.ro/grup-led" title="Grup Led" data-category-id="542">Grup Led</a></li><li data-depth="1"><a href="https://aleximtop.ro/becuri" title="Becuri" data-category-id="395">Becuri</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar395"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar395">
    <ul><li data-depth="2"><a href="https://aleximtop.ro/halogen" title="Halogen" data-category-id="585">Halogen</a></li><li data-depth="2"><a href="https://aleximtop.ro/lanterna" title="Lanterna" data-category-id="586">Lanterna</a></li><li data-depth="2"><a href="https://aleximtop.ro/led" title="Led" data-category-id="587">Led</a></li></ul></div></li><li data-depth="1"><a href="https://aleximtop.ro/corpuri-de-lumina" title="Corpuri de lumina " data-category-id="392">Corpuri de lumina </a></li><li data-depth="1"><a href="https://aleximtop.ro/conectori-led" title="Conectori Led" data-category-id="362">Conectori Led</a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/surse-de-iluminat-lanterne-alimentate-cu-baterii" title="Lanterne Led" data-category-id="337">Lanterne Led</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar337"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar337">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/lanterne-de-mana" title="Lanterne de Mana " data-category-id="579">Lanterne de Mana </a></li><li data-depth="1"><a href="https://aleximtop.ro/lanterne-de-cap" title="Lanterne de Cap" data-category-id="578">Lanterne de Cap</a></li><li data-depth="1"><a href="https://aleximtop.ro/lampi-de-lucru" title="Lampi de Lucru " data-category-id="341">Lampi de Lucru </a></li></ul></div></li><li data-depth="0"><a href="https://aleximtop.ro/proiectoare-cu-led" title="Proiectoare cu Led" data-category-id="396">Proiectoare cu Led</a><div class="navbar-toggler collapse-icons" data-toggle="collapse" data-target="#exCollapsingNavbar396"><i class="material-icons add"></i><i class="material-icons remove"></i></div><div class="category-sub-menu collapse" id="exCollapsingNavbar396">
    <ul><li data-depth="1"><a href="https://aleximtop.ro/proiectoare-220v" title="Proiectoare 220V " data-category-id="580">Proiectoare 220V </a></li><li data-depth="1"><a href="https://aleximtop.ro/proiectoare-12v" title="Proiectoare 12V " data-category-id="581">Proiectoare 12V </a></li><li data-depth="1"><a href="https://aleximtop.ro/module-led-smd-cob" title="Module Led SMD &amp; COB  " data-category-id="582">Module Led SMD &amp; COB  </a></li></ul></div></li></ul>
    </div>
    """

    #Start measuring the runtime
    start_time = time.time()

    # Create an empty list to store the product information
    products = []

    result = scrape_categories(html)

    # Create a new ChromeDriver instance
    

    for category in result:
        print("Parent Category:", category['parent_category'])

        chromedriver_path = "./chromedriver.exe"
        os.environ["webdriver.chrome.driver"] = chromedriver_path
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        driver = webdriver.Chrome(options=chrome_options)

        for subcategory in category['subcategories']:
            print("  - Subcategory Name:", subcategory['name'])
            print("    Subcategory URL:", subcategory['url'])

            # Scrape information from the subcategory page and add it to the products list
            extra_pages_list = extra_pages(subcategory['url'], driver) # check how many pages extra
            nr_pages = len(extra_pages_list) # How many pages there are (-1)

            scrape_category_page(subcategory['url'], subcategory['name'], category['parent_category'], driver, products, extra_pages_list, nr_pages)

            write_to_excel(products) # add the category to excel (save progress)
            products.clear() # clear the list
        
        # Quit the driver after scraping the category
        driver.quit()

    # End measuring the runtime
    end_time = time.time()
    runtime = end_time - start_time
    print(f"Main function finished in {runtime:.2f} seconds")


# Run code
if __name__ == "__main__":
    main()
