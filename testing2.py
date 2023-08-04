import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import openpyxl
import time



# Main code
def main():

    ## HTML BLOCK PENTRU MENIU cu categorii
    html = """
    <div class="grid-buy-button">
                            <a class="btn add-to-cart details-link" href="https://aleximtop.ro/carbuni-perii-colectoare/-perii-colectoare-carbuni-5x8x115-a">
                <span>Detalii</span>
              </a>
            </div>
    """

    #Start measuring the runtime
    start_time = time.time()

    # Create an empty list to store the product information
    soup = BeautifulSoup(html, 'lxml')

    chromedriver_path = "./chromedriver.exe"
    os.environ["webdriver.chrome.driver"] = chromedriver_path
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(options=chrome_options)

    button_element = soup.find(class_="grid-buy-button")
    button_text = button_element.find("span").text.strip()
    print(button_text)
    if "Cumpăra" in button_text:
        stoc = 1000
    else:
        stoc = 0
    print(stoc)
    driver.quit()

    # End measuring the runtime
    end_time = time.time()
    runtime = end_time - start_time
    print(f"Main function finished in {runtime:.2f} seconds")

# Run code
if __name__ == "__main__":
    main()