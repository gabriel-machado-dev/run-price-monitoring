import openpyxl
import os
from time import sleep
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
from selenium.webdriver.chrome.options import Options
import logging as lg
from tqdm import tqdm


# Simulate a loading bar
print("Loading Bot:")
for _ in tqdm(range(100), desc="Loading", ncols=75):
    sleep(0.05)


RED = '\033[91m'
GREEN = '\033[92m'
YELLOW = '\033[93m'
BLUE = '\033[94m'
MAGENTA = '\033[95m'
CYAN = '\033[96m'
RESET = '\033[0m'
BOLD = '\033[1m'

print(CYAN + BOLD + r'''
                                                          _____                         
                    _|_|_|      _|__     _        _      |     |   \     /
                    _|    _|    _|  |    |        |      |      |   \   /
                    _|    _|    _|__      |_    _|       |  ___|     \_/
                    _|    _|    _|         |    |        |            |
                    _|_|_|      -|__|       |__|       .  |           |  
                                                                              
                                                                           
                                  Price Monitoring Bot 
''' + RESET)

def run_price_monitoring():
    # run driver and wait for the page to load
    def run_driver(site_url):
        print("Initializing the driver")
        sleep(1)
        try:
            chorme_options = Options()

            arguments = [
                "--lang=pt-BR",
                "--window-size=1300,1000",
                "--disable-notifications",
                "--incognito"
            ]
            for argument in arguments:
                chorme_options.add_argument(argument)

            # desabilitar pop-up de navegador controlado por automacao
            chorme_options.add_experimental_option(
                "excludeSwitches", ["enable-automation"]
            )

            # Using experimental settings
            chorme_options.add_experimental_option(
                "prefs",
                {
                    # Alterar o local padrão de ‘download’ de arquivos
                    "download.default_directory": "C:\\Users\\gabri\\OneDrive\\Área de Trabalho\\projetos "
                    "pyautogui\\selenium_dev_aprender",
                    # notificar o Google chrome sobre essa alteração
                    "download.directory_upgrade": True,
                    # Desabilitar a confirmação de ‘download’
                    "download.prompt_for_download": False,
                    # Desabilitar notificações
                    "profile.default_content_setting_values.notifications": 2,
                    # Permitir multiplos downloads
                    "profile.default_content_setting_values.automatic_downloads": 1,
                },
            )

            driver = webdriver.Chrome(options=chorme_options)
            driver.get(site_url)

            wait = WebDriverWait(
                driver,
                10,
                poll_frequency=1,
                ignored_exceptions=[
                    NoSuchElementException,
                    ElementNotVisibleException,
                    ElementNotSelectableException,
                    TimeoutException,
                ],
            )

            return driver, wait
        except Exception as e:
            lg.error(
                f"Error occurred while initializng driver: {type(e).__name__} - {e}"
            )
            return None, None

    # acess website buscape
    def search_product():
        while True:
            search_product_ = input("Enter the product you want to search: ")
            if not search_product_ or search_product_.isspace():
                print("Invalid product name. Please try again.")
            else:
                break

        driver, wait = run_driver(
            f"https://www.buscape.com.br/search?q={search_product_}"
        )
        
        if driver is None or wait is None:
            print("Failed to initialize the driver")
            return None, None

        # a list to store the product_name, data_atual, product_price and link
        print("Extracting product data")
        sleep(2)
        
        products_data = []

        try:
            product = wait.until(
                EC.visibility_of_any_elements_located(
                    (By.XPATH, '//div[@data-testid="product-card"]')
                )
            )

            for item in product:
                product_name = item.find_element(By.XPATH, ".//h2").text
                data_atual = datetime.now().strftime("%d/%m/%Y")
                product_price = (
                    item.find_element(By.XPATH, "./a//div/p")
                    .text.split()[1]
                    .replace(",", ".")
                )
                link = item.find_element(By.XPATH, "./a").get_attribute("href")

                products_data.append((product_name, data_atual, product_price, link))
            
            sleep(1)    
            print('Product data extracted successfully')
            sleep(2)
            
        except TimeoutException as e:
            lg.error(f"TimeoutException occurred while searching for the product: {e}")
            print("Timeout occurred while extracting the product data. Please try again.")
            return None, None

        except Exception as e:
            lg.error(
                f"Error occurred while searching for the product: {type(e).__name__} - {e}"
            )
            print("An error occurred while extracting the product data")
            return None, None

        return products_data, search_product_

    def clean_price(price_str):
        # Remove any commas and convert to float
        return float(
            price_str.replace(",", "").replace(".", "", price_str.count(".") - 1)
        )

    # save data to excel file
    def save_data_excel():
       
        products_data, search_product_ = search_product()
        if not products_data:
            return None

        # Sort products_data by price (assuming price is the third element in the tuple)
        products_data = sorted(products_data, key=lambda x: clean_price(x[2]))

        file_path = f"{search_product_.strip()}.xlsx"

        try:
            # check if the file exists
            if os.path.exists(file_path):
                wb = openpyxl.load_workbook(file_path)
            else:
                wb = openpyxl.Workbook()
                sheet = wb.active
                sheet.append(["Product", "Date", "Price", "Link"])

            for product in products_data:
                # save the data to the file
                sheet.append(product)
                
            print("Saving data to excel")
            sleep(2)
            wb.save(file_path)
            
            print("Data saved successfully")
            sleep(2)
            
            print(f"Data saved to the folder: {os.getcwd()}")
            
            open_file = input("Do you want to open the file? (y/n): ")
            if open_file.lower() == "y":
                if os.name == "posix":
                    os.system(f"open {file_path}")
                elif os.name == "nt":
                    os.startfile(file_path)
                else:
                    print("Unsupported operating system.")
            
            
        except Exception as e:
            lg.error(
                f"Error occurred while saving data to excel: {type(e).__name__} - {e}"
            )
            print("An error occurred while saving the data to the excel file")
            return None
        finally:
            wb.close()
            
            return file_path
    
    save_data_excel()

# Run the price monitoring bot if the user wants to run it again after the first run 
while True:
    run_bot = input("Do you want to run the price monitoring bot? (y/n): ")
    if run_bot.lower() == "y":
        run_price_monitoring()
    else:
        break


