import openpyxl
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
from selenium.webdriver.chrome.options import Options
import logging as lg


def run_price_monitoring():
    # run driver and wait for the page to load
    def run_driver(site_url):
        try:
            chorme_options = Options()

            arguments = [
                "--lang=pt-BR",
                "--window-size=1300,1000",
                "--disable-notifications",
                "--incognito",
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
            return None

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

        # a list to store the product_name, data_atual, product_price and link
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

                print(
                    f"Product: {product_name}\n"
                    f"Date: {data_atual}\n"
                    f"Price: R$ {product_price}\n"
                    f"Link: {link}\n"
                )

        except Exception as e:
            lg.error(
                f"Error occurred while searching for the product: {type(e).__name__} - {e}"
            )
            print("An error occurred while extracting the product data")
            return None

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

        file_path = f"{search_product_}.xlsx"

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

            wb.save(file_path)
            print("Data saved successfully")
            print(
                "Copy the path below and paste it into the file explorer to open the file"
            )
            print(
                f'The path where the file was saved: {os.path.abspath("products.xlsx")}'
            )
        except Exception as e:
            lg.error(
                f"Error occurred while saving data to excel: {type(e).__name__} - {e}"
            )
            print("An error occurred while saving the data to the excel file")
            return None
        finally:
            wb.close()
        return True

    save_data_excel()


run_price_monitoring()
