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

          arguments = ['--lang=pt-BR', '--window-size=1300,1000', '--disable-notifications', '--incognito']
          for argument in arguments:
              chorme_options.add_argument(argument)

          # desabilitar pop-up de navegador controlado por automacao
          chorme_options.add_experimental_option(
              'excludeSwitches', ['enable-automation'])

          # Using experimental settings
          chorme_options.add_experimental_option('prefs', {
              # Alterar o local padrão de ‘download’ de arquivos
              'download.default_directory': 'C:\\Users\\gabri\\OneDrive\\Área de Trabalho\\projetos '
                                            'pyautogui\\selenium_dev_aprender',
              # notificar o Google chrome sobre essa alteração
              'download.directory_upgrade': True,
              # Desabilitar a confirmação de ‘download’
              'download.prompt_for_download': False,
              # Desabilitar notificações
              'profile.default_content_setting_values.notifications': 2,
              # Permitir multiplos downloads
              'profile.default_content_setting_values.automatic_downloads': 1,

          })

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
                  TimeoutException
              ]
          )

          return driver, wait
      except Exception as e:
          lg.error(f'Error occurred while initializng driver: {type(e).__name__} - {e}')
          return None

  # acess website buscape
  def search_product():
    while True:
      search_product = input('Enter the product you want to search: ')
      if not search_product or search_product.isspace():
          print('Invalid product name. Please try again.')
      else:
          break

    
    driver, wait = run_driver(f'https://www.buscape.com.br/search?q={search_product}')
    
    # a list to store the product_name, data_atual, product_price and link
    products_data = []

    try:
      product = wait.until(EC.visibility_of_any_elements_located((By.XPATH, '//div[@data-testid="product-card"]')))

      
      for item in product:
          product_name = item.find_element(By.XPATH, './/h2').text
          data_atual = datetime.now().strftime('%d/%m/%Y')
          product_price = item.find_element(By.XPATH, './a//div/p').text.split()[1].replace(',', '.')
          link = item.find_element(By.XPATH, './a').get_attribute('href')

          products_data.append((product_name, data_atual, product_price, link))

          print(f'Product: {product_name}\n'  
                f'Date: {data_atual}\n'
                f'Price: R$ {product_price}\n'
                f'Link: {link}\n')


    except Exception as e:
      lg.error(f'Error occurred while searching for the product: {type(e).__name__} - {e}')
      print('An error occurred while extracting the product data')
      return None

    return products_data

  # save data to excel file
  def save_data_excel():
    products_data = search_product()
    if not products_data:
      return None
    
    file_path = 'products.xlsx'
    
    try:
      # check if the file exists
      if os.path.exists(file_path):
          wb = openpyxl.load_workbook(file_path)
          sheet = wb.active
          
          # clear the data in the file before saving the new data
          for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
              for cell in row:
                  cell.value = None

      else:
          # create a new file if it does not exist and add the headers
          wb = openpyxl.Workbook()
          sheet = wb.active
          sheet.append(['Product', 'Date', 'Price', 'Link'])
      
      for product in products_data:
          # save the data to the file
          sheet.append(product)

      
      wb.save('products.xlsx')
      print('Data saved successfully')
    except Exception as e:
      lg.error(f'Error occurred while saving data to excel: {type(e).__name__} - {e}')
      print('An error occurred while saving the data to the excel file')
      return None
    finally:
      wb.close()
    return True
  save_data_excel()
