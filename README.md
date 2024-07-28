# Project run-price-monitoring

- it's a Bot that get the data of the product of your wish and store in a Excel file

#### Funcionalities

1. Search for products: The user type the product of his wish.
2. Extract the data: Extract the data as it follows:  Product name, extraction date, product price and its link
3. Store data Excel: Get all the data extract of your product and stores in a excel file named: **products.xlsx**

#### Tecnologies

- `selenium`, `openpyxl`, `datetime`, `os`, `logging`

#### How to use:

- Install depedencies
  - git clone this repository
  - pip install -r requirementes.txt
  - run app.py

Run app.py and in the terminal type the name of the product and the magic happens, a file excel is created with all the information described above.

*OBS: The script will run every 30 min
