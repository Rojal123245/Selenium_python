# Selenium Web Scraping Script Documentation

This documentation provides a detailed explanation of the script used to scrape data from a website using Selenium and update an Excel file with the retrieved data. The script is written in Python and utilizes various libraries to achieve the desired functionality.

## Prerequisites

Before running the script, ensure you have the following libraries installed:

- pandas
- selenium
- openpyxl

You can install these libraries using pip:

```bash
pip install pandas selenium openpyxl
```

Additionally, you will need to download the appropriate WebDriver for your browser. For example, if you are using Google Chrome, download the ChromeDriver from [here](https://googlechromelabs.github.io/chrome-for-testing/#stable) . Make sure the WebDriver is in your system's PATH or specify its location in the script.

Script Overview

The script performs the following tasks:

Loads data from an Excel file.
Uses Selenium to navigate to a website and retrieve data based on a PIN number.
Updates the Excel file with the retrieved data.
Importing the Required Libraries
```bash
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import openpyxl
Loading the Excel File
Specify the path to your Excel file and load it into a pandas DataFrame.
```

```bash
file_path = 'data.xlsx'
df = pd.read_excel(file_path) 
```
Initializing the Selenium WebDriver
Ensure you have the correct WebDriver for your browser (e.g., ChromeDriver for Google Chrome).

```bash
driver = webdriver.Chrome()  # You can use other drivers like FirefoxDriver or EdgeDrive
```


Defining the Data Retrieval Function
This function navigates to the website, enters the PIN number, and retrieves the necessary data.

```bash
def get_data_from_website(pin_number):
    driver.get("url")

    # Find the PIN input field and enter the PIN number
    pin_input = driver.find_element(By.NAME, "property_key")
    pin_input.clear()
    pin_input.send_keys(pin_number)
    pin_input.send_keys(Keys.RETURN)

    # Allow time for the page to load
    time.sleep(5)

    # Extract the required fields
    data = {}
    try:
        data['Tax Year'] = driver.find_element(By.XPATH, "//div[contains(text(), 'Tax Year')]/following-sibling::div/div").text
        data['Net Taxable Value'] = driver.find_element(By.XPATH, "//div[contains(text(), 'Net Taxable Value')]/following-sibling::div").text
        data['Tax status'] = driver.find_element(By.XPATH, "//div[contains(text(), 'Tax Status')]/following-sibling::div").text
        data['Comments'] = driver.find_element(By.XPATH, "//div[contains(text(), 'Owner Name & Address')]/following-sibling::div").text
        data['TAXCODE'] = driver.find_element(By.XPATH, "//div[contains(text(), 'Tax Code')]/following-sibling::div").text
        data['name of owner'] = driver.find_element(By.XPATH, "//div[contains(text(), 'Owner Name & Address')]/following-sibling::div").text
    except Exception as e:
        print(f"Error while extracting data for PIN {pin_number}: {e}")
    return data
```
Looping Through the Excel Data and Updating It
The script loops through each PIN in the Excel file, retrieves the data from the website, and updates the DataFrame.

```bash
for index, row in df.iterrows():
    pin_number = row['PIN']
    if pd.isna(pin_number) or pin_number == '':
        print("Empty PIN encountered, stopping the loop.")
        break
    if pd.notna(row['Tax Year']):
        print(f"Data already present for PIN {pin_number}, skipping.")
        continue
    data = get_data_from_website(pin_number)
    print(data)
    for key, value in data.items():
        if key in df.columns:
            df.at[index, key] = value
            print(index)
        else:
            print("Not found")
```


Saving the Updated Excel File
After updating the DataFrame, save it back to the Excel file.

``` bash 
df.to_excel(file_path, index=False)
```
Closing the WebDriver
Finally, close the Selenium WebDriver.

```bash
driver.quit()
```
Running the Script

To run the script, simply execute it using Python:

```bash
python script_name.py
```
Make sure to replace script_name.py with the actual name of your Python script.


