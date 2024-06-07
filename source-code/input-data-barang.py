# Input Data Barang
# ================

# This script is used to input data barang into the system.

# Import required libraries
import time
import sys

# Selenium is used for web automation
from selenium import webdriver

# OpenPyXL is used for reading Excel files
from openpyxl import load_workbook

# Select is used for selecting options from dropdown menus
from selenium.webdriver.support.ui import Select

# By is used for specifying the type of element to find
from selenium.webdriver.common.by import By

# EC is used for specifying the expected conditions for an element
from selenium.webdriver.support import expected_conditions as EC

# WebDriverWait is used for waiting for an element to be available
from selenium.webdriver.support.ui import WebDriverWait

# TimeoutException is used for handling timeouts
from selenium.common.exceptions import TimeoutException

# Load the Excel workbook and select the sheet
# wb: The Excel workbook object
# sheetRange: The range of cells in the Excel sheet that contains the data
wb = load_workbook(filename="ADJUST WITH YOUR XLSX FILE (excel format) PATH")
sheetRange = wb['Sheet3']

# Create a Chrome webdriver instance
# driver: The Selenium WebDriver object
driver = webdriver.Chrome()

# Navigate to the login page and maximize the window
driver.get("http://")
driver.maximize_window()
driver.implicitly_wait(10)

# Find the username and password input fields
# username_input: The username input field element
# password_input: The password input field element
username_input = driver.find_element(By.ID, "username")
password_input = driver.find_element(By.ID, "password")

# Enter the username and password
username_input.send_keys("username")
password_input.send_keys("password")

# Create a Select object
# level_select: The level select element
select = Select(driver.find_element(By.ID, "level"))

# Select an option by value
select.select_by_value("superadmin")

# Find the login button and click it
# login_button: The login button element
login_button = driver.find_element(By.ID, "login")
login_button.click()

# Initialize a counter variable
# i: The counter variable
i = 2

while i <= len(sheetRange['A']):
    # Get the values from the current row
    # kode_barang: The kode barang value
    # nama_barang: The nama barang value
    # jenis_barang: The jenis barang value
    # satuan_barang: The satuan barang value
    kode_barang = sheetRange['A' + str(i)].value
    nama_barang = sheetRange['B' + str(i)].value
    jenis_barang = sheetRange['C' + str(i)].value
    satuan_barang = sheetRange['D' + str(i)].value

    try:
        # Navigate to the tambahgudang page
        driver.get("http://")

        # Fill in the form fields
        # kode_barang_input: The kode barang input field element
        # nama_barang_input: The nama barang input field element
        # jenis_barang_input: The jenis barang input field element
        # satuan_barang_input: The satuan barang input field element
        kode_barang_input = driver.find_element(By.ID, "kode_barang")
        nama_barang_input = driver.find_element(By.ID, "nama_barang")
        jenis_barang_input = driver.find_element(By.ID, "jenis_barang")
        satuan_barang_input = driver.find_element(By.ID, "satuan_barang")

        kode_barang_input.send_keys(kode_barang)
        nama_barang_input.send_keys(nama_barang)
        jenis_barang_input.send_keys(jenis_barang)
        satuan_barang_input.send_keys(satuan_barang)

        # Click the simpan button
        # simpan_button: The simpan button element
        simpan_button = driver.find_element(By.ID, "simpan")
        simpan_button.click()

        # Wait for 1 second before proceeding
        time.sleep(1)
    except:
        print('Form tidak muncul')
        sys.exit(0)

    # Increment the counter
    i += 1

print("Data Selesai Di Eksekusi!!!!!!!")