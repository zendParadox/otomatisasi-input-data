import time
import sys

from selenium import webdriver
from openpyxl import load_workbook
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException

wb = load_workbook(filename="../automatisasi-input-data/automatisasi-data/dataset/data-barang-masuk.xlsx")
sheetRange = wb['Sheet3']

driver = webdriver.Chrome()
driver.get("http://localhost/inventory/login.php")
driver.maximize_window()
driver.implicitly_wait(10)

# Find the username and password input fields
username_input = driver.find_element(By.ID, "username")
password_input = driver.find_element(By.ID, "password")
level_select = driver.find_element(By.ID, "level")
login_button = driver.find_element(By.ID, "login")

# Enter the username and password
username_input.send_keys("zendparadoxxx")
password_input.send_keys("admin123")
# Create a Select object
select = Select(level_select)
# Select an option by value
select.select_by_value("superadmin")  # Replace with the desired value (e.g. "superadmin", "admin", "petugas")

# Find the login button and click it
login_button.click()



i = 2
# driver.get("http://localhost/inventory/index3.php?page=gudang&aksi=tambahgudang")

while i <= len(sheetRange['A']):
  kode_barang = sheetRange['A' + str(i)].value
  nama_barang = sheetRange['B' + str(i)].value
  jenis_barang = sheetRange['C' + str(i)].value
  satuan_barang = sheetRange['D' + str(i)].value

  try:
    # WebDriverWait(driver, 10).until(EC.visibility_of_element_located(By.))
    driver.get("http://localhost/inventory/index3.php?page=gudang&aksi=tambahgudang")
    driver.find_element(By.ID, "kode_barang").send_keys(kode_barang)
    driver.find_element(By.ID, "nama_barang").send_keys(nama_barang)
    driver.find_element(By.ID, "jenis_barang").send_keys(jenis_barang)
    driver.find_element(By.ID, "satuan_barang").send_keys(satuan_barang)
    driver.find_element(By.ID, "simpan").click()
    
    time.sleep(1)
  except:
    print('Form tidak muncul')
    sys.exit(0)
  i += 1

print("Data Selesai Di Eksekusi!!!!!!!")