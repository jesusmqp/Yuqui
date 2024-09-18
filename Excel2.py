import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import signal
import sys
import logging
import os

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Variable global para el driver
driver = None

def signal_handler(sig, frame):
    logging.info('Interrupción detectada. Cerrando el navegador...')
    if driver:
        driver.quit()
    sys.exit(0)

# Registrar el manejador de señales
signal.signal(signal.SIGINT, signal_handler)

def login_to_admin():
    driver.get('https://appyuquitas.compas.studio/admin')
    username = driver.find_element(By.NAME, 'username')
    password = driver.find_element(By.NAME, 'password')

    username.send_keys('admin-yucax86')  # Cambia 'tu_usuario' por tu nombre de usuario de Django Admin
    password.send_keys('),S4Pq$334Bz')  # Cambia 'tu_contraseña' por tu contraseña de Django Admin
    password.send_keys(Keys.RETURN)
    logging.info("Inicio de sesión completado")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

def select_brand(brand_name):
    brand_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'lookup_id_brand')))
    brand_field.click()
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
    driver.switch_to.window(driver.window_handles[1])
    search_brand = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'searchbar')))
    search_brand.send_keys(brand_name)
    search_brand.send_keys(Keys.RETURN)
    result_link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.result-list a')))
    result_link.click()
    driver.switch_to.window(driver.window_handles[0])

def add_product(row):
    driver.get('https://appyuquitas.compas.studio/admin/products/product/add/')
    logging.info(f"Añadiendo producto: {row['title']}")

    select_brand(row['brand'])
    
    category_select = Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'category'))))
    category_select.select_by_value('ea5d8525-e79a-4876-a189-f528a622ca23')
    
    nombre_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'name')))
    nombre_field.send_keys(row['title'])
    
    # Aquí puedes añadir más campos del producto si es necesario

    guardar_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, '_save')))
    guardar_btn.click()
    
    logging.info(f"Producto {row['title']} añadido")

def add_store_to_product(product_row, store_row):
    # Asumiendo que estamos en la página de edición del producto después de añadirlo
    logging.info(f"Añadiendo tienda {store_row['store']} al producto {product_row['title']}")
    
    # Encuentra el botón para añadir una nueva tienda y haz clic en él
    add_store_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.add-row a')))
    add_store_btn.click()
    
    # Rellena los campos para la nueva tienda
    store_select = Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'productstore_set-0-store'))))
    store_select.select_by_visible_text(store_row['store'])
    
    url_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'productstore_set-0-url')))
    url_field.send_keys(store_row['url'])
    
    code_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'productstore_set-0-code')))
    code_field.send_keys(store_row['code'])
    
    price_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'productstore_set-0-price')))
    price_field.send_keys(str(store_row['price']))
    
    # Guarda los cambios
    save_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, '_save')))
    save_btn.click()
    
    logging.info(f"Tienda {store_row['store']} añadida al producto {product_row['title']}")

def process_excel_files(directory):
    for filename in os.listdir(directory):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(directory, filename)
            df = pd.read_excel(file_path)
            logging.info(f"Procesando archivo: {filename}")
            
            for index, row in df.iterrows():
                try:
                    add_product(row)
                    # Asumiendo que cada fila tiene información de una tienda
                    add_store_to_product(row, row)
                except Exception as e:
                    logging.error(f"Error procesando fila {index} en {filename}: {str(e)}")
                    driver.save_screenshot(f"error_screenshot_{filename}_{index}.png")

def main():
    global driver
    try:
        driver = webdriver.Chrome()
        logging.info("Navegador Chrome iniciado")
        
        login_to_admin()
        
        excel_directory = 'C:\\Users\\Miguel\\Documents\\Yuqui\\metro.xlsx'  # Ajusta esta ruta
        process_excel_files(excel_directory)

    except Exception as e:
        logging.error(f"Error general: {str(e)}")
    finally:
        if driver:
            driver.quit()
        logging.info("Script finalizado")

if __name__ == "__main__":
    main()