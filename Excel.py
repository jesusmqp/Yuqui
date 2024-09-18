import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Configurar el controlador de Selenium (Chrome en este caso)
driver = webdriver.Chrome()  # Asegúrate de tener ChromeDriver instalado y en el PATH

# Cargar el archivo Excel con la ruta correcta en Windows
df = pd.read_excel('C:\\Users\\Miguel\\Documents\\Yuqui\\metro.xlsx')  # Cambia esta ruta por la correcta en tu sistema

# Abre la página de administración de Django
driver.get('https://appyuquitas.compas.studio/admin')

# Inicia sesión
username = driver.find_element(By.NAME, 'username')
password = driver.find_element(By.NAME, 'password')

username.send_keys('admin-yucax86')  # Cambia 'tu_usuario' por tu nombre de usuario de Django Admin
password.send_keys('),S4Pq$334Bz')  # Cambia 'tu_contraseña' por tu contraseña de Django Admin
password.send_keys(Keys.RETURN)

# Espera para asegurarse de que la página ha cargado completamente
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

# Iterar sobre cada fila del DataFrame de pandas y rellenar el formulario
for index, row in df.iterrows():
    # Navegar a la página de agregar producto
    driver.get('https://appyuquitas.compas.studio/admin/products/product/add/')

    try:
        # Selección de Marca (Brand) usando el buscador
        brand_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'lookup_id_brand')))
        brand_field.click()

        # Cambiar al nuevo popup de búsqueda de marca
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(driver.window_handles[1])

        # Buscar la marca que corresponde a la columna del Excel 'brand'
        search_brand = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'searchbar')))
        search_brand.send_keys(row['brand'])  # Ajusta al nombre de la columna que contiene la marca
        search_brand.send_keys(Keys.RETURN)

        # Espera que cargue los resultados y selecciona el primero
        result_link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.result-list a')))
        result_link.click()

        # Volver al popup original
        driver.switch_to.window(driver.window_handles[0])

        # Seleccionar la categoría fija
        category_select = Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'category'))))
        category_select.select_by_value('ea5d8525-e79a-4876-a189-f528a622ca23')  # Asegúrate de que este valor sea correcto

        # Localizar los campos del formulario y rellenarlos
        nombre_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'name')))
        url_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'productstore_set-0-url')))
        precio_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'productstore_set-0-code')))
        
        nombre_field.send_keys(row['title'])  # Usa el nombre correcto de la columna del Excel
        url_field.send_keys(row['url'])
        precio_field.send_keys(str(row['price']))  # Asegúrate de que el precio esté en el formato adecuado

        # Guardar el producto
        guardar_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, '_save')))
        guardar_btn.click()

    except (TimeoutException, NoSuchElementException) as e:
        print(f"Error en la fila {index}: {str(e)}")
        continue  # Salta a la siguiente iteración del bucle

# Cerrar el navegador después de terminar
driver.quit()