#scraping-2

#script scraping

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
from webdriver_manager.firefox import GeckoDriverManager

# Configuración de Selenium
def init_driver():
    options = webdriver.FirefoxOptions()
    # Quitar el modo headless para depuración
    # options.add_argument("--headless")  # Ejecuta Firefox en modo sin cabeza.
    driver = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()), options=options)
    return driver

def login(driver, login_url):
    print("Navegando a la página de login...")
    driver.get(login_url)
    time.sleep(5)  # Esperar un poco para que cargue la página
    print(f"URL actual: {driver.current_url}")
    input("Ingresa las credenciales y presiona Enter para continuar...")

def click_button_with_js(driver, button_selector):
    try:
        print(f"Intentando hacer clic en el botón con el selector: {button_selector}")
        button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, button_selector))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", button)
        time.sleep(1)  # Esperar un momento después de desplazar
        
        # Intentar hacer clic usando JavaScript
        driver.execute_script("arguments[0].click();", button)
        print(f"Botón con el selector {button_selector} clicado exitosamente.")
    except Exception as e:
        print(f"No se pudo hacer clic en el botón con el selector {button_selector}: {e}")


def close_modal_if_present(driver):
    try:
        print("Verificando si hay un modal presente...")
        # Intentar cerrar el modal si hay un botón de cierre
        modal_close_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-success"))
        )
        modal_close_button.click()
        print("Modal cerrado.")
    except Exception as e:
        print("No se encontró un modal o no se pudo cerrar:", e)

def apply_filters(driver):
    
    # Asegurarse de que el botón esté visible y hacer clic en él usando JavaScript
    click_button_with_js(driver, "button.btn.btn-success.btn-sm")

    # Asegurarse de que el botón "OK" en el modal de filtros esté visible y hacer clic en él
    click_button_with_js(driver, "button.btn.btn-success")  # Reemplaza con el selector correcto si es diferente
  
    
    # Cerrar el modal si está presente
    close_modal_if_present(driver)     

    
    
def select_order_option(driver):
    print("Seleccionando la opción de ordenes con productos...")
    # Esperar y hacer clic en el menú desplegable
    dropdown_toggle = WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "mdi.mdi-chevron-down"))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", dropdown_toggle)
    time.sleep(2)  # Espera adicional para asegurar que el elemento esté listo

    # Mover el cursor al elemento antes de hacer clic
    ActionChains(driver).move_to_element(dropdown_toggle).perform()
    time.sleep(1)  # Espera adicional para asegurar que el elemento esté listo

    #hacer click en menú desplegable
    dropdown_toggle.click()

    # Seleccionar la opción "Ordenes con productos (Un producto por fila)"
    orders_option = WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "button.dropdown-item[title='Ordenes con productos (Un producto por fila)']"))
    )

    # Desplazar el elemento a la vista si es necesario
    driver.execute_script("arguments[0].scrollIntoView(true);", orders_option)
    time.sleep(1)  # Esperar un momento después de desplazar

    # Hacer clic usando JavaScript para evitar cualquier interferencia
    driver.execute_script("arguments[0].click();", orders_option)

def download_excel(driver):
    time.sleep(10)  # Esperar a que el archivo se descargue
    print("Archivo Excel descargado.")

def main():
    try:
        driver = init_driver()
        login_url = 'https://app.dropi.cl/auth/login'
        orders_url = 'https://app.dropi.cl/dashboard/orders/supplier'
        
        login(driver, login_url)
        
        print("Redirigiendo a la página de órdenes de pedido...")
        driver.get(orders_url)
        time.sleep(5)  # Esperar un poco para que cargue la página
        print(f"URL actual: {driver.current_url}")
        
        apply_filters(driver)
        select_order_option(driver)
        download_excel(driver)
    
    except Exception as e:
        print(f"Error: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()

