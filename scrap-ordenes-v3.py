import pandas as pd
import os
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
from webdriver_manager.firefox import GeckoDriverManager
from datetime import datetime
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime
import locale


locale.setlocale(locale.LC_TIME, "es_ES.UTF-8")  # Establecer la localización en español (esto puede variar según el sistema operativo)
# Configuración de Selenium
def init_driver():
    options = webdriver.FirefoxOptions()
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
        
        # Asegurarse de que el botón esté visible en la vista
        driver.execute_script("arguments[0].scrollIntoView(true);", button)
        WebDriverWait(driver, 5).until(EC.visibility_of(button))
        
        # Intentar hacer clic usando JavaScript
        driver.execute_script("arguments[0].click();", button)
        print(f"Botón con el selector {button_selector} clicado exitosamente.")
    except Exception as e:
        print(f"No se pudo hacer clic en el botón con el selector {button_selector}: {e}")

def is_filters_modal_open(driver):
    try:
        driver.find_element(By.CSS_SELECTOR, "div.modal-content")
        return True
    except NoSuchElementException:
        return False

def open_filters_modal(driver):
    if not is_filters_modal_open(driver):
        print("Abriendo la ventana de filtros...")
        try:
            click_button_with_js(driver, "button.btn.btn-success.btn-sm")
            WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "div.modal-content"))
            )
            print("Ventana de filtros abierta.")
        except Exception as e:
            print(f"Error al intentar abrir el modal de filtros: {e}")
    else:
        print("La ventana de filtros ya está abierta.")

def open_filters_modal_once(driver):
    if not is_filters_modal_open(driver):
        try:
            print("Abriendo la ventana de filtros...")
            click_button_with_js(driver, "button.btn.btn-success.btn-sm")
            WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "div.modal-content"))
            )
            print("Ventana de filtros abierta.")
        except Exception as e:
            print(f"Error al intentar abrir el modal de filtros: {e}")
    else:
        print("La ventana de filtros ya está abierta.")

def select_date(driver, calendar_button_selector, prev_button_selector, target_month, date_selector):
    try:
        # Abrir el calendario usando JavaScript
        click_button_with_js(driver, calendar_button_selector)
        print(f"Calendario abierto usando el selector: {calendar_button_selector}")

        # Navegar hasta el mes objetivo
        while True:
            current_month = driver.find_element(By.CSS_SELECTOR, "div.mat-calendar-controls").text
            print(f"Mes actual visible en el calendario: {current_month}")

            if current_month == target_month:
                print(f"Se alcanzó el mes objetivo: {target_month}")
                break
            else:
                prev_button = driver.find_element(By.CSS_SELECTOR, prev_button_selector)
                driver.execute_script("arguments[0].click();", prev_button)
                time.sleep(1)

        # Seleccionar la fecha
        date_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, date_selector))
        )
        driver.execute_script("arguments[0].click();", date_element)
        print(f"Fecha seleccionada con éxito usando el selector: {date_selector}")
    except Exception as e:
        print(f"Error al seleccionar la fecha: {e}")

def select_dates(driver):
    try:
        # Seleccionar la fecha "desde"
        select_date(driver,
                    'span.mat-mdc-button-touch-target',  # Botón de calendario "de"
                    'button.mat-calendar-previous-button.mdc-icon-button.mat-mdc-icon-button.mat-unthemed.mat-mdc-button-base',  # Botón para navegar meses
                    'DIC DE 2023',  # Mes objetivo
                    'button.mat-calendar-body-cell[aria-label="1 de diciembre de 2023"]')  # Día 1 de diciembre 2023
               
        # Confirmar las fechas seleccionadas
        confirm_modal(driver, "button.btn.btn-success")
        # Esperar a que el modal se cierre
        #wait_for_modal_close(driver, "div.modal-content")

    except Exception as e:
        print(f"Error al seleccionar las fechas: {e}")



def remove_overlay(driver):
    try:
        overlay = driver.find_element(By.CSS_SELECTOR, ".cdk-overlay-backdrop")
        driver.execute_script("arguments[0].style.display = 'none';", overlay)
        print("Overlay eliminado con éxito.")
    except NoSuchElementException:
        print("No se encontró el overlay. Continuando sin eliminar.")




def confirm_modal(driver, selector):
    try:
        confirm_button = driver.find_element(By.CSS_SELECTOR, selector)
        confirm_button.click()
        print(f"Modal confirmado con éxito usando el selector: {selector}")
    except Exception as e:
        print(f"Error al confirmar el modal: {e}")

def wait_for_modal_close(driver, modal_selector):
    print("Esperando a que el modal se cierre...")
    WebDriverWait(driver, 30).until(
        EC.invisibility_of_element((By.CSS_SELECTOR, modal_selector))
    )
    print("Modal cerrado.")

def select_order_option(driver):
    print("Seleccionando la opción de órdenes...")
    try:
        dropdown_toggle = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "a.dropdown-toggle.btn.btn-light"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", dropdown_toggle)
        time.sleep(2)  # Espera adicional para asegurar que el elemento esté listo

        # Solucionar la intercepción del clic
        try:
            driver.execute_script("arguments[0].click();", dropdown_toggle)
        except ElementClickInterceptedError:
            print("El clic fue interceptado, intentando de nuevo después de desplazar el elemento a la vista.")
            driver.execute_script("arguments[0].scrollIntoView(true);", dropdown_toggle)
            driver.execute_script("arguments[0].click();", dropdown_toggle)

        orders_option = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "button.dropdown-item[title='Órdenes (Una orden por fila)']"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", orders_option)
        time.sleep(1)  # Esperar un momento después de desplazar

        driver.execute_script("arguments[0].click();", orders_option)
    except Exception as e:
        print(f"Error al seleccionar la opción de órdenes: {e}")

def confirm_report_modal(driver):
    print("Confirmando el modal de reporte...")
    click_button_with_js(driver, "button.swal2-confirm.swal2-styled")
    print("Modal de reporte confirmado.")

def download_report(driver):
    print("Descargando el reporte...")
    WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.btn.btn-outline-info.ng-star-inserted"))
    ).click()
    print("Reporte descargado.")

def update_master_table(folder_path, master_table_path):
    excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))

    if os.path.exists(master_table_path):
        master_table = pd.read_excel(master_table_path)
    else:
        master_table = pd.DataFrame()

    for file_path in excel_files:
        new_data = pd.read_excel(file_path)
        master_table = pd.concat([master_table, new_data]).drop_duplicates(subset=['ID'], keep='last')

    master_table.to_excel(master_table_path, index=False)
    print(f"Tabla maestra actualizada y guardada en {master_table_path}")

def countdown(seconds):
    while seconds > 0:
        mins, secs = divmod(seconds, 60)
        timer = '{:02d}:{:02d}'.format(mins, secs)
        print(timer, end="\r")
        time.sleep(120)
        seconds -= 1
    print("Tiempo de espera terminado!")

def main():
    try:
        driver = init_driver()
        driver.maximize_window()
        login_url = 'https://app.dropi.cl/auth/login'
        orders_url = 'https://app.dropi.cl/dashboard/orders/supplier'
        
        login(driver, login_url)
        
        print("Redirigiendo a la página de órdenes de pedido...")
        driver.get(orders_url)
        time.sleep(5)  # Esperar un poco para que cargue la página
        print(f"URL actual: {driver.current_url}")
        
        open_filters_modal_once(driver)  # Abrir el modal de filtros automáticamente

        # Seleccionar fechas en el calendario
        select_dates(driver)
        
        confirm_modal(driver, "button.btn.btn-success")  # Confirmar el modal de filtros

        wait_for_modal_close(driver, "div.modal-dialog")  # Esperar a que el modal de filtros se cierre

        select_order_option(driver)  # Seleccionar la opción de órdenes
        confirm_report_modal(driver)  # Confirmar el modal de reporte
        download_report(driver)  # Descargar el reporte

        folder_path = "C:/Users/Jortizme/Desktop/PROYECTOS PERSONALES/Liquidambar/excel/"
        master_table_path = "C:/Users/Jortizme/Desktop/PROYECTOS PERSONALES/Liquidambar/tabla_maestra.xlsx"
        update_master_table(folder_path, master_table_path)
    
    except NoSuchElementException as e:
        print(f"Error: No se encontró un elemento: {e}")
    except Exception as e:
        print(f"Error general: {e}")
    finally:
        driver.quit()


if __name__ == "__main__":
    while True:
        main()
        countdown(1)


