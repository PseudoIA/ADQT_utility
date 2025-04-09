import pickle
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from datetime import datetime, timezone,timedelta
import tkinter as tk
import threading
import json
import pygetwindow as gw
from selenium.common.exceptions import WebDriverException
from create_driver import create_driver
import logging
import win32gui
import win32con


logging.basicConfig(
    filename="application.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] [%(funcName)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def log_event(level, function_name, message):
    """
    Registra eventos en el archivo de log y los imprime en consola.
    """
    log_message = f"[{function_name}] {message}"
    if level == "info":
        logging.info(log_message)
    elif level == "warning":
        logging.warning(log_message)
    elif level == "error":
        logging.error(log_message)
    
    print(log_message)  


class create_case:
 
    def __init__(self, root):
        self.driver = None
        self.root = root
        self.options = None
        self.case_lock = threading.Lock()
        self.is_creating = False 
        self.open_case = None 
        self.driver_initialized = threading.Event()
        self.stop_thread = False
    

    def ensure_driver_visibility(self):
        """
        Verifica si el navegador está maximizado y en primer plano.
        Si no, lo maximiza y lo trae al frente.
        """
        try:
            hwnd = self.driver.current_window_handle  # Obtiene el identificador de la ventana
            hwnd = win32gui.FindWindow(None, self.driver.title)  # Busca la ventana por título
            
            if hwnd:

                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                win32gui.SetForegroundWindow(hwnd)

                log_event("info", "ensure_driver_visibility", "Navegador maximizado y en primer plano.")
            else:
                log_event("warning", "ensure_driver_visibility", "No se pudo encontrar la ventana del navegador.")
        
        except Exception as e:
            log_event("error", "ensure_driver_visibility", f"Error al gestionar visibilidad del navegador: {e}")

    def create_case(self,item_type, name_user, name_service, name_category, location, group, registryType, affair, origin, impact, urgency):

        if self.case_lock.locked():
            log_event("warning", "create_case", "Ya se está creando un caso. Esperando...")

            def check_lock():
                time.sleep(120)  
                if self.case_lock.locked():  
                    log_event("error", "check_lock", "Tiempo de espera excedido. Reiniciando el driver.")
                    self.stop_driver_thread()
                    self.open_driver()

                    if self.driver:
                        return self.create_case(item_type, name_user, name_service, name_category, location, group, registryType, affair, origin, impact, urgency)
                    else:
                        log_event("error", "check_lock", "Error crítico: No se pudo reiniciar el driver.")
                        return

            threading.Thread(target=check_lock, daemon=True).start()
            return

        try:
            if self.open_case:
                _ = self.driver.title 
                self.driver.refresh()  
                self.ensure_driver_visibility() 

        except WebDriverException as e:
            log_event("error", "create_case", f"Error del WebDriver: {e}")
            self.stop_driver_thread()  
            self.open_driver()  

            if self.driver:  
                log_event("info", "create_case", "Driver reiniciado correctamente. Reintentando operación...")
                return self.create_case(item_type, name_user, name_service, name_category, location, group, registryType, affair, origin, impact, urgency)
            else:
                log_event("error", "create_case", "Error crítico: No se pudo reiniciar el driver.")
                return

        if self.open_case:
            self.driver.refresh()

        with self.case_lock:
            try:
                log_event("info", "create_case", f"Iniciando creación de caso para: {name_user}")
                self.create_case_steps(self,item_type, name_user, name_service, name_category, location, group, registryType, affair, origin, impact, urgency)
                self.open_case = True
                log_event("info", "create_case", "Caso creado exitosamente.")

            except Exception as e:
                print(e)
                log_event("error", "create_case", f"Error al crear el caso: {e}")

                try:
                    self.driver.refresh()
                    log_event("info", "create_case", "Reintentando crear el caso después de refrescar el driver...")
                    self.create_case_steps(item_type, name_user, name_service, name_category, location, group, registryType, affair, origin, impact, urgency)
                    
                    self.open_case = True
                    log_event("info", "create_case", "Caso creado exitosamente en el segundo intento.")

                except Exception as e:
                    print(e)
                    log_event("error", "create_case", f"Error después del segundo intento: {e}")
                    self.driver.quit()  
            finally:
                self.is_creating = False



    def create_case_steps(self,item_type, name_user, name_service, name_category, location, group, registryType, affair, origin, impact, urgency):
        self.is_creating = True
        log_event("info", "create_case_steps", f"Iniciando la creación del caso para el usuario: {name_user}")  

        try:
            window = gw.getWindowsWithTitle("Aranda Service Management - Google Chrome")[0]
            if not window.isMaximized:
                window.restore()  
                window.maximize()
            window.activate() 
        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo forzar la ventana al frente: {e}")


        try:
            self.click_elements("//*[@data-testid='create-button-case']")
            log_event("info", "create_case_steps", "Se hizo clic en el botón de crear caso.")

        except Exception as e:
            log_event("error", "create_case_steps", f"Error al hacer clic en el botón de crear caso: {e}")
            time.sleep(1)  # Esperar un segundo y reintentar

 
        try:    
            self.click_elements("//span[@id='itemType']//button")
            log_event("info", "create_case_steps", "Se seleccionó el tipo de ítem.")
        except Exception as e:
            log_event("error", "create_case_steps", f"Error al seleccionar itemType: {e}")

        try:
            project_input = self.driver.find_element(By.ID, "project")
            current_value = project_input.get_attribute("value").strip()

            if current_value == "CO - COMFENALCO ANTIOQUIA":
                log_event("info", "lists_with_options", "El proyecto ya está seleccionado, no es necesario cambiarlo.")
            else:
                self.lists_with_options("project", ".//li", "CO - COMFENALCO ANTIOQUIA", "aria-owns")
                log_event("info", "create_case_steps", f"Se seleccionó el proyecto: CO - COMFENALCO ANTIOQUIA.")
        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo seleccionar itemType: {e}")


        try:
            self.lists_with_options("itemType", ".//li", item_type, '//ul[@role="listbox"]/li')
            log_event("info", "create_case_steps", "Se seleccionó el ítem tipo 'Requerimientos de Servicio'.")
        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo seleccionar itemType: {e}")

        try:
            self.lists_with_options("service", ".//li", name_service, "aria-owns")
            log_event("info", "create_case_steps", f"Se seleccionó el servicio: {name_service}.")

        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo seleccionar service: {e}")
          
        try:
            self.lists_with_options("category", "./td", name_category, "//tbody[@class='k-table-tbody']/tr")
            log_event("info", "create_case_steps", f"Se seleccionó la categoría: {name_category}.")
        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo seleccionar category: {e}")
        
        try:
            self.lists_with_options("applicant", ".//li", name_user, "aria-owns")
            log_event("info", "create_case_steps", f"Se seleccionó el solicitante: {name_user}.")
        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo seleccionar applicant: {e}")
        

        try:
            self.lists_with_options("customer", ".//li", name_user, "aria-owns")
            log_event("info", "create_case_steps", f"Se seleccionó el solicitante: {name_user}.")
        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo seleccionar applicant: {e}")

        try:
            self.lists_with_options("location", ".//li", location, "aria-owns")
            log_event("info", "create_case_steps", f"Se seleccionó la ubicación: {location}.")
        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo seleccionar group: {e}")

        try:
            self.lists_with_options("group", ".//li", group, "aria-owns")
            log_event("info", "create_case_steps", f"Se seleccionó el grupo: {group}.")
        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo seleccionar group: {e}")

        try:
            self.lists_with_options("registryType", ".//li", registryType, "aria-owns")
            log_event("info", "create_case_steps", f"Se seleccionó el tipo de registro: {registryType}.")
        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo seleccionar registryType: {e}")

        try:
            self.add_text_to_editable_div(affair)
            log_event("info", "create_case_steps", f"Se agregó el texto del asunto: {affair}.")
        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo agregar el texto del asunto: {e}")

        try:
            self.lists_with_options("44265", ".//li", origin, "aria-owns")
            log_event("info", "create_case_steps", f"Se seleccionó el origen: {origin}.")
        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo seleccionar origin: {e}")

        try:
            self.lists_with_options("impact", ".//li", impact, "aria-owns")
            log_event("info", "create_case_steps", f"Se seleccionó el impacto: {impact}.")
        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo seleccionar impact: {e}")

        try:
            self.lists_with_options("urgency", ".//li", urgency, "aria-owns")
            log_event("info", "create_case_steps", f"Se seleccionó la urgencia: {urgency}.")
        except Exception as e:
            log_event("warning", "create_case_steps", f"No se pudo seleccionar urgency: {e}")


    def lists_with_options(self, name_id, name_xpath_options, key_send, value_atribute):
        driver = self.driver
        
        try:
            if name_id == "category":
                log_event("info", "lists_with_options", f"Iniciando selección para 'category' con clave: {key_send}")

       
                element = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.ID, name_id))
                )


                start_time = time.time()
                while element.get_attribute("aria-disabled") == "true":
                    log_event("warning", "lists_with_options", "El campo 'category' sigue deshabilitado. Esperando...")
                    time.sleep(0.5) 

        
                    if time.time() - start_time > 5:
                        log_event("warning", "lists_with_options", "Forzando habilitación con JavaScript.")
                        driver.execute_script("arguments[0].removeAttribute('disabled')", element)
                        break  

               
                element.click()

             
                rows = WebDriverWait(driver, 5).until(
                    EC.presence_of_all_elements_located((By.XPATH, value_atribute))
                )

              
                log_event("debug", "lists_with_options", "HTML de las opciones:")
                selected = False 

                for index, row in enumerate(rows):
                    try:
                     
                        option_html = row.get_attribute("outerHTML")
                        log_event("debug", "lists_with_options", f"Opción {index + 1}: {option_html}")

                       
                        cell_text = row.find_element(By.XPATH, name_xpath_options).text
                        log_event("debug", "lists_with_options", f"Texto de la opción {index + 1}: {cell_text}")

            
                        if key_send in cell_text:
                            row.click()
                            log_event("info", "lists_with_options", f"Se seleccionó la categoría: {key_send}")
                            selected = True
                            break
                    except Exception as e:
                        log_event("error", "lists_with_options", f"Error al procesar la opción {index + 1}: {e}")


                if not selected:
                    log_event("error", "lists_with_options", "No se encontró la categoría en la lista.")
                

                log_event("info", "lists_with_options", "Confirmada la selección en la UI.")

       
                btn_aceptar = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Aceptar']]"))
                )
                btn_aceptar.click()
                log_event("info", "click_elements", "Se hizo clic en 'Aceptar' después de verificar la selección.")

            elif name_id == "itemType":
                print(key_send)
                log_event("info", "lists_with_options", f"Iniciando selección para 'itemType' con clave: {key_send}")
                elements = WebDriverWait(driver, 20).until(
                    EC.visibility_of_all_elements_located((By.XPATH, value_atribute))
                )
                for option in elements:
                    if option.text.strip()== key_send:
                        option.click()
                        
            elif name_id == "applicant":
                log_event("info", "lists_with_options", f"Iniciando búsqueda para 'applicant' con clave: {key_send}")
                options = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, name_id)))
                options.clear()
                options.send_keys(key_send)
                attribe_value = options.get_attribute(f'{value_atribute}')
                listbox = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"//*[@id='{attribe_value}']")))

                options = listbox.find_elements(By.XPATH, name_xpath_options)
                if not options:
                    log_event("warning", "lists_with_options", "No se encontraron opciones, intentando con 'Operación Sofi'.")
                    options.clear()
                    options.send_keys("Operación Sofi")

                    listbox = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, listbox_id))
                    )
                    options = listbox.find_elements(By.XPATH, name_xpath_options)
                    if options:
                        options[0].click()
                        log_event("info", "lists_with_options", "Se seleccionó 'Operación Sofi'.")
                    else:
                        log_event("warning", "lists_with_options", "No se encontró la opción 'Operación Sofi'.")
                else:
                    options[0].click()
                    log_event("info", "lists_with_options", f"Se seleccionó el solicitante: {key_send}")

            elif name_id == "customer":
                log_event("info", "lists_with_options", f"Iniciando búsqueda para 'customer' con clave: {key_send}")
                element = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID, name_id))
                )
                
                current_value = element.get_attribute("value").strip()
                if current_value == "":
                    element.clear()
                    element.send_keys(Keys.SPACE)
                    listbox_id = element.get_attribute(value_atribute)
                
                    listbox = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, listbox_id))
                    )
                
                    options = listbox.find_elements(By.XPATH, name_xpath_options)
                
                    # Busca la opción exacta en la lista
                    exact_option = next((o for o in options if o.text.strip() == key_send.strip()), None)

                    if exact_option:
                        exact_option.click()
                        log_event("info", "lists_with_options", f"Se seleccionó el cliente: {key_send}")
                    else:
                        if options:
                            options[0].click()
                            log_event("info", "lists_with_options", f"No se encontró coincidencia exacta; se seleccionó la primera opción: {options[0].text.strip()}")
                        else:
                            log_event("warning", "lists_with_options", f"No hay opciones disponibles para: {key_send}")

                else:
                    log_event("info", "lists_with_options", "El valor del campo 'customer' ya está establecido, no se realiza búsqueda.")


            else:
                log_event("info", "lists_with_options", f"Iniciando búsqueda general con clave: {key_send}")
                options = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, name_id)))
                options.click()
                options.send_keys(Keys.SPACE)

                attribe_value = options.get_attribute(f'{value_atribute}')
                listbox = WebDriverWait(driver, 600).until(EC.element_to_be_clickable((By.XPATH, f"//*[@id='{attribe_value}']")))

                options = listbox.find_elements(By.XPATH, name_xpath_options)

                selected = False

                for option in options:
                    option_text = option.text.strip()
                    if key_send == option_text:  # Coincidencia exacta
                        option.click()
                        selected = True
                        break

                if not selected:
                    for option in options:
                        if key_send in option.text:
                            option.click()
                            selected = True
                            break
                if not selected:
                    log_event("warning", "lists_with_options", f"No se encontró coincidencia para: {key_send}")

        except Exception as e:
            log_event("error", "lists_with_options", f"Error durante la selección de opciones: {e}")

           
    def click_elements(self, locator):
        driver = self.driver

        try:
            log_event("info", "click_elements", f"Iniciando búsqueda del elemento con el locator: {locator}")
            element = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, locator))
            )
            element.click()
            log_event("info", "click_elements", f"Elemento con el locator '{locator}' fue clickeado con éxito.")
        except Exception as e:
            log_event("error", "click_elements", f"Error al intentar hacer clic en el elemento con locator '{locator}': {e}")

    def add_text_to_editable_div(self, text):
        driver = self.driver

        try:
            log_event("info", "add_text_to_editable_div", "Iniciando el envío de texto al div editable.")

            wait = WebDriverWait(driver, 1)  # Reduce el tiempo de espera a 5s en lugar de los 10s por defecto.

            # Espera a que el iframe esté presente y luego cambia a él
            iframe = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'k-iframe')))
            driver.switch_to.frame(iframe)

            # Espera a que el div editable esté interactuable
            editable_div = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.k-content.ProseMirror')))
            editable_div.send_keys(text)

            log_event("info", "add_text_to_editable_div", f"Texto '{text}' enviado exitosamente al div editable.")

            driver.switch_to.default_content()

        except Exception as e:
            log_event("error", "add_text_to_editable_div", f"Error al intentar enviar texto al div editable: {e}")

    def validate_date_expiration(self):
        current_time = datetime.now(timezone(timedelta(hours=-5)))

        try:
            with open("cookies.pkl", "rb") as file:
                cookies = pickle.load(file)
            log_event("info", "validate_date_expiration", "Cookies cargadas correctamente desde 'cookies.pkl'.")
        except FileNotFoundError:
            log_event("error", "validate_date_expiration", "No se encontró el archivo 'cookies.pkl'.")
            return False
        except Exception as e:
            log_event("error", "validate_date_expiration", f"Error al cargar el archivo de cookies: {e}")
            return False

        expired = False

        for cookie in cookies:
            domain = cookie.get('domain', '')
            expiry = cookie.get('expiry')

            # Solo procesar cookies del dominio 'itsm.sonda.com' o '.itsm.sonda.com'
            if domain.endswith("itsm.sonda.com") and expiry:
                cookie_name = cookie.get('name', 'Unknown')
                
                expiry_date = datetime.fromtimestamp(expiry, timezone.utc)
                log_event("info", "validate_date_expiration", f"Fecha de expiración de la cookie '{cookie_name}' del dominio '{domain}': {expiry_date}")

                if current_time > expiry_date:
                    log_event("info", "validate_date_expiration", f"La cookie '{cookie_name}' ha expirado.")
                    expired = True
                else:
                    log_event("info", "validate_date_expiration", f"La cookie '{cookie_name}' aún no ha expirado.")

        if expired:
            try:
                os.remove("cookies.pkl")
                log_event("info", "validate_date_expiration", "El archivo 'cookies.pkl' ha sido eliminado debido a cookies expiradas.")
            except Exception as e:
                log_event("error", "validate_date_expiration", f"Error al eliminar el archivo 'cookies.pkl': {e}")
            return False

        log_event("info", "validate_date_expiration", "Las cookies no han expirado.")
        return True


    def open_driver(self):
        try:
            self.options = webdriver.ChromeOptions()
            
            self.options.add_argument("--disable-background-timer-throttling")
            self.options.add_argument("--disable-backgrounding-occluded-windows")
            self.options.add_argument("--disable-renderer-backgrounding")
            self.options.add_argument("--start-maximized")
            self.driver = create_driver.setup_webdriver(options=self.options)
            
            log_event("info", "open_driver", "Driver de Chrome inicializado correctamente.")
            log_event("info", "open_driver", f"Driver details: {self.driver}")
            
            self.open_page()
            self.driver_initialized.set() 
            log_event("info", "open_driver", "Página abierta correctamente.")
        except Exception as e:
            log_event("error", "open_driver", f"Error al inicializar el driver de Chrome: {e}")
            print(f"Error initializing Chrome driver: {e}")
            return

    def open_page(self):
        log_event("info", "open_page", "Intentando ingresar a la página...")

        if os.path.exists("cookies.pkl"):
            log_event("info", "open_page", "Existe archivo de cookies.")

            cookies_valid = self.validate_date_expiration()
            log_event("info", "open_page", "Validando fecha de expiración de las cookies...")

            self.driver.get("https://itsm.sonda.com/asmsspecialist/index.html#/home/list")  

            if cookies_valid:
                with open("cookies.pkl", "rb") as file:
                    cookies = pickle.load(file)
                
                for cookie in cookies:
                    if "domain" in cookie and (cookie["domain"] == "itsm.sonda.com" or cookie["domain"] == ".itsm.sonda.com"):
                        try:
                            self.driver.add_cookie(cookie)
                            log_event("info", "open_page", f"Cookie añadida: {cookie}")
                        except Exception as e:
                            log_event("error", "open_page", f"No se pudo agregar cookie: {cookie} - {e}")
                
                self.driver.get("https://itsm.sonda.com/asmsspecialist/index.html#/home/list")

                try:
                    WebDriverWait(self.driver, 3).until(
                        EC.url_to_be("https://itsm.sonda.com/asmsspecialist/index.html#/home/list"))
                    log_event("info", "open_page", "URL final confirmada: https://itsm.sonda.com/asmsspecialist/index.html#/home/list")
                except Exception as e:
                    log_event("error", "open_page", f"La URL final no es la esperada. Cookies posiblemente inválidas. Error: {e}")
                    os.remove("cookies.pkl")
                    self.authenticate_and_store_cookies()

            else:
                log_event("info", "open_page", "Cookies no válidas, procediendo a autenticación.")
                self.authenticate_and_store_cookies()  

        else:
            log_event("info", "open_page", "No existe archivo de cookies, procediendo a autenticación.")
            self.authenticate_and_store_cookies()

    def authenticate_and_store_cookies(self):
        
        if os.path.exists("cookies.pkl"):
            try:
                os.remove("cookies.pkl")

                print("Archivo eliminado correctamente.")
            except Exception as e:
                print(f"Error al eliminar el archivo: {e}")

        log_event("info", "authenticate_and_store_cookies", "Iniciando autenticación y almacenamiento de cookies.")

        self.driver.get("https://itsm.sonda.com/asmsspecialist/index.html#/") 
        log_event("info", "authenticate_and_store_cookies", "Esperando autenticación de doble factor...")

        visited_urls = set()  
        all_cookies = []  

        try:
            while True:
                current_url = self.driver.current_url

                if current_url not in visited_urls:
                    visited_urls.add(current_url)
                    log_event("info", "authenticate_and_store_cookies", f"Visitando: {current_url}")

                    cookies = self.driver.get_cookies()

                    if cookies:  # Solo guardar si hay cookies
                        all_cookies.extend(cookies)  # Agregar sin modificar estructura

                        # Guardar cookies en un archivo JSON (sobrescribiendo cada vez)
                        with open("cookies.json", "w", encoding="utf-8") as file:
                            json.dump(all_cookies, file, indent=4, ensure_ascii=False)
                            log_event("info", "authenticate_and_store_cookies", "Cookies guardadas en formato JSON.")

                        # Guardar cookies en formato binario con Pickle
                        with open("cookies.pkl", "wb") as file:
                            pickle.dump(all_cookies, file)
                            log_event("info", "authenticate_and_store_cookies", "Cookies guardadas en formato Pickle.")

                # Verificar si ya estamos en la página final
                if current_url == "https://itsm.sonda.com/asmsspecialist/index.html#/home/list/":
                    log_event("info", "authenticate_and_store_cookies", "Autenticación completada. Guardando cookies finales...")
                    break

                time.sleep(2)  # Pequeña espera para evitar sobrecarga

            WebDriverWait(self.driver, 2)  # Esperar un momento más antes de cerrar
            log_event("info", "authenticate_and_store_cookies", "Proceso de autenticación y almacenamiento completado.")

        except Exception as e:
            log_event("error", "authenticate_and_store_cookies", f"No se completó la autenticación: {e}")


    def get_state_driver(self):
        log_event("info", "get_state_driver", "Verificando estado del driver...")

        # Espera hasta que el driver esté listo (con timeout de 10s para evitar bloqueos indefinidos)
        if not self.driver_initialized.wait(timeout=1):
            log_event("warning", "get_state_driver", "Tiempo de espera agotado para inicialización del driver.")
            self.driver_status = "cerrado"
            return self.driver_status

        if self.driver is None:
            self.driver_status = "cerrado"
            log_event("info", "get_state_driver", "Driver no encontrado. Estado: cerrado.")
            return self.driver_status
                
        else:
            try:
                _ = self.driver.title  # Verifica que el driver está activo
                self.driver_status = "abierto"
                log_event("info", "get_state_driver", f"Driver activo. Estado: {self.driver_status}")
            except WebDriverException:
                self.driver_status = "cerrado"
                self.stop_driver_thread()
                log_event("error", "get_state_driver", "Error con el driver. Estado: cerrado.")
            
        print(f"Estado del Driver: {self.driver_status}")
        return self.driver_status


    def stop_driver_thread(self):
        log_event("info", "stop_driver_thread", "Intentando detener el driver...")

        if self.driver is not None:
            try:
                self.driver.quit()
                log_event("info", "stop_driver_thread", "Driver detenido correctamente.")
                print("Driver detenido correctamente")

            except WebDriverException as e:
                log_event("error", "stop_driver_thread", f"Error crítico al detener el driver: {e}")
                print(f"Error crítico al detener el driver: {e}")
      
            except Exception as e:
                log_event("error", "stop_driver_thread", f"Error inesperado al detener el driver: {e}")
                print(f"Error inesperado al detener el driver: {e}")
            finally:
                self.driver = None 

                log_event("info", "stop_driver_thread", "Referencia al driver eliminada.")

        else:
            log_event("warning", "stop_driver_thread", "Intento de detener un driver que ya es None.")
            print("Intento de detener un driver que ya es None.")
            self.driver.quit()


"""if __name__ == "__main__":
    demo =create_case(root=None)
 
    demo.open_page()
    input("ingrese texto")"""

