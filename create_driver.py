import os
import json
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import WebDriverException

class create_driver:
    # Nombre del archivo de validación
    VALIDATION_FILE = "validacion.json"
    LOG_FILE = "log.txt"  # Archivo para registrar eventos

    # Ruta donde debería estar el Chrome Portable y el WebDriver
    CHROME_FOLDER = "chrome-win64"

    @staticmethod
    def log_message(message):
        """Guarda un mensaje en log.txt con fecha y hora."""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(create_driver.LOG_FILE, "a") as log_file:
            log_file.write(f"[{timestamp}] {message}\n")

    @staticmethod
    def read_validation_file():
        """Lee el archivo de validación. Si no existe, lo crea con estado 'unknown'."""
        if not os.path.exists(create_driver.VALIDATION_FILE):
            create_driver.write_validation_file("unknown")  # Se crea con estado desconocido
        with open(create_driver.VALIDATION_FILE, "r") as file:
            return json.load(file)

    @staticmethod
    def write_validation_file(status):
        """Escribe el estado en el archivo de validación y lo registra en el log."""
        if isinstance(status, dict):  # Asegura que status no sea anidado
            data = status
        else:
            data = {"driver_status": status}

        with open(create_driver.VALIDATION_FILE, "w") as file:
            json.dump(data, file, indent=4)  # Agrega indentación para legibilidad

        create_driver.log_message(f"Estado actualizado en {create_driver.VALIDATION_FILE}: {data}")

    @staticmethod
    def get_local_chromedriver():
        """Obtiene la ruta del WebDriver local y Chrome Portable"""
        driver_path = os.path.join(create_driver.CHROME_FOLDER, "chromedriver.exe")
        chrome_path = os.path.join(create_driver.CHROME_FOLDER, "chrome.exe")

        if os.path.exists(driver_path) and os.path.exists(chrome_path):
            return driver_path, chrome_path
        return None, None

    @classmethod
    def setup_webdriver(cls, options=None):
        """Configura el WebDriver según el estado guardado en validacion.json"""
        validation = cls.read_validation_file()
        driver_status = validation.get("driver_status", "unknown")
        driver_path = validation.get("driver_path", None)  # Cargar ruta si ya se instaló

        if driver_status == "ChromeDriverManager_si" and driver_path and os.path.exists(driver_path):
            cls.log_message("[INFO] Reutilizando ChromeDriver ya instalado.")
            try:
                service = Service(driver_path)
                return webdriver.Chrome(service=service, options=options)
            except WebDriverException:
                cls.log_message("[ERROR] WebDriver reutilizado falló, intentando reinstalar...")
        
        if driver_status == "unknown":
            cls.log_message("[INFO] No se ha definido un método de ejecución. Probando ChromeDriverManager...")
            try:
                driver_path = ChromeDriverManager().install().replace("/", "\\")  # Asegurar backslashes en Windows
                cls.write_validation_file({"driver_status": "ChromeDriverManager_si", "driver_path": driver_path})
                service = Service(driver_path)
                return webdriver.Chrome(service=service, options=options)
            except WebDriverException:
                cls.log_message("[ERROR] Falló la descarga del WebDriver. Usando Chrome Portable.")
                cls.write_validation_file("ChromeDriverManager_no")
                
        if driver_status == "ChromeDriverManager_no":
            cls.log_message("[INFO] Usando WebDriver local con Chrome Portable...")
            driver_path, chrome_path = cls.get_local_chromedriver()
            if driver_path and chrome_path:
                service = Service(driver_path)
                if options is None:
                    options = webdriver.ChromeOptions()
                options.binary_location = chrome_path
                return webdriver.Chrome(service=service, options=options)
            else:
                cls.log_message("[ERROR] No se encontró Chrome Portable o ChromeDriver.")
                return None

