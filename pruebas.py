import subprocess
import sys

# Lista de librerías necesarias
libraries = [
    "tk",
    "pandas",
    "requests",
    "selenium",
    "webdriver-manager",
    "pygetwindow",
    "pywin32",
    "ttkwidgets",
    "groq"
]

# Instalación de librerías
def install_libraries():
    for lib in libraries:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", lib])
            print(f"✅ {lib} instalado correctamente.")
        except subprocess.CalledProcessError:
            print(f"❌ Error al instalar {lib}")

if __name__ == "__main__":
    install_libraries()
