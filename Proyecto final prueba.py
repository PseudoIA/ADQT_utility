
import subprocess
import tkinter as tk
import os
from tkinter import ttk
from tkinter import messagebox
import time
from groq import Groq
import pandas as pd 
import webbrowser
import json
import threading
import sys
from boot import SkynetBootAssistant
from interfaz_req_aranda_demo import create_interface_aranda



class Qery_ad:
    click_count = 0 
    
    def __init__(self, root,username,password, config_file):
        
        self.root = root   

        self.username_ad=username
        self.password_ad=password
        self.create_interfaz()
        self.root.lift()
        self.root.focus_force()    
        self.query_count = 0
        self.text_window = None
        self.config_file=config_file
        self.config_domain = config_file.get("dominio", {}).get("domain_ad", "")
        self.case2 = create_interface_aranda(self.root)
        

    def create_interfaz(self):
        
        script_directory=os.path.dirname(os.path.abspath(__file__))
        os.chdir(script_directory)
        self.click_count = 0
        self.root.title("Mini AD")
        self.root.geometry("550x220")
        self.root.config(bg='black')
        self.root.attributes('-alpha', 1)
        self.root.iconbitmap('sonda.ico')
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self.close_principal) 
        self.create_widgets()       
        self.root.bind("<Configure>", self.move_window_main)
        self.root.bind("<Unmap>", self.minimize_secondary_windows)
        self.root.bind("<Map>", self.restore_secondary_windows)
              
    def close_principal(self):
        self.show_info_message("Finalizando asistente;)")
        
        readme_content = """# ADQT - AD Query Toolkit

                ## Descripción de la herramienta
                Este proyecto consiste en una aplicación de escritorio desarrollada en Python utilizando la biblioteca Tkinter. La aplicación permite realizar consultas a Active Directory (AD) para buscar información de usuarios y desbloquear cuentas bloqueadas. Además, cuenta con una funcionalidad para registrar el uso y gestionar el límite de consultas, así como una interfaz para la corrección de texto con integración de un asistente de mesa de servicio.

                ## Funcionalidades
                - **Consulta de Usuarios en AD**: Permite buscar usuarios en el Active Directory por cédula o nombre de usuario y muestra información relevante como el nombre completo, correo electrónico y estado de la cuenta.
                - **Desbloqueo de Cuentas**: Permite desbloquear cuentas de usuario en AD mediante un comando PowerShell.
                - **Registro de Uso**: Mantiene un registro de consultas realizadas y muestra un mensaje cuando se alcanza el límite de 500 consultas.
                - **Temporizador**: Implementa un temporizador en la interfaz para monitorear el tiempo de uso.
                - **Corrección de Texto**: Proporciona una interfaz para la corrección de texto con un asistente de clsmesa de servicio integrado.

                ## Requisitos
                - Python 3.x
                - Biblioteca Tkinter (incluida en la instalación estándar de Python)
                - Biblioteca `pymssql`
                - PowerShell
                - Acceso a Active Directory
                - Imagen de icono (`Agregar icono`)
                - Archivo de imagen para la interfaz (`Agregar imagen de interfaz corporativa`)

                ## Instalación
                1. Clona el repositorio o descarga el archivo del proyecto.

                2. Asegúrate de tener Python 3.x instalado en tu sistema.

                3. Instala las dependencias necesarias:

                    ```bash
                    pip install pymssql
                    ```

                4. Asegúrate de tener acceso al Active Directory y a una base de datos en Azure SQL.

                ## Uso
                - **Consulta de Usuarios**:
                1. Ingresa una cédula o nombre de usuario en los campos correspondientes.
                2. Presiona el botón de búsqueda para realizar la consulta.
                3. La información del usuario se mostrará en la ventana de texto.

                - **Desbloqueo de Cuentas**:
                1. Ingresa el nombre de usuario o usa el usuario previamente consultado.
                2. Presiona el botón de desbloqueo para desbloquear la cuenta.

                - **Corrección de Texto**:
                1. Abre la ventana de corrección de texto desde la interfaz principal.
                2. Ingresa el texto a corregir y presiona el botón "Enviar" para obtener la corrección.
                3. Usa el botón "Limpiar" para borrar el texto ingresado.

                - **Temporizador**:
                1. El temporizador se activará automáticamente al abrir la ventana de temporización.
                2. El tiempo restante se mostrará en la ventana emergente.

                ## Estructura del Código
                - **main.py**: Contiene la lógica principal de la aplicación, incluyendo:
                - Interfaz Gráfica: Implementación de la interfaz de usuario usando Tkinter.
                - Consulta de Usuarios: Funciones para ejecutar comandos PowerShell y procesar la salida.
                - Desbloqueo de Cuentas: Función para desbloquear cuentas de usuario en AD.
                - Registro de Uso: Función para registrar el uso de la aplicación y gestionar el límite de consultas.
                - Temporizador: Funciones para mostrar y actualizar el temporizador.

                - **database.py**: Maneja las operaciones con la base de datos, incluyendo:
                - Conexión a la Base de Datos: Establecimiento de la conexión a Azure SQL.
                - Registro de Consultas: Inserción de registros en la tabla de uso.
                - Ejecutar Comandos: Ejecución de comandos PowerShell y almacenamiento de resultados.

                ## Ejemplo de Uso
                A continuación se muestra un ejemplo básico de cómo ejecutar el script:

                    bash
                    python main.py

                ## Contribuciones

                Si deseas contribuir al proyecto, por favor realiza un fork del repositorio y envía tus cambios a través de una solicitud de pull (pull request). Asegúrate de seguir las prácticas de codificación y documentación del proyecto.

                ## Licencia

                Este código es de uso privado y está protegido bajo los derechos de propiedad intelectual de la empresa GADGET HUB. La distribución, modificación, uso, copia y cualquier otra acción que infrinja los derechos de propiedad intelectual están prohibidas.

                Cualquier violación de estos derechos puede resultar en acciones legales, que pueden incluir, pero no se limitan a, reclamaciones por daños y perjuicios, sanciones financieras, y medidas cautelares para prevenir el uso no autorizado del software.

                ## Contacto

                Para cualquier pregunta o comentario, por favor contacta al desarrollador principal a través del correo electrónico: abarreram22@gmail.com
                """
        
        with open('readme.txt', 'w') as file:
            file.write(readme_content)
        self.root.destroy()

    def open_windowd(self):
        
        if self.text_window is not None and self.text_window.winfo_exists():
            self.save_and_close()

        self.text_history = []
        self.history_index = -1
        self.text_window = tk.Toplevel(self.root)
        self.text_window.title("Resultado de busqueda")
        self.text_window.geometry("550x220")

        self.text_window.iconbitmap('sonda.ico')
        self.text_widget = tk.Text(self.text_window, 
                                   wrap='word', 
                                   #height=15, 
                                   #width=48, 
                                
                                   bg=self.bg_color, 
                                   fg=self.fg_color, 
                                   insertbackground='#33ff42', 
                                   font=("Helvetica", 13))
    
        self.text_widget.grid(row=0, column=0, sticky='nsew')
        self.text_window.grid_rowconfigure(0, weight=1)
        self.text_window.grid_columnconfigure(0, weight=1)

        x_principal = self.root.winfo_x()
        y_principal = self.root.winfo_y()
        ancho_principal = self.root.winfo_width()
        alto_principal = self.root.winfo_height()
        x_secundaria = x_principal 
        y_secundaria = y_principal + alto_principal + 30
        self.text_window.geometry(f"550x220+{x_secundaria}+{y_secundaria}")
        
        def store_text(event):
            self.current_text = self.text_widget.get("1.0", tk.END).strip()
            if self.history_index == -1 or self.current_text != self.text_history[self.history_index]:
                self.text_history.append(self.current_text)
                self.history_index += 1 

            if len(self.text_history) > 1000: 
                self.text_history.pop(0)
                self.history_index = -1
                self.text_history = []

        def restore_text(event):
            if self.history_index > 0:
                self.history_index -= 1
                self.text_widget.delete("1.0", tk.END)
                self.text_widget.insert(tk.END, self.text_history[self.history_index])

        self.text_window.bind("<KeyRelease>",  store_text)
        self.text_widget.bind('<Control-z>', restore_text)

            

        self.text_window.protocol("WM_DELETE_WINDOW", self.save_and_close)
        
    def create_widgets(self):
     
        self.image_tk = tk.PhotoImage(file="images.gif")  
        self.label_img = tk.Label(self.root, image=self.image_tk)
        self.label_img.place(relwidth=1, relheight=1)

        self.bg_color = 'black'
        self.fg_color = '#33ff42'

        labels = ["Cédula", "Usuario","Celular", "Ciudad", "Sede","Bloqueado","Pass expired"]
        
        for i, label_text in enumerate(labels):
            label = tk.Label(self.root, text=label_text, bg="black", fg="#77ff33", font=('Arial', 10, 'bold'))
            label.grid(row=i, column=1, padx=10, pady=0, sticky='w')
        
        self.ciudades_sedes = {
            "AMALFI": [
                "Ludoteca Amalfi"
            ],
            "ANDES": ["Ecoparque Regional Mario Aramburo Restrepo","Unidad De Servicios Suroeste"],
            "APARTADÓ": [
                "Agencia Gestión Y Colocación De Empleo",
                "Centro Atención Integral A La Infancia",
                "Centro De Servicios Plaza Del Río",
                "Parque De Los Encuentros",
                "Unidad De Servicios Nuevo Apartado"
            ],
            "BELLO": [
                "Biblioteca Ud. Aburrá Norte Niquía",
                "Centro De Servicios Puerta Del Norte"
            ],
            "BOLOMBOLO": [
                "Centro Atención Integral A La Infancia"
            ],
            "CALDAS": [
                "Ludoteca Caldas",
                "Sala De Venta Volare"
            ],
            "CAÑASGORDAS": [
                "Agencia Gestión Y Colocación De Empleo"
            ],
            "CAREPA": [
                "Centro Atención Integral A La Infancia"
            ],
            "CARMEN DE VIBORAL": [
                "Ludoteca Carmen De Viboral",
                "Recinto Quirama"
            ],
            "CAUCASIA": [
                "Ludoteca Caucasia",
                "Unidad De Servicios Caucasia"
            ],
            "CHIGORODÓ": [
                "Centro Atención Integral A La Infancia",
                "Parque Recreativo Ilur",
                "Sala De Ventas Río De Guaduas"
            ],
            "CISNEROS": [
                "Agencia Gestión Y Colocación De Empleo"
            ],
            "CIUDAD BOLÍVAR": [
                "Agencia Gestión Y Colocación De Empleo",
                "Colegio Cooperativo Ciudad Bolívar"
            ],
            "DON MATÍAS": [
                "Ludoteca Don Matías"
            ],
            "DORADAL": [
                "Centro De Servicios Doradal"
            ],
            "EL BAGRE": [
                "Centro De Servicios El Bagre"
            ],
            "EL PEÑOL": [
                "Sala De Ventas Sueños De Vida"
            ],
            "EL RETIRO": [
                "Parque Ecológico Los Salados",
                "Sala De Ventas El Claustro Vis"
            ],
            "ENVIGADO": [
                "Centro De Servicios Vivo Envigado",
                "Parque Ecoturístico El Salado",
                "Unidad De Servicios Aburrá Sur"
            ],
            "ITAGÜÍ": [
                "Acuaparque Ditaires",
                "Biblioteca La Aldea",
                "Centro De Servicios Magorca",
                "Sala De Ventas Arboleda De San Antonio"
            ],
            "JARDÍN": [
                "Hotel Hacienda Balandú",
                "Teatro Municipal De Jardín"
            ],
            "LA PINTADA": [
                "Camping Los Farallones",
                "Centro Atención Integral A La Infancia",
                "Hostería Los Farallones"
            ],
            "MARINILLA": [
                "Sala De Ventas Vermonte"
            ],
            "MEDELLÍN": [
                "Acción Plus",
                "Agencia Gestión Y Colocación De Empleo",
                "Biblioteca Castillo",
                "Biblioteca Centro Occidental",
                "Casa De La Lectura Infantil",
                "Centro De Desarrollo Cultural Moravia",
                "Centro De Distribución Guayabal",
                "Centro De Servicios Belén",
                "Centro De Servicios Punto Clave",
                "Club Edad Dorada",
                "Complementos Humanos",
                "Creorser Son Cristóbal",
                "Edificio Palomar Fondo De Empleados",
                "Empresa Tiempos",
                "Gobernación De Antioquia",
                "Hilar La Vida Castillo",
                "Jardín Infantil Mamá Chila",
                "Magisterio",
                "Otrabanda",
                "Parque Biblioteca Belén",
                "Parque Club Comfenalco Guayabal",
                "Sede Administrativa Y Servicios Palace",
                "Sede Educativa Girardot",
                "Sede La Playa"
            ],
            "NECHI": [
                "Estrategia Nechi"
            ],
            "OTROS": [
                "Automático Correo",
                "Residencio Del Colaborador"
            ],
            "PUERTO BERRÍO": [
                "Edatel Puerto Berrío",
                "Estrategia Puerto Berrío",
                "Unidad De Servicios Magdalena Medio"
            ],
            "RIONEGRO": [
                "Agencia Gestión Y Colocación De Empleo",
                "Sala De Ventas Río Campestre",
                "Sede Educativa Regional Oriente",
                "Unidad De Servicios Oriente"
            ],
            "SAN JERÓNIMO": [
                "Parque Los Tamarindos"
            ],
            "SAN LUIS": [
                "Ludoteca San Luis"
            ],
            "SAN ROQUE": [
                "Ludoteca Experimental Son Roque"
            ],
            "SANTA ELENA": [
                "Hotel Piedras Blancas",
                "Parque Piedras Blancas"
            ],
            "SANTA FE DE ANTIOQUIA": [
                "Unidad De Servicios"
            ],
            "SANTA ROSA DE OSOS": [
                "Unidad De Servicios"
            ],
            "SEGOVIA": [
                "Centro De Servicios Segovia"
            ],
            "TURBO": [
                "Agencia Gestión Y Colocación De Empleo",
                "Centro De Servicios Regional Urabá",
                "Hogar Infantil Centro Desarrollo Vecinal",
                "Hogar Infantil El León",
                "Hogar Infantil María Elena De Crovo"
            ],
            "YARUMAL": [
                "Agencia Gestión Y Colocación De Empleo"
            ]
        }
        
        self.ciudades = list(self.ciudades_sedes.keys())
        self.sede =list(self.ciudades_sedes.values())


        self.entry_cedula = tk.Entry(self.root, width=22, fg='#77ff33', bg='black', font=("Helvetica", 10),insertbackground="#77ff33")
        self.entry_usuario = tk.Entry(self.root, width=22, fg='#77ff33', bg='black', font=("Helvetica", 10),insertbackground="#77ff33")
        self.entry_celular = tk.Entry(self.root, width=22, fg='#77ff33', bg='black', font=("Helvetica", 10),insertbackground="#77ff33")
        self.entry_bloqueado = tk.Entry(self.root, justify='center', width=22, fg='#77ff33', bg='black', font=("Helvetica", 10),insertbackground="#77ff33")
        self.entry_expired_pass = tk.Entry(self.root, justify='center', width=22, fg='#77ff33', bg='black', font=("Helvetica", 10),insertbackground="#77ff33")
        
        style = ttk.Style()
        style.theme_use('clam')  
        style.configure("TCombobox", fieldbackground="black", foreground="#77ff33", font=("Helvetica", 10))

        self.combobox_city = ttk.Combobox(self.root, values=self.ciudades, width=22, style="TCombobox")
        self.combobox_sede = ttk.Combobox(self.root, values=self.sede, width=22, style="TCombobox")

        

        self.entry_cedula.grid(row=0, column=2, padx=10, pady=5, sticky='w')
        self.entry_celular.grid(row=2, column=2, padx=10, pady=5, sticky='w')
        self.entry_usuario.grid(row=1, column=2, padx=10, pady=5, sticky='w')
        self.combobox_city.grid(row=3, column=2, padx=10, pady=5, sticky='w')
        self.combobox_sede.grid(row=4, column=2, padx=10, pady=5, sticky='w')
        self.entry_bloqueado.grid(row=5, column=2, padx=10, pady=5, sticky='w')
        self.entry_expired_pass.grid(row=6, column=2, padx=10, pady=5, sticky='w')

        self.combobox_city.bind("<<ComboboxSelected>>", self.update_sedes)
        self.combobox_city.bind("<KeyRelease>", self.search_word_list)

        self.root.bind('<Return>', self.on_entry_local_key)
        self.root.bind('<Return>', self.on_entry_local_key)

        self.root.bind('<Control-Return>', self.on_entry_ad_key)
        self.root.bind('<Control-Return>', self.on_entry_ad_key)
        
        
        self.button_search_local = tk.Button(self.root, text="Buscar localmente",
                                            command=self.on_button_search_local,
                                            bg="black",
                                            fg="#77ff33",
                                            width=17,
                                            cursor='hand2')
        self.button_search_local.grid(row=0, column=3, sticky='w')

        self.button_buscar = tk.Button(self.root, text="Buscar en AD",
                                        command=lambda: self.start_thread(self.on_button_search_click),
                                        bg="black",
                                        fg="#77ff33",
                                        width=17,
                                        cursor='hand2')
        self.button_buscar.grid(row=1, column=3, sticky='w')

        self.button_asist = tk.Button(self.root, text="Iniciar asistente", 
                                      bg="#77ff33",
                                      fg="black", 
                                      command=self.open_assistant, 
                                      width=17,cursor='hand2')
        
        self.button_asist.grid(row=2, column=3, sticky='w')

        self.button_new_clock=tk.Button(self.root, text="Iniciar temporizador",
                                        command=self.new_open_clock,
                                        bg="black",
                                        fg="#77ff33", 
                                        width=17,
                                        cursor='hand2').grid(row=3, column=3, sticky='w')
        
        self.button_new_temp=tk.Button(self.root, text="Limpiar datos",
                                       command=self.clear_entries,
                                       bg="black",
                                       fg="#77ff33", 
                                       width=17,
                                       cursor='hand2').grid(row=4, column=3, sticky='w')
        
        self.button_quick_asist=tk.Button(self.root, text="Abrir asistencia rapida",
                                        bg="black",
                                        fg="#77ff33",
                                        command=self.open_quick_asist, 
                                        width=17,cursor='hand2').grid(row=5, column=3, sticky='w')
    
        self.button_aranda_remote=tk.Button(self.root,
                                            text="Abrir Aranda remoto",
                                             bg="black",
                                            fg="#77ff33",
                                            command=self.open_aranda_asist, 
                                            width=17,cursor='hand2').grid(row=6, column=3, sticky='w')                                          
        
        self.button_desbloquear = tk.Button(self.root, text="Desbloquear",
                                        bg="black",
                                        fg="#77ff33",  
                                        command=self.on_button_unlock_click, 
                                        width=17,
                                        cursor='hand2').grid(row=0, column=4, sticky='w')
                
        self.button_changepass_clic =tk.Button(self.root, text="Cambiar Contraseña", 
                                            command=self.on_button_changepass_clic, 
                                            anchor="center",
                                            bg="black",
                                            fg="#77ff33", 
                                            width=17,cursor='hand2' ).grid(row=1,column=4,pady=2,sticky='w')
        self.button_extend_pass =tk.Button(self.root, text="Ampliar vigencia psw", 
                                               command=lambda:self.start_thread(self.enable_expired),
                                               bg="black",
                                               fg="#77ff33", 
                                               anchor="center", 
                                               width=17,cursor='hand2' ).grid(row=2,column=4,pady=2,sticky='w')
        self.download_users =tk.Button(self.root, text="Actualizar usuarios", 
                                               command=lambda: self.start_thread(descargar_usuarios()),
                                               bg="black",
                                               fg="#77ff33", 
                                               anchor="center", 
                                               width=17,cursor='hand2' ).grid(row=3,column=4,pady=2,sticky='w')
        
        self.download_users =tk.Button(self.root, text="Crear Requerimiento", 
                                               command= lambda: self.start_thread(self.create_req),
                                               bg="black",
                                               fg="#77ff33", 
                                               anchor="center", 
                                               width=17,cursor='hand2' ).grid(row=4,column=4,pady=2,sticky='w')

    def on_entry_local_key(self, event):
        self.on_button_search_local()

    def on_entry_ad_key(self, event):
        self.on_button_search_click()

    def open_assistant(self):
        assistant = SkynetBootAssistant(self.root, self.config_file)
        assistant.Create_interface_boot()
    
    def save_and_close(self):
        
        user_dir = os.path.expanduser('~')
        file_path = os.path.join(user_dir, 'Downloads','registros.txt')

        with open(file_path, "a") as file:
            file.write("--------------------------------------------------\n")
            file.write(self.text_widget.get("1.0", tk.END))   
        
        if self.text_window is not None and self.text_window.winfo_exists():
            self.text_window.destroy()
            
        self.entry_bloqueado.delete(0, tk.END)
        self.entry_expired_pass.delete(0, tk.END)
        self.user_get = None
        self.cedula_get = None

    def on_button_changepass_clic(self):
            self.force_change_pass = tk.IntVar()
            try:
                if hasattr(self, 'window_changepsw') and self.window_changepsw.winfo_exists():
                    return  
            except tk.TclError:
                pass 
            
            if not self.entry_cedula.get().strip() and not self.entry_usuario.get().strip()  :
                self.show_error_message("Debe ingresar un numero de cedula o un usuario valido para poder cambiar la contraseña")
                return
            else:
                self.window_changepsw = tk.Toplevel(self.root)
                self.window_changepsw.title("Cambiar password")            
                self.window_changepsw.config(bg='black')
                self.window_changepsw.attributes('-alpha', 1)
                self.window_changepsw.iconbitmap('sonda.ico')
                image_tk = tk.PhotoImage(file="images.gif")  
                label_img = tk.Label(self.window_changepsw, image=self.image_tk)
                label_img.place(relwidth=1, relheight=1)
               
                x_principal = self.root.winfo_x()
                y_principal = self.root.winfo_y()
                ancho_principal = self.root.winfo_width()
                alto_principal = self.root.winfo_height()
                x_secundaria = x_principal + ancho_principal
                y_secundaria = y_principal 
                self.window_changepsw.geometry(f"+{x_secundaria}+{y_secundaria}")       


    
                style = ttk.Style()
                
                style.configure("Custom.TLabel",background = "black", foreground="#77ff33", font=("Arial", 9, "bold"))
                style.configure("Custom.TEntry",fieldbackground="white", foreground="black", font=("Arial", 13, "bold"))
                style.configure("Custom.TButton", background = "black", foreground="#77ff33", font=("Arial", 9, "bold"))

                label1 = ttk.Label(self.window_changepsw, text="Ingrese contraseña", style="Custom.TLabel",anchor="w")
                label1.grid(row=0, column=0, padx=10, pady=10)

                label2 = ttk.Label(self.window_changepsw, text="Repita la contraseña", 
                                   style="Custom.TLabel",
                                   anchor="w",
                                   background = "black", 
                                   foreground="#77ff33")
                
                label2.grid(row=1, column=0, padx=10, pady=10)

                self.pass_entry = ttk.Entry(self.window_changepsw, 
                                            show="*",
                                            style="Custom.TEntry")
                self.pass_entry.grid(row=0, column=1, padx=10, pady=10)

                self.pass_entry_confirm = ttk.Entry(self.window_changepsw, 
                                                    show="*",
                                                    style="Custom.TEntry")
                self.pass_entry_confirm.grid(row=1, column=1, padx=10, pady=10)

            
                Checkbutton = tk.Checkbutton(
                                        self.window_changepsw,
                                        text="Usuario debe cambiar contraseña \n en el proximo inicio de sesión",
                                        bg="black",        
                                        fg="#77ff33",        
                                        font=("Arial", 9, "bold"), 
                                        variable=self.force_change_pass, 
                                        selectcolor="black", 
                                        activebackground="white", 
                                        activeforeground="#77ff33", 
                                    )
                Checkbutton.grid(row=2, column=0, padx=10, pady=10)
            
                
                
                def press_change_psw(event):
                    self.submmit_change_pass(self.pass_entry.get(), self.pass_entry_confirm.get(), self.window_changepsw)
                
                self.pass_entry_confirm.bind("<Return>", press_change_psw)
                button_submit = ttk.Button(self.window_changepsw, text="Cambiar contraseña",
                                           command=lambda: self.submmit_change_pass(self.pass_entry.get(), self.pass_entry_confirm.get(), self.window_changepsw),
                                            style="Custom.TButton", width=17
                                            )
                button_submit.configure(cursor="hand2")
                button_submit.grid(row=2, column=1, padx=8, pady=7)       

    def align_secondary_windows(self):

        x_principal = self.root.winfo_x()
        y_principal = self.root.winfo_y()
        ancho_principal = self.root.winfo_width()
        alto_principal = self.root.winfo_height()

        if hasattr(self, 'text_window') and self.text_window and self.text_window.winfo_exists():
            x_secundaria = x_principal
            y_secundaria = y_principal + alto_principal + 30
            self.text_window.geometry(f"550x220+{x_secundaria}+{y_secundaria}")

        if hasattr(self, 'window_changepsw') and self.window_changepsw and self.window_changepsw.winfo_exists():
            x_secundaria = x_principal + ancho_principal
            y_secundaria = y_principal
            self.window_changepsw.geometry(f"+{x_secundaria}+{y_secundaria}")

        if hasattr(self.case2, 'create_req') and self.case2.create_req and self.case2.create_req.winfo_exists():
            x_secundaria = x_principal + ancho_principal
            y_secundaria = y_principal
            self.case2.create_req.geometry(f"+{x_secundaria}+{y_secundaria}")

    def move_window_main(self, event=None):
        self.align_secondary_windows()

        if hasattr(self, 'text_window') and self.text_window and self.text_window.winfo_exists():
            x_principal = self.root.winfo_x()
            y_principal = self.root.winfo_y()
            ancho_principal = self.root.winfo_width()
            alto_principal = self.root.winfo_height()
            x_secundaria = x_principal 
            y_secundaria = y_principal + alto_principal + 30
            self.text_window.geometry(f"550x220+{x_secundaria}+{y_secundaria}")
    
        if hasattr(self, 'window_changepsw') and self.window_changepsw and self.window_changepsw.winfo_exists():
            x_principal = self.root.winfo_x()
            y_principal = self.root.winfo_y()
            ancho_principal = self.root.winfo_width()
            alto_principal = self.root.winfo_height()
            x_secundaria = x_principal + ancho_principal 
            y_secundaria = y_principal 
            self.window_changepsw.geometry(f"+{x_secundaria}+{y_secundaria}")

        if hasattr(self.case2, 'create_req') and self.case2.create_req and self.case2.create_req.winfo_exists():
            x_principal = self.root.winfo_x()
            y_principal = self.root.winfo_y()
            ancho_principal = self.root.winfo_width()
            x_secundaria = x_principal + ancho_principal
            y_secundaria = y_principal
            self.case2.create_req.geometry(f"+{x_secundaria}+{y_secundaria}")

    def minimize_secondary_windows(self, event):
        if hasattr(self, 'text_window') and self.text_window and self.text_window.winfo_exists():
            self.text_window.withdraw()
        if  hasattr(self, 'window_changepsw') and self.window_changepsw and self.window_changepsw.winfo_exists():
            self.window_changepsw.withdraw()
        if hasattr(self.case2, 'create_req') and self.case2.create_req and self.case2.create_req.winfo_exists():
            self.case2.create_req.withdraw()

    def restore_secondary_windows(self, event):
        """Restaura todas las ventanas secundarias y las realinea con la principal."""
        if hasattr(self, 'text_window') and self.text_window and self.text_window.winfo_exists():
            self.text_window.deiconify()

        if hasattr(self, 'window_changepsw') and self.window_changepsw and self.window_changepsw.winfo_exists():
            try:
                self.window_changepsw.deiconify()
            except _tkinter.TclError:
                pass

        if hasattr(self.case2, 'create_req') and self.case2.create_req and self.case2.create_req.winfo_exists():
            self.case2.create_req.deiconify()

        self.align_secondary_windows()

    def submmit_change_pass(self, pass_user, pass_user_confirm, window_changepsw):

        if self.force_change_pass.get() == 1:
            force_pass = '-ChangePasswordAtLogon $true'
        else:
            force_pass = '-ChangePasswordAtLogon $false'

        pass_user= self.pass_entry.get()
        
        pass_user_confirm = self.pass_entry_confirm.get()
        
        if not pass_user and not pass_user_confirm:
            self.show_error_message("Los campos no pueden estar vacios")
            return


        if pass_user == pass_user_confirm:

            if self.entry_cedula.get().strip():
                ced_input = self.request_valid_input()
                if ced_input is not None:
                    change_psw = f"""powershell.exe -Command "$cred = New-Object System.Management.Automation.PSCredential('{self.username_ad}', (ConvertTo-SecureString '{self.password_ad}' -AsPlainText -Force));\
                                $descriptionValue = '{ced_input}';\
                                $adUser = Get-ADUser -Filter {{ Description -eq $descriptionValue }} -Credential $cred -Server {self.config_domain} -Properties SamAccountName;\
                                Set-ADAccountPassword -Identity $adUser.SamAccountName -NewPassword (ConvertTo-SecureString '{pass_user}' -AsPlainText -Force) -Server {self.config_domain} -Reset;\
                                Set-ADUser -Identity $adUser.SamAccountName {force_pass}" """
                
                    
                    #process = subprocess.run(change_psw, capture_output=True, text=True, shell=True)
            elif self.entry_usuario.get().strip():

                user_input = self.entry_usuario.get().strip()
                change_psw = f"""powershell.exe -Command "$cred = New-Object System.Management.Automation.PSCredential('{self.username_ad}', (ConvertTo-SecureString '{self.password_ad}' -AsPlainText -Force));\
                Set-ADAccountPassword -Identity '{user_input}' -NewPassword (ConvertTo-SecureString '{pass_user}' -AsPlainText -Force) -Reset -Credential $cred -Server {self.config_domain};\
                Set-ADUser -Identity '{user_input}' {force_pass}" """
            
        else:
            self.show_error_message("Las contraseñas no coinciden.")
            self.pass_entry.delete(0, tk.END)
            return
        
        change_psw_process = subprocess.run(change_psw, capture_output=True, text=True, shell=True)


        if change_psw_process.returncode != 0:
                
                    if self.entry_cedula.get().strip():
                        
                            
                        change_psw = f'powershell.exe -Command "$descriptionValue = \'{self.entry_cedula.get().strip()}\'; $adUser = Get-ADUser -Filter {{ Description -eq $descriptionValue }} -Properties SamAccountName; Set-ADAccountPassword -Identity $adUser.SamAccountName  -NewPassword (ConvertTo-SecureString \'{pass_user}\' -AsPlainText -Force) -Server {self.config_domain} -Reset;Set-ADUser -Identity $adUser.SamAccountName {force_pass}"'

                    elif self.entry_usuario.get().strip():
                       
                        change_psw = f'powershell.exe -Command "Set-ADAccountPassword -Identity {self.entry_usuario.get().strip()} -NewPassword (ConvertTo-SecureString \'{pass_user}\' -AsPlainText -Force) -Server {self.config_domain} -Reset; Set-ADUser -Identity {self.entry_usuario.get().strip()} {force_pass}"'
                   
                    change_psw_process = subprocess.run(change_psw, capture_output=True, text=True, shell=True)
                    
                    
                    
                    if change_psw_process.returncode != 0:
       
                        change_psw_process = subprocess.run(change_psw, capture_output=True, text=True, shell=True)

                        code_change_psw_process = change_psw_process.stderr.strip()
                        self.show_info_message(f"Se ha presentado un error en la ejecución de la solicitud del AD. Detalles del error: \n {code_change_psw_process}")
                    else:
                        self.show_info_message("Se ha cambiando la contraseña del usuario")
                        self.window_changepsw.destroy()
                        
                        
        else:
            self.show_info_message("Se ha cambiando la contraseña del usuario")
            self.window_changepsw.destroy()
            
    def on_button_search_local(self):
       
        celular_value = self.entry_celular.get()
        ciudad_value = self.combobox_city.get()
        sede_value = self.combobox_sede.get()

        try:
            df = pd.read_csv('AD_users.csv')
        except FileNotFoundError:
            messagebox.showerror("Error", "El archivo AD_users.csv no se encontró.")
            return
        cedula_input = str(self.entry_cedula.get())   
        user_input =str(self.entry_usuario.get())
        self.resultado_local = None


        if cedula_input:
            resultado = df[df['Description'].notna() & df['Description'].str.contains(cedula_input, case=False)]
        elif user_input:
            resultado = df[df['SamAccountName'].notna() & df['SamAccountName'].str.contains(user_input, case=False)]
        else:
            messagebox.showwarning("No encontrado", "No se puede dejar espacios en blanco")
            return

        if not resultado.empty:
            name_get = resultado['CN'].values[0]
            cedula_get = resultado['Description'].values[0]
            mail_get = resultado['UserPrincipalName'].values[0]
            cargo_get = resultado['Title'].values[0]
            user_get = resultado['SamAccountName'].values[0]
                                
            self.paragraph = (f"Nombre completo: {name_get}\nCédula: {cedula_get}\nCorreo electrónico: {mail_get}\nCargo: {cargo_get}\n"
                                    f"Celular: {celular_value} \nUsuario: {user_get}\nCiudad: {ciudad_value} \nSede: {sede_value}\n"
                                    "Descripción de la solicitud:\n")
            self.open_windowd()
            self.text_widget.insert(tk.END, self.paragraph)
            resultado = None
            if name_get and self.paragraph:
                self.get_parameter_mainclass(name_get, self.paragraph)
            else:
                messagebox.showwarning("Error", "No hay datos válidos para enviar.")

        else:
            messagebox.showwarning("No encontrado", "No se encontraron resultados para la búsqueda.")
            return

    def on_button_search_click(self):
        
            
            global click_count 
            click_count =+ 1
            
            if self.click_count == 1:
                
                self.open_windowd(self.text_widget.delete('1.0', tk.END))
                self.click_count = 0
            
            self.consulta = None
            
            if self.entry_cedula.get().strip():

                ced_input = self.request_valid_input()

                if ced_input is not None:

                    
                    self.consulta = f"""powershell.exe -Command "$cred = New-Object System.Management.Automation.PSCredential('{self.username_ad}', (ConvertTo-SecureString '{self.password_ad}' -AsPlainText -Force));\
                                Get-ADUser -Filter {{Description -Like '{ced_input}'}} -Credential $cred -Server {self.config_domain} -Properties CN,Description,SamAccountName,UserPrincipalName,LockedOut,Title,PasswordExpired |\
                                Format-List CN,Description,SamAccountName,UserPrincipalName,LockedOut,Title,PasswordExpired" """
                              

            elif self.entry_usuario.get().strip():
                user_input = self.entry_usuario.get().strip()
                self.consulta = f"""powershell.exe -Command "$cred = New-Object System.Management.Automation.PSCredential('{self.username_ad}', (ConvertTo-SecureString '{self.password_ad}' -AsPlainText -Force));\
                                Get-ADUser -Filter {{SamAccountName -Like '{user_input}'}} -Credential $cred -Server {self.config_domain} -Properties CN,Description,SamAccountName,UserPrincipalName,LockedOut,Title,PasswordExpired |\
                                Format-List CN,Description,SamAccountName,UserPrincipalName,LockedOut,Title,PasswordExpired" """

            else:
                self.show_error_message("Debe ingresar un numero de cedula o un usuario valido")
                
                return

            if not self.consulta:
                return
            
            else:
                process_nombre = subprocess.run(self.consulta, capture_output=True, text=True, shell=True)

                if process_nombre.returncode != 0:
                
                    if self.entry_cedula.get().strip():
                        
                        self.consulta = f'powershell.exe -Command "Get-ADUser -Filter {{Description -Like \\"{ced_input}\\"}} -Server {self.config_domain} -Properties CN,Description,SamAccountName,UserPrincipalName,LockedOut,Title,PasswordExpired  | Format-List CN,Description,SamAccountName,UserPrincipalName,LockedOut, Title,PasswordExpired"'
                        
                    elif self.entry_usuario.get().strip():
                        self.consulta = f'powershell.exe -Command "Get-ADUser -Filter {{SamAccountName -Like \\"{user_input}\\"}} -Server {self.config_domain} -Properties CN,Description,SamAccountName,UserPrincipalName,LockedOut,Title,PasswordExpired  | Format-List CN,Description,SamAccountName,UserPrincipalName,LockedOut,Title,PasswordExpired"'

                    process_nombre = subprocess.run(self.consulta, capture_output=True, text=True, shell=True)
                    
                    if process_nombre.returncode != 0:
                        error = process_nombre.stderr.strip()
                        self.show_info_message(f"Se ha presentado un error en la ejecución de la solicitud del AD. Detalles del error: \n{error}")
           
            output_nombre = process_nombre.stdout
            self.lines_nombre = self.parse_aduser_output(output_nombre)
       
            self.open_windowd()
            
            celular_value = self.entry_celular.get()
            ciudad_value = self.combobox_city.get()
            sede_value = self.combobox_sede.get()

            if len(self.lines_nombre) != 0:
                name_get = self.lines_nombre[0].strip()
                self.cedula_get = self.lines_nombre[1].strip()
                self.user_get = self.lines_nombre[2].strip()
                mail_get = self.lines_nombre[3].strip()
                estado_get = self.lines_nombre[4].strip()
                Cargo_get = self.lines_nombre[5].strip()
                pass_expired=self.lines_nombre[6].strip()
               


                estado_get = "Bloqueado" if estado_get == "True" else "Desbloqueado"
                pass_expired = "Si" if pass_expired == "True" else "No"

                self.entry_bloqueado.insert(0, estado_get)
                self.entry_expired_pass.insert(0, pass_expired)

                paragraph = (f"Nombre completo: {name_get}\nCédula: {self.cedula_get}\nCorreo electrónico: {mail_get}\nCargo: {Cargo_get}\n"
                            f"Celular: {celular_value} \nUsuario: {self.user_get}\nCiudad: {ciudad_value} \nSede: {sede_value}\n"
                            "Descripción de la solicitud:\n")
                self.query_count += 1

                self.get_parameter_mainclass(name_get, paragraph)
        
            else:
                paragraph = "No se encontraron registros\n"

            
            self.text_widget.insert(tk.END, paragraph)

    def get_parameter_mainclass(self, name,paragraph):
        self.name_send=name
        self.paragraph_send = paragraph    
       
        if self.case2:
            self.case2.update_parameters(self.name_send, self.paragraph_send)

    def check_and_convert(self):
        value = self.entry_cedula.get()
        try:
            cedula_value = int(value)
            return cedula_value
        except ValueError:
            self.entry_cedula.delete(0, tk.END)
            self.show_error_message("El valor ingresado no es un numero de entero valido.")
            return None

    def request_valid_input(self):
        ced_input = self.check_and_convert()
        if ced_input is not None:
            return ced_input
        else:
            return None

    def on_button_unlock_click(self):
        try:
            process = ""

            if self.entry_cedula.get().strip() and self.entry_usuario.get().strip():
                self.show_error_message("Se debe agregar un usuario o un numero de cédula.")
                return

            if self.entry_cedula.get().strip():
                ced_input = self.request_valid_input()
                if ced_input is not None:
                    desbloquear = f"""powershell.exe -Command "$cred = New-Object System.Management.Automation.PSCredential('{self.username_ad}', (ConvertTo-SecureString '{self.password_ad}' -AsPlainText -Force));\
                                        $descriptionValue = '{ced_input}';\
                                        $adUser = Get-ADUser -Filter {{ Description -eq $descriptionValue }} -Credential $cred -Server {self.config_domain} -Properties SamAccountName;\
                                        Unlock-ADAccount -Identity $adUser.SamAccountName -Server {self.config_domain}" """
                    process = subprocess.run(desbloquear, capture_output=True, text=True, shell=True)
                    output = process.stdout.strip()


                    if process.returncode != 0:
                        desbloquear = f"""powershell.exe -Command "$descriptionValue = '{ced_input}';\
                                        $adUser = Get-ADUser -Filter {{ Description -eq $descriptionValue }} -Properties SamAccountName;\
                                        Unlock-ADAccount -Identity $adUser.SamAccountName -Server {self.config_domain}" """
                        process = subprocess.run(desbloquear, capture_output=True, text=True, shell=True)
                        output = process.stdout.strip()

            elif self.entry_usuario.get().strip():
                user_input = self.entry_usuario.get().strip()
                desbloquear = f"""powershell.exe -Command "$cred = New-Object System.Management.Automation.PSCredential('{self.username_ad}', (ConvertTo-SecureString '{self.password_ad}' -AsPlainText -Force)); Unlock-ADAccount -Identity '{user_input}' -Credential $cred -Server {self.config_domain}" """
                process = subprocess.run(desbloquear, capture_output=True, text=True, shell=True)

                if process.returncode != 0:
                    desbloquear = f"""powershell.exe -Command "Unlock-ADAccount -Identity '{user_input}' -Server {self.config_domain}" """
                    process = subprocess.run(desbloquear, capture_output=True, text=True, shell=True)
                    output = process.stdout.strip()


            if process.returncode == 0:
                self.show_info_message(f"Cuenta {user_input if self.entry_usuario.get().strip() else ced_input} desbloqueada con éxito.")
            else:
                error = process.stderr.strip()
                self.show_info_message(f"Se ha presentado un error en la ejecución de la solicitud del AD. Detalles del error: \n{error}")

            self.entry_bloqueado.delete(0, tk.END)

        except Exception as e:
            self.show_error_message(f"Se debe agregar un usuario o cedula valido.{e}" )
            return
                
    def enable_expired(self):

        
        if self.entry_cedula.get():
            ced_expired =  str(self.request_valid_input())
        if self.entry_usuario.get():
            user_expired = self.entry_usuario.get()
        

        if ced_expired:       
            desbloquear = f"""powershell.exe -Command "$cred = New-Object System.Management.Automation.PSCredential('{self.username_ad}', (ConvertTo-SecureString '{self.password_ad}' -AsPlainText -Force));\
                                $descriptionValue = '{ced_expired}';\
                                $adUser = Get-ADUser -Filter {{ Description -eq $descriptionValue }} -Credential $cred -Server {self.config_domain} -Properties SamAccountName;\
                                Set-ADUser -Credential $cred -Identity $adUser -ChangePasswordAtLogon $true;\
                                Unlock-ADAccount -Credential $cred -Identity $adUser -Server {self.config_domain};\
                                Set-ADUser -Credential $cred -Identity $adUser -ChangePasswordAtLogon $false -Server {self.config_domain}" """
            
            process = subprocess.run(desbloquear, capture_output=True, text=True, shell=True)

            if process.returncode != 0:
                desbloquear = f"""powershell.exe -Command "$descriptionValue = '{ced_expired}';\
                                $adUser = Get-ADUser -Filter {{ Description -eq $descriptionValue }} -Server {self.config_domain} -Properties SamAccountName;\
                                Set-ADUser -Identity $adUser -Server {self.config_domain} -ChangePasswordAtLogon $true;\
                                Unlock-ADAccount -Identity $adUser -Server {self.config_domain};\
                                Set-ADUser -Identity $adUser -ChangePasswordAtLogon $false -Server {self.config_domain}" """
                process = subprocess.run(desbloquear, capture_output=True, text=True, shell=True)
            if process.returncode == 0:
                    self.show_info_message(f"Se ha restablecido la vigencia del password de la cuenta con éxito.")
            
            else:
                error = process.stderr.strip()
                self.show_info_message(f"Se ha presentado un error en la ejecución de la solicitud del AD. Detalles del error: \n{error}")
                return
        
        if user_expired:
            desbloquear = f"""powershell.exe -Command "$cred = New-Object System.Management.Automation.PSCredential('{self.username_ad}', (ConvertTo-SecureString '{self.password_ad}' -AsPlainText -Force));\
                                Set-ADUser -Credential $cred -Identity {user_expired} -ChangePasswordAtLogon $true;\
                                Unlock-ADAccount -Credential $cred -Identity {user_expired} -Server {self.config_domain};\
                                Set-ADUser -Credential $cred -Identity {user_expired} -ChangePasswordAtLogon $false -Server {self.config_domain}" """
            process = subprocess.run(desbloquear, capture_output=True, text=True, shell=True)

            if process.returncode != 0:

                desbloquear = f"""powershell.exe -Command "Set-ADUser -Identity {user_expired} -ChangePasswordAtLogon $true -Server {self.config_domain};\
                                Unlock-ADAccount  -Identity {user_expired} -Server {self.config_domain};\
                                Set-ADUser -Identity {user_expired} -ChangePasswordAtLogon $false -Server {self.config_domain}" """
                process = subprocess.run(desbloquear, capture_output=True, text=True, shell=True)

            if process.returncode == 0:
                    self.show_info_message(f"Cuenta {user_expired} desbloqueada con éxito.")
            else:
                error = process.stderr.strip()
                self.show_info_message(f"Se ha presentado un error en la ejecución de la solicitud del AD. Detalles del error: \n{error}")
                return
        else:
            error = process.stderr.strip()
            self.show_info_message(f"Se ha presentado un error en la ejecución de la solicitud del AD. Detalles del error: \n{error}")
            return
        
        user_expired = "" if ced_expired or user_expired else user_expired

        self.entry_expired_pass.delete(0, tk.END)

    def new_open_clock(self):
    
        if hasattr(self, 'clock') and self.clock.winfo_exists():
            self.clock.destroy()  
        self.clock_timer()

    def update_sedes(self, event):
            
            ciudad_seleccionada = self.combobox_city.get()

            sedes = self.ciudades_sedes.get(ciudad_seleccionada, [])

            self.combobox_sede['values'] = sedes
            self.combobox_sede.set('')  

    def search_word_list(self, event=None):
        ciudad_actual = self.combobox_city.get()
        
        matching_opciones = [ciudad for ciudad in self.ciudades_sedes.keys() if ciudad.lower().startswith(ciudad_actual.lower())]

    
        self.combobox_city['values'] = matching_opciones

  
        if matching_opciones:
            self.combobox_city.event_generate('<Down>')
        else:
            self.combobox_city.event_generate('<Escape>')     
    
    def parse_aduser_output(self, output):
        valores = []
        lines = output.splitlines()
        for line in lines:
            if ':' in line:
                key, valor = line.split(':', 1)
                valores.append(valor.strip())
        return valores

    def show_error_message(self, message):

        messagebox.showerror("Error", message)

    def show_info_message(self, message):
        messagebox.showinfo("Alerta", message)
        
    def clear_entries(self):
        self.entry_cedula.delete(0, tk.END)
        self.entry_celular.delete(0, tk.END)
        self.entry_usuario.delete(0, tk.END)
        self.combobox_city.delete(0, tk.END)
        self.combobox_sede.delete(0, tk.END)
        self.entry_bloqueado.delete(0, tk.END)
        self.entry_expired_pass.delete(0, tk.END)

    def open_quick_asist(self):
    
        exe_path = r'C:\windows\system32\quickassist.exe'
    
        cmd_command = f'cmd /c start "" "{exe_path}"'

        try:
            subprocess.Popen(exe_path)

        except Exception as e:

            try:
                subprocess.Popen(cmd_command, shell=True)
            except Exception as e:


                try:
                    subprocess.Popen(['start', 'ms-quick-assist://'], shell=True)
                except Exception as e:
                    return

    def open_aranda_asist(self):

        dir = os.path.join(os.getcwd(), 'C:/Program Files (x86)/Aranda/ADM Utils')
        
        if os.path.exists(dir):
            url = "https://comfenalcoadm.sonda.com/adm/Pages/Login.aspx?ReturnUrl=%2fadm%2fPages%2fDefault.aspx"
            webbrowser.open(url)
        else:

            try:
                dir = os.path.join(os.getcwd(), "manuales")
                exe_path = os.path.join(dir, "Aranda.ADM.Utils.Installer.9.21.2.15.exe")
                exefile= os.path.join(dir, "PASO A PASO ARANDA REMOTO (1).rtf")
                print(exe_path)
                os.startfile(exefile)
                os.startfile(exe_path)

            except Exception as e:
                error_message = f"Ocurrió un error: {str(e)}"
                self.show_error_message(error_message)

    def create_inc(self):
        print("en proceso")
        
    def clock_timer(self):

        if hasattr(self, 'clock') and self.clock.winfo_exists():
            return  
        self.clock = tk.Toplevel()
        self.clock.title("Tiempo de llamada")
        self.clock.geometry("200x50") 
        self.clock.attributes('-topmost', True) 
        self.clock.iconbitmap('sonda.ico')
        self.clock.resizable(False, False)
        
 
        self.clock.update_idletasks()
        screen_width = self.clock.winfo_screenwidth()
        screen_height = self.clock.winfo_screenheight()
        self.clock.geometry(f"200x50+{screen_width-230}+{screen_height-140}")
        

        self.time_label = tk.Label(self.clock, font=('Helvetica', 22), bg='black', fg='Green')
        self.time_label.pack(expand=True, fill='both')

        self.clock.bind("<Button-1>", self.close_timer)
        
   
        self.start_time = time.time()

        self.start_time = time.time()
        self.timer_id = None
        self.update_timer()

    def close_timer(self, event):
        self.clock.destroy()
    
    def update_timer(self):
                
                if not self.clock.winfo_exists():  
                    return
                
                elapsed_time = time.time() - self.start_time
                remaining_time = max(0, 14*60 - elapsed_time)  
                
              
                minutes, seconds = divmod(int(remaining_time), 60)
                milliseconds = int((remaining_time - int(remaining_time)) * 1000)
                
               
                if remaining_time <= 60:
                    if int(elapsed_time * 2) % 2 == 0:  
                        self.time_label.config(fg='red')
                    else:
                        self.time_label.config(fg='black')
                else:
                    self.time_label.config(fg='green')
                
                # Actualizar el texto de la etiqueta
                self.time_label.config(text=f"{minutes:02}:{seconds:02}:{milliseconds:03}")
                
                if remaining_time > 0:
                    # Programar la actualización del temporizador después de 50 milisegundos
                    self.clock.after(50, self.update_timer)
                else:
                    # Si el tiempo ha terminado, actualizar el texto y cerrar el programa
                    self.time_label.config(text="¡Tiempo terminado!", font=('Helvetica', 20))
                    self.clock.quit()  # Cierra la ventana de Tkinter
  
    def clean_txt(self):
        self.entry_izquierda.delete("1.0", tk.END)  
        self.text_derecha.config(state=tk.NORMAL) 
        self.text_derecha.delete("1.0", tk.END)  
        self.text_derecha.config(state=tk.DISABLED)  

    def key_events(self, event):
        if event.state & 0x0004 and event.keysym == 'z':
                self.text_widget.edit_undo()  
                return "break"
    
    def create_req(self):       

        if hasattr(self.case2, 'create_req') and self.case2.create_req and self.case2.create_req.winfo_exists():
            self.case2.create_req.lift()  # La trae al frente si ya existe
            return 
        
        self.case2.create_widgets()
        self.move_window_main()

    def start_thread(self,proces_name):
        hilo = threading.Thread(target=proces_name, daemon=True)
        hilo.start()

def preguntar_descarga():

    ruta_actual = os.path.dirname(os.path.abspath(__file__))
    ruta_archivo = os.path.join(ruta_actual, 'AD_users.csv')
    
    if os.path.exists(ruta_archivo):
        return  
    
    respuesta = messagebox.askquestion("Descargar Usuarios", 
                                        "¿Desea hacer la búsqueda de usuarios localmente?\n\n"
                                        "Para eso debe descargar los usuarios en un archivo.\n"
                                        "Clic en 'Sí' para descargar, clic en 'No' para omitir y hacer la búsqueda con el AD.")

    if respuesta == 'yes':
        descargar_usuarios()
    else:
        return

def descargar_usuarios():

    ruta_actual = os.path.dirname(os.path.abspath(__file__))
    ruta_archivo = os.path.join(ruta_actual, 'AD_users.csv')

    with open("config.json", "r", encoding="utf-8") as file:
                    config =json.load(file)
    config_domain = config.get("dominio", {}).get("domain_ad", "")
    consulta = [
        "powershell.exe", "-Command",
        """if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) { \
                Write-Host 'El módulo ActiveDirectory no está instalado. Instalando...' -ForegroundColor Yellow; \
                Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0; \
                Write-Host 'Módulo instalado exitosamente.' -ForegroundColor Green; } """  
        f"Get-ADUser -Filter * -Server {config_domain} -Properties CN,Description,SamAccountName,UserPrincipalName,Title | " 
        "Select-Object CN,Description,SamAccountName,UserPrincipalName,Title | "
        f"Export-Csv -Path '{ruta_archivo}' -NoTypeInformation"
    ]

    subprocess.run(consulta, capture_output=True, text=True)
    messagebox.showinfo("Éxito", f"Consulta completada.")

def show_credentials_window():

    script_directory=os.path.dirname(os.path.abspath(__file__))

    os.chdir(script_directory)
    
    def submit():
        nonlocal username, password  
        username = username_entry.get()
        password = password_entry.get()

        if username and password:
            credentials_window.destroy()
        else:
            messagebox.showwarning("Input Error", "Por favor, ingrese ambos campos")
            username_entry.delete(0, tk.END)
            password_entry.delete(0,tk.END)


    def press_pass_login(event):
         submit()

    username = None
    password = None

    credentials_window = tk.Tk()


    style = ttk.Style()

    style.configure("TLabel",
                    background="black",  
                    foreground="white",
                    font=("Arial", 9, "bold",)) 


    credentials_window.iconbitmap('sonda.ico')
    credentials_window.title("Login AD")
    credentials_window.config(bg='black')


    ttk.Label(credentials_window, text="Usuario de dominio",style="TLabel").grid(row=0, column=0, padx=10, pady=10)
    username_entry = ttk.Entry(credentials_window)
    username_entry.grid(row=0, column=1, padx=10, pady=10,)
    

    ttk.Label(credentials_window, text="Contraseña",style="TLabel").grid(row=1, column=0, padx=10, pady=10, sticky='w')
    password_entry = ttk.Entry(credentials_window, show="x")
    password_entry.grid(row=1, column=1, padx=10, pady=10)
    

    password_entry.bind("<Return>", press_pass_login)

    ttk.Button(credentials_window,
                text="Iniciar sesión",
                command=submit,
                cursor="hand2").grid(row=2, column=0, columnspan=2, padx=10, pady=10)

    credentials_window.mainloop()

    return username, password

def show_error_message(message):
    messagebox.showerror("Error", message)

def validate_config():
    base_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_path, "config.json")
    if not os.path.exists(config_path):
        show_error_message(f"Se requiere el archivo {config_path} para iniciar el programa.")
        sys.exit(1)
    else: 
        with open(config_path, "r", encoding="utf-8") as file:
            config = json.load(file)
        return config
 
if __name__ == "__main__":
    
    config_file=validate_config()

    username, password = show_credentials_window()

    if username and password:
        preguntar_descarga()
        interfaz = tk.Tk()
        app = Qery_ad(interfaz, username,  password,config_file)
        interfaz.mainloop()
