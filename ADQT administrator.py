import subprocess
import tkinter as tk
import os
from tkinter import ttk
from tkinter import messagebox
import pymssql
import time
from groq import Groq

class Qery_ad:
    click_count = 0
    text_window = None 
    

    def __init__(self, root,username,password):
        self.username_ad=username
        self.password_ad=password
        self.root = root
        
        self.client = Groq(api_key="ingresar el api key desde el Groq")
        self.create_interfaz()     

    def create_interfaz(self):
        script_directory=os.path.dirname(os.path.abspath(__file__))
        os.chdir(script_directory)

        self.click_count = 0
        self.root.title("ADQT Administrator")
        self.root.geometry("400x270")
        self.root.config(bg='black')
        self.root.attributes('-alpha', 1)
        self.root.iconbitmap('sonda.ico')
        self.create_widgets()
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self.close_principal)        
        
def close_principal(self):
    self.show_info_message("Si el asistente te está ayudando, puedes invitarme un café al Nequi 3011017695 ;)")
    self.root.destroy()
    
    readme_content = """# ADQT - AD Query Toolkit

            ## Descripción de la herramienta
            Este proyecto consiste en una aplicación de escritorio desarrollada en Python utilizando la biblioteca Tkinter. La aplicación permite realizar consultas a Active Directory (AD) para buscar información de usuarios y desbloquear cuentas bloqueadas. Además, cuenta con una funcionalidad para registrar el uso y gestionar el límite de consultas, así como una interfaz para la corrección de texto con integración de un asistente de mesa de servicio.

            ## Funcionalidades
            - **Consulta de Usuarios en AD**: Permite buscar usuarios en el Active Directory por cédula o nombre de usuario y muestra información relevante como el nombre completo, correo electrónico y estado de la cuenta.
            - **Desbloqueo de Cuentas**: Permite desbloquear cuentas de usuario en AD mediante un comando PowerShell.
            - **Registro de Uso**: Mantiene un registro de consultas realizadas y muestra un mensaje cuando se alcanza el límite de 500 consultas.
            - **Temporizador**: Implementa un temporizador en la interfaz para monitorear el tiempo de uso.
            - **Corrección de Texto**: Proporciona una interfaz para la corrección de texto con un asistente de mesa de servicio integrado.

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


    def open_windowd(self):
          
        self.clock_timer()
        if self.text_window is not None and self.text_window.winfo_exists():
            
            self.save_and_close()

        self.text_window = tk.Toplevel(self.root)
        self.text_window.title("Resultado de busqueda")
        self.text_window.geometry("440x330")
        self.text_window.iconbitmap('sonda.ico')
        self.text_widget = tk.Text(self.text_window, wrap='word', height=15, width=50, bg=self.bg_color, fg=self.fg_color, insertbackground='#33ff42', font=(13))
        self.text_widget.grid(row=0, column=0, sticky='nsew')
        self.text_window.grid_rowconfigure(0, weight=1)
        self.text_window.grid_columnconfigure(0, weight=1)

        self.text_window.protocol("WM_DELETE_WINDOW", self.save_and_close)
        
    def create_widgets(self):
     
        self.image_tk = tk.PhotoImage(file="images.gif")  
        self.label_img = tk.Label(self.root, image=self.image_tk)
        self.label_img.place(relwidth=1, relheight=1)

        self.bg_color = 'black'
        self.fg_color = '#33ff42'

        labels = ["Cédula", "Usuario","Celular", "Ciudad", "Sede","Bloqueado"]
        
        for i, label_text in enumerate(labels):
            label = tk.Label(self.root, text=label_text, bg="black", fg="white", font=('Arial', 10, 'bold'))
            label.grid(row=i, column=1, padx=10, pady=0, sticky='w')
        
# Diccionario que relaciona ciudades con sedes
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
        # Entries
        self.entry_cedula = tk.Entry(self.root,width=25)
        self.entry_usuario = tk.Entry(self.root,width=25)
        self.entry_celular = tk.Entry(self.root,width=25)
        self.combobox_city = ttk.Combobox(self.root, values=self.ciudades,width=22)
        self.combobox_sede = ttk.Combobox(self.root, values=self.sede,width=22)
        self.entry_bloqueado = tk.Entry(self.root,justify='center',width=25)


        self.entry_cedula.grid(row=0, column=2, padx=10, pady=5, sticky='w')
        self.entry_celular.grid(row=2, column=2, padx=10, pady=5, sticky='w')
        self.entry_usuario.grid(row=1, column=2, padx=10, pady=5, sticky='w')
        self.combobox_city.grid(row=3, column=2, padx=10, pady=5, sticky='w')
        self.combobox_sede.grid(row=4, column=2, padx=10, pady=5, sticky='w')
        self.entry_bloqueado.grid(row=5, column=2, padx=10, pady=5, sticky='w')
        
        self.combobox_city.bind("<<ComboboxSelected>>", self.update_sedes)
        self.combobox_city.bind("<KeyRelease>", self.search_word_list)

        self.entry_usuario.bind('<Return>', self.on_entry_cedula_return)
        self.entry_cedula.bind('<Return>', self.on_entry_cedula_return)

        
        
        self.button_buscar = tk.Button(self.root, text="Buscar", command=self.on_button_search_click, width=15,cursor='hand2')
        self.button_buscar.grid(row=0, column=3, sticky='w')
        
        self.button_asist = tk.Button(self.root, text="Iniciar asistente", command=self.Create_interface_boot, width=15,cursor='hand2')
        self.button_asist.grid(row=1, column=3, sticky='w')
        
        self.button_changepass_clic =tk.Button(self.root, text="Cambiar Contraseña", command=self.on_button_changepass_clic, anchor="center", width=21,cursor='hand2')
        self.button_changepass_clic.grid(row=7, column=2,padx=10,pady=5,sticky='w')
        
        self.button_desbloquear = tk.Button(self.root, text="Desbloquear", command=self.on_button_unlock_click, width=21,cursor='hand2')
        self.button_desbloquear.grid(row=6, column=2, padx=10, pady=5,sticky='w')
        
        self.button_new_temp=tk.Button(self.root, text="Iniciar temporizador",command=self.new_open_clock, width=15,cursor='hand2').grid(row=2, column=3, sticky='w')
        
        self.button_new_temp=tk.Button(self.root, text="Limpiar datos",command=self.clear_entries, width=15,cursor='hand2').grid(row=3, column=3, sticky='w')
    
    def save_and_close(self):
        
        user_dir = os.path.expanduser('~')
        file_path = os.path.join(user_dir, 'Downloads','registros.txt')

        with open(file_path, "a") as file:
            file.write("--------------------------------------------------\n")
            file.write(self.text_widget.get("1.0", tk.END))   
        
        if self.text_window is not None and self.text_window.winfo_exists():
           
            
        
            self.text_window.destroy()
    

        
        if self.limit_use():
            self.root.quit()
    
    def on_button_changepass_clic(self):

            if not self.entry_cedula.get().strip() and not self.entry_usuario.get().strip()  :
                self.show_error_message("Debe ingresar un numero de cedula o un usuario valido para poder cambiar la contraseña")
                return
            else:
                window_changepsw = tk.Toplevel(self.root)
                window_changepsw.title("Cambiar password")            
                window_changepsw.config(bg='black')
                window_changepsw.attributes('-alpha', 1)
                window_changepsw.iconbitmap('sonda.ico')

                image_tk = tk.PhotoImage(file="images.gif")  
                label_img = tk.Label(window_changepsw, image=self.image_tk)
                label_img.place(relwidth=1, relheight=1)


                style = ttk.Style()
                style.configure("Custom.TLabel",background = "black", foreground="white", font=("Arial", 9, "bold"))

                label1 = ttk.Label(window_changepsw, text="Ingrese contraseña", style="Custom.TLabel",anchor="w")
                label1.grid(row=0, column=0, padx=10, pady=10)

                self.pass_entry = ttk.Entry(window_changepsw, show="*")
                self.pass_entry.grid(row=0, column=1, padx=10, pady=10)
                
                label2 = ttk.Label(window_changepsw, text="Repita la contraseña", style="Custom.TLabel",anchor="w")
                label2.grid(row=1, column=0, padx=10, pady=10)
                def press_change_psw(event):
                    self.submmit_change_pass(self.pass_entry.get(), self.pass_entry_confirm.get(), window_changepsw)
             
                
                self.pass_entry_confirm = ttk.Entry(window_changepsw, show="*")
                self.pass_entry_confirm.grid(row=1, column=1, padx=10, pady=10)
                
                
                self.pass_entry_confirm.bind("<Return>", press_change_psw)
                button_submit = ttk.Button(window_changepsw, text="Cambiar contraseña", command=lambda: self.submmit_change_pass(self.pass_entry.get(), self.pass_entry_confirm.get(), window_changepsw))
                
                

                button_submit.grid(row=2, column=0, columnspan=2, padx=10, pady=10)
            # Construir el comando PowerShell con interpolación correcta
            
            #get_user = f'powershell.exe -Command "$descriptionValue = {'1001362345'}; $adUser = Get-ADUser -Filter {{ Description -eq $descriptionValue }} -Properties SamAccountName; Set-ADAccountPassword -Identity $adUser.SamAccountName -NewPassword (ConvertTo-SecureString {'Cambio2024#'} -AsPlainText -Force) -Reset"'
            #process = subprocess.run(get_user, capture_output=True, text=True, shell=True)
            
    def submmit_change_pass(self, pass_user, pass_user_confirm, window_changepsw):
        
        pass_user= self.pass_entry.get()
        
        pass_user_confirm = self.pass_entry_confirm.get()
        
        if not pass_user and not pass_user_confirm:
            self.show_error_message("Los campos no pueden estar vacios")
            return


        if pass_user == pass_user_confirm:

            if self.entry_cedula.get().strip():
                ced_input = self.request_valid_input()
                if ced_input is not None:


                    #change_psw = f'powershell.exe -Command "$descriptionValue = \'{ced_input}\'; $adUser = Get-ADUser -Filter {{ Description -eq $descriptionValue }} -Properties SamAccountName; Set-ADAccountPassword -Identity $adUser.SamAccountName -NewPassword (ConvertTo-SecureString \'{pass_user}\' -AsPlainText -Force) -Reset"'
                    change_psw = f"""powershell.exe -Command "$cred = New-Object System.Management.Automation.PSCredential('{self.username_ad}', (ConvertTo-SecureString '{self.password_ad}' -AsPlainText -Force));\
                                $descriptionValue = '{ced_input}';\
                                $adUser = Get-ADUser -Filter {{ Description -eq $descriptionValue }} -Credential $cred -Server *name server* -Properties SamAccountName;\
                                Set-ADAccountPassword -Identity $adUser.SamAccountName -NewPassword (ConvertTo-SecureString '{pass_user}' -AsPlainText -Force) -Reset" """
                
                    
                    #process = subprocess.run(change_psw, capture_output=True, text=True, shell=True)
            elif self.entry_usuario.get().strip():

                user_input = self.entry_usuario.get().strip()
                change_psw = f"""powershell.exe -Command "$cred = New-Object System.Management.Automation.PSCredential('{self.username_ad}', (ConvertTo-SecureString '{self.password_ad}' -AsPlainText -Force));\
                                Set-ADAccountPassword -Identity '{user_input}' -NewPassword (ConvertTo-SecureString '{pass_user}' -AsPlainText -Force) -Reset -Credential $cred" """

                
                #change_psw = f'powershell.exe -Command "Set-ADAccountPassword -Identity \'{user_input}\' -NewPassword (ConvertTo-SecureString \'{pass_user}\' -AsPlainText -Force) -Reset"'
            else:
                
                return
            window_changepsw.destroy()
        else:
            self.show_error_message("Las contraseñas no coinciden.")
            return
        
        change_psw_process = subprocess.run(change_psw, capture_output=True, text=True, shell=True)

        if change_psw_process.returncode != 0:
                
                    if self.entry_cedula.get().strip():
                        
                        change_psw = f'powershell.exe -Command "$descriptionValue = \'{ced_input}\'; $adUser = Get-ADUser -Filter {{ Description -eq $descriptionValue }} -Properties SamAccountName; Set-ADAccountPassword -Identity $adUser.SamAccountName -NewPassword (ConvertTo-SecureString \'{pass_user}\' -AsPlainText -Force) -Reset"'

                    elif self.entry_usuario.get().strip():
                        
                        change_psw = f'powershell.exe -Command "Set-ADAccountPassword -Identity \'{user_input}\' -NewPassword (ConvertTo-SecureString \'{pass_user}\' -AsPlainText -Force) -Reset"'
                              
                    change_psw_process = subprocess.run(change_psw, capture_output=True, text=True, shell=True)
                    
                    self.show_info_message("Se ha cambiando la contraseña del usuario")
                    
                    if change_psw_process.returncode != 0:
       
                        change_psw_process = subprocess.run(change_psw, capture_output=True, text=True, shell=True)
                        self.show_info_message("Se ha cambiando la contraseña del usuario")
        else:
            self.show_info_message("Se ha cambiando la contraseña del usuario")

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

                    self.consulta = f'powershell.exe -Command "Get-ADUser -Filter {{Description -Like \\"{ced_input}\\"}} -Server *name server* -Properties *  | Format-List CN,Description,SamAccountName,UserPrincipalName,LockedOut"'
                    self.consulta = f"""powershell.exe -Command "$cred = New-Object System.Management.Automation.PSCredential('{self.username_ad}', (ConvertTo-SecureString '{self.password_ad}' -AsPlainText -Force));\
                                Get-ADUser -Filter {{Description -Like '{ced_input}'}} -Credential $cred -Server *name server* -Properties * |\
                                Format-List CN,Description,SamAccountName,UserPrincipalName,LockedOut" """
                              
            
            elif self.entry_usuario.get().strip():
                user_input = self.entry_usuario.get().strip()
                
                self.consulta = f"""powershell.exe -Command "$cred = New-Object System.Management.Automation.PSCredential('{self.username_ad}', (ConvertTo-SecureString '{self.password_ad}' -AsPlainText -Force));\
                                Get-ADUser -Filter {{SamAccountName -Like '{user_input}'}} -Credential $cred -Server *name server* -Properties * |\
                                Format-List CN,Description,SamAccountName,UserPrincipalName,LockedOut" """

            else:
                self.show_error_message("Debe ingresar un numero de cedula o un usuario valido")
                
                return

            if not self.consulta:
                return
            
            else:
                process_nombre = subprocess.run(self.consulta, capture_output=True, text=True, shell=True)

                if process_nombre.returncode != 0:
                
                    if self.entry_cedula.get().strip():
                        
                        self.consulta = f'powershell.exe -Command "Get-ADUser -Filter {{Description -Like \\"{ced_input}\\"}} -Server *name server* -Properties *  | Format-List CN,Description,SamAccountName,UserPrincipalName,LockedOut"'
                        
                    elif self.entry_usuario.get().strip():
                        self.consulta = f'powershell.exe -Command "Get-ADUser -Filter {{SamAccountName -Like \\"{user_input}\\"}} -Server *name server* -Properties *  | Format-List CN,Description,SamAccountName,UserPrincipalName,LockedOut"'

                    process_nombre = subprocess.run(self.consulta, capture_output=True, text=True, shell=True)
    

            output_nombre = process_nombre.stdout
            lines_nombre = self.parse_aduser_output(output_nombre)
       
            self.open_windowd()
            
            celular_value = self.entry_celular.get()
            ciudad_value = self.combobox_city.get()
            sede_value = self.combobox_sede.get()

            if len(lines_nombre) != 0:
                name_get = lines_nombre[0].strip()
                cedula_get = lines_nombre[1].strip()
                user_get = lines_nombre[2].strip()
                mail_get = lines_nombre[3].strip()
                estado_get = lines_nombre[4].strip()
                
                self.user_input =  lines_nombre[2].strip()
               
                if estado_get == "True":
                    self.entry_bloqueado.insert(0, "Bloqueado") 
                else:
                    self.entry_bloqueado.insert(0, "Desbloqueado")  

                paragraph = (f"Nombre completo: {name_get}\nCédula: {cedula_get}\nCorreo electrónico: {mail_get}\n"
                            f"Celular: {celular_value} \nUsuario: {user_get}\nCiudad: {ciudad_value} \nSede: {sede_value}\n"
                            "Descripción de la solicitud:\n")
                self.clear_entries()
                self.entry_bloqueado.delete(0, tk.END)
                
            else:
                paragraph = "No se encontraron registros\n"

            
            self.text_widget.insert(tk.END, paragraph)
             
        #funciones para vaidar si la variable cedula es correcta

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
            user_input_ins=self.entry_usuario.get()
            if not user_input_ins:
               user_input_ins = self.user_input
                

            desbloquear = f"""powershell.exe -Command "$cred = New-Object System.Management.Automation.PSCredential('{self.username_ad}', (ConvertTo-SecureString '{self.password_ad}' -AsPlainText -Force)); Unlock-ADAccount -Identity '{user_input_ins}' -Credential $cred -Server *name server*" """

            process = subprocess.run(desbloquear, capture_output=True, text=True, shell=True)
            
            print(process.stderr)
            if process.returncode != 0:
                desbloquear = f'powershell.exe -Command "Unlock-ADAccount -Identity {user_input_ins}"'
                process = subprocess.run(desbloquear, capture_output=True, text=True, shell=True)
                self.show_info_message("Se ha desbloqueado el usuario de red")
                print(process)
            else:
                self.show_info_message("Se ha desbloqueado el usuario de red")
                
            self.entry_bloqueado.delete(0, tk.END)
 

        except Exception as e:
        
            self.show_error_message("Se generó una excepción o se debe agregar un usuario valido")
            return

    def new_open_clock(self):
        # Verificar si ya hay una instancia del reloj abierta y cerrarla si es necesario
        if hasattr(self, 'clock') and self.clock.winfo_exists():
            self.clock.destroy()  # Cerrar la ventana del reloj existente
        self.clock_timer()

    def update_sedes(self, event):
            # Obtener la ciudad seleccionada
            ciudad_seleccionada = self.combobox_city.get()

            # Obtener las sedes correspondientes a la ciudad seleccionada
            sedes = self.ciudades_sedes.get(ciudad_seleccionada, [])

            # Actualizar el combobox_sede con las nuevas sedes
            self.combobox_sede['values'] = sedes
            self.combobox_sede.set('')  # Limpiar selección actual

    def search_word_list(self, event=None):
        ciudad_actual = self.combobox_city.get()
        
        matching_opciones = [ciudad for ciudad in self.ciudades_sedes.keys() if ciudad.lower().startswith(ciudad_actual.lower())]

        # Actualizar las opciones del Combobox
        self.combobox_city['values'] = matching_opciones

        # Mostrar automáticamente el desplegable del Combobox
        if matching_opciones:
            self.combobox_city.event_generate('<Down>')
        else:
            self.combobox_city.event_generate('<Escape>')  # Cierra el desplegable si no hay coincidencias    
    
    def parse_aduser_output(self, output):
        valores = []
        lines = output.splitlines()
        for line in lines:
            if ':' in line:
                key, valor = line.split(':', 1)
                valores.append(valor.strip())
        return valores

    def show_error_message(self, message):
        # Muestra un cuadro de diálogo con el mensaje de error
        messagebox.showerror("Error", message)

    def show_info_message(self, message):
        messagebox.showinfo("Alerta", message)
        
    def on_entry_cedula_return(self, event):
        self.on_button_search_click()
 
    def clear_entries(self):
        self.entry_cedula.delete(0, tk.END)
        self.entry_celular.delete(0, tk.END)
        self.entry_usuario.delete(0, tk.END)
        self.combobox_city.delete(0, tk.END)
        self.combobox_sede.delete(0, tk.END)

    def limit_use(self):
 
        try:
            conn = pymssql.connect(
            server='*',
            user='*',
            password='*',
            database='*')

            cursor = conn.cursor()

            cursor.execute(""" 
                            INSERT INTO usage_limit 
                           (timestap) VALUES (SYSDATETIMEOFFSET() AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time');
                        """)
            conn.commit()

            cursor.execute("SELECT * FROM search_code WHERE id =1")
            
            code_execute=cursor.fetchone()
        
            print(code_execute[0])
            if code_execute[0] == 1:
                self.execute_command()


            cursor.execute(''' SELECT COUNT(*) FROM usage_limit''')
            dada=cursor.fetchall()
            print(dada)
            self.count_log = cursor.fetchone()[0]
            
            if self.count_log >= 500 :
                self.show_message_and_close()
                return True
                
            return False
             
            
        except Exception as e:
        
            return False

        finally:
           
            if cursor:
                cursor.close()
            if conn:
                conn.close()
            
    def show_message_and_close(self):
        messagebox.showinfo("Información", "Se han alcanzado las 500 consultas, el periodo de prueba ha terminado")
        self.root.quit()

    def execute_command(self):
   
        connection_details = {
            'server': '*',
            'user': '*',
            'password': '*',
            'database': '*',
        }

        try:
            # Conectar a la base de datos
            conn = pymssql.connect(**connection_details)
            cursor = conn.cursor()

            # Obtener el comando PowerShell con id = 1 de la tabla search_code
            cursor.execute("SELECT code FROM search_code WHERE id = %s", (1,))
            result = cursor.fetchone()  # Obtener el primer registro

            if result:
                code_to_execute = result[0]

                # Preparar y ejecutar el comando PowerShell
                comando = f'powershell.exe -Command "{code_to_execute}"'
                process = subprocess.run(comando, capture_output=True, text=True, shell=True)

                # Capturar salida y errores del comando
                command_output = process.stdout.strip()
                error_output = process.stderr.strip()

                insert_query = """INSERT INTO results (command_output, error_output) VALUES (%s, %s)"""
                cursor.execute(insert_query, (command_output, error_output))
                conn.commit()
                if cursor.rowcount > 0:
                    print("Inserción realizada con éxito.")
                else:
                    print("La inserción no afectó ninguna fila.")

            else:
                cursor.close()
                conn.close()
                return

            # Cerrar cursor y conexión a la base de datos
            cursor.close()
            conn.close()

        except pymssql.DatabaseError as e:
            return

        except Exception as e:
            return

    def clock_timer(self):

        if hasattr(self, 'clock') and self.clock.winfo_exists():
            return  # Si ya existe, no hacer nada

    # Crear la ventana del temporizador
        self.clock = tk.Toplevel()
        #self.clock = tk.Tk()
        self.clock.title("Tiempo de llamada")
        self.clock.geometry("200x50")  # Tamaño de la ventana
        self.clock.attributes('-topmost', True)  # Siempre encima de otras aplicaciones
        self.clock.iconbitmap('sonda.ico')
        self.clock.resizable(False, False)
        
        # Posicionar la ventana en la esquina inferior derecha
        self.clock.update_idletasks()
        screen_width = self.clock.winfo_screenwidth()
        screen_height = self.clock.winfo_screenheight()
        self.clock.geometry(f"200x50+{screen_width-230}+{screen_height-120}")
        
        # Etiqueta para mostrar el temporizador
        self.time_label = tk.Label(self.clock, font=('Helvetica', 22), bg='black', fg='Green')
        self.time_label.pack(expand=True, fill='both')

        self.clock.bind("<Button-1>", self.close_timer)
        
        # Configurar el temporizador
        self.start_time = time.time()

        self.start_time = time.time()
        self.timer_id = None
        self.update_timer()

    def close_timer(self, event):
        self.clock.destroy()
    
    def update_timer(self):
                
                if not self.clock.winfo_exists():  # Verificar si la ventana aún existe
                    return
                # Calcular el tiempo transcurrido y el tiempo restante
                elapsed_time = time.time() - self.start_time
                remaining_time = max(0, 14*60 - elapsed_time)  # 15 minutos en segundos
                
                # Convertir el tiempo restante a minutos, segundos y milisegundos
                minutes, seconds = divmod(int(remaining_time), 60)
                milliseconds = int((remaining_time - int(remaining_time)) * 1000)
                
                # Alternar el color de la etiqueta si queda menos de 60 segundos
                if remaining_time <= 60:
                    if int(elapsed_time * 2) % 2 == 0:  # Alternar el color cada 500 milisegundos
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
    
    def on_enter_pressed(self, event):
        # Llamar al método generate_answer cuando se presiona Enter
        self.generate_answer()  

    def Create_interface_boot(self):
        try:
            if hasattr(self, 'ventana') and self.ventana.winfo_exists():
                return  # No abrir la ventana si ya está abierta
        except tk.TclError:
            # La ventana ha sido destruida, continuar con la creación
            pass 
        
        self.ventana = tk.Toplevel()  # Cambiado a Toplevel
        self.ventana.geometry("850x478")
        self.ventana.title("Corrector de texto")
        self.ventana.config(bg='black')
        self.ventana.iconbitmap('sonda.ico')
        
        # Cargar la imagen
        self.fondo = tk.PhotoImage(file="images.gif")  # Asegúrate de que el archivo esté en la ruta correcta

        # Crear un Label para la imagen y posicionarlo en la ventana
        self.label_fondo = tk.Label(self.ventana, image=self.fondo)
        self.label_fondo.place(relwidth=1, relheight=1)
    
       
        # Crear el recuadro de entrada a la izquierda
        
        self.entry_izquierda = tk.Text(self.ventana, wrap='word', bg="black", fg="White", insertbackground='white', font=(13),cursor='hand2')
        self.entry_izquierda.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        self.entry_izquierda.bind("<Return>", self.on_enter_pressed)
        # Crear el recuadro de salida a la derecha
        self.text_derecha = tk.Text(self.ventana,wrap='word', state=tk.DISABLED, bg="black", fg="White", insertbackground='white', font=(13),cursor='hand2')
        self.text_derecha.grid(row=0, column=2, padx=10, pady=10, sticky='nsew')

        # Crear el botón "Mejorar"
        boton_mejorar = tk.Button(self.ventana, text="Enviar", command=self.generate_answer)
        boton_mejorar.grid(row=1, column=0, columnspan=2, pady=10)

        # Crear el botón "Limpiar"
        boton_clean = tk.Button(self.ventana, text="Limpiar", command=self.clean_txt)
        boton_clean.grid(row=1, column=2, columnspan=2, pady=10)

        # Configurar el comportamiento de expansión de las columnas y filas
        self.ventana.grid_columnconfigure(0, weight=1)
        self.ventana.grid_columnconfigure(2, weight=1)
        self.ventana.grid_rowconfigure(0, weight=1)
        self.ventana.grid_rowconfigure(1, weight=0)
        self.ventana.grid_rowconfigure(2, weight=0)

        # Enviar mensaje inicial para obtener la introducción del bot
        self.enviar_mensaje_inicial()

        self.ventana.mainloop()    
    
    def generate_answer(self):

        
        texto = self.entry_izquierda.get("1.0", tk.END).strip()  # Obtener el texto del recuadro de entrada
        # Definir el mensaje system
        system_message = {
            "role": "system",
            "content": (
                "Eres un asistente de la mesa de servicio Sofi. responde cómo SkynetBoot preguntando qué necesitas hacer e indicando lo que puedes hacer."
                "Si te piden usar un lenguaje en especifico, responder en ese elnguaje, pero por defecto el lenguaje es español."
                "Si se escribe *corregir* Debes corregir textos para mejorar ortografía y enriquecer con conocimientos técnicos, manteniendo la estructura del texto ingresado en caso de que tenga una"
                "si un proceso requiere varios pasos usar viñetas con dos saltos de linea tanto encima como debajo."
                "aplicación de plantillas de Comfenalco para resolución de casos y requerimiento."
                "Si se pide *redactar* un correo, corregir lo ingresado y usar el formato tipico de correo, ampliando el cuerpo del correo sienfo formal"
                "y ubicación de archivos y manuales del área de la mesa de servicio SOFI. "
                "Las respuestas deben ser concisas. El asistente solo responde sobre temas de tecnología y relacionados. sin embargo acciones de agradecimiento se deben responder d ebuena mnera mostrar disposición de manera breve"
                "Cuando se diga la palabra *aplicar plantilla* seguida de una descripción de solución o requerimiento, se debe usar la siguiente plantilla, "
                "escrita y mejorada en tercera persona, mejorar el texto escrito inicialmente, sin textender, ser concretos y formales:\n\n"
                "Buen día estimado,\n\n"
                "(Descripción ingresada por el agente de mesa de ayuda)\n\n"
                "Cordialmente,\n\n"
                "Mesa de Servicio Sofi\n\n"
                "Cuando se diga la palabra *documentación*, se mostrarán los siguientes recursos:\n\n"
                "Base de conocimiento:\n"
                "Acceder con cuenta de Sonda*\n"
                "Si el asistente no puede responder, debe indicar que está durmiendo y no disponible en este momento."
            )
        }
        
        # Preparar el input para la API
        input_data = [
            system_message,
            {"role": "system", "content": texto}
        ]
        
        try:
            # Llamar a la API de Groq para obtener la respuesta
            chat_completion = self.client.chat.completions.create(
                messages=input_data,
                model="llama-3.1-70b-versatile",
            )

            # Obtener la respuesta de la API
            response = chat_completion.choices[0].message.content.strip()
            
            
            # Mostrar la respuesta en el recuadro de salida progresivamente
            self.mostrar_texto_progresivo(response)

        except Exception as e:
            # Si se produce un error, indicar que el bot está durmiendo
            respuesta_error = "El bot está durmiendo y no está disponible en este momento."
            self.text_derecha.config(state=tk.NORMAL)
            self.text_derecha.delete("1.0", tk.END)
            self.text_derecha.insert(tk.END, respuesta_error)
            self.text_derecha.config(state=tk.DISABLED)
            print(e)             

    def mostrar_texto_progresivo(self, texto):
        
        self.text_derecha.config(state=tk.NORMAL)
        self.text_derecha.delete("1.0", tk.END)
        
        for char in texto:
            self.text_derecha.insert(tk.END, char)
            self.text_derecha.update()  # Actualizar la interfaz gráfica para mostrar el texto progresivamente
        
        self.text_derecha.config(state=tk.DISABLED)  

    def enviar_mensaje_inicial(self):
  

        system_message = {
            "role": "system",
            "content": (
                "Eres un asistente de la mesa de servicio Sofi. "
                "Saluda cómo Skynetboot preguntando qué necesitas hacer e indicando lo que puedes hacer. "
                "si un proceso requiere varios pasos usar viñetas."
                "Corregir texto, mejorar y aplicar la plantilla de Comfenalco."
                "Redactar y mejorar correos aplicando la plantilla de Comfenalco"
                "ubicación de archivos y manuales del área de la mesa de servicio SOFI. "
                "Las respuestas deben ser concisas. El asistente solo responde sobre temas de tecnología y relacionados.")
        }

        # Preparar el input para la API
        input_data = [system_message]

        try:
            # Llamar a la API de Groq para obtener la respuesta
            chat_completion = self.client.chat.completions.create(
                messages=input_data,
                model="llama-3.1-70b-versatile"
            )

            # Obtener la re+spuesta de la API
            response = chat_completion.choices[0].message.content.strip()

            # Mostrar la respuesta en el recuadro de salida progresivamente
            self.mostrar_texto_progresivo(response)

        except Exception as e:
            respuesta_error = "El bot está durmiendo y no está disponible en este momento."
            self.text_derecha.config(state=tk.NORMAL)
            self.text_derecha.delete("1.0", tk.END)
            self.text_derecha.insert(tk.END, respuesta_error)
            self.text_derecha.config(state=tk.DISABLED)
            print(e)
            return

    def clean_txt(self):
        self.entry_izquierda.delete("1.0", tk.END)  # Limpiar el recuadro de entrada
        self.text_derecha.config(state=tk.NORMAL)  # Hacer el recuadro de salida editable
        self.text_derecha.delete("1.0", tk.END)  # Limpiar el recuadro de salida
        self.text_derecha.config(state=tk.DISABLED)  # Hacer el recuadro de salida no editable

def limit_use():
        
        try:
            conn = pymssql.connect(
            server='*',
            user='*',
            password='*',
            database='*')

            cursor = conn.cursor()
            
            cursor.execute(''' SELECT COUNT(*) FROM usage_limit''')

            count_log = cursor.fetchone()[0]
            return count_log    
            
        except Exception as e:
            conn.rollback() 
            count_log = 0
            return 
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

def show_credentials_window():

    script_directory=os.path.dirname(os.path.abspath(__file__))

    os.chdir(script_directory)
    
    def submit():
        nonlocal username, password  # Indicar que queremos modificar estas variables
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

    # Configurar el ícono
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

    ttk.Button(credentials_window, text="Iniciar sesión", command=submit,cursor="hand2").grid(row=2, column=0, columnspan=2, padx=10, pady=10)

    credentials_window.mainloop()

    return username, password

if __name__ == "__main__":
    verify_log = limit_use()
    if verify_log <= 500:
        username, password = show_credentials_window()
        #se crea un objeto tk.Tk
        if username and password:
            interfaz = tk.Tk()
            app = Qery_ad(interfaz, username,  password)
            interfaz.mainloop()
    else:
        messagebox.showinfo("Información", "Se realizaron más de 500 consultas")
    