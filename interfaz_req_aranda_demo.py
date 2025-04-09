import tkinter as tk
from tkinter import ttk
import json
import threading    
from create_case_demo import create_case
from ttkwidgets.autocomplete import AutocompleteCombobox


class create_interface_aranda:


    def __init__(self, root):

        self.root = root
        self.name = None
        self.affair = None
        self.root.title("Mini AD")
        self.root.geometry("550x220")
        self.root.config(bg='black')
        self.root.iconbitmap('sonda.ico')
        
        self.case = create_case(root=None)
         

    def create_widgets(self):
       
        if hasattr(self, 'create_req') and self.create_req and self.create_req.winfo_exists():
            self.create_req.lift()
            return  
        
        self.create_req = tk.Toplevel(self.root)  
        self.create_req.title("Create req")
        self.create_req.configure(bg='black')

        with open('categorias_y_ubicaciones.json', 'r', encoding='utf-8') as file:
            self.data = json.load(file)

        
        #self.selected_values = {}
   
        style_combobox = ttk.Style(self.create_req)
        style_combobox.theme_use('default')
        style_combobox.configure("TCombobox", 
                                foreground="black",
                                font=("Helvetica", 10))

     
        tk.Label(self.create_req, text="Tipo de servicio", bg='black', fg='#33ff42', font=("Helvetica", 10)).grid(row=0, column=0, padx=10, pady=10, sticky='w')

        self.combobox_service_type = ttk.Combobox(self.create_req, width=30, state="readonly")
        self.combobox_service_type.grid(row=0, column=1, padx=10, pady=5, sticky='w')

        self.combobox_service_type["values"] = [item["name"] for item in self.data["tipo servicio"]]
       
        self.combobox_service_type.bind("<<ComboboxSelected>>", self.on_service_type_change)

        # Combobox de Servicio  
        tk.Label(self.create_req, text="Servicio", bg='black', fg='#33ff42', font=("Helvetica", 10)).grid(row=1, column=0, padx=10, pady=10, sticky='w')
        self.combobox_service = ttk.Combobox(self.create_req, width=80, state="readonly")
        self.combobox_service.grid(row=1, column=1, padx=10, pady=5, sticky='w')
        self.combobox_service.bind("<<ComboboxSelected>>", self.update_categories)
        self.combobox_service["height"] = 30


        # Combobox de Categoría
        tk.Label(self.create_req, text="Categoría", bg='black', fg='#33ff42', font=("Helvetica", 10)).grid(row=2, column=0, padx=10, pady=5, sticky='w')
        self.combobox_category = ttk.Combobox(self.create_req, width=80, state="readonly")
        self.combobox_category.grid(row=2, column=1, padx=10, pady=5, sticky='w')


        default_service_type = self.combobox_service_type["values"][0] if self.combobox_service_type["values"] else None
        if default_service_type:
            self.combobox_service_type.set(default_service_type)
        self.on_service_type_change(None)  
        
        default_service = self.combobox_service["values"][0] if self.combobox_service["values"] else None
        if default_service:
            self.combobox_service.set(default_service)
        self.update_categories(None)  
        
        fields = [ "ubicación", "grupo_resolutor", "tipo_registro", "sede", "impacto", "urgencia"]
        
        self.comboboxes = {}

        row_index = 3  

        for field in fields:
            tk.Label(self.create_req, text=field.replace("_", " ").capitalize(), 
                    bg='black', fg='#33ff42', font=("Helvetica", 10)).grid(row=row_index, column=0, padx=10, pady=5, sticky='w')

        
            combobox = ttk.Combobox(self.create_req, width=30, state="readonly", 
                                    values=self.get_field_values(field), style="TCombobox")
            combobox.grid(row=row_index, column=1, padx=10, pady=5, sticky='w')

            

            combobox.set(self.get_default_value(field))  
            
            self.comboboxes[field] = combobox


            row_index += 1 
        
        tk.Button(self.create_req, width=24, text="Crear Req", command=self.monitor_state_driver,  bg='#33ff42', fg='black', font=("Helvetica", 10)).grid(row=9, column=1, padx=10, pady=5, sticky='w')             

        self.create_req.protocol("WM_DELETE_WINDOW", self.cerrar_create_req)


    def cerrar_create_req(self):
        
        self.create_req.destroy()
        self.create_req = None


    def on_service_type_change(self, event):
        selected_type = self.combobox_service_type.get()
    

        # Convertir el nombre del tipo de servicio al formato correcto de la clave en self.data
        mapping = {
            "Requerimientos de Servicio": "requerimientos de servicios",
            "Incidentes": "incidentes de servicios"
        }
        selected_key = mapping.get(selected_type, "")

        if selected_key in self.data:
            services = list(self.data[selected_key].keys())  # Extrae los nombres de los servicios
        else:
            services = []

        if services:
            self.combobox_service["values"] = services
            self.combobox_service.set(services[0])  # Seleccionar el primero por defecto
        else:
            self.combobox_service["values"] = []
            self.combobox_service.set("")

        # Llamar a update_categories con el servicio seleccionado
        self.update_categories(None)


    def update_categories(self, event):
        selected_type = self.combobox_service_type.get()
        selected_service = self.combobox_service.get()

        # Mapeo del tipo de servicio a la clave del JSON
        mapping = {
            "Requerimientos de Servicio": "requerimientos de servicios",
            "Incidentes": "incidentes de servicios"
        }
        selected_key = mapping.get(selected_type, "")

        categories = []
        if selected_key in self.data and selected_service in self.data[selected_key]:
            categories = [item["hierarchy"] for item in self.data[selected_key][selected_service]]

        if categories:
            self.combobox_category["values"] = categories
            self.combobox_category.set(categories[0])  # Seleccionar la primera por defecto
        else:
            self.combobox_category["values"] = []
            self.combobox_category.set("")

    def get_field_values(self, field):

        if field in self.data:
            return [item["name"] for item in self.data[field]]
        return []

    def get_default_value(self, field):
  
            field_defaults = {
                "ubicación": "Mesa de Servicios TI",
                "grupo_resolutor": "MESA DE SERVICIO TI",
                "tipo_registro": "Teléfono",
                "sede": "Otros-Automático Correo",
                "impacto": "5-Una persona",
                "urgencia": "3-Rápidamente"
            }

       
            default_value = field_defaults.get(field)
            if default_value:
                return default_value 
            
 
            values = [item["name"] for item in self.data[field]]
            if values:
                return values[0]  
    
    def update_parameters(self, name, affair):
        self.name = name  
        self.affair = affair 
    
    def get_items_req(self):
        
        if self.name is None and self.affair is None:
            self.name= "Operación Sofi"
            self.affair = "Agregue texto"

        item_type = self.combobox_service_type.get()
        service = self.combobox_service.get()
        category = self.combobox_category.get()
        
        selected_values = {}
        for field, combobox in self.comboboxes.items():
            selected_values[field] = combobox.get()
        
        selected_values["asunto"]=""

        selected_values["nombre"]=self.name
        selected_values["tipo de servicio"]=item_type
        selected_values["servicio"] = service
        selected_values["categoría"] = category
        selected_values["asunto"]=self.affair
        
        self.send_to_create_case(selected_values)

    def send_to_create_case(self, selected_values):

        threading.Thread(target=self.case.create_case, args=(
            selected_values["tipo de servicio"],
            selected_values["nombre"],
            selected_values["servicio"],
            selected_values["categoría"],
            selected_values["ubicación"],
            selected_values["grupo_resolutor"],
            selected_values["tipo_registro"],
            selected_values["asunto"],
            selected_values["sede"],
            selected_values["impacto"],
            selected_values["urgencia"]
        )).start()
    
    def start_thread(self, process_name):
        # Si no existe el hilo o si ya terminó, lo creamos.
        if not hasattr(self, '_open_driver_thread') or not self._open_driver_thread.is_alive():
            self._open_driver_thread = threading.Thread(target=process_name, daemon=True)
            self._open_driver_thread.start()

    def monitor_state_driver(self):
        estado = self.case.get_state_driver()
        if estado == "cerrado":
            # Solo se creará un nuevo hilo si no hay uno activo
            self.start_thread(self.case.open_driver)
        else:
            self.get_items_req()

