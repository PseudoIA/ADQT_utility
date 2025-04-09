import tkinter as tk
from tkinter import messagebox
from groq import Groq
import pandas as pd
import json
import os
import threading
import time
from datetime import datetime

class SkynetBootAssistant:
    def __init__(self, root, config):
        
        self.config_key = config.get("APY_KEY_BOOT", {}).get("key", "")
        self.config_version_model = config.get("APY_KEY_BOOT", {}).get("modelo", "")
        
        self.client = Groq(api_key=str(self.config_key) )
        self.root = root

    
    def Create_interface_boot(self):
        """try:
            if hasattr(self, 'ventana') and self.ventana.winfo_exists():
                return  
        except tk.TclError:
            pass """
        
        self.keywords = ["aplicaciones", "zeus", "tips", "info", "manuales", "notas", "repositorios", "matriz", "crear", "plantilla", "aplicar plantilla", "ayuda", "crear"]
        self.ventana = tk.Toplevel(self.root)  
        self.ventana.geometry("570x600")
        self.ventana.title("Skynetboot asistente")
        self.ventana.config(bg='black')
        self.ventana.iconbitmap('sonda.ico')
        self.ventana.grid_rowconfigure(0, weight=1)
        self.ventana.grid_columnconfigure(0, weight=1)

        self.ventana.rowconfigure(0, weight=1) 
        self.ventana.rowconfigure(1, weight=2)
        
        
        self.frame1 = tk.Frame(self.ventana, bg="#545454", bd=0, relief="solid")
        self.frame1.grid(row=0, column=0, sticky="nsew")
        
        self.frame1.grid_rowconfigure(0, weight=1)
        self.frame1.grid_columnconfigure(0, weight=1)
        self.crear_canvas_redondeado(self.frame1)
        
  
        self.container = tk.Frame(self.ventana, bg="#545454")
        self.container.grid(row=1, column=0, sticky="nsew")


        self.container.columnconfigure(0, weight=7) 
        self.container.columnconfigure(1, weight=1)  
        self.container.rowconfigure(0, weight=1)  
        
        
      
        self.frame2 = tk.Frame(self.container, bg="#545454", bd=0, relief="solid")
        self.frame2.grid(row=0, column=0, sticky="nsew")
        self.frame2.grid_propagate(False) 
        self.frame2.grid_rowconfigure(0, weight=1)
        self.frame2.grid_columnconfigure(0, weight=1)

        self.crear_canvas_redondeado_entry(self.frame2)

        
        self.frame3_container = tk.Frame(self.container, bg="#545454", bd=0, relief="solid")
        self.frame3_container.grid(row=0, column=1, sticky="nsew")

        self.frame3_container.grid_rowconfigure(0, weight=1)
        self.frame3_container.grid_columnconfigure(0, weight=1)
        self.frame3_container.grid_propagate(False)

        
        self.canvas_frame3 = tk.Canvas(
            self.frame3_container,
            borderwidth=0, highlightthickness=0,
            bg="#1e1e1e"
        )
        self.canvas_frame3.grid(row=0, column=0, sticky="nsew")

        self.subframe_botones = tk.Frame(self.frame3_container, bg="#444444")
        self.subframe_botones.place(relx=0, rely=0.09, relwidth=1, relheight=0.87)


        self.subframe_botones.grid_rowconfigure(0, weight=1)
        self.subframe_botones.grid_rowconfigure(1, weight=1)
        self.subframe_botones.grid_rowconfigure(2, weight=1)
        self.subframe_botones.grid_columnconfigure(0, weight=1)

        self.boton_mejorar_frame = tk.Frame(self.subframe_botones, bg="#444444")
        self.boton_mejorar_frame.grid(row=0, column=0, sticky="nsew")

        self.boton_mejorar = tk.Button(
            self.boton_mejorar_frame,
            text="Enviar",
            bg="#3c3c3c", fg="white",
            borderwidth=0, relief="flat", cursor="hand2",
            command=self.create_hilo, font=("Segoe Ui Semibold", 8)
        )
        # El botón se expande para llenar el contenedor
        self.boton_mejorar.pack(fill="both", expand=True)

        # Borde blanco inferior (simula un borde sólo en la parte de abajo)
        self.borde_mejorar = tk.Frame(self.boton_mejorar_frame, bg="#848484", height=1.2)
        self.borde_mejorar.place(relx=0.5, rely=1.0, relwidth=0.75, anchor="s")

        self.boton_clean_frame = tk.Frame(self.subframe_botones, bg="#444444")
        self.boton_clean_frame.grid(row=1, column=0, sticky="nsew")
    
        self.boton_clean = tk.Button(
            self.boton_clean_frame,
            text="Limpiar",
            bg="#3c3c3c", fg="white",
            borderwidth=0, relief="flat", cursor="hand2",
            command=self.clean_txt, font=("Segoe Ui Semibold", 8)
        )
        self.boton_clean.pack(fill="both", expand=True)
               # Borde blanco inferior (simula un borde sólo en la parte de abajo)
        self.borde_clean = tk.Frame(self.boton_clean_frame, bg="#848484", height=1.2)
        self.borde_clean.place(relx=0.5, rely=1.0, relwidth=0.75, anchor="s")


        self.boton_voice_frame = tk.Frame(self.subframe_botones, bg="#444444")
        self.boton_voice_frame.grid(row=2, column=0, sticky="nsew")

        self.boton_voice_text = tk.Button(self.subframe_botones,
                                        text="Voice\nText",
                                        bg="#3c3c3c", fg="white",
                                        borderwidth=0,cursor="hand2", relief="flat",font=("Segoe Ui Semibold", 8))
        self.boton_voice_text.grid(row=2, column=0, sticky="nsew")


        self.frame3_container.bind("<Configure>", self.actualizar_canvas)


        self.enviar_mensaje_inicial()
    
    def actualizar_canvas(self, event):
        self.canvas_frame3.delete("all")  # Limpia el canvas
        width = self.frame3_container.winfo_width()
        height = self.frame3_container.winfo_height()
        radius = 18  # Radio de las esquinas

        if width < 2*radius or height < 2*radius:
            return


        self.canvas_frame3.create_arc(0, 0, 2*radius, 2*radius, start=90, extent=90, fill="#3c3c3c", outline="#3c3c3c")
        self.canvas_frame3.create_arc(width-2*radius, 0, width, 2*radius, start=0, extent=90, fill="#3c3c3c", outline="#3c3c3c")
        self.canvas_frame3.create_arc(0, height-2*radius, 2*radius, height, start=180, extent=90, fill="#3c3c3c", outline="#3c3c3c")
        self.canvas_frame3.create_arc(width-2*radius, height-2*radius, width, height, start=270, extent=90, fill="#3c3c3c", outline="#3c3c3c")

  
        self.canvas_frame3.create_rectangle(radius, 0, width-radius, height, fill="#3c3c3c", outline="#3c3c3c")
        self.canvas_frame3.create_rectangle(0, radius, width, height-radius, fill="#3c3c3c", outline="#3c3c3c")


        self.canvas_frame3.configure(width=width, height=height)
        

    def crear_canvas_redondeado_entry(self, frame):
        self.canvas_entry = tk.Canvas(frame, borderwidth=0, highlightthickness=0, bg="#1e1e1e")
        self.canvas_entry.grid(row=0, column=0, sticky="nsew")
        

        # Text widget con bordes redondeados
        self.entry_izquierda = tk.Text(frame, wrap='word', state=tk.NORMAL, bg="#444444",
                                    fg="#eeeeee", insertbackground='white', font=("Segoe Ui Semibold", 10),
                                    cursor='xterm', borderwidth=0)
        self.entry_izquierda.grid(row=0, column=0, sticky="nsew")
        self.entry_izquierda.tag_configure("highlight", foreground="#444444")
        self.entry_izquierda.bind('<Control-Return>', self.key_press)
        self.entry_izquierda.bind('<KeyRelease>', self.change_color_key)
        self.canvas_entry.bind("<Configure>", self.redibujar_canvas_entry)
        
    def redibujar_canvas_entry(self, event):


        self.canvas_entry.delete("all")
        width, height = event.width, event.height
        radius = 18
    
        self.canvas_entry.create_arc(0, 0, 2*radius, 2*radius, start=90, extent=90, fill="#444444",outline="#444444")
        self.canvas_entry.create_arc(width-2*radius, 0, width, 2*radius, start=0, extent=90,  fill="#444444",outline="#444444")
        self.canvas_entry.create_arc(0, height-2*radius, 2*radius, height, start=180, extent=90,  fill="#444444",outline="#444444")
        self.canvas_entry.create_arc(width-2*radius, height-2*radius, width, height, start=270, extent=90, fill="#444444",outline="#444444")
        
        self.canvas_entry.create_rectangle(radius, 0, width-radius, height, fill="#444444",outline="#444444")
        self.canvas_entry.create_rectangle(0, radius, width, height-radius, fill="#444444",outline="#444444")
            # Ajustar el entry_izquierda al nuevo tamaño del canvas
        self.entry_izquierda.place(x=radius, y=radius, width=width - 2*radius, height=height - 2*radius)

    def crear_canvas_redondeado(self, frame):
        self.canvas = tk.Canvas(frame, borderwidth=0, highlightthickness=0, bg="#1e1e1e")
        self.canvas.grid(row=0, column=0, sticky="nsew")
        

    
        self.text_derecha = tk.Text(frame, wrap='word', state=tk.NORMAL, bg="#242424",
                            fg="white", insertbackground='white', font=("Segoe Ui Semibold", 10),
                            cursor='xterm', borderwidth=0)
        self.text_derecha.grid(row=0, column=0, sticky="nsew")
        self.canvas.bind("<Configure>", self.redibujar_canvas)
            
    def redibujar_canvas(self, event):
        self.canvas.delete("all")
        width, height = event.width, event.height
        radius = 18

        borde_color = "#242424"  

        self.canvas.create_arc(0, 0, 2*radius, 2*radius, start=90, extent=90, fill="#242424", outline=borde_color)
        self.canvas.create_arc(width-2*radius, 0, width, 2*radius, start=0, extent=90, fill="#242424", outline=borde_color)
        self.canvas.create_arc(0, height-2*radius, 2*radius, height, start=180, extent=90, fill="#242424", outline=borde_color)
        self.canvas.create_arc(width-2*radius, height-2*radius, width, height, start=270, extent=90, fill="#242424", outline=borde_color)
        
         
        self.canvas.create_rectangle(radius, 0, width-radius, height, fill="#242424", outline="#242424")
        self.canvas.create_rectangle(0, radius, width, height-radius, fill="#242424", outline="#242424")

   
        self.text_derecha.place(x=radius, y=radius, width=width - 2*radius, height=height - 2*radius)
    
    def create_hilo(self):
        self.running = True  
        hilo = threading.Thread(target=self.generate_answer)
        hilo.start()

    def change_color_key(self, event):
        self.entry_izquierda.tag_configure("highlight", foreground="#00e6da") 
        for word in self.keywords:
            start = "1.0"
            while True:
                start = self.entry_izquierda.search(word, start, stopindex=tk.END)
                if not start:
                    break
                end = f"{start}+{len(word)}c"
                self.entry_izquierda.tag_add("highlight", start, end)  # Usa el tag "highlight"
                start = end  # Mover el índice de búsqueda para evitar bucles infinitos


    def key_press(self, event):
        if event.keysym == 'Return' and (event.state & 0x0004):
            self.generate_answer()

    def open_file(self,text):
        dir = os.path.join(os.getcwd(), "manuales")

        if not os.path.exists(dir):
            os.makedirs(dir)
        
        files = os.listdir(dir)

        system_message = {
            "role": "system",
            "content": f"Analiza el texto del usuario y verifica si hay una relación con los archivos listados a continuación: \n {files}. \n "
                    "Si encuentras una coincidencia, devuelve únicamente el nombre del archivo junto a su extensión tal como aparece en la lista ya que se intentara abrir. "
                    "Si no hay relación, responde con 'No se encontró ningún archivo.' solo eso "
                    "NOTA!!!!!!!! al boot: unicamente debes responder con el nombre del archivo con su extensión tal como aparece en la lista así se pidan otras cosas en el texto ingresado, solo debes buscar la relación entre el texto ingresado y los nombres de los archivos listados."
        }

        input_data = [
            system_message,
            {"role": "user", "content": text}
        ]

        try:
            chat_completion = self.client.chat.completions.create(
                messages=input_data,
                model=str(self.config_version_model),
            )

            result_search = chat_completion.choices[0].message.content.strip()
            

            if result_search != "No se encontró ningún archivo en la base de conocimiento.":
                file_path = os.path.join(dir, result_search)
                try:
                    if os.name == 'nt':  
                        os.startfile(file_path)
                        return result_search

                except Exception as e:
                    return result_search

            else:
                return result_search

        except Exception as e:
            return ""

    def search_keyword(self, input_text):
            if not input_text:
                return ""
            
            if ':' in input_text:
                text = input_text.split(':', 1)[0]
            else:
                text = input_text

            first_keyword = text.split()[0]
            
            return first_keyword
    
    def search_key(self,text):
        app_name = []
        with open(r'.\Base de conocimiento.JSON', 'r', encoding='utf-8') as file:
            self.apps_lists = json.load(file) 
        app_name = [app["Nombre de Aplicación"] for app in self.apps_lists["matriz"]["Aplicacion"]]
        
        system_message ={"role": "system",
                        "content": f"Analiza el texto del usuario y verifica si hay una relación con las aplicaciones listadas a continuación: {app_name}. "
                        "Si encuentras una coincidencia, devuelve únicamente el nombre de la aplicación tal como aparece en la lista. "
                        "Si no hay relación, responde con 'No se encontró la aplicación en la base de conocimiento.' solo eso ya que se le enviara "
                        "NOTA!!!!!!!! al boot: unicamente debes responder con el nombre de la aplicación tal como aparece en la lista así se pidan otras cosas en el texto ingresado, solo debes buscar la relación entre el nombre de la aplicación ingresada y el nombre de la aplicación de la lista enviada."} 

        input_data = [
            system_message,
            {"role": "user", "content": text}
        ]
        
        try:
            chat_completion = self.client.chat.completions.create(
                messages=input_data,
                model=str(self.config_version_model),
            )

            result_search = chat_completion.choices[0].message.content.strip()
            print(result_search)
            
            if result_search != "No se encontró la aplicación en la base de conocimiento.":
                result_app= [app for app in self.apps_lists["matriz"]["Aplicacion"] if app["Nombre de Aplicación"] == (result_search)]    
                return result_app       
            else:
                return result_search
                

        except Exception as e:
            return None
        
    def search_pin(self, text):
        if ':' in text:
            pin = text.split(':', 1)[1].strip()
        else:
            parts = text.split()
            if len(parts) > 1 and parts[0] == "pines":
                pin = parts[1].strip()
            else:
                return "Comando inválido. Usa 'pines: 'nombre de usuario'' o 'pines 'nombre de usuario''."
      
        if pin in self.data["pines"]:
            pin_value = self.data["pines"][pin] 
            return f"{pin}: {pin_value} es el pin de impresión." 
        else:
            return f"Usuario '{pin}' no encontrado en los pines." 

    def generate_answer(self):
        
        try:
            with open("Base de conocimiento.JSON", "r", encoding="utf-8") as json_file:
                        self.data = json.load(json_file)
        except Exception as e:
                self.text_derecha.insert(tk.END, e)              
        
        text = self.entry_izquierda.get("1.0", tk.END).strip()
        
        keyword = str(self.search_keyword(self.entry_izquierda.get("1.0", tk.END).strip()))

        if keyword == "matriz":

            answer_system=f"{self.data.get('rol', 'Información de rol no encontrada.')}\n{self.search_key(text) }"       
            
        elif keyword == "manuales":
            manual=self.open_file(text)
  
            answer_system=f"{self.data.get('rol', 'Información de rol no encontrada.')}\n {manual}"
        
        elif keyword == "pines":
            
            answer_system=f"{self.data.get('rol', 'Información de rol no encontrada.')}\n{self.search_pin(text) }"    
            print(answer_system)

        elif keyword == "crear":
        
            new_data = self.create_personal_note(text)

            answer_system =f"se ha ingresado la siguiente infomrmación: {text}"
        
        elif keyword in self.keywords:
            answer_system = f"{self.data.get('rol', 'Información de rol no encontrada.')}\n{self.data.get(keyword, 'Información no encontrada.')}"   

        else:
                answer_system = f"{self.data.get('rol', 'Información de rol no encontrada.')}"

        system_message ={"role": "system",
                        "content": answer_system}
        
        input_data = [
            system_message,
            {"role": "user", "content": text}
        ]
        
        try:
            chat_completion = self.client.chat.completions.create(
                messages=input_data,
                model=str(self.config_version_model),
            )

            response = chat_completion.choices[0].message.content.strip()
            self.mostrar_texto_progresivo(response)

        except Exception as e:
            respuesta_error = f"El bot no está disponible en este momento, se presentó el siguiente error. \n{e}"
            self.text_derecha.config(state=tk.NORMAL)
            self.text_derecha.delete("1.0", tk.END)
            self.text_derecha.insert(tk.END, respuesta_error)
            self.text_derecha.config(state=tk.DISABLED)
            print(e)

    def create_personal_note(self, text):
        
        if text.startswith("crear:"):
            text = text.replace("crear:", "", 1).strip()
        elif text.startswith("crear "):
            text = text.replace("crear ", "", 1).strip()
        else:
            return None
        
        if text:
            new_note = {
                "nueva nota": text 
            }
           
            if "notas" in self.data:
                self.data["notas"].append(new_note)
            else:
                self.data["notas"] = [new_note]

            try:
                with open("Base de conocimiento.JSON", "w", encoding="utf-8") as json_file:
                    json.dump(self.data, json_file, ensure_ascii=False, indent=4)

                return text
            
            except Exception as e:
                return None
        else:
            return None

    def mostrar_texto_progresivo(self, texto):
        self.text_derecha.config(state=tk.NORMAL)
        self.text_derecha.delete("1.0", tk.END)
        
        for char in texto:
            self.text_derecha.insert(tk.END, char)
            self.text_derecha.update()  
        
        self.text_derecha.config(state=tk.DISABLED)  

    def enviar_mensaje_inicial(self):
        try:
            with open("Base de conocimiento.JSON", "r", encoding="utf-8") as json_file:
                data = json.load(json_file)
                bienvenida = data.get("bienvenida", "Mensaje de bienvenida no encontrado.")
        except Exception as e:
            bienvenida = f"Error al cargar el archivo JSON: {e}"

        system_message = {
            "role": "system",
            "content": (str(bienvenida))
        }

        input_data = [system_message]

        try:
            chat_completion = self.client.chat.completions.create(
                messages=input_data,
                model=str(self.config_version_model)
            )

            response = chat_completion.choices[0].message.content.strip()
            self.mostrar_texto_progresivo(response)

        except Exception as e:
            respuesta_error = f"El bot está durmiendo y no está disponible en este momento. \n{e}"
            self.text_derecha.config(state=tk.NORMAL)
            self.text_derecha.delete("1.0", tk.END)
            self.text_derecha.insert(tk.END, respuesta_error)
            self.text_derecha.config(state=tk.DISABLED)
            print(e)
            return

    def clean_txt(self):
        self.entry_izquierda.delete("1.0", tk.END)  
        self.text_derecha.config(state=tk.NORMAL) 
        self.text_derecha.delete("1.0", tk.END)  
        self.text_derecha.config(state=tk.DISABLED)


def cargar_configuracion():
    try:
        with open("config.json", "r") as file:
            return json.load(file)
    except FileNotFoundError:
        return {"APY_KEY_BOOT": {"key": "", "modelo": ""}}

def main():
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal (si solo usas Toplevel)
    
    config = cargar_configuracion()
    
    asistente = SkynetBootAssistant(root, config)
    asistente.Create_interface_boot()
    
    root.mainloop()

if __name__ == "__main__":
    main()