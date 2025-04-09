import requests
import json

class CategoryService:
    def __init__(self):
        # Encabezados de la solicitud
        self.headers = {
            "x-authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJBdWRpZW5jZSI6IlNlc3Npb24iLCJBdXRoZW50aWNhdGlvbkNvbnRleHRDbGFzcyI6bnVsbCwiQXV0aGVudGljYXRpb25NZXRob2RzIjpudWxsLCJBdXRoZW50aWNhdGlvblRpbWUiOiJcL0RhdGUoLTYyMTM1NTk2ODAwMDAwKVwvIiwiQXV0aG9yaXplZFBhcnR5IjpudWxsLCJFeHBpcmF0aW9uVGltZSI6IlwvRGF0ZSgxNzQxMDQxMDU0MjMwKVwvIiwiSXNzdWVkVGltZSI6IlwvRGF0ZSgxNzQxMDI0NDA0MjMwKVwvIiwiSXNzdWVySWRlbnRpZmllciI6ImFzbXMuYXJhbmRhc29mdC5jb20iLCJOb25jZSI6bnVsbCwiU3ViamVjdCI6Ins2OTQzYmRkNy1lOTBjLTQwZGUtOGEwMy03NDFjNGM0ZmU4MDV9IiwiU3ViamVjdElkZW50aWZpZXIiOiI5OTAyMDkiLCJVc2VyTmFtZSI6bnVsbH0.s1uN70fRBMnmvQxfRYeL1ZAkkJN29oQhbbPfenR14D0"}

    def obtener_servicios(self, type):
        """Consulta los servicios disponibles."""
        payload = {
            "criteria": " ",
            "projectId": 942,
            "type": type,
            "contractId": 1925,
        }
        url = "https://itsm.sonda.com/asmsapi/api/v9/service/contract/search"

        response = requests.post(url, headers=self.headers, json=payload)
        if response.status_code == 200:
            servicios = response.json()
            return servicios.get("content", [])
        else:
            print(f"Error al realizar la solicitud: {response.status_code}")
            print(response.text)
            return []

    def obtener_categorias(self, service_id,type):
        """Consulta las categorías de un servicio específico."""
        url = "https://itsm.sonda.com/asmsapi/api/v9/item/categoriesHierarchy"
        payload = {
            "orderField": "Hierarchy",
            "orderType": "Desc",
            "pageIndex": 0,
            "pageSize": 100,
            "ItemType": type,
            "ServiceId": service_id,
            "search": "",
        }

        response = requests.post(url, headers=self.headers, json=payload)
        if response.status_code == 200:
            return response.json().get("content", [])
        else:
            print(f"Error al obtener categorías para el servicio {service_id}: {response.status_code}")
            print(response.text)
            return []

    def peticiones_get(self, url):
        response = requests.get(url, headers=self.headers)
        if response.status_code == 200:
            return response.json().get("content", [])
        else:
            print(f"Error al obtener ubicaciones: {response.status_code}")
            print(response.text)
            return []


    def procesar_servicios_y_categorias(self):
        """Procesa todos los servicios, categorías y ubicaciones, y guarda los datos en un archivo JSON consolidado."""
        servicios = self.obtener_servicios(4)
        incidentes = self.obtener_servicios(1)

        ubicaciones =self.peticiones_get("https://itsm.sonda.com/asmsapi/api/v9/project/942/locations?criteria=%20&locationId=")
        grupo =self.peticiones_get("https://itsm.sonda.com/asmsapi/api/v9/service/7005/state/29547/group/list")
        tipo_registro=self.peticiones_get("https://itsm.sonda.com/asmsapi/api/v9/catalog/registry_type?language=0")
        responsables=self.peticiones_get("https://itsm.sonda.com/asmsapi/api/v9/group/19667/project/942/specialists?available=true")
        sede=self.peticiones_get("https://itsm.sonda.com/asmsapi/api/v9/additionalsfields/44265/type/3/values?dataType=&parentId=&catalogId=&userId=965961")
        impacto=self.peticiones_get("https://itsm.sonda.com/asmsapi/api/v9/catalog/impact?language=0")
        urgencia =self.peticiones_get("https://itsm.sonda.com/asmsapi/api/v9/catalog/urgency?language=0")
        
        resultado = {
            "requerimientos de servicios": {},
            "incidentes de servicios":{},
            "tipo servicio": [{"name": "Requerimientos de Servicio"},{"name": "Incidentes"}],
            "grupo_resolutor": grupo,
            "ubicación": ubicaciones,
            "responsable": responsables,
            "tipo_registro":tipo_registro,
            "sede":sede,
            "impacto":impacto,
            "urgencia":urgencia
        }

        for servicio in servicios:
            service_id = servicio.get("id")
            service_name = servicio.get("name", "Sin nombre")
            print(f"Procesando servicio: {service_name} (ID: {service_id})")
            

            categorias = self.obtener_categorias(service_id,4)
            resultado["requerimientos de servicios"][service_name] = categorias


        for servicio in incidentes:
            service_id = servicio.get("id")
            service_name = servicio.get("name", "Sin nombre")
            print(f"Procesando incidentes: {service_name} (ID: {service_id})")
            
  
            categorias = self.obtener_categorias(service_id,1)
            resultado["incidentes de servicios"][service_name] = categorias


        with open("categorias_y_ubicaciones.json", "w", encoding="utf-8") as file:
            json.dump(resultado, file, indent=4, ensure_ascii=False)
        print("Archivo 'categorias_y_ubicaciones.json' creado con éxito.")


if __name__ == "__main__":
    service = CategoryService()
    service.procesar_servicios_y_categorias()
