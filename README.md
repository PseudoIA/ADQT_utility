# Proyecto de Gestión de Usuarios y Soporte Técnico

Este proyecto proporciona una solución para la gestión de usuarios en Active Directory (AD) y la interacción con bases de datos SQL Server, integrando también una API de inteligencia artificial para la mejora de la comunicación. Incluye una interfaz gráfica de usuario (GUI) y un temporizador para la gestión de tiempos de respuesta y cumplimiento de SLAs.

## Funcionalidades Principales

### Autenticación en Active Directory
- **Entrada de Credenciales**: Permite a los usuarios ingresar sus credenciales para autenticarse en el entorno de AD.
- **Gestión de Usuarios**: Facilita tareas administrativas en AD, como cambiar contraseñas y desbloquear cuentas.

### Consultas y Gestión de Bases de Datos
- **Conexión a SQL Server**: Establece una conexión con bases de datos SQL Server para realizar consultas.
- **Control de Uso**: Limita el número de consultas realizadas para gestionar recursos y evitar abusos.

### Integración con API de Inteligencia Artificial
- **Generación de Respuestas**: Utiliza una API para generar respuestas automáticas, corrección ortográfica y enriquecimiento técnico.
- **Asistente Virtual**: Proporciona asistencia en la redacción de correos electrónicos y la mejora de textos.

### Interfaz Gráfica de Usuario (GUI)
- **Ventana de Credenciales**: Interfaz para ingresar credenciales de manera segura.
- **Experiencia del Usuario**: Interfaz amigable con elementos visuales claros y simulación de escritura progresiva.

### Temporizador de Gestión de Tiempo
- **Visualización del Tiempo**: Muestra un temporizador para indicar el tiempo restante para la solución de problemas o cumplimiento de SLAs.
- **Alertas de Tiempo**: Cambia el color de la visualización cuando el tiempo restante es crítico.

## Instalación

1. **Clona el Repositorio**
   ```sh
   git clone <url-del-repositorio>
Instala las Dependencias Asegúrate de tener Python y los siguientes paquetes instalados:

tkinter
pymssql
groq
Puedes instalar las dependencias necesarias usando pip:

sh
Copiar código
pip install pymssql groq
Configura las Credenciales Asegúrate de reemplazar las credenciales de la API de Groq en el código en Qery_ad con tus propias credenciales.

Uso
Ejecuta el Proyecto Navega a la carpeta del proyecto y ejecuta el archivo principal:

sh
Copiar código
python <nombre-del-archivo>.py
Interacción

Buscar: Usa el botón "Buscar" para consultar datos en la interfaz.
Cambiar Contraseña: Usa el botón "Cambiar Contraseña" para cambiar la contraseña de un usuario.
Desbloquear: Usa el botón "Desbloquear" para desbloquear una cuenta.
Iniciar Temporizador: Usa el botón "Iniciar temporizador" para comenzar a contar el tiempo.
Contribuciones
Las contribuciones son bienvenidas. Por favor, sigue los siguientes pasos para contribuir al proyecto:

Haz un fork del repositorio.
Crea una nueva rama para tu característica o corrección de errores.
Envía un pull request con una descripción detallada de los cambios.
Licencia
Este proyecto está licenciado bajo la Licencia MIT.

Contacto
Para cualquier consulta, puedes contactar a [tu-email@dominio.com].

