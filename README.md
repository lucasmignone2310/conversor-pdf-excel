
# Conversor de PDF a Excel

Una herramienta automatizada con interfaz gráfica en Python para extraer tablas estructuradas de archivos PDF y exportarlas a planillas de Excel (`.xlsx`).

## Características

- **Interfaz Gráfica (GUI):** Ventana interactiva y amigable para el usuario.
- **Flujo dinámico:** Utiliza el explorador nativo de Windows para elegir qué PDF abrir y dónde guardar el Excel resultante.
- **Portabilidad:** Posibilidad de empaquetar el código en un ejecutable `.exe` para usar en cualquier computadora sin necesidad de instalar Python.

## ⚠️ Alcance y Limitaciones (Versión 1.0)

- El programa solo extrae información de **tablas nativas estructuradas** (con bordes o celdas definidas lógicamente en el PDF).
- **No** procesa correctamente texto alineado en columnas mediante espacios o tabulaciones.
- **No** funciona con PDFs escaneados o imágenes (la integración con OCR está planeada para futuras versiones).

## Tecnologías utilizadas

- Python 3.11
- pdfplumber (Extracción de datos y lectura de coordenadas espaciales)
- pandas (Estructuración de datos y transformación)
- openpyxl (Motor de escritura para archivos Excel)
- Tkinter (Desarrollo de la interfaz gráfica)
- PyInstaller (Empaquetado del software)

## Uso para el usuario final

1. Abrir el programa (ejecutando el script o el `.exe`).
2. Hacer clic en "Seleccionar PDF" y buscar el archivo origen.
3. Elegir la carpeta de destino y el nombre para el nuevo archivo Excel.
4. Esperar la ventana de confirmación de éxito.

## Cómo preparar el entorno local (Desarrolladores)

1. Clonar el repositorio.
2. Crear un entorno virtual: `python -m venv venv`
3. Activar el entorno e instalar dependencias: `pip install pdfplumber pandas openpyxl pyinstaller`
4. Ejecutar el código fuente: `python main.py`
5. Para compilar el ejecutable por tu cuenta, correr: `pyinstaller --onefile --windowed main.py`
