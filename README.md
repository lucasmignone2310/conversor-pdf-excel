# Conversor Multiformato: PDF a Excel y Word

Una herramienta automatizada con interfaz gráfica en Python para extraer información de archivos PDF y exportarla a planillas de Excel (`.xlsx`) o documentos de Word (`.docx`).

## Características

- **Interfaz Gráfica (GUI):** Ventana interactiva y amigable con botones dedicados para cada formato de exportación.
- **Flujo dinámico:** Utiliza el explorador nativo de Windows para elegir qué PDF abrir y dónde guardar el archivo resultante.
- **Conversión a Word:** Reconstruye el texto, las tablas y el diseño del PDF original a un formato Word nativo.
- **Portabilidad:** Empaquetado en un ejecutable `.exe` para usar en cualquier computadora con Windows sin necesidad de instalar Python.

## ⚠️ Alcance y Limitaciones (Versión 2.0)

- **Para Excel:** El programa solo extrae información de *tablas nativas estructuradas* (con bordes o celdas definidas lógicamente). No procesa correctamente texto alineado con espacios/tabulaciones.
- **Para Word:** La precisión de la conversión depende del diseño del PDF. Al pasar de coordenadas fijas a texto fluido, pueden ocurrir ligeras variaciones de formato.
- **Limitación General:** No funciona con PDFs escaneados o imágenes (la integración con OCR está planeada para futuras versiones).

## Tecnologías utilizadas

- **Python 3.11**
- **pdfplumber:** Extracción de datos y lectura de coordenadas espaciales.
- **pdf2docx:** Motor de conversión estructural de PDF a Word.
- **pandas:** Estructuración de datos y transformación.
- **openpyxl:** Motor de escritura para archivos Excel.
- **Tkinter:** Desarrollo de la interfaz gráfica.
- **PyInstaller:** Empaquetado del software.

## Uso para el usuario final

1. Abrir el programa (ejecutando el script o el `.exe`).
2. Hacer clic en "Convertir a Excel" o "Convertir a Word" según la necesidad.
3. Buscar y seleccionar el archivo PDF origen.
4. Elegir la carpeta de destino y el nombre para el nuevo archivo.
5. Esperar la ventana de confirmación de éxito.

## Cómo preparar el entorno local (Desarrolladores)

1. Clonar el repositorio.
2. Crear un entorno virtual: `python -m venv venv`
3. Activar el entorno e instalar dependencias: `pip install pdfplumber pandas openpyxl pdf2docx pyinstaller`
4. Ejecutar el código fuente: `python main.py`
5. Para compilar el ejecutable por tu cuenta, correr: `pyinstaller --onefile --windowed main.py`