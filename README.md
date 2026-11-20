# Conversor Multiformato: PDF a Excel y Word (v3.0)

Herramienta de escritorio con interfaz gráfica en Python para extraer información de archivos PDF —tanto nativos como escaneados— y exportarla a planillas de Excel (`.xlsx`) o documentos de Word (`.docx`).

## Características

- **Interfaz Gráfica (GUI):** Ventana intuitiva con botones dedicados para cada formato y un selector (*Checkbox*) para activar el modo OCR.
- **Soporte para PDFs Escaneados (OCR):** Integra **Tesseract OCR** y **PyMuPDF** para renderizar documentos escaneados en alta resolución y transcribir su contenido automáticamente a Word o Excel.
- **Flujo dinámico:** Utiliza el explorador nativo de Windows para elegir qué PDF abrir y dónde guardar el archivo resultante.
- **Manejo de PDFs sin formato (Fallback):** Si el documento no posee tablas estructuradas (o proviene de OCR), el sistema realiza un volcado limpio de texto en la Columna A, listo para ordenar con "Texto en columnas" en Excel.
- **Conversión a Word:** Reconstruye texto, tablas y diseño en PDFs nativos, o genera un documento limpio con el texto transcrito en modo OCR.
- **Limpieza automática:** Elimina los archivos de imagen temporales generados durante el proceso de OCR al finalizar cada conversión.
- **100% Portable:** Empaquetado en un único ejecutable `.exe` que incluye el motor de Tesseract embebido, listo para usar en Windows sin instalar Python ni dependencias externas.

## ⚠️ Alcance y Limitaciones (Versión 3.0)

- **Extracción de páginas:** La conversión a Excel (nativa y OCR) y la conversión a Word con OCR procesan actualmente la **primera página** del documento. La conversión nativa a Word procesa el documento completo.
- **Tablas en Excel:** En modo nativo, extrae tablas estructuradas con bordes definidos; si no las detecta (o si se usa modo OCR), exporta los renglones limpios en la Columna A.
- **Calidad del OCR:** La precisión en documentos escaneados depende de la nitidez de la imagen original (configurado por defecto para idioma español).

## Tecnologías utilizadas

- **Python 3.11**
- **Tkinter:** Interfaz gráfica de usuario.
- **pdfplumber:** Extracción de tablas y texto en PDFs nativos.
- **pdf2docx:** Conversión estructural de PDF nativo a Word.
- **PyMuPDF (`fitz`):** Renderizado de páginas PDF a imagen de alta resolución.
- **pytesseract + Tesseract-OCR:** Reconocimiento óptico de caracteres (OCR).
- **python-docx:** Generación de documentos Word a partir de texto OCR.
- **pandas & openpyxl:** Transformación de datos y exportación a Excel.
- **PyInstaller:** Empaquetado en ejecutable portable.

## Uso para el usuario final

1. Descargar `ConversorPDF_v3.exe` desde la sección **Releases** de este repositorio y ejecutarlo.
2. *(Opcional)* Tildar la casilla **"Usar OCR para PDF escaneados"** si el archivo es una imagen o documento escaneado.
3. Hacer clic en **"Convertir a Excel"** o **"Convertir a Word"**.
4. Seleccionar el archivo PDF de origen y elegir dónde guardar el archivo resultante.
5. Esperar el cartel de confirmación de éxito.

## Cómo preparar el entorno local (Desarrolladores)

1. Clonar el repositorio.
2. Instalar [Tesseract-OCR](https://github.com/UB-Mannheim/tesseract/wiki) en Windows (incluyendo el paquete de idioma español `spa`) en la ruta por defecto: `C:\Program Files\Tesseract-OCR\`.
3. Crear y activar un entorno virtual:
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```
4. Instalar las dependencias:
   ```bash
   pip install -r requirements.txt
   ```
5. Ejecutar el código fuente:
   ```bash
   python main.py
   ```
6. Para compilar el ejecutable portable con Tesseract embebido:
   ```bash
   pyinstaller --noconsole --onefile --name="ConversorPDF_v3" --add-data "C:\Program Files\Tesseract-OCR;Tesseract-OCR" main.py
   ```