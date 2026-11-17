import pymupdf as fitz # pip install PyMuPDF. Saca información de la imagen de los pdfs  (renderizado de pixeles)
import pytesseract # pip install pytesseract. LLeva la imagen al motor y trae el texto. Es un motor OCR (Optical Character Recognition) (Mensajero)
import docx # pip install python-docx

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # Ruta al ejecutable de Tesseract. La r antes de las comillas es para indicar que es una cadena cruda (raw string) y evitar problemas con las barras invertidas en la ruta.

pdf_escaneado = fitz.open("imagen a archivo.pdf") # Abrimos el archivo PDF escaneado. fitz es la librería que nos permite abrir y manipular archivos PDF.

primera_pagina = pdf_escaneado[0] # Seleccionamos la primera página del PDF. Los PDFs pueden tener múltiples páginas, y aquí estamos accediendo a la primera.

zoom = fitz.Matrix(3, 3)

imagen = primera_pagina.get_pixmap(matrix = zoom) # Obtenemos un "pixmap" de la primera página. Un pixmap es una representación de la imagen de la página en memoria.

imagen.save("pagina1.png") # Guardamos la imagen de la primera página como un archivo PNG. Esto nos permite tener una representación visual de la página para el OCR.

pdf_transcrito = pytesseract.image_to_string("pagina1.png", lang="spa") # Usamos Tesseract para realizar OCR en la imagen guardada. Esto convierte la imagen en texto.
doc_word = docx.Document()

doc_word.add_paragraph(pdf_transcrito) # Agregamos el texto transcrito al documento Word. Cada párrafo en el documento Word corresponde a un bloque de texto extraído de la imagen.

doc_word.save("pagina1.docx") # Guardamos el documento Word resultante.

print(pdf_transcrito)
