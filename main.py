import pdfplumber # pip install pdfplumber. Extrae tablas de los pdfs y las convierte en listas de listas (estructura de datos en memoria)
import pandas as pd # pip install pandas. Permite manipular y analizar datos de manera eficiente, especialmente en forma de tablas (DataFrames).
import tkinter as tk # pip install tkinter. Interfaz gráfica de usuario (ventanas, botones, etc.)
from tkinter import filedialog, messagebox #Permite interactuar con la interfaz gráfica del usuario (ventanas, botones, etc.)
from pdf2docx import Converter # pip install pdf2docx. Convierte pdf a word manteniendo el formato, tablas y diseño del documento.
import pymupdf as fitz # pip install PyMuPDF. Saca información de la imagen de los pdfs  (renderizado de pixeles)
import pytesseract # pip install pytesseract. LLeva la imagen al motor y trae el texto. Es un motor OCR (Optical Character Recognition) (Mensajero)
import docx # pip install python-docx
import os # Librería estándar de Python para interactuar con el sistema operativo (manejo de archivos, rutas, etc.)

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # Ruta al ejecutable de Tesseract. La r antes de las comillas es para indicar que es una cadena cruda (raw string) y evitar problemas con las barras invertidas en la ruta.

def convertir_documento(ruta_pdf, ruta_excel):
    """
    Extrae la tabla de la primera página del PDF especificado 
    y exporta los datos limpios a un archivo Excel en la ruta de destino.
    """
    # Usamos 'with' para abrir el archivo en modo lectura de forma segura.
    # Esto garantiza que el archivo se cierre automáticamente al terminar, liberando memoria.
    with pdfplumber.open(ruta_pdf) as pdf:
    # Las páginas en pdfplumber son una lista, el índice 0 es la primera hoja
        primera_pagina = pdf.pages[0]
        # extract_table() mapea las coordenadas del PDF y devuelve los datos en una lista de listas
        tabla = primera_pagina.extract_table()
        # --- TRANSFORMACIÓN Y CARGA ---
        if tabla :
            #Modo normal 
            #Convertimos la lista de listas en un DataFrame (estructura de tabla virtual en memoria)
            df = pd.DataFrame(tabla)
            
            mensaje_titulo = "Conversión Exitosa"
            mensaje_texto = "El PDF ha sido convertido exitosamente a Excel."
        else:
            #Alternativa si no se detecta tabla, extraemos el texto y lo limpiamos
            texto_crudo = primera_pagina.extract_text() 
            
            renglones = texto_crudo.split("\n") #Separa el texto en renglones usando el salto de línea como delimitador
            
            datos_crudos = [] #Lista vacía para almacenar los datos extraídos
            
            for renglon in renglones:
                renglon_limpio = renglon.strip() #Elimina los espacios en blanco al principio y al final del renglón
                if renglon_limpio :
                    datos_crudos.append([renglon_limpio]) #Agrega el renglón limpio a la lista de datos crudos
            # Creamos un DataFrame con los datos crudos, cada renglón será una fila en la columna A
            df = pd.DataFrame(datos_crudos)
            mensaje_titulo = "Conversión Exitosa"
            mensaje_texto =  "El PDF no tenía tablas detectables.\nSe exportó el texto en crudo en la Columna A para que puedas ordenarlo desde Excel."

        # Exportamos el DataFrame a un archivo Excel físico.
        # index=False y header=False evitan que Pandas agregue números de fila y columna automáticos.
        df.to_excel(ruta_excel, index=False, header=False)
        # Mostramos un mensaje de confirmación al usuario
        messagebox.showinfo(mensaje_titulo, mensaje_texto)

def convertir_a_word(ruta_pdf, ruta_word):
    """
    Convierte un archivo PDF a formato Word (.docx) intentando mantener 
    el formato, las tablas y el diseño original del documento.
    
    Parámetros:
    ruta_pdf: Ubicación del archivo PDF original.
    ruta_word: Ubicación y nombre del archivo Word a generar.
    """
    # Creamos el objeto conversor
    cv = Converter(ruta_pdf)
    # Ejecutamos la conversión (por defecto convierte todas las páginas)
    cv.convert(ruta_word)
    cv.close()

#convertir_documento("Prueba.pdf")

def convertir_a_word_con_ocr(ruta_pdf, ruta_word):
    """
    Convierte un archivo PDF escaneado a formato Word (.docx) usando OCR.
    
    Parámetros:
    ruta_pdf: Ubicación del archivo PDF escaneado.
    ruta_word: Ubicación y nombre del archivo Word a generar.
    """
    pdf_escaneado = fitz.open(ruta_pdf) # Abrimos el archivo PDF escaneado. fitz es la librería que nos permite abrir y manipular archivos PDF.

    primera_pagina = pdf_escaneado[0] # Seleccionamos la primera página del PDF. Los PDFs pueden tener múltiples páginas, y aquí estamos accediendo a la primera.
    
    zoom = fitz.Matrix(3, 3)
    
    imagen = primera_pagina.get_pixmap(matrix = zoom) # Obtenemos un "pixmap" de la primera página. Un pixmap es una representación de la imagen de la página en memoria.
    
    imagen.save("pagina1.png") # Guardamos la imagen de la primera página como un archivo PNG. Esto nos permite tener una representación visual de la página para el OCR.
    
    
    pdf_transcrito = pytesseract.image_to_string("pagina1.png", lang="spa") # Usamos Tesseract para realizar OCR en la imagen guardada. Esto convierte la imagen en texto.
    doc_word = docx.Document()
    
    doc_word.add_paragraph(pdf_transcrito) # Agregamos el texto transcrito al documento Word. Cada párrafo en el documento Word corresponde a un bloque de texto extraído de la imagen.
    
    doc_word.save(ruta_word) # Guardamos el documento Word resultante.

    pdf_escaneado.close()  # Cerramos el archivo PDF para liberar recursos
    os.remove("pagina1.png")

def convertir_a_excel_con_ocr(ruta_pdf, ruta_excel):
    """
    Convierte un archivo PDF escaneado a formato Excel (.xlsx) usando OCR.
    
    Parámetros:
    ruta_pdf: Ubicación del archivo PDF escaneado.
    ruta_excel: Ubicación y nombre del archivo Excel a generar.
    """
    pdf_escaneado = fitz.open(ruta_pdf) # Abrimos el archivo PDF escaneado. fitz es la librería que nos permite abrir y manipular archivos PDF.

    primera_pagina = pdf_escaneado[0] # Seleccionamos la primera página del PDF. Los PDFs pueden tener múltiples páginas, y aquí estamos accediendo a la primera.
    
    zoom = fitz.Matrix(3, 3)
    
    imagen = primera_pagina.get_pixmap(matrix = zoom) # Obtenemos un "pixmap" de la primera página. Un pixmap es una representación de la imagen de la página en memoria.
    
    imagen.save("pagina1.png") # Guardamos la imagen de la primera página como un archivo PNG. Esto nos permite tener una representación visual de la página para el OCR.
    
    
    pdf_transcrito = pytesseract.image_to_string("pagina1.png", lang="spa") # Usamos Tesseract para realizar OCR en la imagen guardada. Esto convierte la imagen en texto.
    
    renglones = pdf_transcrito.split("\n") # Separa el texto en renglones usando el salto de línea como delimitador
    
    datos_crudos = [] # Lista vacía para almacenar los datos extraídos
    
    for renglon in renglones:
        renglon_limpio = renglon.strip() # Elimina los espacios en blanco al principio y al final del renglón
        if renglon_limpio:
            datos_crudos.append([renglon_limpio]) # Agrega el renglón limpio a la lista de datos crudos
    
    # Creamos un DataFrame con los datos crudos, cada renglón será una fila en la columna A
    df = pd.DataFrame(datos_crudos)
    
    # Exportamos el DataFrame a un archivo Excel físico.
    df.to_excel(ruta_excel, index=False, header=False)
    
    pdf_escaneado.close()  # Cerramos el archivo PDF para liberar recursos
    os.remove("pagina1.png")
    messagebox.showinfo("Conversión Exitosa", "Se aplicó OCR y se exportó el texto en la Columna A del Excel.")

    

def seleccionar_archivo():
    """
    Abre el explorador de Windows para que el usuario seleccione un PDF.
    Luego, abre una ventana para elegir dónde guardar el Excel resultante.
    Si ambos pasos son exitosos, ejecuta la conversión.
    """
    #abrimos el explorador de archivos para seleccionar un PDF
    ruta_pdf = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
    # Si el usuario eligió un archivo (y no apretó "Cancelar")
    if ruta_pdf:
        # Abrimos otra ventana para seleccionar dónde guardar el archivo Excel
        ruta_excel = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
        if ruta_excel:
            # Si el usuario eligió un lugar (y no apretó "Cancelar")
            # Llamamos a la función para convertir el PDF
            if usar_ocr.get():
                 convertir_a_excel_con_ocr(ruta_pdf, ruta_excel)
            else:
                 convertir_documento(ruta_pdf, ruta_excel)
        else:
            print("El usuario canceló la operación.")

def seleccionar_archivo_word():
    """
    Abre el explorador de Windows para que el usuario seleccione un PDF.
    Luego, abre una ventana para elegir dónde guardar el Word resultante.
    Si ambos pasos son exitosos, ejecuta la conversión.
    """

    #Abrimos el explorador de archivos para seleccionar un PDF
    ruta_pdf = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
    # Si el usuario eligió un archivo (y no apretó "Cancelar")
    if ruta_pdf:
        #abrimos otra ventana para seleccionar dónde guardar el archivo Word
        ruta_word = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Archivos Word", "*.docx")])
        if ruta_word:
            # Si el usuario eligió un lugar (y no apretó "Cancelar")
            # Llamamos a la función para convertir el PDF
            if usar_ocr.get():
                convertir_a_word_con_ocr(ruta_pdf, ruta_word)
            else:
               convertir_a_word(ruta_pdf, ruta_word)
            messagebox.showinfo("Éxito", "El archivo se convirtió a Word correctamente.")
        else:
            print("El usuario canceló la operación.")

# Inicializamos el objeto principal que representa la ventana de la aplicación
ventana = tk.Tk()
# Asignamos el título que aparecerá en la barra superior
ventana.title("Conversor de PDF a Excel")
#Ajustamos el tamaño de la ventana a 400x250 píxeles
ventana.geometry("400x300")
ventana.config(bg="#f0f0f0")  # Cambiamos el color de fondo a un gris claro
# Agregamos un label (etiqueta) para dar instrucciones al usuario. 
tk.Label(ventana, text="Transforma tus PDF a archivos Excel y Word. ", bg="#f0f0f0", font=("Segoe UI", 12)).pack(pady=20) 

#Variable de estado para el checkbox. Es un objeto que almacena el estado del checkbox (marcado o desmarcado).
usar_ocr = tk.BooleanVar(value=False)  # Inicializamos la variable en False (desmarcado)

#widget checkbox para permitir al usuario decidir si quiere usar OCR.
check_ocr = tk.Checkbutton(ventana, text= "Usar OCR para PDF escaneados", variable=usar_ocr, bg="#f0f0f0", font=("Segoe UI", 9), cursor="hand2")

check_ocr.pack(pady=10)  # Empaquetamos el checkbox en la ventana con un margen vertical de 10 píxeles

# Creamos el botón principal. El parámetro 'command' es el puente que conecta el clic del usuario con nuestra función 'seleccionar_archivo' (sin paréntesis)
boton_excel = tk.Button(ventana, text="Convertir a Excel", command=seleccionar_archivo, bg="#28a745", fg="white", font=("Segoe UI", 10, "bold"), cursor="hand2")
# pack() renderiza el botón en la ventana. 
# padx y pady le dan márgenes (aire) horizontal y verticalmente.
boton_word = tk.Button(ventana, text="Convertir a Word", command=seleccionar_archivo_word, bg="#007bff", fg="white", font=("Segoe UI", 10, "bold"), cursor="hand2")

boton_excel.pack(padx=30, pady=15)
boton_word.pack(padx=30, pady=10)

# Iniciamos el bucle de eventos (mainloop). Esto le dice a Python que mantenga la ventana abierta y se quede escuchando las interacciones del usuario.
ventana.mainloop()