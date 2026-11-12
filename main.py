import pdfplumber 
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter

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
        #Convertimos la lista de listas en un DataFrame (estructura de tabla virtual en memoria)
        df = pd.DataFrame(tabla)
        # Exportamos el DataFrame a un archivo Excel físico.
        # index=False y header=False evitan que Pandas agregue números de fila y columna automáticos.
        df.to_excel(ruta_excel, index=False, header=False)
        # Mostramos un mensaje de confirmación al usuario
        messagebox.showinfo("Conversión Exitosa", "El PDF ha sido convertido exitosamente a Excel.")

def convertir_a_word(ruta_pdf, ruta_word):
    print(f"Convirtiendo {ruta_pdf} a Word...")

#convertir_documento("Prueba.pdf")

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
            convertir_documento(ruta_pdf, ruta_excel)
        else:
            print("El usuario canceló la operación.")

# Inicializamos el objeto principal que representa la ventana de la aplicación
ventana = tk.Tk()
# Asignamos el título que aparecerá en la barra superior
ventana.title("Conversor de PDF a Excel")
#Ajustamos el tamaño de la ventana a 400x250 píxeles
ventana.geometry("400x250")
ventana.config(bg="#f0f0f0")  # Cambiamos el color de fondo a un gris claro
# Agregamos un label (etiqueta) para dar instrucciones al usuario. 
tk.Label(ventana, text="Transforma tus PDF de tablas a Excel con un clic ", bg="#f0f0f0", font=("Segoe UI", 12)).pack(pady=20) 
# Creamos el botón principal. El parámetro 'command' es el puente que conecta el clic del usuario con nuestra función 'seleccionar_archivo' (sin paréntesis)
boton_seleccionar = tk.Button(ventana, text="Seleccionar PDF", command=seleccionar_archivo, bg="#28a745", fg="white", font=("Segoe UI", 10, "bold"), cursor="hand2")
# pack() renderiza el botón en la ventana. 
# padx y pady le dan márgenes (aire) horizontal y verticalmente.
boton_seleccionar.pack(padx=30, pady=30)
# Iniciamos el bucle de eventos (mainloop). Esto le dice a Python que mantenga la ventana abierta y se quede escuchando las interacciones del usuario.
ventana.mainloop()