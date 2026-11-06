import pdfplumber
import pandas as pd

ruta_pdf = "Prueba.pdf"
    # Usamos 'with' para abrir el archivo en modo lectura de forma segura.
    # Esto garantiza que el archivo se cierre automáticamente al terminar, liberando memoria.
with pdfplumber.open(ruta_pdf) as pdf:
    # Las páginas en pdfplumber son una lista, el índice 0 es la primera hoja
    primera_pagina = pdf.pages[0]
    # extract_table() mapea las coordenadas del PDF y devuelve los datos en una lista de listas
    tabla = primera_pagina.extract_table()
    print(tabla)
    # --- TRANSFORMACIÓN Y CARGA ---
    #Convertimos la lista de listas en un DataFrame (estructura de tabla virtual en memoria)
    df = pd.DataFrame(tabla)
    # Exportamos el DataFrame a un archivo Excel físico.
    # index=False y header=False evitan que Pandas agregue números de fila y columna automáticos.
    df.to_excel("Prueba.xlsx", index=False, header=False)

    print("Archivo Excel generado exitosamente!")