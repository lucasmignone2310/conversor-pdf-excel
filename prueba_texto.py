import re
import pdfplumber
import pandas as pd

ruta_pdf = "Prueba sin tablas.pdf"

with pdfplumber.open(ruta_pdf) as pdf:
    primera_pagina = pdf.pages[0]
    texto_crudo = primera_pagina.extract_text() 

    renglones = texto_crudo.split("\n") #Separa el texto en renglones usando el salto de línea como delimitador

    datos_crudos = [] #Lista vacía para almacenar los datos extraídos

    for renglon in renglones:
        renglon_limpio = renglon.strip() #Elimina los espacios en blanco al principio y al final del renglón
        if renglon_limpio :
            datos_crudos.append(renglon_limpio) #Agrega el renglón limpio a la lista de datos crudos

    df = pd.DataFrame(datos_crudos)
    df.to_excel("Prueba sin tablas.xlsx", index=False, header=False) #Exporta el DataFrame a un archivo Excel