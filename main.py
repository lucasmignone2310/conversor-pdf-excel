import pdfplumber

ruta_pdf = "Prueba.pdf"
    # Usamos 'with' para abrir el archivo en modo lectura de forma segura.
    # Esto garantiza que el archivo se cierre automáticamente al terminar, liberando memoria.
with pdfplumber.open(ruta_pdf) as pdf:
    # Las páginas en pdfplumber son una lista, el índice 0 es la primera hoja
    primera_pagina = pdf.pages[0]
    # extract_table() mapea las coordenadas del PDF y devuelve los datos en una lista de listas
    tabla = primera_pagina.extract_table()
    print(tabla)