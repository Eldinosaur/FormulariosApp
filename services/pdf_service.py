import os
from template.pdf_template import create_pdf

# Ruta del archivo PDF
PDF_PATH = 'pdf'

os.makedirs(PDF_PATH, exist_ok=True)  # Crear la carpeta 'pdf' si no existe



# Función para crear el archivo PDF con los datos del formulario
def create_pdf_file(data):
    
    # Crear el PDF con los datos del formulario
    create_pdf(data)