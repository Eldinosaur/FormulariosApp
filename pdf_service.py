import os
from fpdf import FPDF
from datetime import datetime

PDF_PATH = 'pdf'

# Crear la carpeta para almacenar los archivos PDF si no existe
if not os.path.exists(PDF_PATH):
    os.makedirs(PDF_PATH)

# Función para crear un archivo PDF con los datos proporcionados
def create_pdf(data):
    data_date = datetime.now().strftime('%Y-%m-%d')
    
    # Crear un nuevo PDF
    pdf = FPDF()
    pdf.add_page()
    
    # Agregar un título al PDF
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Registro de Entrevista Social", ln=True)
    
    # Configurar la fuente y el tamaño
    pdf.set_font("Arial", size=12)
    # Agregar el contenido al PDF
    pdf.cell(200, 10, txt=f"ID: {data['id']}", ln=True)
    pdf.cell(200, 10, txt=f"Name: {data['name']}", ln=True)
    pdf.cell(200, 10, txt=f"Email: {data['email']}", ln=True)

    # Guardar el PDF en un archivo
    pdf.output(f"{PDF_PATH}/formulario_{data['id']}.pdf")