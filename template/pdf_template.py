from fpdf import FPDF

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Formulario de Entrevista Social', 0, 1, 'C')

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Pagina {self.page_no()}', 0, 0, 'C')

def create_pdf(data):
    pdf = PDF()
    pdf.add_page()
    pdf.set_font('Arial', '', 12)

    pdf.cell(0, 10, f"ID: {data['id']}", ln=True)
    pdf.cell(0, 10, f"Name: {data['name']}", ln=True)
    pdf.cell(0, 10, f"Email: {data['email']}", ln=True)
    pdf.cell(0, 10, f"Date: {data['date']}", ln=True)

    # Guardar el PDF con un nombre basado en el ID
    pdf.output(f"pdf/{data['id']}.pdf")