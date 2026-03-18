from fpdf import FPDF
import os
from PIL import Image, ExifTags

PAGE_WIDTH = 190


class PDF(FPDF):

    def header(self):

        self.set_font("Arial", "B", 14)
        

        if self.page_no() == 1:
            self.cell(0, 10, "FICHA SOCIOECONÓMICA", 0, 1, "C")
            self.rect(165, 10, 30, 35)
            
            if hasattr(self, "photo_path") and self.photo_path and os.path.exists(self.photo_path):
                self.image(self.photo_path, x=165, y=10, w=30, h=35)

            self.set_y(50)

        else:
            self.ln(10)

    def footer(self):

        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Página {self.page_no()}", 0, 0, "C")


# -----------------------------
# HELPERS
# -----------------------------

def fix_image_orientation(image_path):

    try:
        image = Image.open(image_path)

        for orientation in ExifTags.TAGS.keys():
            if ExifTags.TAGS[orientation] == 'Orientation':
                break

        exif = image._getexif()

        if exif is not None:
            orientation_value = exif.get(orientation)

            if orientation_value == 3:
                image = image.rotate(180, expand=True)
            elif orientation_value == 6:
                image = image.rotate(270, expand=True)
            elif orientation_value == 8:
                image = image.rotate(90, expand=True)

        fixed_path = "temp_fixed.jpg"
        image.save(fixed_path)

        return fixed_path

    except Exception:
        return image_path

def checkbox(value):
    return "[X]" if value else "[ ]"


def section_title(pdf, title):
    pdf.set_font("Arial", "B", 10)
    pdf.cell(PAGE_WIDTH, 8, title, border=1, ln=True)


def table_row(pdf, widths, data, height=8):

    # evita errores por desalineación
    if len(data) < len(widths):
        data = list(data) + [""] * (len(widths) - len(data))

    for i in range(len(widths)):
        pdf.cell(widths[i], height, str(data[i]), border=1)

    pdf.ln(height)
    
def table_row_mixed(pdf, widths, data, styles, height=8):

    # Ajuste por seguridad
    if len(data) < len(widths):
        data = list(data) + [""] * (len(widths) - len(data))

    if len(styles) < len(widths):
        styles = list(styles) + [""] * (len(widths) - len(styles))

    for i in range(len(widths)):

        pdf.set_font("Arial", styles[i], 10)
        pdf.cell(widths[i], height, str(data[i]), border=1)

    pdf.ln(height)

    # Reset
    pdf.set_font("Arial", "", 10)

def table_row_multiline(pdf, widths, data, styles, line_height=6):

    if len(data) < len(widths):
        data = list(data) + [""] * (len(widths) - len(data))

    if len(styles) < len(widths):
        styles = list(styles) + [""] * (len(widths) - len(styles))

    # calcular número de líneas REAL por celda
    line_counts = []

    for i in range(len(data)):
        text = str(data[i])
        col_width = widths[i]

        # dividir por saltos manuales
        lines = text.split("\n")
        total_lines = 0

        for line in lines:
            text_width = pdf.get_string_width(line)

            # calcular cuántas líneas ocupa según ancho
            lines_needed = max(1, int(text_width / (col_width - 2)) + 1)
            total_lines += lines_needed

        line_counts.append(total_lines)

    # altura máxima de la fila
    max_lines = max(line_counts)
    row_height = max_lines * line_height

    # dibujar celdas con MISMA altura
    x_start = pdf.get_x()
    y_start = pdf.get_y()

    for i in range(len(widths)):

        pdf.set_xy(x_start, y_start)
        pdf.set_font("Arial", styles[i], 10)

        x_current = pdf.get_x()
        y_current = pdf.get_y()

        pdf.multi_cell(widths[i], line_height, str(data[i]), border=1)

        # volver arriba para siguiente celda
        pdf.set_xy(x_current + widths[i], y_current)

        x_start += widths[i]

    pdf.ln(row_height)

    pdf.set_font("Arial", "", 10)

def check_page_space(pdf, space_needed):
    if pdf.get_y() + space_needed > 270:
        pdf.add_page()


def box(pdf, text, h=20):
    pdf.multi_cell(PAGE_WIDTH, h / 4, str(text), border=1)


# -----------------------------
# SECCIONES
# -----------------------------
def fecha_visita(pdf, data):
    table_row_mixed(pdf, [40, 150], [
        "Fecha de Visita", data.get("visit_date", "")
    ], ["B", ""])

def datos_personales(pdf, data):

    section_title(pdf, "1. DATOS PERSONALES DEL COLABORADOR")

    table_row_mixed(pdf, [40, 150], [
        "Nombre", data.get("name", "")
    ],["B",""])

    table_row_mixed(pdf, [40, 55, 40, 55], [
        "Cédula", data.get("ci", ""),
        "Edad", data.get("age", "")
    ],["B","","B",""])
    
    table_row_mixed(pdf,[80,110], [
        "Lugar y Fecha de Nacimiento", data.get("birth_date_place", "")
    ],["B",""])
    
    masculino = checkbox(data.get("gender") == "masculino")
    femenino = checkbox(data.get("gender") == "femenino")
    otro = checkbox(data.get("gender") == "otro")   
    table_row_mixed(pdf, [40, 150], [
        "Genero", f"{masculino} Masculino     {femenino} Femenino     {otro} Otro"
    ],["B",""])

    casado = checkbox(data.get("marital_status") == "casado")
    soltero = checkbox(data.get("marital_status") == "soltero")
    separado = checkbox(data.get("marital_status") == "separado")
    divorciado = checkbox(data.get("marital_status") == "divorciado")
    unionlibre = checkbox(data.get("marital_status") == "unión libre")
    viudo = checkbox(data.get("marital_status") == "viudo")
    table_row_mixed(pdf, [40, 150], [
        "Estado civil",
        f"{casado} Casado     {soltero} Soltero     {separado} Separado     {divorciado} Divorciado     {unionlibre} Unión Libre     {viudo} Viudo"
    ],["B",""])

    table_row_mixed(pdf, [40, 150], [
        "Email", data.get("email", "")
    ],["B",""])
    
    table_row_mixed(pdf, [40,50, 50, 50],[
        "Área de trabajo", data.get("work_area", ""), 
        "Puesto que desempeña", data.get("job_title", "")
    ], ["B","","B",""])
    
    table_row_mixed(pdf, [55, 135],[
        "Fecha de Ingreso", data.get("entry_date", "")
    ], ["B",""])
    
    table_row_mixed(pdf, [55,15,25,35,40,20],[
        "Tiene alguna discapacidad", data.get("disability", ""),
        "Tipo", data.get("disability_type", ""),
        "Porcentaje", data.get("disability_percentage", "")
    ],["B","","B","","B",""])

    profesional = checkbox(data.get("educational_level") == "profesional")
    bachiller = checkbox(data.get("educational_level") == "bachiller")
    basica = checkbox(data.get("educational_level") == "basica")
    primaria = checkbox(data.get("educational_level") == "primaria")
    nuloestudio = checkbox(data.get("educational_level") == "nuloestudio")
    table_row_mixed(pdf, [40, 150], [
        "Nivel de Educación", 
        f"{profesional} Profesional     {bachiller} Bachiller     {basica} Básica     {primaria} Primaria     {nuloestudio} Sin Estudios"
    ],["B",""])
    
    table_row_mixed(pdf,[40,150],[
        "Observaciones:", data.get("educational_observations", "")
    ],["B",""])
    
    table_row_mixed(pdf, [50,15,30,35,30,30],[
        "¿Actualmente estudia?", data.get("currently_studying", ""),
        "Carrera", data.get("study_career", ""),
        "Nivel", data.get("study_level", "")
    ],["B","","B","","B",""])

    table_row_mixed(pdf, [40, 150], [
        "Dirección Domiciliaria", data.get("address", "")
    ],["B",""])
    

    table_row_mixed(pdf, [40, 150], [
        "Teléfono", data.get("phone", "")
    ],["B",""])

    table_row_mixed(pdf, [60, 130], [
        "Contacto de emergencia", data.get("emergency_contact", "")
    ],["B",""])


def miembros_hogar(pdf, data):

    section_title(pdf, "2. MIEMBROS DEL HOGAR")
    table_row_mixed(pdf, [190], [
        "Célula Social - Composición Familiar del Colaborador (viven dentro del hogar)"
    ],["B"])

    headers = ["N°","Nombre", "Edad", "Parentesco", "Ocupación", "Ingreso"]
    widths = [10,50, 20, 40, 40, 30]

    table_row_mixed(pdf, widths, headers,["B","B","B","B","B","B"])

    miembros = data.get("family_members", [])

    # SI NO HAY DATOS
    if not miembros:
        table_row(pdf, [190], ["No se registran miembros del hogar"])
        return

    # SOLO RECORRE LOS QUE EXISTEN
    for i, m in enumerate(miembros, start=1):

        row = [
            str(i),
            str(m.get("name", "")),
            str(m.get("age", "")),
            str(m.get("relation", "")),
            str(m.get("job", "")),
            str(m.get("income", ""))
        ]

        # calcular altura dinámica
        line_counts = []

        for j in range(len(row)):
            text = row[j]
            col_width = widths[j]
            text_width = pdf.get_string_width(text)

            lines = max(1, int(text_width / col_width) + 1)
            line_counts.append(lines)

        row_height = max(line_counts) * 6

        check_page_space(pdf, row_height)

        for j in range(len(row)):
            x_current = pdf.get_x()
            y_current = pdf.get_y()

            pdf.multi_cell(widths[j], 6, row[j], border=1)
            pdf.set_xy(x_current + widths[j], y_current)

        pdf.ln(row_height)
def info_familia(pdf, data):
    
    table_row_mixed(pdf, [190], [
        "Información de la Familia:"
    ],["B"])

    nuclear = checkbox(data.get("family_type") == "nuclear")
    cosanguinea = checkbox(data.get("family_type") == "cosanguinea")
    afines = checkbox(data.get("family_type") == "afines")
    ampliada = checkbox(data.get("family_type") == "ampliada")
    materna = checkbox(data.get("family_type") == "materna")
    paterna = checkbox(data.get("family_type") == "paterna")
    unipersonal = checkbox(data.get("family_type") == "unipersonal")
    table_row_multiline(pdf, [40, 150],
    [
        "Tipo de Familia:" "\n  ",
        f"{nuclear} Nuclear     {cosanguinea} Extensa Cosanguínea     {afines} Extensa Afines     {ampliada} Ampliada\n"
        f"{materna} Monoparental Materna     {paterna} Monoparental Paterna     {unipersonal} Unipersonal"
    ],
    ["B", ""]
    )

    siotroshijos = checkbox(data.get("children_outside_marriage") == "si")
    nootroshijos = checkbox(data.get("children_outside_marriage") == "no")
    table_row_mixed(pdf, [60,30,60,40],[
        "N° Hijos del colaborador", data.get("children_count", ""),
        "Existen hijos fuera del Matrimonio", f"{siotroshijos} Sí     {nootroshijos} No"
    ],["B","","B",""])
    table_row_mixed(pdf, [60,130],[
        "Paga pensión alimenticia", data.get("pays_alimony", "")
    ],["B",""])
    table_row_mixed(pdf, [70,120],[
        "N° de Matrimonio o Relación Actual", data.get("marriage_number", "")
    ],["B",""])
    table_row_mixed(pdf, [110,80],[
        "Comparte con su familia tiempo de descanso y recreación", data.get("family_time", "")
    ],["B",""])
    table_row_mixed(pdf, [40,150],[
        "Frecuencia", data.get("family_time_frequency", "")
    ],["B",""])
    table_row_mixed(pdf, [40,150],[
        "Actividades:", data.get("family_activities", "")
    ],["B",""])
    table_row_multiline(pdf, [80,110],[
        "Qué actividades realiza\n"
        "fuera de su jornada laboral:", data.get("other_activities", "\n  ")
    ],["B",""])
    otrasactividades = checkbox(data.get("other_activities") == "si")
    nootrasactividades = checkbox(data.get("other_activities") == "no")
    table_row_mixed(pdf, [75,25,40,50],[
        "En su domicilio realiza otras actividades:",
        f"{otrasactividades} Sí     {nootrasactividades} No",
        "Especificar:", data.get("other_activities_specify", "")
    ],["B","", "B",""])
    table_row_mixed(pdf, [75,115],[
        "Describa cuales son sus hobbies:", data.get("hobbies", "")
    ],["B",""])
    table_row_mixed(pdf, [75,115],[
        "Que tiempo le dedica a sus hobbies:", data.get("hobbies_time", "")
    ],["B",""])

    muybuenarelacionpareja = checkbox(data.get("partner_relationship") == "muy buena")
    buena = checkbox(data.get("partner_relationship") == "buena")
    regular = checkbox(data.get("partner_relationship") == "regular")
    mala = checkbox(data.get("partner_relationship") == "mala")
    table_row_mixed(pdf, [90,100],[
        "Como califica la relación de pareja:",
        f"{muybuenarelacionpareja} Muy buena     {buena} Buena     {regular} Regular     {mala} Mala"
    ],["B",""])
    table_row_mixed(pdf, [40,150],[
        "¿Por qué?", data.get("partner_relationship_reason", "")
    ],["B",""])

    muybuenarelacionhijos = checkbox(data.get("children_relationship") == "muy buena")
    buena = checkbox(data.get("children_relationship") == "buena")
    regular = checkbox(data.get("children_relationship") == "regular")
    mala = checkbox(data.get("children_relationship") == "mala")
    table_row_mixed(pdf, [90,100],[
        "Como califica la relación con los hijos:",
        f"{muybuenarelacionhijos} Muy buena     {buena} Buena     {regular} Regular     {mala} Mala"
    ],["B",""])
    table_row_mixed(pdf, [40,150],[
        "¿Por qué?", data.get("children_relationship_reason", "")
    ],["B",""])
    divorcios = checkbox(data.get("family_problems") == "divorcios")
    separaciones = checkbox(data.get("family_problems") == "separaciones")
    resentimientos = checkbox(data.get("family_problems") == "resentimientos")
    otros = checkbox(data.get("family_problems") == "otros")
    ningunos = checkbox(data.get("family_problems") == "ningunos")
    table_row_mixed(pdf, [40,150],[
        "Problemas familiares:",
        f"{divorcios} Divorcios     {separaciones} Separaciones     {resentimientos} Resentimientos     {otros} Otros     {ningunos} Ninguno"
    ],["B",""])
    table_row_mixed(pdf, [40,150],[
        "Recibio ayuda", data.get("family_problems_help", "")
    ],["B",""])
    sisalieronpais = checkbox(data.get("family_migration") == "si")
    nosalieronpais = checkbox(data.get("family_migration") == "no")
    table_row_mixed(pdf, [90,100],[
        "Miembros del hogar que salieron del país:", f"{sisalieronpais} Sí     {nosalieronpais} No"
    ],["B",""])

    sirecibedinero = checkbox(data.get("family_migration_received", "") == "si")
    nosirecibedinero = checkbox(data.get("family_migration_received", "") == "no")
    table_row_mixed(pdf, [90,100],[
        "Recibió dinero del exterior:", f"{sirecibedinero} Sí     {nosirecibedinero} No"
    ],["B",""])

def vivienda(pdf, data):
    
    section_title(pdf, "3. VIVIENDA")

    urbano = checkbox(data.get("sector") == "urbano")
    rural = checkbox(data.get("sector") == "rural")

    table_row_mixed(pdf, [50, 140], [
        "Sector",
        f"{urbano} Urbano     {rural} Rural"
    ],["B",""])
    propia = checkbox(data.get("house_ownership") == "propia")
    arrendada = checkbox(data.get("house_ownership") == "arrendada")
    hipotecada = checkbox(data.get("house_ownership") == "hipotecada")
    prestada = checkbox(data.get("house_ownership") == "prestada")
    otros = checkbox(data.get("house_ownership") == "otros")
    table_row(pdf,[190],[
        f"{propia} Propia     {arrendada} Arrendada     {hipotecada} Hipotecada     {prestada} Prestada     {otros} Otros"
    ])
    table_row_mixed(pdf, [30,20,50,30,40,20], [
        "N° habitantes", data.get("household_size", ""),
        "Que tiempo vive en el sector", data.get("time_living_sector", ""),
        "Se considera seguro", data.get("is_safe", "")
    ],["B","","B","","B",""])

    casa = checkbox(data.get("house_type") == "casa")
    departamento = checkbox(data.get("house_type") == "departamento")
    cuarto = checkbox(data.get("house_type") == "cuarto")
    mediagua = checkbox(data.get("house_type") == "mediagua")
    table_row_mixed(pdf,[50, 140],[
        "Tipo de vivienda",
        f"{casa} Casa     {departamento} Departamento     {cuarto} Cuarto     {mediagua} Mediagua"
    ],["B",""])
    precaria = checkbox(data.get("house_class") == "precaria")
    elemental = checkbox(data.get("house_class") == "elemental")
    completa = checkbox(data.get("house_class") == "completa")
    table_row_mixed(pdf,[50, 140],[
        "Clase de vivienda",
        f"{precaria} Precaria     {elemental} Elemental     {completa} Completa"
    ],["B",""])
    table_row_mixed(pdf,[50, 140],[
        "Avaluo de casa:",
        data.get("house_valuation", "")
    ],["B",""])
    posesion = checkbox(data.get("house_possession") == "si")
    no_posesion = checkbox(data.get("house_possession") == "no")
    table_row_mixed(pdf,[50,30,20,20,30,40],[
        "¿Posee la vivienda?",
        f"{posesion} Sí     {no_posesion} No",
        "Avaluo:", data.get("house_valuation", ""),
        "Observaciones:", data.get("house_observations", "")
    ],["B","","B","","B",""])
    table_row_mixed(pdf,[190],[
        "Distribucion de la vivienda"
    ],["B"])
    dormitorio = checkbox(data.get("house_distribution") == "dormitorio")
    camas = checkbox(data.get("house_distribution") == "camas")
    cocina = checkbox(data.get("house_distribution") == "cocina")
    comedor = checkbox(data.get("house_distribution") == "comedor")
    sala = checkbox(data.get("house_distribution") == "sala")
    bano = checkbox(data.get("house_distribution") == "bano")
    patio = checkbox(data.get("house_distribution") == "patio")
    jardin = checkbox(data.get("house_distribution") == "jardin")
    terraza = checkbox(data.get("house_distribution") == "terraza")
    garaje = checkbox(data.get("house_distribution") == "garaje")
    table_row_multiline(pdf, [190], [
        f"{dormitorio} Dormitorio     {camas} Camas     {cocina} Cocina     "
        f"{comedor} Comedor     {sala} Sala     \n{bano} Baño     {patio} Patio     " 
        f"{jardin} Jardín     {terraza} Terraza     {garaje} Garaje"
    ],[""])
    table_row_mixed(pdf, [190], [
        "Techo o cubierta"
    ],["B"])
    loza = checkbox(data.get("roof_type") == "loza")
    eternit = checkbox(data.get("roof_type") == "eternit")
    zinc = checkbox(data.get("roof_type") == "zinc")
    teja = checkbox(data.get("roof_type") == "teja")
    plastico = checkbox(data.get("roof_type") == "plastico")
    bueno_techo = checkbox(data.get("roof_status") == "bueno")
    regular_techo = checkbox(data.get("roof_status") == "regular")
    malo_techo = checkbox(data.get("roof_status") == "malo")
    table_row_multiline(pdf, [95,95], [
        f"{loza} Loza     \n{eternit} Eternit     \n{zinc} Zinc     \n{teja} Teja     \n{plastico} Plástico     \nOtros:{data.get("roof_type_other", "")}", 
        f"Estado:   \n\n{bueno_techo} Bueno     \n{regular_techo} Regular     \n{malo_techo} Malo       \n      "
    ],["", ""])

def servicios_basicos(pdf, data):

    section_title(pdf, "4. SERVICIOS BÁSICOS")

    agua = checkbox(data.get("water"))
    luz = checkbox(data.get("electricity"))
    internet = checkbox(data.get("internet"))

    table_row(pdf, [190], [
        f"{agua} Agua potable      {luz} Energía eléctrica      {internet} Internet"
    ])


def economia(pdf, data):

    section_title(pdf, "5. SITUACIÓN ECONÓMICA")

    shared = checkbox(data.get("shared_expenses"))

    table_row(pdf, [50, 140], [
        "Gastos compartidos",
        f"{shared} Sí"
    ])

    pdf.cell(190, 6, "¿Quiénes aportan?", border=1, ln=True)
    box(pdf, data.get("contributors", ""))

    pdf.cell(190, 6, "Gastos principales", border=1, ln=True)
    box(pdf, data.get("expenses", ""))

    pdf.cell(190, 6, "Deudas", border=1, ln=True)
    box(pdf, data.get("debts", ""))


def salud(pdf, data):

    section_title(pdf, "6. SALUD")

    sport = checkbox(data.get("sports"))

    table_row(pdf, [50, 140], [
        "Actividad física",
        f"{sport} Sí"
    ])

    pdf.cell(190, 6, "Enfermedades", border=1, ln=True)
    box(pdf, data.get("disease", ""))


def laboral(pdf, data):

    section_title(pdf, "7. SITUACIÓN LABORAL")

    pdf.cell(190, 6, "Funciones actuales", border=1, ln=True)
    box(pdf, data.get("functions", ""), 25)

    pdf.cell(190, 6, "Plan de vida", border=1, ln=True)
    box(pdf, data.get("life_plan", ""), 25)


def observaciones(pdf, data):

    section_title(pdf, "OBSERVACIONES")
    box(pdf, data.get("observations", ""), 30)


def firmas(pdf):

    check_page_space(pdf, 50)

    section_title(pdf, "FIRMAS")

    pdf.ln(10)

    pdf.cell(90, 8, "_____________________________", 0, 0, "C")
    pdf.cell(10)
    pdf.cell(90, 8, "_____________________________", 0, 1, "C")

    pdf.cell(90, 6, "Firma del entrevistado", 0, 0, "C")
    pdf.cell(10)
    pdf.cell(90, 6, "Firma del trabajador social", 0, 1, "C")

    pdf.cell(90, 6, "Nombre", 0, 0, "C")
    pdf.cell(10)
    pdf.cell(90, 6, "Nombre ", 0, 1, "C")


# -----------------------------
# MAIN
# -----------------------------

def create_pdf(data):

    pdf = PDF()
    
    photo = data.get("photo")
    if photo:
        pdf.photo_path = fix_image_orientation(photo)

    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

    fecha_visita(pdf, data)
    pdf.ln(3)

    datos_personales(pdf, data)
    pdf.ln(3)

    miembros_hogar(pdf, data)
    pdf.ln(3)

    info_familia(pdf, data)
    pdf.ln(3)

    vivienda(pdf, data)
    pdf.ln(3)

    servicios_basicos(pdf, data)
    pdf.ln(3)

    economia(pdf, data)
    pdf.ln(3)

    salud(pdf, data)
    pdf.ln(3)

    laboral(pdf, data)
    pdf.ln(3)

    observaciones(pdf, data)
    pdf.ln(5)

    firmas(pdf)

    os.makedirs("pdf", exist_ok=True)

    filename = f"pdf/{data.get('ci','sin_ci')}.pdf"
    pdf.output(filename)

    return filename