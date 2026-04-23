from fpdf import FPDF
import os
from datetime import datetime
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

# ========== TÍTULO DE SECCIÓN CON FONDO CELESTE ==========
def section_title(pdf, title):
    """Título de sección con fondo celeste y texto azul oscuro"""
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(200, 220, 255)
    pdf.set_text_color(0, 0, 150)
    pdf.cell(PAGE_WIDTH, 8, title, border=1, ln=True, fill=True)
    pdf.set_text_color(0, 0, 0)

def table_row(pdf, widths, data, height=8):
    if len(data) < len(widths):
        data = list(data) + [""] * (len(widths) - len(data))
    for i in range(len(widths)):
        pdf.cell(widths[i], height, str(data[i]), border=1)
    pdf.ln(height)
    
def table_row_mixed(pdf, widths, data, styles, height=8):
    if len(data) < len(widths):
        data = list(data) + [""] * (len(widths) - len(data))
    if len(styles) < len(widths):
        styles = list(styles) + [""] * (len(widths) - len(styles))
    for i in range(len(widths)):
        pdf.set_font("Arial", styles[i], 10)
        pdf.cell(widths[i], height, str(data[i]), border=1)
    pdf.ln(height)
    pdf.set_font("Arial", "", 10)

# ========== FUNCIÓN MULTILÍNEA CORREGIDA ==========
def table_row_multiline(pdf, widths, data, styles, line_height=5):
    """Fila que maneja texto largo - TODAS las celdas tienen la MISMA altura"""
    if len(data) < len(widths):
        data = list(data) + [""] * (len(widths) - len(data))
    if len(styles) < len(widths):
        styles = list(styles) + [""] * (len(widths) - len(styles))

    # Calcular número de líneas REAL por celda
    line_counts = []
    pdf.set_font("Arial", "", 8)

    for i in range(len(data)):
        text = str(data[i])
        col_width = widths[i]
        
        if not text.strip():
            line_counts.append(1)
            continue

        lines = text.split("\n")
        total_lines = 0
        for line in lines:
            if line.strip():
                text_width = pdf.get_string_width(line)
                lines_needed = max(1, int(text_width / (col_width - 3)) + 1)
                total_lines += lines_needed
            else:
                total_lines += 1
        line_counts.append(max(1, total_lines))

    # Altura MÁXIMA de la fila
    max_lines = max(line_counts)
    row_height = max_lines * line_height

    if pdf.get_y() + row_height > pdf.page_break_trigger:
        pdf.add_page()

    x_start = pdf.get_x()
    y_start = pdf.get_y()

    # Dibujar cada celda y luego rellenar hasta la altura máxima
    for i in range(len(widths)):
        pdf.set_xy(x_start, y_start)
        pdf.set_font("Arial", styles[i], 8)
        
        # Guardar posición X actual
        x_current = pdf.get_x()
        
        # Dibujar el contenido de la celda SIN borde
        pdf.multi_cell(widths[i], line_height, str(data[i]), border=0)
        
        # Dibujar el borde completo de la celda con la altura máxima
        pdf.rect(x_current, y_start, widths[i], row_height)
        
        # Avanzar X para la siguiente celda
        x_start += widths[i]
        # Volver a la posición Y inicial
        pdf.set_y(y_start)

    # Mover al final de la fila
    pdf.set_y(y_start + row_height)
    pdf.set_x(pdf.l_margin)
    pdf.set_font("Arial", "", 10)

def check_page_space(pdf, space_needed):
    if pdf.get_y() + space_needed > 270:
        pdf.add_page()

def box(pdf, text, h=20):
    """Función box corregida - asegura que h sea número"""
    if isinstance(h, str):
        try:
            h = float(h)
        except:
            h = 20
    pdf.set_font("Arial", "", 8)
    pdf.multi_cell(PAGE_WIDTH, h / 4, str(text), border=1)
    pdf.set_font("Arial", "", 10)


# -----------------------------
# SECCIONES
# -----------------------------

def fecha_visita(pdf, data):
    table_row_multiline(pdf, [40, 150], [
        "Fecha de Visita", data.get("visit_date", "")
    ], ["B", ""])

def datos_personales(pdf, data):
    section_title(pdf, "1. DATOS PERSONALES DEL COLABORADOR")

    table_row_multiline(pdf, [40, 150], [
        "Nombre", data.get("name", "")
    ], ["B", ""])

    table_row_multiline(pdf, [40, 55, 40, 55], [
        "Cédula", data.get("ci", ""),
        "Edad", data.get("age", "")
    ], ["B", "", "B", ""])
    
    table_row_multiline(pdf, [80, 110], [
        "Lugar y Fecha de Nacimiento", data.get("birth_date_place", "")
    ], ["B", ""])
    
    masculino = checkbox(data.get("gender") == "masculino")
    femenino = checkbox(data.get("gender") == "femenino")
    otro = checkbox(data.get("gender") == "otro")   
    table_row_multiline(pdf, [40, 150], [
        "Genero", f"{masculino} Masculino     {femenino} Femenino     {otro} Otro"
    ], ["B", ""])

    casado = checkbox(data.get("marital_status") == "casado")
    soltero = checkbox(data.get("marital_status") == "soltero")
    separado = checkbox(data.get("marital_status") == "separado")
    divorciado = checkbox(data.get("marital_status") == "divorciado")
    unionlibre = checkbox(data.get("marital_status") == "unión libre")
    viudo = checkbox(data.get("marital_status") == "viudo")
    table_row_multiline(pdf, [40, 150], [
        "Estado civil",
        f"{casado} Casado     {soltero} Soltero     {separado} Separado     {divorciado} Divorciado     {unionlibre} Unión Libre     {viudo} Viudo"
    ], ["B", ""])

    table_row_multiline(pdf, [40, 150], [
        "Email", data.get("email", "")
    ], ["B", ""])
    
    table_row_multiline(pdf, [40, 50, 50, 50], [
        "Área de trabajo", data.get("work_area", ""), 
        "Puesto que desempeña", data.get("job_title", "")
    ], ["B", "", "B", ""])
    
    table_row_multiline(pdf, [55, 135], [
        "Fecha de Ingreso", data.get("entry_date", "")
    ], ["B", ""])
    
    table_row_multiline(pdf, [55, 15, 25, 35, 40, 20], [
        "Tiene alguna discapacidad", data.get("disability", ""),
        "Tipo", data.get("disability_type", ""),
        "Porcentaje", data.get("disability_percentage", "")
    ], ["B", "", "B", "", "B", ""])

    profesional = checkbox(data.get("educational_level") == "profesional")
    bachiller = checkbox(data.get("educational_level") == "bachiller")
    basica = checkbox(data.get("educational_level") == "basica")
    primaria = checkbox(data.get("educational_level") == "primaria")
    nuloestudio = checkbox(data.get("educational_level") == "nuloestudio")
    table_row_multiline(pdf, [40, 150], [
        "Nivel de Educación", 
        f"{profesional} Profesional     {bachiller} Bachiller     {basica} Básica     {primaria} Primaria     {nuloestudio} Sin Estudios"
    ], ["B", ""])
    
    table_row_multiline(pdf, [40, 150], [
        "Observaciones:", data.get("educational_observations", "")
    ], ["B", ""])
    
    table_row_multiline(pdf, [50, 15, 30, 35, 30, 30], [
        "¿Actualmente estudia?", data.get("currently_studying", ""),
        "Carrera", data.get("study_career", ""),
        "Nivel", data.get("study_level", "")
    ], ["B", "", "B", "", "B", ""])

    table_row_multiline(pdf, [40, 150], [
        "Dirección Domiciliaria", data.get("address", "")
    ], ["B", ""])
    
    table_row_multiline(pdf, [40, 150], [
        "Teléfono", data.get("phone", "")
    ], ["B", ""])

    table_row_multiline(pdf, [60, 130], [
        "Contacto de emergencia", data.get("emergency_contact", "")
    ], ["B", ""])

def miembros_hogar(pdf, data):
    section_title(pdf, "2. MIEMBROS DEL HOGAR")
    table_row_multiline(pdf, [190], [
        "Célula Social - Composición Familiar del Colaborador (viven dentro del hogar)"
    ], ["B"])

    headers = ["N°", "Nombre", "Edad", "Parentesco", "Ocupación", "Ingreso"]
    widths = [10, 50, 20, 40, 40, 30]

    table_row_multiline(pdf, widths, headers, ["B", "B", "B", "B", "B", "B"])

    miembros = data.get("family_members", [])

    if not miembros:
        table_row(pdf, [190], ["No se registran miembros del hogar"])
        return

    for i, m in enumerate(miembros, start=1):
        ingreso = m.get("income", "")
        if ingreso:
            try:
                ingreso = f"${float(ingreso):,.2f}"
            except:
                ingreso = str(ingreso)
        
        table_row_multiline(pdf, widths, [
            str(i),
            str(m.get("name", "")),
            str(m.get("age", "")),
            str(m.get("relation", "")),
            str(m.get("job", "")),
            ingreso
        ], ["", "", "", "", "", ""])

def info_familia(pdf, data):
    section_title(pdf, "3. INFORMACIÓN DE LA FAMILIA")
    
    table_row_multiline(pdf, [190], [
        "Información de la Familia:"
    ], ["B"])

    nuclear = checkbox(data.get("family_type") == "nuclear")
    cosanguinea = checkbox(data.get("family_type") == "cosanguinea")
    afines = checkbox(data.get("family_type") == "afines")
    ampliada = checkbox(data.get("family_type") == "ampliada")
    materna = checkbox(data.get("family_type") == "materna")
    paterna = checkbox(data.get("family_type") == "paterna")
    unipersonal = checkbox(data.get("family_type") == "unipersonal")
    
    table_row_multiline(pdf, [40, 150], [
        "Tipo de Familia:",
        f"{nuclear} Nuclear     {cosanguinea} Extensa Cosanguínea     {afines} Extensa Afines\n{ampliada} Ampliada     {materna} Monoparental Materna     {paterna} Monoparental Paterna     {unipersonal} Unipersonal"
    ], ["B", ""])

    siotroshijos = checkbox(data.get("children_outside_marriage") == "si")
    nootroshijos = checkbox(data.get("children_outside_marriage") == "no")
    table_row_multiline(pdf, [60, 30, 60, 40], [
        "N° Hijos del colaborador", data.get("children_count", ""),
        "Existen hijos fuera del Matrimonio", f"{siotroshijos} Sí     {nootroshijos} No"
    ], ["B", "", "B", ""])
    
    table_row_multiline(pdf, [60, 130], [
        "Paga pensión alimenticia", data.get("pays_alimony", "")
    ], ["B", ""])
    
    table_row_multiline(pdf, [70, 120], [
        "N° de Matrimonio o Relación Actual", data.get("marriage_number", "")
    ], ["B", ""])
    
    table_row_multiline(pdf, [110, 80], [
        "Comparte con su familia tiempo de descanso y recreación", data.get("family_time", "")
    ], ["B", ""])
    
    table_row_multiline(pdf, [40, 150], [
        "Frecuencia", data.get("family_time_frequency", "")
    ], ["B", ""])
    
    table_row_multiline(pdf, [40, 150], [
        "Actividades:", data.get("family_activities", "")
    ], ["B", ""])
    
    table_row_multiline(pdf, [80, 110], [
        "Qué actividades realiza\nfuera de su jornada laboral:", data.get("other_activities", " ")
    ], ["B", ""])
    
    otrasactividades = checkbox(data.get("other_activities") == "si")
    nootrasactividades = checkbox(data.get("other_activities") == "no")
    table_row_multiline(pdf, [75, 25, 40, 50], [
        "En su domicilio realiza otras actividades:",
        f"{otrasactividades} Sí     {nootrasactividades} No",
        "Especificar:", data.get("other_activities_specify", "")
    ], ["B", "", "B", ""])
    
    table_row_multiline(pdf, [75, 115], [
        "Describa cuales son sus hobbies:", data.get("hobbies", "")
    ], ["B", ""])
    
    table_row_multiline(pdf, [75, 115], [
        "Que tiempo le dedica a sus hobbies:", data.get("hobbies_time", "")
    ], ["B", ""])

    muybuenarelacionpareja = checkbox(data.get("partner_relationship") == "muy buena")
    buena = checkbox(data.get("partner_relationship") == "buena")
    regular = checkbox(data.get("partner_relationship") == "regular")
    mala = checkbox(data.get("partner_relationship") == "mala")
    table_row_multiline(pdf, [90, 100], [
        "Como califica la relación de pareja:",
        f"{muybuenarelacionpareja} Muy buena     {buena} Buena     {regular} Regular     {mala} Mala"
    ], ["B", ""])
    
    table_row_multiline(pdf, [40, 150], [
        "¿Por qué?", data.get("partner_relationship_reason", "")
    ], ["B", ""])

    muybuenarelacionhijos = checkbox(data.get("children_relationship") == "muy buena")
    buena = checkbox(data.get("children_relationship") == "buena")
    regular = checkbox(data.get("children_relationship") == "regular")
    mala = checkbox(data.get("children_relationship") == "mala")
    table_row_multiline(pdf, [90, 100], [
        "Como califica la relación con los hijos:",
        f"{muybuenarelacionhijos} Muy buena     {buena} Buena     {regular} Regular     {mala} Mala"
    ], ["B", ""])
    
    table_row_multiline(pdf, [40, 150], [
        "¿Por qué?", data.get("children_relationship_reason", "")
    ], ["B", ""])
    
    divorcios = checkbox(data.get("family_problems") == "divorcios")
    separaciones = checkbox(data.get("family_problems") == "separaciones")
    resentimientos = checkbox(data.get("family_problems") == "resentimientos")
    otros = checkbox(data.get("family_problems") == "otros")
    ningunos = checkbox(data.get("family_problems") == "ningunos")
    table_row_multiline(pdf, [40, 150], [
        "Problemas familiares:",
        f"{divorcios} Divorcios     {separaciones} Separaciones     {resentimientos} Resentimientos     {otros} Otros     {ningunos} Ninguno"
    ], ["B", ""])
    
    table_row_multiline(pdf, [40, 150], [
        "Recibio ayuda", data.get("family_problems_help", "")
    ], ["B", ""])
    
    sisalieronpais = checkbox(data.get("family_migration") == "si")
    nosalieronpais = checkbox(data.get("family_migration") == "no")
    table_row_multiline(pdf, [90, 100], [
        "Miembros del hogar que salieron del país:", f"{sisalieronpais} Sí     {nosalieronpais} No"
    ], ["B", ""])

    sirecibedinero = checkbox(data.get("family_migration_received", "") == "si")
    nosirecibedinero = checkbox(data.get("family_migration_received", "") == "no")
    table_row_multiline(pdf, [90, 100], [
        "Recibió dinero del exterior:", f"{sirecibedinero} Sí     {nosirecibedinero} No"
    ], ["B", ""])

def vivienda(pdf, data):
    section_title(pdf, "4. VIVIENDA")

    urbano = checkbox(data.get("sector") == "urbano")
    rural = checkbox(data.get("sector") == "rural")
    table_row_multiline(pdf, [50, 140], [
        "Sector",
        f"{urbano} Urbano     {rural} Rural"
    ], ["B", ""])
    
    propia = checkbox(data.get("house_ownership") == "propia")
    arrendada = checkbox(data.get("house_ownership") == "arrendada")
    hipotecada = checkbox(data.get("house_ownership") == "hipotecada")
    prestada = checkbox(data.get("house_ownership") == "prestada")
    otros = checkbox(data.get("house_ownership") == "otros")
    table_row_multiline(pdf, [190], [
        f"{propia} Propia     {arrendada} Arrendada     {hipotecada} Hipotecada     {prestada} Prestada     {otros} Otros"
    ], [""])
    
    table_row_multiline(pdf, [30, 20, 50, 30, 40, 20], [
        "N° habitantes", data.get("household_size", ""),
        "Que tiempo vive en el sector", data.get("time_living_sector", ""),
        "Se considera seguro", data.get("is_safe", "")
    ], ["B", "", "B", "", "B", ""])

    casa = checkbox(data.get("house_type") == "casa")
    departamento = checkbox(data.get("house_type") == "departamento")
    cuarto = checkbox(data.get("house_type") == "cuarto")
    mediagua = checkbox(data.get("house_type") == "mediagua")
    table_row_multiline(pdf, [50, 140], [
        "Tipo de vivienda",
        f"{casa} Casa     {departamento} Departamento     {cuarto} Cuarto     {mediagua} Mediagua"
    ], ["B", ""])
    
    precaria = checkbox(data.get("house_class") == "precaria")
    elemental = checkbox(data.get("house_class") == "elemental")
    completa = checkbox(data.get("house_class") == "completa")
    table_row_multiline(pdf, [50, 140], [
        "Clase de vivienda",
        f"{precaria} Precaria     {elemental} Elemental     {completa} Completa"
    ], ["B", ""])
    
    table_row_multiline(pdf, [50, 140], [
        "Avaluo de casa:",
        data.get("house_valuation", "")
    ], ["B", ""])
    
    posesion = checkbox(data.get("house_possession") == "si")
    no_posesion = checkbox(data.get("house_possession") == "no")
    table_row_multiline(pdf, [50, 30, 20, 20, 30, 40], [
        "¿Posee la vivienda?",
        f"{posesion} Sí     {no_posesion} No",
        "Avaluo:", 
        data.get("house_valuation", ""),
        "Observaciones:", 
        data.get("house_observations", "")
    ], ["B", "", "B", "", "B", ""])
    
    table_row_multiline(pdf, [190], [
        "Distribucion de la vivienda"
    ], ["B"])
    
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
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Techo o cubierta"
    ], ["B"])
    
    loza = checkbox(data.get("roof_type") == "loza")
    eternit = checkbox(data.get("roof_type") == "eternit")
    zinc = checkbox(data.get("roof_type") == "zinc")
    teja = checkbox(data.get("roof_type") == "teja")
    plastico = checkbox(data.get("roof_type") == "plastico")
    bueno_techo = checkbox(data.get("roof_status") == "bueno")
    regular_techo = checkbox(data.get("roof_status") == "regular")
    malo_techo = checkbox(data.get("roof_status") == "malo")
    
    table_row_multiline(pdf, [95, 95], [
        f"{loza} Loza     \n{eternit} Eternit     \n{zinc} Zinc     \n{teja} Teja     \n{plastico} Plástico     \nOtros:{data.get("roof_type_other", "")}", 
        f"Estado:   \n\n{bueno_techo} Bueno     \n{regular_techo} Regular     \n{malo_techo} Malo       \n      "
    ], ["", ""])

    hormigon = checkbox(data.get("wall_type") == "hormigon")
    bloque = checkbox(data.get("wall_type") == "bloque")
    adobe = checkbox(data.get("wall_type") == "adobe")
    madera = checkbox(data.get("wall_type") == "madera")
    bueno_pared = checkbox(data.get("wall_status") == "bueno")
    regular_pared = checkbox(data.get("wall_status") == "regular")
    malo_pared = checkbox(data.get("wall_status") == "malo")
    
    table_row_multiline(pdf, [190], [
        "Tipo de pared"
    ], ["B"])
    
    table_row_multiline(pdf, [95, 95], [
        f"{hormigon} Hormigón     \n{bloque} Bloque     \n{adobe} Adobe     \n{madera} Madera     \nOtros:{data.get("wall_type_other", "")}",
        f"Estado:   \n{bueno_pared} Bueno     \n{regular_pared} Regular     \n{malo_pared} Malo       \n      "
    ], ["", ""])
    
    table_row_multiline(pdf, [190], [
        "Tipo de piso"
    ], ["B"])
    
    entablado = checkbox(data.get("floor_type") == "entablado")
    baldosa = checkbox(data.get("floor_type") == "baldosa")
    cemento = checkbox(data.get("floor_type") == "cemento")
    tierra = checkbox(data.get("floor_type") == "tierra")
    bueno_piso = checkbox(data.get("floor_status") == "bueno")
    regular_piso = checkbox(data.get("floor_status") == "regular")
    malo_piso = checkbox(data.get("floor_status") == "malo")

    table_row_multiline(pdf, [95, 95], [
        f"{entablado} Entablado     \n{baldosa} Baldosa     \n{cemento} Cemento     \n{tierra} Tierra     \nOtros:{data.get("floor_type_other", "")}",
        f"Estado:   \n{bueno_piso} Bueno     \n{regular_piso} Regular     \n{malo_piso} Malo       \n      "
    ], ["", ""])

    table_row_multiline(pdf, [190], [
        "Tipo de estructura"
    ], ["B"])
    
    hormigon_est = checkbox(data.get("structure_type") == "hormigon")
    hierro = checkbox(data.get("structure_type") == "hierro")
    madera_est = checkbox(data.get("structure_type") == "madera")
    bloque_est = checkbox(data.get("structure_type") == "bloque")
    bueno_estructura = checkbox(data.get("structure_status") == "bueno")
    regular_estructura = checkbox(data.get("structure_status") == "regular")
    malo_estructura = checkbox(data.get("structure_status") == "malo")

    table_row_multiline(pdf, [95, 95], [
        f"{hormigon_est} Hormigón     \n{hierro} Hierro     \n{madera_est} Madera     \n{bloque_est} Bloque     \nOtros:{data.get("structure_type_other", "")}",
        f"Estado:   \n{bueno_estructura} Bueno     \n{regular_estructura} Regular     \n{malo_estructura} Malo       \n      "
    ], ["", ""])
    
    table_row_multiline(pdf, [190], [
        "Servicios básicos"
    ], ["B"])
    
    # Agua
    potable = checkbox(data.get("water_type") == "potable")
    cisterna = checkbox(data.get("water_type") == "cisterna")
    vertiente = checkbox(data.get("water_type") == "vertiente")
    repartidor = checkbox(data.get("water_type") == "repartidor")
    sequia = checkbox(data.get("water_type") == "sequia")
    table_row_multiline(pdf, [30, 160], [
        "Agua",
        f"{potable} Potable     {cisterna} Cisterna     {vertiente} Vertiente     {repartidor} Repartidor     {sequia} Sequía"
    ], ["B", ""])
    
    # Luz
    permanente = checkbox(data.get("light_type") == "permanente")
    temporal = checkbox(data.get("light_type") == "temporal")
    noenergia = checkbox(data.get("light_type") == "noenergia")
    table_row_multiline(pdf, [30, 160], [
        "Luz",
        f"{permanente} Permanente     {temporal} Temporal     {noenergia} No Energía"
    ], ["B", ""])
    
    # SSHH
    sshhpropio = checkbox(data.get("sshh_type") == "propio")
    sshhcompartido = checkbox(data.get("sshh_type") == "compartido")
    sshhpozo = checkbox(data.get("sshh_type") == "pozo")
    sshhlibre = checkbox(data.get("sshh_type") == "libre")
    table_row_multiline(pdf, [30, 160], [
        "SSHH",
        f"{sshhpropio} Propio     {sshhcompartido} Compartido     {sshhpozo} Pozo     {sshhlibre} Libre",
    ], ["B", ""])
    
    table_row_multiline(pdf, [30, 160], [
        "Observaciones:", data.get("basic_services_other", "")
    ], ["B", ""])
    
    sihacinamiento = checkbox(data.get("hacinamiento") == "si")
    no_hacinamiento = checkbox(data.get("hacinamiento") == "no")
    table_row_multiline(pdf, [30, 160], [
        "Hacinamiento",
        f"{sihacinamiento} Sí     {no_hacinamiento} No"
    ], ["B", ""])
    
    table_row_multiline(pdf, [50, 140], [
        "Manejo de desechos", data.get("waste_management", "")
    ], ["B", ""])
    
    transportepublico = checkbox(data.get("transport_type") == "publico")
    transporteempresa = checkbox(data.get("transport_type") == "empresa")
    transporteprivado = checkbox(data.get("transport_type") == "privado")
    table_row_multiline(pdf, [30, 160], [
        "Transporte",
        f"{transportepublico} Público     {transporteempresa} Empresa     {transporteprivado} Privado     Otro: {data.get("transport_type_other", "")}"
    ], ["B", ""])
    
    table_row_multiline(pdf, [60, 130], [
        "Que electrodomesticos posee:", data.get("appliances", "")
    ], ["B", ""])
    
    si_internet = checkbox(data.get("internet") == "si")
    no_internet = checkbox(data.get("internet") == "no")
    table_row_multiline(pdf, [60, 130], [
        "Dispone de servicio de Internet",
        f"{si_internet} Sí     {no_internet} No"
    ], ["B", ""])
    
    recarga = checkbox(data.get("internet_type") == "recarga")
    mensual = checkbox(data.get("internet_type") == "mensual")
    satelital = checkbox(data.get("internet_type") == "satelital")
    fibraoptica = checkbox(data.get("internet_type") == "fibraoptica")
    table_row_multiline(pdf, [60, 130], [
        "Tipo de Internet",
        f"{recarga} Recarga     {mensual} Mensual     {satelital} Satelital     {fibraoptica} Fibra Óptica"
    ], ["B", ""])
    
    sianimales = checkbox(data.get("animals") == "si")
    no_sianimales = checkbox(data.get("animals") == "no")
    table_row_multiline(pdf, [40, 30, 30, 40, 30, 20], [
        "Tiene animales",
        f"{sianimales} Sí     {no_sianimales} No",
        "Tipo",
        f"{data.get('animal_type', '')}",
        "Cantidad",
        data.get('animal_quantity', '')
    ], ["B", "", "", "", "", ""])
    
    si_peste = checkbox(data.get("peste") == "si")
    no_peste = checkbox(data.get("peste") == "no")
    table_row_multiline(pdf, [40, 30, 75, 45], [
        "Zona de peste",
        f"{si_peste} Sí     {no_peste} No",
        "Lugar de tenencia (dentro o fuera del hogar)",
        f"{data.get('animal_location', '')}"
    ], ["B", "", "B", ""])
    
    table_row_multiline(pdf, [50, 140], [
        "Observaciones:", data.get('animal_observations', '')
    ], ["B", ""])

def economia(pdf, data):
    section_title(pdf, "5. SITUACIÓN ECONÓMICA")

    gastos_compartidos = checkbox(data.get("shared_expenses") == "si")
    gastos_no_compartidos = checkbox(data.get("shared_expenses") == "no")
    table_row_multiline(pdf, [100, 90], [
        "Los gastos en el hogar son compartidos",
        f"{gastos_compartidos} Sí     {gastos_no_compartidos} No"
    ], ["B", ""])
    
    table_row_multiline(pdf, [190], [
        "Personas que aportan economicamente"
    ], ["B"])
    
    table_row_multiline(pdf, [95, 95], [
        f"Padre: \nMadre: \nHermanos: \nColaborador: \nConyuge: \nHijos: \nOtros:",
        f"${data.get('father_contribution', '')} \n${data.get('mother_contribution', '')} \n${data.get('siblings_contribution', '')} \n${data.get('collaborators_contribution', '')} \n${data.get('spouse_contribution', '')} \n${data.get('children_contribution', '')} \n${data.get('other_contribution', '')}"   
    ], ["B", ""])
    
    table_row_multiline(pdf, [95, 95], [
        "TOTAL:", f"${data.get('total_contribution', '')}"
    ], ["B", ""])
    
    table_row_multiline(pdf, [95, 95], [
        "Monto de deudas:", f"${data.get('debt_amount', '')}"
    ], ["B", ""])
    
    prestamos_formales = checkbox(data.get("formal_loans"))
    table_row_multiline(pdf, [95, 95], [
        f"Prestamos: {prestamos_formales} Formales",
        f"${data.get('formal_loans_amount', '')}"
    ], ["B", ""])
    
    prestamos_informales = checkbox(data.get("informal_loans"))
    table_row_multiline(pdf, [95, 95], [
        f"Prestamos: {prestamos_informales} Informales"
        f"\nFamiliares:"
        f"\nChulqueros"
        f"\n\nOtros:",
        f"${data.get('informal_loans_amount', '')}"
        f"\n${data.get('informal_loans_family_amount', '')}"
        f"\n${data.get('informal_loans_moneylender_amount', '')}"
        f"\n\n${data.get('informal_loans_other_amount', '')}"
    ], ["B", ""])
    
    sitarjetas = checkbox(data.get("credit_cards") == 'si')
    notarjetas = checkbox(data.get("credit_cards") == 'no')
    table_row_multiline(pdf, [95, 95], [
        "Posee tarjetas de crédito?",
        f"{sitarjetas} Sí     {notarjetas} No"
    ], ["B", ""])
    
    table_row_multiline(pdf, [190], [
        "Gastos"
    ], ["B"])
    
    table_row_multiline(pdf, [95, 95], [
        f"Alimentación:\n"
        f"Educación:\n"
        f"Vivienda:\n"
        f"Vestimenta:\n"
        f"Salud:\n"
        f"Transporte:\n"
        f"Servicios básicos:\n"
        f"Internet:\n"
        f"Tv Cable:\n"
        f"Plan Celular:\n"
        f"Prestamos:\n"
        f"Prestamos Quirografarios:\n"
        f"Tarjetas de crédito:\n"
        f"Pension de Alimentos:\n"
        f"Locales Comerciales:\n"
        f"Apoyo economico a terceras personas:\n"
        f"Otros Gastos:\n",
        f"${data.get("food_support", "")}\n"
        f"${data.get("education_support", "")}\n"
        f"${data.get("housing_support", "")}\n"
        f"${data.get("clothing_support", "")}\n"
        f"${data.get("health_support", "")}\n"
        f"${data.get("transport_support", "")}\n"
        f"${data.get("basic_services_support", "")}\n"
        f"${data.get("internet_support", "")}\n"
        f"${data.get("cable_tv_support", "")}\n"
        f"${data.get("cell_plan_support", "")}\n"
        f"${data.get("loans_support", "")}\n"
        f"${data.get("unsecured_loans_support", "")}\n"
        f"${data.get("credit_cards_support", "")}\n"
        f"${data.get("alimony_support", "")}\n"
        f"${data.get("commercial_properties_support", "")}\n"
        f"${data.get("financial_support_others", "")}\n"
        f"${data.get("other_expenses_support", "")}\n"
    ], ["B", ""])
    
    table_row_multiline(pdf, [95, 95], [    
        "TOTAL GASTOS:", f"${data.get("total_expenses", "")}"
    ], ["B", ""])
    
    sitransporte = checkbox(data.get("transportation") == 'si')
    notransporte = checkbox(data.get("transportation") == 'no')
    table_row_multiline(pdf, [95, 95], [
        "Posee vehiculo o algun medio de transporte",
        f"{sitransporte} Sí     {notransporte} No"
    ], ["B", ""])
    
    table_row_multiline(pdf, [50, 140], [
        "Descripcion", data.get("transportation_description", "")
    ], ["B", ""])
    
    table_row_multiline(pdf, [190], [
        "Actividad económica adicional (ingresos, horario en que labora):"        
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("additional_economic_activity", "")
    ], [""])

    si_crianza_animales = checkbox(data.get("animal_breeding") == 'si')
    no_crianza_animales = checkbox(data.get("animal_breeding") == 'no')
    table_row_multiline(pdf, [90, 100], [
        "Se dedica a la crianza de animales?",
        f"{si_crianza_animales} Sí     {no_crianza_animales} No"
    ], ["B", ""])
    
    table_row_multiline(pdf, [190], [
        "Resumen de la actividad economica de la familia"
    ], ["B"])
    
    table_row_multiline(pdf, [35, 30, 35, 30, 30, 30], [
        "Ingresos:", data.get("total_income", ""), 
        "Egresos:", data.get("total_expenses", ""), 
        "Saldo:", data.get("balance", "")
    ], ["B", "", "B", "", "B", ""])

def salud(pdf, data):
    section_title(pdf, "6. SALUD")

    practica_deporte = checkbox(data.get("sports") == 'si')
    no_practica_deporte = checkbox(data.get("sports") == 'no')
    table_row_multiline(pdf, [120, 70], [
        "¿El colaborador practica algun deporte o actividad física?",
        f"{practica_deporte} Sí     {no_practica_deporte} No"
    ], ["B", ""])
    
    table_row_multiline(pdf, [50, 140], [
        "¿Cual?", data.get("sports_description", "")
    ], ["B", ""])
    
    table_row_multiline(pdf, [50, 140], [
        "¿Con que frecuencia?", data.get("sports_frequency", "")
    ], ["B", ""])
    
    table_row_multiline(pdf, [90, 100], [
        "¿El colaborador tiene alguna enfermedad?", data.get("disease", "")
    ], ["B", ""])
    
    familiar_problemas_salud = checkbox(data.get("family_health_problems") == 'si')
    no_familiar_problemas_salud = checkbox(data.get("family_health_problems") == 'no')
    table_row_multiline(pdf, [120, 70], [
        "¿La familia tiene problemas de salud?",
        f"{familiar_problemas_salud} Sí     {no_familiar_problemas_salud} No"
    ], ["B", ""])
    
    table_row_multiline(pdf, [50, 140], [
        "¿Cual?", data.get("family_health_problems_description", "")
    ], ["B", ""])
    
    familiar_discapacidad = checkbox(data.get("family_disability") == 'si')
    no_familiar_discapacidad = checkbox(data.get("family_disability") == 'no')
    table_row_multiline(pdf, [120, 70], [
        "¿La familia tiene algún tipo de discapacidad?",
        f"{familiar_discapacidad} Sí     {no_familiar_discapacidad} No"
    ], ["B", ""])
    
    table_row_multiline(pdf, [30, 30, 35, 30, 35, 30], [
        "Tipo:", f"{data.get("family_disability_type", "")}",
        "Porcentaje:", f"{data.get("family_disability_percentage", "")}",
        "Parentezco:", f"{data.get("family_disability_relationship", "")}"
    ], ["B", "", "B", "", "B", ""])

def laboral(pdf, data):
    section_title(pdf, "7. SITUACIÓN LABORAL")

    table_row_multiline(pdf, [190], [
        "Antes de ingresar a Comercial Yolanda Salazar a que se dedicaba (tiempo):"
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("previous_occupation", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Funciones que desempeña actualmente:"
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("current_functions", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Cree que su trabajo se relaciona con su formación, habilidades y experiencia:"
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("job_relation", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Como es la relación con sus compañeros, suelen compartir experiencias y conocimientos:"
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("colleague_relationship", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Que podría mejorar:",
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("improvement_suggestions", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Los conflictos en su área se resuelven de forma abierta y eficaz:"
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("conflict_resolution", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Su trabajo es desgastante:"
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("job_exhaustion", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Siente presión laboral",
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("job_pressure", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Le genera estrés:",
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("job_stress", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Le alcanza el tiempo para estar al día en su trabajo:"
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("job_time_management", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Le parece que su trabajo es reconocido por su jefe inmediato:"
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("job_manager_recognition", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Se siente reconocido con Comercial Yolanda Salazar:"
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("job_recognition", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Cuál es su proyección, su plan de vida o aspiraciones en Comercial Yolanda Salazar:"
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("job_projection", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Alguna vez, se ha sentido discriminado en su lugar de trabajo:"
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("job_discrimination", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Qué haría para mejorar como trabajador:"
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("job_improvement", "")
    ], [""])
    
    table_row_multiline(pdf, [190], [
        "Que tipos de beneficios considera que Comercial Yolanda Salazar podría implementar:"
    ], ["B"])
    
    table_row_multiline(pdf, [190], [
        data.get("job_benefits", "")
    ], [""])

# ========== FIRMAS CON NOMBRES AUTOMÁTICOS ==========
def firmas(pdf, data):
    check_page_space(pdf, 50)
    section_title(pdf, "8. FIRMAS")
    pdf.ln(10)

    # Obtener nombre del entrevistado desde los datos
    nombre_entrevistado = data.get("name", "")
    # Nombre fijo del trabajador social (puedes cambiarlo)
    nombre_trabajador_social = "Mgs. Ana Lucía Pérez"

    # Líneas de firmas
    pdf.cell(90, 8, "_____________________________", 0, 0, "C")
    pdf.cell(10)
    pdf.cell(90, 8, "_____________________________", 0, 1, "C")

    pdf.cell(90, 6, "Firma del entrevistado", 0, 0, "C")
    pdf.cell(10)
    pdf.cell(90, 6, "Firma del trabajador social", 0, 1, "C")

    # Nombres debajo de las firmas
    pdf.cell(90, 6, nombre_entrevistado, 0, 0, "C")
    pdf.cell(10)
    pdf.cell(90, 6, nombre_trabajador_social, 0, 1, "C")


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

    economia(pdf, data)
    pdf.ln(3)

    salud(pdf, data)
    pdf.ln(3)

    laboral(pdf, data)
    pdf.ln(3)

    firmas(pdf, data)

    os.makedirs("pdf", exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"pdf/{data.get('ci','sin_ci')}_{timestamp}.pdf"
    pdf.output(filename)

    return filename