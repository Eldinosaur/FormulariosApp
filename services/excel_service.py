import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime

FILE_PATH = 'data/data.xlsx'


def create_excel_file(data):
    """Crea o actualiza el archivo Excel con todos los datos del formulario"""
    
    # Definir todos los headers (completos)
    HEADERS = [
        
        # ========== 16. METADATOS ==========
        "Fecha registro",
        # ========== 1. DATOS PERSONALES ==========
        "CI", "Nombre completo", "Edad", "Fecha de nacimiento", "Género", 
        "Estado civil", "Email", "Teléfono", "Dirección", "Contacto emergencia",
        
        # ========== 2. DATOS LABORALES ==========
        "Área de trabajo", "Puesto", "Fecha de ingreso",
        
        # ========== 3. EDUCACIÓN ==========
        "Nivel educativo", "Observaciones educación", "Estudia actualmente", 
        "Carrera", "Nivel de estudio",
        
        # ========== 4. DISCAPACIDAD ==========
        "Tiene discapacidad", "Tipo discapacidad", "Porcentaje discapacidad",
        
        # ========== 5. VIVIENDA ==========
        "Sector", "Tipo vivienda", "Clase vivienda", "Tenencia", 
        "N° habitantes", "Tiempo en sector", "¿Seguro?", "Avalúo", "Observaciones vivienda",
        
        # ========== 6. DISTRIBUCIÓN VIVIENDA ==========
        "Dormitorio", "Camas", "Cocina", "Comedor", "Sala", "Baño", "Patio", "Jardín", "Terraza", "Garaje",
        
        # ========== 7. MATERIALES ==========
        "Techo material", "Techo estado", "Techo otros",
        "Pared material", "Pared estado", "Pared otros",
        "Piso material", "Piso estado", "Piso otros",
        "Estructura material", "Estructura estado", "Estructura otros",
        
        # ========== 8. SERVICIOS BÁSICOS ==========
        "Agua", "Luz", "Baño (SSHH)", "Observaciones servicios", "Hacinamiento",
        "Manejo desechos", "Transporte", "Transporte otro", "Electrodomésticos",
        "Tiene Internet", "Tipo Internet",
        
        # ========== 9. ANIMALES ==========
        "Tiene animales", "Tipo animales", "Cantidad animales", 
        "Zona peste", "Lugar tenencia", "Observaciones animales",
        
        # ========== 10. INFORMACIÓN FAMILIAR ==========
        "Tipo familia", "N° hijos", "Hijos fuera matrimonio", "Paga pensión",
        "N° matrimonio", "Comparte tiempo familia", "Frecuencia", "Actividades familiares",
        "Actividades fuera trabajo", "Otras actividades domicilio", "Especificar actividades",
        "Hobbies", "Tiempo hobbies", "Relación pareja", "Por qué relación pareja",
        "Relación hijos", "Por qué relación hijos", "Problemas familiares", "Recibió ayuda",
        "Migración familiar", "Recibió dinero exterior",
        
        # ========== 11. SITUACIÓN ECONÓMICA ==========
        "Gastos compartidos",
        "Aporte padre", "Aporte madre", "Aporte hermanos", "Aporte colaborador",
        "Aporte cónyuge", "Aporte hijos", "Aporte otros", "Total aportes",
        "Monto deudas", "Préstamos formales", "Monto formales",
        "Préstamos informales", "Monto informales", "Préstamos familiares", 
        "Préstamos chulqueros", "Otros préstamos informales", "Tarjetas crédito",
        
        # ========== 12. GASTOS MENSUALES ==========
        "Gasto alimentación", "Gasto educación", "Gasto vivienda", "Gasto vestimenta",
        "Gasto salud", "Gasto transporte", "Gasto servicios básicos", "Gasto internet",
        "Gasto TV cable", "Gasto plan celular", "Gasto préstamos", "Gasto préstamos quirografarios",
        "Gasto tarjetas crédito", "Gasto pensión alimentos", "Gasto locales comerciales",
        "Gasto apoyo terceros", "Otros gastos", "Total gastos",
        
        # ========== 13. INGRESOS Y BALANCE ==========
        "Total ingresos", "Saldo", "Posee vehículo", "Descripción vehículo",
        "Actividad económica adicional", "Crianza animales",
        
        # ========== 14. SALUD ==========
        "Practica deporte", "Cuál deporte", "Frecuencia deporte",
        "Tiene enfermedad", "Problemas salud familia", "Cuáles problemas salud",
        "Discapacidad familia", "Tipo discapacidad familia", "Porcentaje discapacidad familia",
        "Parentezco discapacidad",
        
        # ========== 15. SITUACIÓN LABORAL ==========
        "Ocupación anterior", "Funciones actuales", "Relación con formación",
        "Relación compañeros", "Qué mejoraría", "Resolución conflictos",
        "Trabajo desgastante", "Presión laboral", "Genera estrés", "Tiempo suficiente",
        "Reconocimiento jefe", "Reconocimiento empresa", "Proyección", "Discriminación",
        "Mejora como trabajador", "Beneficios sugeridos"
        
       
    ]
    
    # Agregar headers de familiares (6 máximo, 4 campos cada uno)
    for i in range(1, 7):
        HEADERS.extend([
            f"F{i}_Nombre",
            f"F{i}_Edad",
            f"F{i}_Parentesco",
            f"F{i}_Ocupación",
            f"F{i}_Ingreso"
        ])
    
    # Crear archivo si no existe
    if not os.path.exists(FILE_PATH):
        os.makedirs('data', exist_ok=True)
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Fichas Socioeconómicas"
        
        # Estilos para headers
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="2563eb", end_color="2563eb", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Escribir headers con estilos
        for col_idx, header in enumerate(HEADERS, 1):
            cell = worksheet.cell(row=1, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border
        
        # Ajustar anchos de columnas
        for col_idx, header in enumerate(HEADERS, 1):
            column_letter = worksheet.cell(row=1, column=col_idx).column_letter
            worksheet.column_dimensions[column_letter].width = min(max(len(str(header)) + 3, 12), 30)
        
        workbook.save(FILE_PATH)
    
    # Cargar archivo existente
    workbook = load_workbook(FILE_PATH)
    worksheet = workbook.active
    
    # ========== CONSTRUIR FILA CON TODOS LOS DATOS ==========
    row = []
    
    # 16. METADATOS
    row.append(data.get("visit_date", datetime.now().strftime("%d/%m/%Y")))
    
    # 1. DATOS PERSONALES
    row.append(data.get("ci", ""))
    row.append(data.get("name", ""))
    row.append(data.get("age", ""))
    row.append(data.get("birth_date_place", ""))
    row.append(data.get("gender", ""))
    row.append(data.get("marital_status", ""))
    row.append(data.get("email", ""))
    row.append(data.get("phone", ""))
    row.append(data.get("address", ""))
    row.append(data.get("emergency_contact", ""))
    
    # 2. DATOS LABORALES
    row.append(data.get("work_area", ""))
    row.append(data.get("job_title", ""))
    row.append(data.get("entry_date", ""))
    
    # 3. EDUCACIÓN
    row.append(data.get("educational_level", ""))
    row.append(data.get("educational_observations", ""))
    row.append(data.get("currently_studying", ""))
    row.append(data.get("study_career", ""))
    row.append(data.get("study_level", ""))
    
    # 4. DISCAPACIDAD
    row.append(data.get("disability", ""))
    row.append(data.get("disability_type", ""))
    row.append(data.get("disability_percentage", ""))
    
    # 5. VIVIENDA
    row.append(data.get("sector", ""))
    row.append(data.get("house_type", ""))
    row.append(data.get("house_class", ""))
    row.append(data.get("house_ownership", ""))
    row.append(data.get("household_size", ""))
    row.append(data.get("time_living_sector", ""))
    row.append(data.get("is_safe", ""))
    row.append(data.get("house_valuation", ""))
    row.append(data.get("house_observations", ""))
    
    # 6. DISTRIBUCIÓN VIVIENDA
    distribution = data.get("house_distribution", [])
    for item in ["dormitorio", "camas", "cocina", "comedor", "sala", 
                 "baño", "patio", "jardín", "terraza", "garaje"]:
        row.append("Sí" if item in distribution else "No")
    
    # 7. MATERIALES
    row.append(data.get("roof_type", ""))
    row.append(data.get("roof_status", ""))
    row.append(data.get("roof_type_other", ""))
    row.append(data.get("wall_type", ""))
    row.append(data.get("wall_status", ""))
    row.append(data.get("wall_type_other", ""))
    row.append(data.get("floor_type", ""))
    row.append(data.get("floor_status", ""))
    row.append(data.get("floor_type_other", ""))
    row.append(data.get("structure_type", ""))
    row.append(data.get("structure_status", ""))
    row.append(data.get("structure_type_other", ""))
    
    # 8. SERVICIOS BÁSICOS
    row.append(data.get("water_type", ""))
    row.append(data.get("light_type", ""))
    row.append(data.get("sshh_type", ""))
    row.append(data.get("basic_services_other", ""))
    row.append(data.get("hacinamiento", ""))
    row.append(data.get("waste_management", ""))
    row.append(data.get("transport_type", ""))
    row.append(data.get("transport_type_other", ""))
    row.append(data.get("appliances", ""))
    row.append(data.get("internet", ""))
    row.append(data.get("internet_type", ""))
    
    # 9. ANIMALES
    row.append(data.get("animals", ""))
    row.append(data.get("animal_type", ""))
    row.append(data.get("animal_quantity", ""))
    row.append(data.get("peste", ""))
    row.append(data.get("animal_location", ""))
    row.append(data.get("animal_observations", ""))
    
    # 10. INFORMACIÓN FAMILIAR
    row.append(data.get("family_type", ""))
    row.append(data.get("children_count", ""))
    row.append(data.get("children_outside_marriage", ""))
    row.append(data.get("pays_alimony", ""))
    row.append(data.get("marriage_number", ""))
    row.append(data.get("family_time", ""))
    row.append(data.get("family_time_frequency", ""))
    row.append(data.get("family_activities", ""))
    row.append(data.get("other_activities", ""))
    row.append(data.get("other_activities", ""))  # ¿Otras actividades?
    row.append(data.get("other_activities_specify", ""))
    row.append(data.get("hobbies", ""))
    row.append(data.get("hobbies_time", ""))
    row.append(data.get("partner_relationship", ""))
    row.append(data.get("partner_relationship_reason", ""))
    row.append(data.get("children_relationship", ""))
    row.append(data.get("children_relationship_reason", ""))
    row.append(data.get("family_problems", ""))
    row.append(data.get("family_problems_help", ""))
    row.append(data.get("family_migration", ""))
    row.append(data.get("family_migration_received", ""))
    
    # 11. SITUACIÓN ECONÓMICA - APORTES
    row.append(data.get("shared_expenses", ""))
    row.append(data.get("father_contribution", ""))
    row.append(data.get("mother_contribution", ""))
    row.append(data.get("siblings_contribution", ""))
    row.append(data.get("collaborators_contribution", ""))
    row.append(data.get("spouse_contribution", ""))
    row.append(data.get("children_contribution", ""))
    row.append(data.get("other_contribution", ""))
    row.append(data.get("total_contribution", ""))
    row.append(data.get("debt_amount", ""))
    row.append(data.get("formal_loans", ""))
    row.append(data.get("formal_loans_amount", ""))
    row.append(data.get("informal_loans", ""))
    row.append(data.get("informal_loans_amount", ""))
    row.append(data.get("informal_loans_family_amount", ""))
    row.append(data.get("informal_loans_moneylender_amount", ""))
    row.append(data.get("informal_loans_other_amount", ""))
    row.append(data.get("credit_cards", ""))
    
    # 12. GASTOS MENSUALES
    row.append(data.get("food_support", ""))
    row.append(data.get("education_support", ""))
    row.append(data.get("housing_support", ""))
    row.append(data.get("clothing_support", ""))
    row.append(data.get("health_support", ""))
    row.append(data.get("transport_support", ""))
    row.append(data.get("basic_services_support", ""))
    row.append(data.get("internet_support", ""))
    row.append(data.get("cable_tv_support", ""))
    row.append(data.get("cell_plan_support", ""))
    row.append(data.get("loans_support", ""))
    row.append(data.get("unsecured_loans_support", ""))
    row.append(data.get("credit_cards_support", ""))
    row.append(data.get("alimony_support", ""))
    row.append(data.get("commercial_properties_support", ""))
    row.append(data.get("financial_support_others", ""))
    row.append(data.get("other_expenses_support", ""))
    row.append(data.get("total_expenses", ""))
    
    # 13. INGRESOS Y BALANCE
    row.append(data.get("total_income", ""))
    row.append(data.get("balance", ""))
    row.append(data.get("transportation", ""))
    row.append(data.get("transportation_description", ""))
    row.append(data.get("additional_economic_activity", ""))
    row.append(data.get("animal_breeding", ""))
    
    # 14. SALUD
    row.append(data.get("sports", ""))
    row.append(data.get("sports_description", ""))
    row.append(data.get("sports_frequency", ""))
    row.append(data.get("disease", ""))
    row.append(data.get("family_health_problems", ""))
    row.append(data.get("family_health_problems_description", ""))
    row.append(data.get("family_disability", ""))
    row.append(data.get("family_disability_type", ""))
    row.append(data.get("family_disability_percentage", ""))
    row.append(data.get("family_disability_relationship", ""))
    
    # 15. SITUACIÓN LABORAL
    row.append(data.get("previous_occupation", ""))
    row.append(data.get("current_functions", ""))
    row.append(data.get("job_relation", ""))
    row.append(data.get("colleague_relationship", ""))
    row.append(data.get("improvement_suggestions", ""))
    row.append(data.get("conflict_resolution", ""))
    row.append(data.get("job_exhaustion", ""))
    row.append(data.get("job_pressure", ""))
    row.append(data.get("job_stress", ""))
    row.append(data.get("job_time_management", ""))
    row.append(data.get("job_manager_recognition", ""))
    row.append(data.get("job_recognition", ""))
    row.append(data.get("job_projection", ""))
    row.append(data.get("job_discrimination", ""))
    row.append(data.get("job_improvement", ""))
    row.append(data.get("job_benefits", ""))
    
    
    
    # ========== FAMILIARES (MÁXIMO 6) ==========
    family_members = data.get("family_members", [])
    
    for i in range(6):
        if i < len(family_members):
            member = family_members[i]
            row.extend([
                member.get("name", ""),
                member.get("age", ""),
                member.get("relation", ""),
                member.get("job", ""),
                member.get("income", "")
            ])
        else:
            # Campos vacíos si no hay suficientes familiares
            row.extend(["", "", "", "", ""])
    
    # Verificar que la fila tenga la longitud correcta
    expected_length = len(HEADERS)
    if len(row) != expected_length:
        print(f"Advertencia: Longitud de fila ({len(row)}) no coincide con headers ({expected_length})")
        # Ajustar longitud
        if len(row) < expected_length:
            row.extend([""] * (expected_length - len(row)))
        else:
            row = row[:expected_length]
    
    # Agregar la fila al worksheet
    worksheet.append(row)
    
    # Autoajustar anchos de columnas (solo si es nueva hoja o cada cierto tiempo)
    for col_idx, header in enumerate(HEADERS, 1):
        column_letter = worksheet.cell(row=1, column=col_idx).column_letter
        max_length = len(str(header))
        
        # Revisar algunas filas para ajustar ancho
        for row_idx in range(2, min(worksheet.max_row, 10)):
            cell_value = worksheet.cell(row=row_idx, column=col_idx).value
            if cell_value:
                max_length = max(max_length, len(str(cell_value)))
        
        worksheet.column_dimensions[column_letter].width = min(max_length + 3, 40)
    
    # Guardar el archivo
    workbook.save(FILE_PATH)
    
    return FILE_PATH


def get_all_records():
    """Obtiene todos los registros del archivo Excel"""
    if not os.path.exists(FILE_PATH):
        return []
    
    workbook = load_workbook(FILE_PATH)
    worksheet = workbook.active
    
    records = []
    headers = [cell.value for cell in worksheet[1]]
    
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        if any(row):  # Si la fila no está vacía
            record = dict(zip(headers, row))
            records.append(record)
    
    return records


def get_record_by_ci(ci):
    """Busca un registro por cédula"""
    records = get_all_records()
    for record in records:
        if record.get("CI") == ci:
            return record
    return None