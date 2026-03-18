import os
from openpyxl import Workbook, load_workbook

FILE_PATH = 'data/data.xlsx'

HEADERS = [
    # Datos principales
    "CI", "Nombre", "Edad", "Email", "Dirección", "Teléfono", "Fecha",

    # Vivienda
    "Sector", "Tipo Vivienda", "Tenencia", "N° Personas", "Tiempo Sector",

    # Servicios
    "Agua", "Luz", "Internet",

    # Economía
    "Gastos Compartidos", "Aportantes", "Gastos", "Deudas",

    # Salud y otros
    "Salud", "Funciones", "Plan Vida", "Observaciones",
]

# Agregar headers de familiares (6 máximo)
for i in range(1, 7):
    HEADERS.extend([
        f"F{i}_Nombre",
        f"F{i}_Edad",
        f"F{i}_Parentesco",
        f"F{i}_Ocupacion"
    ])


def create_excel_file(data):

    # CI = ID
    ci = data.get("ci") or data.get("id")

    # Crear archivo si no existe
    if not os.path.exists(FILE_PATH):
        os.makedirs('data', exist_ok=True)
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(HEADERS)
        workbook.save(FILE_PATH)

    workbook = load_workbook(FILE_PATH)
    worksheet = workbook.active

    # -------------------------
    # BASE DATA
    # -------------------------
    row = [
        ci,
        data.get("name"),
        data.get("age"),
        data.get("email"),
        data.get("address"),
        data.get("phone"),
        data.get("date"),

        data.get("sector"),
        data.get("house_type"),
        data.get("ownership"),
        data.get("people_count"),
        data.get("time_sector"),

        str(data.get("water")),
        str(data.get("electricity")),
        str(data.get("internet")),

        str(data.get("shared_expenses")),
        data.get("contributors"),
        data.get("expenses"),
        data.get("debts"),

        data.get("disease"),
        data.get("functions"),
        data.get("life_plan"),
        data.get("observations"),
    ]

    # -------------------------
    # FAMILIARES (MAX 6)
    # -------------------------
    family = data.get("family_members", [])

    for i in range(6):
        if i < len(family):
            member = family[i]
            row.extend([
                member.get("name", ""),
                member.get("age", ""),
                member.get("relation", ""),
                member.get("job", "")
            ])
        else:
            # Vacíos si no hay suficientes
            row.extend(["", "", "", ""])

    worksheet.append(row)
    workbook.save(FILE_PATH)