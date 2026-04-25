import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime

from services import pdf_service as pdf
from services import excel_service as xlsx


class App:

    def __init__(self, root):

        self.root = root
        self.root.title("Ficha Socioeconómica")
        self.root.geometry("1400x850")
        self.root.configure(bg="#f0f2f5")

        self.entries = {}
        self.family_rows = []
        self.MAX_FAMILY_MEMBERS = 6
        self.photo_path = tk.StringVar()

        # Colores modernos
        self.colors = {
            "primary": "#2563eb",
            "primary_dark": "#1d4ed8",
            "secondary": "#64748b",
            "success": "#059669",
            "danger": "#dc2626",
            "warning": "#d97706",
            "bg": "#f1f5f9",
            "white": "#ffffff",
            "card_border": "#e2e8f0",
            "tab_selected": "#2563eb",
            "tab_normal": "#e2e8f0",
            "tab_text_selected": "#ffffff",
            "tab_text_normal": "#1e293b"
        }

        # =========================
        # BARRA SUPERIOR
        # =========================
        self.top_bar = tk.Frame(root, bg=self.colors["primary"], height=60)
        self.top_bar.pack(fill="x")
        self.top_bar.pack_propagate(False)

        tk.Label(self.top_bar, text="📋 FICHA SOCIOECONÓMICA", 
                font=("Segoe UI", 16, "bold"),
                fg="white", bg=self.colors["primary"]).pack(side="left", padx=20, pady=15)

        tk.Label(self.top_bar, text="Developed by: Eldinosaur", 
                font=("Segoe UI", 10),
                fg="#bfdbfe", bg=self.colors["primary"]).pack(side="left", padx=10)

        # Botón guardar
        save_btn = tk.Button(self.top_bar, text="💾 GUARDAR", 
                            command=self.save_all,
                            bg=self.colors["success"], fg="white",
                            font=("Segoe UI", 10, "bold"),
                            relief="flat", padx=20, pady=8)
        save_btn.pack(side="right", padx=20, pady=10)

        # =========================
        # NOTEBOOK PERSONALIZADO (sin ttk)
        # =========================
        self.notebook_frame = tk.Frame(root, bg=self.colors["bg"])
        self.notebook_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Frame para las pestañas (tabs)
        self.tab_bar = tk.Frame(self.notebook_frame, bg=self.colors["bg"], height=45)
        self.tab_bar.pack(fill="x")
        self.tab_bar.pack_propagate(False)

        # Frame para el contenido de las pestañas
        self.content_frame = tk.Frame(self.notebook_frame, bg=self.colors["bg"])
        self.content_frame.pack(fill="both", expand=True)

        # Diccionario para almacenar las pestañas
        self.tabs = {}
        self.current_tab = None

        # Crear pestañas personalizadas
        self.create_custom_tab("personales", "👤 DATOS PERSONALES", self.build_personales)
        self.create_custom_tab("familia", "👨‍👩‍👧‍👦 FAMILIA", self.build_familia)
        self.create_custom_tab("vivienda", "🏠 VIVIENDA", self.build_vivienda)
        self.create_custom_tab("economia", "💰 ECONOMÍA", self.build_economia)
        self.create_custom_tab("salud", "❤️ SALUD", self.build_salud)
        self.create_custom_tab("laboral", "💼 LABORAL", self.build_laboral)

        # Seleccionar primera pestaña
        self.select_tab("personales")

    def create_custom_tab(self, tab_id, title, builder_func):
        """Crea una pestaña personalizada que no se queda en blanco"""
        
        # Crear botón de pestaña
        tab_button = tk.Button(
            self.tab_bar, 
            text=title,
            font=("Segoe UI", 10, "bold"),
            bg=self.colors["tab_normal"],
            fg=self.colors["tab_text_normal"],
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2",
            command=lambda: self.select_tab(tab_id)
        )
        tab_button.pack(side="left", padx=(0, 2))
        
        # Crear frame para el contenido
        tab_content = tk.Frame(self.content_frame, bg=self.colors["bg"])
        
        # Canvas con scroll para el contenido
        canvas = tk.Canvas(tab_content, bg=self.colors["bg"], highlightthickness=0)
        scrollbar = tk.Scrollbar(tab_content, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors["bg"])
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e, c=canvas: c.configure(scrollregion=c.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        tab_content.grid_rowconfigure(0, weight=1)
        tab_content.grid_columnconfigure(0, weight=1)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Responsive
        def on_configure(event):
            canvas.itemconfig(1, width=event.width - 20)
        
        canvas.bind("<Configure>", on_configure)
        
        # Mouse wheel
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind("<MouseWheel>", on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", on_mousewheel)
        
        # Guardar referencias
        self.tabs[tab_id] = {
            "button": tab_button,
            "content": tab_content,
            "builder": builder_func,
            "frame": scrollable_frame
        }
        
        # Construir el contenido
        builder_func(scrollable_frame)

    def select_tab(self, tab_id):
        """Selecciona una pestaña y actualiza los estilos"""
        
        # Ocultar todas las pestañas
        for tid, tab in self.tabs.items():
            tab["content"].pack_forget()
            # Restaurar estilo normal
            tab["button"].config(
                bg=self.colors["tab_normal"],
                fg=self.colors["tab_text_normal"]
            )
        
        # Mostrar la pestaña seleccionada
        self.tabs[tab_id]["content"].pack(fill="both", expand=True)
        # Aplicar estilo seleccionado
        self.tabs[tab_id]["button"].config(
            bg=self.colors["tab_selected"],
            fg=self.colors["tab_text_selected"]
        )
        
        self.current_tab = tab_id

    def create_card(self, parent, title, icon="📄"):
        """Tarjeta con título y contenido"""
        frame = tk.Frame(parent, bg=self.colors["white"], 
                        relief="flat", bd=1,
                        highlightbackground=self.colors["card_border"],
                        highlightthickness=1)
        frame.pack(fill="x", padx=15, pady=(0, 15))
        
        # Header
        header = tk.Frame(frame, bg=self.colors["primary"], height=40)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        tk.Label(header, text=f"{icon} {title}", 
                font=("Segoe UI", 11, "bold"),
                fg="white", bg=self.colors["primary"]).pack(side="left", padx=15)
        
        # Contenido
        content = tk.Frame(frame, bg=self.colors["white"], padx=20, pady=15)
        content.pack(fill="x", expand=True)
        
        # Configurar grid responsivo
        content.grid_columnconfigure(0, weight=1)
        content.grid_columnconfigure(2, weight=1)
        
        return content

    def create_field(self, parent, label, key, row, column=0, width=35, required=False):
        """Campo de entrada"""
        label_text = f"{label}{' *' if required else ''}"
        
        tk.Label(parent, text=label_text, font=("Segoe UI", 9),
                bg=self.colors["white"], fg=self.colors["secondary"],
                anchor="w").grid(row=row, column=column, sticky="ew", 
                                pady=(6, 2), padx=(0, 10))
        
        entry = tk.Entry(parent, font=("Segoe UI", 10), 
                        width=width, relief="solid", bd=1,
                        highlightbackground=self.colors["card_border"],
                        highlightthickness=1)
        entry.grid(row=row, column=column+1, sticky="ew", pady=(6, 2), padx=(0, 20))
        
        self.entries[key] = entry
        return entry

    def create_number_field(self, parent, label, key, row, column=0, width=15):
        """Campo numérico con validación"""
        tk.Label(parent, text=label, font=("Segoe UI", 9),
                bg=self.colors["white"], fg=self.colors["secondary"]).grid(
                row=row, column=column, sticky="w", pady=(6, 2), padx=(0, 10))
        
        entry = tk.Entry(parent, font=("Segoe UI", 10), 
                        width=width, relief="solid", bd=1,
                        highlightbackground=self.colors["card_border"],
                        highlightthickness=1)
        entry.grid(row=row, column=column+1, sticky="w", pady=(6, 2), padx=(0, 20))
        
        # Validar que solo ingresen números
        def validate_number(P):
            if P == "" or P.isdigit() or (P.replace('.', '', 1).isdigit() and P.count('.') <= 1):
                return True
            return False
        
        validate_cmd = parent.register(validate_number)
        entry.config(validate="key", validatecommand=(validate_cmd, "%P"))
        
        self.entries[key] = entry
        return entry

    def create_combo(self, parent, label, key, values, row, column=0, width=30):
        """Combobox"""
        tk.Label(parent, text=label, font=("Segoe UI", 9),
                bg=self.colors["white"], fg=self.colors["secondary"]).grid(
                row=row, column=column, sticky="w", pady=(6, 2), padx=(0, 10))
        
        combo = ttk.Combobox(parent, values=values, width=width,
                            font=("Segoe UI", 10))
        combo.grid(row=row, column=column+1, sticky="w", pady=(6, 2), padx=(0, 20))
        
        self.entries[key] = combo
        return combo

    def create_radio_group(self, parent, label, key, options, row, column=0):
        """Grupo de radio buttons"""
        tk.Label(parent, text=label, font=("Segoe UI", 9),
                bg=self.colors["white"], fg=self.colors["secondary"]).grid(
                row=row, column=column, sticky="nw", pady=(8, 2), padx=(0, 10))
        
        frame = tk.Frame(parent, bg=self.colors["white"])
        frame.grid(row=row, column=column+1, sticky="w", pady=(8, 2))
        
        var = tk.StringVar(value=options[0][1] if options else "")
        self.entries[key] = var
        
        for text, value in options:
            rb = tk.Radiobutton(frame, text=text, variable=var, value=value,
                               bg=self.colors["white"], font=("Segoe UI", 9))
            rb.pack(side="left", padx=(0, 15))
        
        return var

    # =========================
    # CONSTRUCCIÓN DE PESTAÑAS
    # =========================
    def build_personales(self, parent):
        s1 = self.create_card(parent, "Datos Personales", "👤")
        
        self.create_field(s1, "Cédula *", "ci", 0, 0, 25, required=True)
        self.create_field(s1, "Nombre completo *", "name", 0, 2, 35, required=True)
        self.create_number_field(s1, "Edad", "age", 1, 0, 15)
        self.create_field(s1, "Lugar y Fecha de Nacimiento", "birth_date_place", 1, 2, 35)
        
        self.create_radio_group(s1, "Género", "gender",
                               [("Masculino", "masculino"), ("Femenino", "femenino"), ("Otro", "otro")], 2, 0)
        
        self.create_radio_group(s1, "Estado Civil", "marital_status",
                               [("Soltero", "soltero"), ("Casado", "casado"), ("Separado", "separado"),
                                ("Divorciado", "divorciado"), ("Unión Libre", "unión libre"), ("Viudo", "viudo")], 3, 0)
        
        self.create_field(s1, "Email", "email", 4, 0, 30)
        self.create_field(s1, "Área de trabajo", "work_area", 4, 2, 25)
        self.create_field(s1, "Puesto que desempeña", "job_title", 5, 0, 25)
        self.create_field(s1, "Fecha de Ingreso", "entry_date", 5, 2, 20)
        
        self.create_radio_group(s1, "¿Tiene discapacidad?", "disability",
                               [("Sí", "si"), ("No", "no")], 6, 0)
        self.create_field(s1, "Tipo de discapacidad", "disability_type", 7, 0, 30)
        self.create_number_field(s1, "Porcentaje", "disability_percentage", 7, 2, 15)
        
        self.create_radio_group(s1, "Nivel de Educación", "educational_level",
                               [("Profesional", "profesional"), ("Bachiller", "bachiller"), 
                                ("Básica", "basica"), ("Primaria", "primaria"), ("Sin estudios", "nuloestudio")], 8, 0)
        
        self.create_field(s1, "Observaciones", "educational_observations", 9, 0, 60)
        
        self.create_radio_group(s1, "¿Actualmente estudia?", "currently_studying",
                               [("Sí", "si"), ("No", "no")], 10, 0)
        self.create_field(s1, "Carrera", "study_career", 11, 0, 30)
        self.create_field(s1, "Nivel", "study_level", 11, 2, 20)
        
        self.create_field(s1, "Dirección Domiciliaria", "address", 12, 0, 50)
        self.create_field(s1, "Teléfono", "phone", 13, 0, 20)
        self.create_field(s1, "Contacto de emergencia", "emergency_contact", 13, 2, 25)
        
        # Foto
        s2 = self.create_card(parent, "Fotografía", "📷")
        photo_frame = tk.Frame(s2, bg=self.colors["white"])
        photo_frame.pack(fill="x")
        
        tk.Entry(photo_frame, textvariable=self.photo_path, 
                font=("Segoe UI", 10), width=60,
                relief="solid", bd=1).pack(side="left", padx=(0, 10), fill="x", expand=True)
        tk.Button(photo_frame, text="📁 Buscar", command=self.load_photo,
                 bg=self.colors["secondary"], fg="white",
                 font=("Segoe UI", 10), relief="flat", padx=15, pady=5).pack(side="right")

    def build_familia(self, parent):
        s1 = self.create_card(parent, "Miembros del Hogar", "👨‍👩‍👧‍👦")
        
        # Indicador de límite
        limit_frame = tk.Frame(s1, bg=self.colors["white"])
        limit_frame.pack(fill="x", pady=(0, 10))
        
        self.family_counter_label = tk.Label(limit_frame, 
            text=f"📊 Familiares registrados: 0 / {self.MAX_FAMILY_MEMBERS}",
            font=("Segoe UI", 9, "bold"),
            fg=self.colors["primary"], bg=self.colors["white"])
        self.family_counter_label.pack(side="left")
        
        # Tabla de miembros
        table_frame = tk.Frame(s1, bg=self.colors["primary"])
        table_frame.pack(fill="x", pady=(0, 10))
        
        headers = ["Nombres y Apellidos", "Edad", "Parentesco", "Estado Civil", "Instrucción", "Ocupación", "Lugar trabajo/estudio", "Ingreso", "Teléfono", ""]
        widths = [18, 6, 10, 10, 10, 12, 14, 8, 10, 4]
        
        for i, header in enumerate(headers):
            tk.Label(table_frame, text=header, font=("Segoe UI", 8, "bold"),
                    fg="white", bg=self.colors["primary"], width=widths[i], wraplength=100).grid(row=0, column=i, padx=2, pady=5)
        
        self.family_container = tk.Frame(s1, bg=self.colors["white"])
        self.family_container.pack(fill="x", pady=5)
        
        # Botón agregar
        self.add_family_btn = tk.Button(s1, text="+ Agregar familiar", 
                                       command=self.add_family,
                                       bg=self.colors["primary"], fg="white",
                                       font=("Segoe UI", 10), relief="flat", 
                                       padx=20, pady=8)
        self.add_family_btn.pack(pady=10)
        
        # Información de la familia
        s2 = self.create_card(parent, "Información de la Familia", "📊")
        
        self.create_radio_group(s2, "Tipo de Familia", "family_type",
                               [("Nuclear", "nuclear"), ("Extensa Cosanguínea", "cosanguinea"),
                                ("Extensa Afines", "afines"), ("Ampliada", "ampliada"),
                                ("Monoparental Materna", "materna"), ("Monoparental Paterna", "paterna"),
                                ("Unipersonal", "unipersonal")], 0, 0)
        
        self.create_number_field(s2, "N° de hijos del colaborador", "children_count", 1, 0, 15)
        
        self.create_radio_group(s2, "¿Hijos fuera del matrimonio?", "children_outside_marriage",
                               [("Sí", "si"), ("No", "no")], 2, 0)
        
        self.create_field(s2, "Paga pensión alimenticia", "pays_alimony", 3, 0, 30)
        self.create_number_field(s2, "N° de matrimonio/relación actual", "marriage_number", 4, 0, 15)
        self.create_field(s2, "¿Comparte tiempo de descanso en familia?", "family_time", 5, 0, 40)
        self.create_field(s2, "Frecuencia", "family_time_frequency", 6, 0, 30)
        self.create_field(s2, "Actividades familiares", "family_activities", 7, 0, 50)
        self.create_field(s2, "Actividades fuera del trabajo", "other_activities", 8, 0, 50)
        
        self.create_radio_group(s2, "¿Otras actividades en domicilio?", "other_activities",
                               [("Sí", "si"), ("No", "no")], 9, 0)
        self.create_field(s2, "Especificar", "other_activities_specify", 10, 0, 50)
        
        self.create_field(s2, "Hobbies", "hobbies", 11, 0, 50)
        self.create_field(s2, "Tiempo dedicado", "hobbies_time", 12, 0, 30)
        
        self.create_radio_group(s2, "Relación de pareja", "partner_relationship",
                               [("Muy buena", "muy buena"), ("Buena", "buena"),
                                ("Regular", "regular"), ("Mala", "mala")], 13, 0)
        self.create_field(s2, "¿Por qué?", "partner_relationship_reason", 14, 0, 50)
        
        self.create_radio_group(s2, "Relación con hijos", "children_relationship",
                               [("Muy buena", "muy buena"), ("Buena", "buena"),
                                ("Regular", "regular"), ("Mala", "mala")], 15, 0)
        self.create_field(s2, "¿Por qué?", "children_relationship_reason", 16, 0, 50)
        
        self.create_radio_group(s2, "Problemas familiares", "family_problems",
                               [("Divorcios", "divorcios"), ("Separaciones", "separaciones"),
                                ("Resentimientos", "resentimientos"), ("Otros", "otros"),
                                ("Ninguno", "ningunos")], 17, 0)
        self.create_field(s2, "¿Recibió ayuda?", "family_problems_help", 18, 0, 50)
        
        self.create_radio_group(s2, "¿Miembros que salieron del país?", "family_migration",
                               [("Sí", "si"), ("No", "no")], 19, 0)
        self.create_radio_group(s2, "¿Recibió dinero del exterior?", "family_migration_received",
                               [("Sí", "si"), ("No", "no")], 20, 0)

    def build_vivienda(self, parent):
        s = self.create_card(parent, "Características de la Vivienda", "🏠")
        
        self.create_radio_group(s, "Sector", "sector",
                               [("Urbano", "urbano"), ("Rural", "rural")], 0, 0)
        
        self.create_radio_group(s, "Tenencia", "house_ownership",
                               [("Propia", "propia"), ("Arrendada", "arrendada"),
                                ("Hipotecada", "hipotecada"), ("Prestada", "prestada"),
                                ("Otros", "otros")], 1, 0)
        
        self.create_number_field(s, "N° de habitantes", "household_size", 2, 0, 15)
        self.create_field(s, "Tiempo en el sector", "time_living_sector", 2, 2, 20)
        
        self.create_radio_group(s, "¿Se considera seguro?", "is_safe",
                               [("Sí", "si"), ("No", "no")], 3, 0)
        
        self.create_radio_group(s, "Tipo de vivienda", "house_type",
                               [("Casa", "casa"), ("Departamento", "departamento"),
                                ("Cuarto", "cuarto"), ("Mediagua", "mediagua")], 4, 0)
        
        self.create_radio_group(s, "Clase", "house_class",
                               [("Precaria", "precaria"), ("Elemental", "elemental"),
                                ("Completa", "completa")], 5, 0)
        
        self.create_field(s, "Avalúo", "house_valuation", 6, 0, 20)
        self.create_field(s, "Observaciones", "house_observations", 7, 0, 50)
        
        # Distribución de la Vivienda - AHORA CON CAMPOS NUMÉRICOS
        s2 = self.create_card(parent, "Distribución de la Vivienda (cantidades)", "🏗️")
        dist_frame = tk.Frame(s2, bg=self.colors["white"])
        dist_frame.pack(fill="x")
        
        # Lista de elementos de la vivienda con sus claves
        distribuciones = [
            ("Dormitorios", "dormitorios"),
            ("Camas", "camas"),
            ("Cocinas", "cocinas"),
            ("Comedores", "comedores"),
            ("Salas", "salas"),
            ("Baños", "banos"),
            ("Patios", "patios"),
            ("Jardines", "jardines"),
            ("Terrazas", "terrazas"),
            ("Garajes", "garajes")
        ]
        
        for i, (label, key) in enumerate(distribuciones):
            tk.Label(dist_frame, text=label, font=("Segoe UI", 9),
                    bg=self.colors["white"], fg=self.colors["secondary"]).grid(row=i//2, column=(i%2)*2, padx=15, pady=5, sticky="w")
            
            entry = tk.Entry(dist_frame, width=10, font=("Segoe UI", 10))
            entry.grid(row=i//2, column=(i%2)*2+1, padx=5, pady=5, sticky="w")
            
            # Validar solo números
            def validate_number(P):
                if P == "" or P.isdigit():
                    return True
                return False
            validate_cmd = dist_frame.register(validate_number)
            entry.config(validate="key", validatecommand=(validate_cmd, "%P"))
            
            self.entries[f"house_distribution_{key}"] = entry
        
        # Materiales
        s3 = self.create_card(parent, "Materiales de Construcción", "🔨")
        
        # Techo
        tk.Label(s3, text="TECHO", font=("Segoe UI", 10, "bold"),
                fg=self.colors["primary"], bg=self.colors["white"]).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))
        
        self.create_radio_group(s3, "Material", "roof_type",
                               [("Loza", "loza"), ("Eternit", "eternit"), ("Zinc", "zinc"),
                                ("Teja", "teja"), ("Plástico", "plastico")], 1, 0)
        self.create_radio_group(s3, "Estado", "roof_status",
                               [("Bueno", "bueno"), ("Regular", "regular"), ("Malo", "malo")], 1, 2)
        self.create_field(s3, "Otros", "roof_type_other", 2, 0, 30)
        
        # Pared
        tk.Label(s3, text="PARED", font=("Segoe UI", 10, "bold"),
                fg=self.colors["primary"], bg=self.colors["white"]).grid(row=3, column=0, columnspan=4, sticky="w", pady=(15, 10))
        
        self.create_radio_group(s3, "Material", "wall_type",
                               [("Hormigón", "hormigon"), ("Bloque", "bloque"),
                                ("Adobe", "adobe"), ("Madera", "madera")], 4, 0)
        self.create_radio_group(s3, "Estado", "wall_status",
                               [("Bueno", "bueno"), ("Regular", "regular"), ("Malo", "malo")], 4, 2)
        self.create_field(s3, "Otros", "wall_type_other", 5, 0, 30)
        
        # Piso
        tk.Label(s3, text="PISO", font=("Segoe UI", 10, "bold"),
                fg=self.colors["primary"], bg=self.colors["white"]).grid(row=6, column=0, columnspan=4, sticky="w", pady=(15, 10))
        
        self.create_radio_group(s3, "Material", "floor_type",
                               [("Entablado", "entablado"), ("Baldosa", "baldosa"),
                                ("Cemento", "cemento"), ("Tierra", "tierra")], 7, 0)
        self.create_radio_group(s3, "Estado", "floor_status",
                               [("Bueno", "bueno"), ("Regular", "regular"), ("Malo", "malo")], 7, 2)
        self.create_field(s3, "Otros", "floor_type_other", 8, 0, 30)
        
        # Estructura
        tk.Label(s3, text="ESTRUCTURA", font=("Segoe UI", 10, "bold"),
                fg=self.colors["primary"], bg=self.colors["white"]).grid(row=9, column=0, columnspan=4, sticky="w", pady=(15, 10))
        
        self.create_radio_group(s3, "Material", "structure_type",
                               [("Hormigón", "hormigon"), ("Hierro", "hierro"),
                                ("Madera", "madera"), ("Bloque", "bloque")], 10, 0)
        self.create_radio_group(s3, "Estado", "structure_status",
                               [("Bueno", "bueno"), ("Regular", "regular"), ("Malo", "malo")], 10, 2)
        self.create_field(s3, "Otros", "structure_type_other", 11, 0, 30)
        
        # Servicios básicos
        s4 = self.create_card(parent, "Servicios Básicos", "💡")
        
        self.create_radio_group(s4, "Agua", "water_type",
                               [("Potable", "potable"), ("Cisterna", "cisterna"),
                                ("Vertiente", "vertiente"), ("Repartidor", "repartidor"),
                                ("Sequía", "sequia")], 0, 0)
        
        self.create_radio_group(s4, "Luz", "light_type",
                               [("Permanente", "permanente"), ("Temporal", "temporal"),
                                ("No tiene", "noenergia")], 1, 0)
        
        self.create_radio_group(s4, "Baño (SSHH)", "sshh_type",
                               [("Propio", "propio"), ("Compartido", "compartido"),
                                ("Pozo", "pozo"), ("Libre", "libre")], 2, 0)
        
        self.create_field(s4, "Observaciones", "basic_services_other", 3, 0, 50)
        
        self.create_radio_group(s4, "¿Hacinamiento?", "hacinamiento",
                               [("Sí", "si"), ("No", "no")], 4, 0)
        
        self.create_field(s4, "Manejo de desechos", "waste_management", 5, 0, 50)
        
        self.create_radio_group(s4, "Transporte", "transport_type",
                               [("Público", "publico"), ("Empresa", "empresa"),
                                ("Privado", "privado")], 6, 0)
        self.create_field(s4, "Otro", "transport_type_other", 7, 0, 30)
        
        self.create_field(s4, "Electrodomésticos", "appliances", 8, 0, 50)
        
        self.create_radio_group(s4, "¿Tiene Internet?", "internet",
                               [("Sí", "si"), ("No", "no")], 9, 0)
        self.create_radio_group(s4, "Tipo", "internet_type",
                               [("Recarga", "recarga"), ("Mensual", "mensual"),
                                ("Satelital", "satelital"), ("Fibra Óptica", "fibraoptica")], 10, 0)
        
        # Animales
        tk.Label(s4, text="ANIMALES", font=("Segoe UI", 10, "bold"),
                fg=self.colors["primary"], bg=self.colors["white"]).grid(row=11, column=0, columnspan=4, sticky="w", pady=(15, 10))
        
        self.create_radio_group(s4, "¿Tiene animales?", "animals",
                               [("Sí", "si"), ("No", "no")], 12, 0)
        self.create_field(s4, "Tipo", "animal_type", 13, 0, 25)
        self.create_number_field(s4, "Cantidad", "animal_quantity", 13, 2, 15)
        
        self.create_radio_group(s4, "¿Zona de peste?", "peste",
                               [("Sí", "si"), ("No", "no")], 14, 0)
        self.create_field(s4, "Lugar de tenencia", "animal_location", 15, 0, 30)
        self.create_field(s4, "Observaciones", "animal_observations", 16, 0, 50)

    def build_economia(self, parent):
        s = self.create_card(parent, "Situación Económica", "💰")
        
        self.create_radio_group(s, "¿Gastos compartidos?", "shared_expenses",
                               [("Sí", "si"), ("No", "no")], 0, 0)
        
        # Aportes
        tk.Label(s, text="APORTES ECONÓMICOS MENSUALES", font=("Segoe UI", 10, "bold"),
                fg=self.colors["primary"], bg=self.colors["white"]).grid(row=1, column=0, columnspan=4, sticky="w", pady=(15, 10))
        
        aportes = [("Padre", "father_contribution"), ("Madre", "mother_contribution"),
                  ("Hermanos", "siblings_contribution"), ("Colaborador", "collaborators_contribution"),
                  ("Cónyuge", "spouse_contribution"), ("Hijos", "children_contribution"),
                  ("Otros", "other_contribution")]
        
        for i, (label, key) in enumerate(aportes):
            tk.Label(s, text=label, font=("Segoe UI", 9),
                    bg=self.colors["white"]).grid(row=2+i, column=0, sticky="w", pady=5, padx=5)
            entry = tk.Entry(s, width=15, font=("Segoe UI", 10))
            entry.grid(row=2+i, column=1, pady=5, padx=5)
            
            # Validar solo números
            def validate_number(P):
                if P == "" or P.isdigit() or (P.replace('.', '', 1).isdigit() and P.count('.') <= 1):
                    return True
                return False
            validate_cmd = s.register(validate_number)
            entry.config(validate="key", validatecommand=(validate_cmd, "%P"))
            
            # Bind para actualizar total
            entry.bind("<KeyRelease>", lambda e: self.calculate_total_contributions())
            self.entries[key] = entry
        
        # TOTAL APORTES
        tk.Label(s, text="TOTAL APORTES", font=("Segoe UI", 9, "bold"),
                bg=self.colors["white"], fg=self.colors["primary"]).grid(row=9, column=0, sticky="w", pady=5, padx=5)
        total_contrib_entry = tk.Entry(s, width=20, font=("Segoe UI", 10, "bold"),
                                       fg=self.colors["primary"], relief="solid", bd=1)
        total_contrib_entry.grid(row=9, column=1, pady=5, padx=5, sticky="w")
        total_contrib_entry.config(state="readonly")
        self.entries["total_contribution"] = total_contrib_entry
        
        self.create_field(s, "Monto de deudas", "debt_amount", 10, 0, 20)
        
        # Préstamos
        tk.Label(s, text="PRÉSTAMOS", font=("Segoe UI", 10, "bold"),
                fg=self.colors["primary"], bg=self.colors["white"]).grid(row=11, column=0, columnspan=4, sticky="w", pady=(15, 10))
        
        self.create_radio_group(s, "Formales", "formal_loans",
                               [("Sí", "si"), ("No", "no")], 12, 0)
        self.create_field(s, "Monto", "formal_loans_amount", 13, 0, 20)
        
        self.create_radio_group(s, "Informales", "informal_loans",
                               [("Sí", "si"), ("No", "no")], 14, 0)
        self.create_field(s, "Monto", "informal_loans_amount", 15, 0, 20)
        self.create_field(s, "Familiares", "informal_loans_family_amount", 16, 0, 20)
        self.create_field(s, "Chulqueros", "informal_loans_moneylender_amount", 17, 0, 20)
        self.create_field(s, "Otros", "informal_loans_other_amount", 18, 0, 20)
        
        self.create_radio_group(s, "¿Tarjetas de crédito?", "credit_cards",
                               [("Sí", "si"), ("No", "no")], 19, 0)
        
        # Gastos
        tk.Label(s, text="GASTOS MENSUALES", font=("Segoe UI", 10, "bold"),
                fg=self.colors["primary"], bg=self.colors["white"]).grid(row=20, column=0, columnspan=4, sticky="w", pady=(15, 10))
        
        gastos = [("Alimentación", "food_support"), ("Educación", "education_support"),
                 ("Vivienda", "housing_support"), ("Vestimenta", "clothing_support"),
                 ("Salud", "health_support"), ("Transporte", "transport_support"),
                 ("Servicios básicos", "basic_services_support"), ("Internet", "internet_support"),
                 ("TV Cable", "cable_tv_support"), ("Plan Celular", "cell_plan_support"),
                 ("Préstamos", "loans_support"), ("Préstamos Quirografarios", "unsecured_loans_support"),
                 ("Tarjetas de crédito", "credit_cards_support"), ("Pensión de alimentos", "alimony_support"),
                 ("Locales Comerciales", "commercial_properties_support"),
                 ("Apoyo a terceros", "financial_support_others"), ("Otros", "other_expenses_support")]
        
        for i, (label, key) in enumerate(gastos):
            tk.Label(s, text=label, font=("Segoe UI", 9),
                    bg=self.colors["white"]).grid(row=21+i, column=0, sticky="w", pady=3, padx=5)
            entry = tk.Entry(s, width=15, font=("Segoe UI", 10))
            entry.grid(row=21+i, column=1, pady=3, padx=5)
            
            # Validar solo números
            def validate_number(P):
                if P == "" or P.isdigit() or (P.replace('.', '', 1).isdigit() and P.count('.') <= 1):
                    return True
                return False
            validate_cmd = s.register(validate_number)
            entry.config(validate="key", validatecommand=(validate_cmd, "%P"))
            
            # Bind para actualizar total
            entry.bind("<KeyRelease>", lambda e: self.calculate_total_expenses())
            self.entries[key] = entry
        
        # TOTAL GASTOS
        tk.Label(s, text="TOTAL GASTOS", font=("Segoe UI", 9, "bold"),
                bg=self.colors["white"], fg=self.colors["primary"]).grid(row=38, column=0, sticky="w", pady=5, padx=5)
        total_expenses_entry = tk.Entry(s, width=20, font=("Segoe UI", 10, "bold"), 
                                        fg=self.colors["primary"], relief="solid", bd=1)
        total_expenses_entry.grid(row=38, column=1, pady=5, padx=5, sticky="w")
        total_expenses_entry.config(state="readonly")
        self.entries["total_expenses"] = total_expenses_entry
        
        # TOTAL INGRESOS
        tk.Label(s, text="TOTAL INGRESOS", font=("Segoe UI", 9, "bold"),
                bg=self.colors["white"], fg=self.colors["primary"]).grid(row=39, column=0, sticky="w", pady=5, padx=5)
        total_income_entry = tk.Entry(s, width=20, font=("Segoe UI", 10, "bold"),
                                      fg=self.colors["primary"], relief="solid", bd=1)
        total_income_entry.grid(row=39, column=1, pady=5, padx=5, sticky="w")
        total_income_entry.config(state="readonly")
        self.entries["total_income"] = total_income_entry
        
        # SALDO
        tk.Label(s, text="SALDO", font=("Segoe UI", 9, "bold"),
                bg=self.colors["white"], fg=self.colors["success"]).grid(row=40, column=0, sticky="w", pady=5, padx=5)
        balance_entry = tk.Entry(s, width=20, font=("Segoe UI", 10, "bold"),
                                 fg=self.colors["success"], relief="solid", bd=1)
        balance_entry.grid(row=40, column=1, pady=5, padx=5, sticky="w")
        balance_entry.config(state="readonly")
        self.entries["balance"] = balance_entry
        
        self.create_radio_group(s, "¿Posee vehículo?", "transportation",
                               [("Sí", "si"), ("No", "no")], 41, 0)
        self.create_field(s, "Descripción", "transportation_description", 42, 0, 50)
        self.create_field(s, "Actividad económica adicional", "additional_economic_activity", 43, 0, 50)
        
        self.create_radio_group(s, "¿Crianza de animales?", "animal_breeding",
                               [("Sí", "si"), ("No", "no")], 44, 0)

    def build_salud(self, parent):
        s = self.create_card(parent, "Salud", "❤️")
        
        self.create_radio_group(s, "¿Practica deporte?", "sports",
                               [("Sí", "si"), ("No", "no")], 0, 0)
        self.create_field(s, "¿Cuál?", "sports_description", 1, 0, 30)
        self.create_field(s, "Frecuencia", "sports_frequency", 2, 0, 20)
        
        self.create_field(s, "¿Tiene alguna enfermedad?", "disease", 3, 0, 50)
        
        self.create_radio_group(s, "¿Problemas de salud en la familia?", "family_health_problems",
                               [("Sí", "si"), ("No", "no")], 4, 0)
        self.create_field(s, "¿Cuáles?", "family_health_problems_description", 5, 0, 50)
        
        self.create_radio_group(s, "¿Discapacidad en la familia?", "family_disability",
                               [("Sí", "si"), ("No", "no")], 6, 0)
        self.create_field(s, "Tipo", "family_disability_type", 7, 0, 25)
        self.create_number_field(s, "Porcentaje", "family_disability_percentage", 7, 2, 15)
        self.create_field(s, "Parentezco", "family_disability_relationship", 8, 0, 25)

    def build_laboral(self, parent):
        s = self.create_card(parent, "Situación Laboral", "💼")
        
        campos = [
            ("Antes de ingresar, ¿a qué se dedicaba?", "previous_occupation"),
            ("Funciones que desempeña actualmente", "current_functions"),
            ("¿Su trabajo se relaciona con su formación?", "job_relation"),
            ("Relación con compañeros", "colleague_relationship"),
            ("¿Qué podría mejorar?", "improvement_suggestions"),
            ("¿Los conflictos se resuelven eficazmente?", "conflict_resolution"),
            ("¿Su trabajo es desgastante?", "job_exhaustion"),
            ("¿Siente presión laboral?", "job_pressure"),
            ("¿Le genera estrés?", "job_stress"),
            ("¿Le alcanza el tiempo para su trabajo?", "job_time_management"),
            ("¿Su trabajo es reconocido por su jefe?", "job_manager_recognition"),
            ("¿Se siente reconocido por la empresa?", "job_recognition"),
            ("Proyección y aspiraciones", "job_projection"),
            ("¿Se ha sentido discriminado?", "job_discrimination"),
            ("¿Qué haría para mejorar como trabajador?", "job_improvement"),
            ("Beneficios que podría implementar la empresa", "job_benefits")
        ]
        
        for i, (label, key) in enumerate(campos):
            self.create_field(s, label, key, i, 0, 80)

    def update_family_counter(self):
        """Actualiza el contador de familiares"""
        count = len(self.family_rows)
        self.family_counter_label.config(text=f"📊 Familiares registrados: {count} / {self.MAX_FAMILY_MEMBERS}")
        
        if count >= self.MAX_FAMILY_MEMBERS:
            self.add_family_btn.config(state="disabled", bg=self.colors["secondary"])
            self.add_family_btn.config(text="🔒 Límite alcanzado (máximo 6 familiares)")
        else:
            self.add_family_btn.config(state="normal", bg=self.colors["primary"])
            self.add_family_btn.config(text=f"+ Agregar familiar ({self.MAX_FAMILY_MEMBERS - count} disponibles)")

    def add_family(self):
        """Agrega un miembro a la tabla familiar"""
        if len(self.family_rows) >= self.MAX_FAMILY_MEMBERS:
            messagebox.showwarning(
                "Límite alcanzado", 
                f"No se pueden agregar más de {self.MAX_FAMILY_MEMBERS} familiares.\n"
                f"Si necesita registrar más, considere agruparlos o usar el campo 'Otros'."
            )
            return
        
        frame = tk.Frame(self.family_container, bg=self.colors["white"], 
                        relief="solid", bd=1,
                        highlightbackground=self.colors["card_border"], 
                        highlightthickness=1)
        frame.pack(pady=5, fill="x", padx=5)

        entries = {}
        
        # Campos para familiares
        fields = [
            ("name", "Nombre", 18),
            ("age", "Edad", 6),
            ("relation", "Parentesco", 10),
            ("marital_status", "Estado Civil", 10),
            ("education", "Instrucción", 10),
            ("job", "Ocupación", 12),
            ("work_or_study_place", "Lugar trabajo/estudio", 14),
            ("income", "Ingreso", 8),
            ("phone", "Teléfono", 10)
        ]
        
        for i, (key, label, width) in enumerate(fields):
            tk.Label(frame, text=label, font=("Segoe UI", 8),
                    bg=self.colors["white"], fg=self.colors["secondary"]).grid(row=0, column=i*2, padx=3, pady=5)
            e = tk.Entry(frame, width=width, font=("Segoe UI", 9))
            e.grid(row=0, column=i*2+1, padx=3, pady=5)
            entries[key] = e
        
        def remove_family_member():
            frame.destroy()
            self.family_rows.remove(entries)
            self.update_family_counter()
        
        tk.Button(frame, text="✖", command=remove_family_member,
                 bg=self.colors["danger"], fg="white",
                 font=("Segoe UI", 9, "bold"), width=3,
                 relief="flat").grid(row=0, column=len(fields)*2, padx=5)

        self.family_rows.append(entries)
        self.update_family_counter()

    def calculate_total_contributions(self, event=None):
        """Calcula el total de aportes y actualiza TOTAL APORTES y TOTAL INGRESOS"""
        total = 0
        contribution_keys = ["father_contribution", "mother_contribution", "siblings_contribution", 
                            "collaborators_contribution", "spouse_contribution", "children_contribution", 
                            "other_contribution"]
        
        for key in contribution_keys:
            if key in self.entries:
                try:
                    value_str = self.entries[key].get()
                    value = float(value_str) if value_str else 0
                    total += value
                except ValueError:
                    pass
        
        # Actualizar TOTAL APORTES
        if "total_contribution" in self.entries:
            self.entries["total_contribution"].config(state="normal")
            self.entries["total_contribution"].delete(0, tk.END)
            self.entries["total_contribution"].insert(0, f"{total:.2f}")
            self.entries["total_contribution"].config(state="readonly")
        
        # Actualizar TOTAL INGRESOS (igual al total de aportes)
        if "total_income" in self.entries:
            self.entries["total_income"].config(state="normal")
            self.entries["total_income"].delete(0, tk.END)
            self.entries["total_income"].insert(0, f"{total:.2f}")
            self.entries["total_income"].config(state="readonly")
        
        # Actualizar balance
        self.calculate_balance()

    def calculate_total_expenses(self, event=None):
        """Calcula el total de gastos y actualiza TOTAL GASTOS"""
        total = 0
        expense_keys = ["food_support", "education_support", "housing_support", "clothing_support",
                       "health_support", "transport_support", "basic_services_support", "internet_support",
                       "cable_tv_support", "cell_plan_support", "loans_support", "unsecured_loans_support",
                       "credit_cards_support", "alimony_support", "commercial_properties_support",
                       "financial_support_others", "other_expenses_support"]
        
        for key in expense_keys:
            if key in self.entries:
                try:
                    value_str = self.entries[key].get()
                    value = float(value_str) if value_str else 0
                    total += value
                except ValueError:
                    pass
        
        # Actualizar TOTAL GASTOS
        if "total_expenses" in self.entries:
            self.entries["total_expenses"].config(state="normal")
            self.entries["total_expenses"].delete(0, tk.END)
            self.entries["total_expenses"].insert(0, f"{total:.2f}")
            self.entries["total_expenses"].config(state="readonly")
        
        # Actualizar balance
        self.calculate_balance()

    def calculate_balance(self, event=None):
        """Calcula el balance (ingresos - egresos)"""
        try:
            # Obtener ingresos
            income = 0
            if "total_income" in self.entries:
                income_str = self.entries["total_income"].get()
                income = float(income_str) if income_str else 0
            
            # Obtener egresos
            expenses = 0
            if "total_expenses" in self.entries:
                expenses_str = self.entries["total_expenses"].get()
                expenses = float(expenses_str) if expenses_str else 0
            
            balance = income - expenses
            
            # Actualizar SALDO
            if "balance" in self.entries:
                self.entries["balance"].config(state="normal")
                self.entries["balance"].delete(0, tk.END)
                self.entries["balance"].insert(0, f"{balance:.2f}")
                # Cambiar color según si es positivo o negativo
                if balance >= 0:
                    self.entries["balance"].config(fg=self.colors["success"])
                else:
                    self.entries["balance"].config(fg=self.colors["danger"])
                self.entries["balance"].config(state="readonly")
        except:
            pass

    def load_photo(self):
        path = filedialog.askopenfilename(filetypes=[("Imagen", "*.jpg *.jpeg *.png")])
        if path:
            self.photo_path.set(path)

    def get_data(self):
        data = {}
        
        for key, entry in self.entries.items():
            if hasattr(entry, 'get'):
                value = entry.get()
                # Convertir valores numéricos
                if key in ["total_income", "total_expenses", "balance", "total_contribution"]:
                    try:
                        value = float(value) if value else 0
                    except:
                        value = 0
                data[key] = value
            else:
                data[key] = entry
        
        # Distribución de vivienda (ahora con cantidades numéricas)
        distribution = {}
        dist_keys = ["dormitorios", "camas", "cocinas", "comedores", "salas", 
                    "banos", "patios", "jardines", "terrazas", "garajes"]
        
        for key in dist_keys:
            entry_key = f"house_distribution_{key}"
            if entry_key in self.entries:
                try:
                    value = int(self.entries[entry_key].get() or 0)
                    if value > 0:
                        distribution[key] = value
                except ValueError:
                    pass
        
        data["house_distribution"] = distribution
        
        data["visit_date"] = datetime.datetime.now().strftime("%d/%m/%Y")
        data["photo"] = self.photo_path.get()
        
        # Recolectar familiares
        family = []
        skipped_count = 0
        
        for member in self.family_rows:
            row_data = {}
            is_complete = True
            
            for k, e in member.items():
                value = e.get().strip()
                row_data[k] = value
                if value == "":
                    is_complete = False
            
            if is_complete and any(row_data.values()):
                # Convertir ingresos a número
                if "income" in row_data and row_data["income"]:
                    try:
                        row_data["income"] = float(row_data["income"])
                    except:
                        pass
                family.append(row_data)
            elif any(row_data.values()):
                skipped_count += 1
        
        if skipped_count > 0:
            messagebox.showwarning(
                "Filas omitidas", 
                f"⚠️ Se omitieron {skipped_count} fila(s) de familiares porque estaban incompletas.\n\n"
                "Para que un familiar se guarde correctamente, debe completar todos los campos.\n"
                "Las filas vacías o incompletas no se incluirán en el PDF."
            )
        
        data["family_members"] = family
    
        return data

    def save_all(self):
        try:
            data = self.get_data()
            
            if not data.get("name") or not data.get("ci"):
                messagebox.showwarning("Campos incompletos", 
                                      "Por favor complete: Nombre y Cédula")
                return
            
            pdf.create_pdf(data)
            xlsx.create_excel_file(data)
            
            messagebox.showinfo("Éxito", 
                               "✅ Datos guardados correctamente!\n\n"
                               "📄 PDF generado en carpeta 'pdf/'\n"
                               "📊 Excel generado en carpeta 'excel/'")
            
        except Exception as e:
            messagebox.showerror("Error", f"❌ Error al guardar: {str(e)}")