import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime

from services import pdf_service as pdf
from services import excel_service as xlsx


class App:

    def __init__(self, root):

        self.root = root
        self.root.title("Sistema Socioeconómico")
        self.root.geometry("1400x850")
        self.root.configure(bg="#f4f6f9")

        self.entries = {}
        self.family_rows = {}
        self.photo_path = tk.StringVar()

        # =========================
        # LAYOUT PRINCIPAL
        # =========================
        self.sidebar = tk.Frame(root, bg="#2c3e50", width=220)
        self.sidebar.pack(side="left", fill="y")

        self.container = tk.Frame(root, bg="#ecf0f1")
        self.container.pack(side="right", fill="both", expand=True)

        # =========================
        # SIDEBAR
        # =========================
        tk.Label(self.sidebar,
                 text="Sistema Socioeconómico",
                 fg="white",
                 bg="#2c3e50",
                 font=("Segoe UI", 12, "bold")).pack(pady=20)

        self.frames = {}

        sections = [
            ("Personales", self.build_personales),
            ("Familia", self.build_familia),
            ("Vivienda", self.build_vivienda),
            ("Economía", self.build_economia),
            ("Salud", self.build_salud),
            ("Laboral", self.build_laboral),
        ]

        for name, builder in sections:
            btn = tk.Button(self.sidebar,
                            text=name,
                            fg="white",
                            bg="#34495e",
                            relief="flat",
                            font=("Segoe UI", 10),
                            command=lambda n=name: self.show_frame(n))
            btn.pack(fill="x", pady=2)

            frame = tk.Frame(self.container, bg="#ecf0f1")
            self.frames[name] = frame
            
            # Crear canvas con scrollbar para cada frame
            canvas = tk.Canvas(frame, bg="#ecf0f1")
            scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg="#ecf0f1")
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e, c=canvas: c.configure(scrollregion=c.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            builder(scrollable_frame)
            setattr(self, f"scrollable_{name}", scrollable_frame)

        self.show_frame("Personales")

        # BOTÓN GLOBAL
        tk.Button(self.sidebar,
                  text="💾 Guardar PDF y Excel",
                  command=self.save_all,
                  bg="#27ae60",
                  fg="white",
                  font=("Segoe UI", 11, "bold"),
                  bd=0).pack(side="bottom", fill="x", pady=10)

    # =========================
    # NAV
    # =========================
    def show_frame(self, name):
        for frame in self.frames.values():
            frame.pack_forget()
        self.frames[name].pack(fill="both", expand=True)

    # =========================
    # UI HELPERS
    # =========================
    def section(self, parent, title):
        frame = tk.Frame(parent, bg="white", bd=1, relief="solid")
        frame.pack(fill="x", padx=20, pady=10)

        tk.Label(frame, text=title,
                 font=("Segoe UI", 11, "bold"),
                 bg="white").pack(anchor="w", padx=10, pady=5)

        content = tk.Frame(frame, bg="white")
        content.pack(fill="x", padx=10, pady=5)

        return content

    def create_entry(self, parent, label, key, row, column=0, width=35):
        tk.Label(parent, text=label, bg="white").grid(row=row, column=column, sticky="w", pady=5, padx=5)
        entry = tk.Entry(parent, width=width)
        entry.grid(row=row, column=column+1, pady=5, padx=5)
        self.entries[key] = entry

    def create_combo(self, parent, label, key, values, row, column=0, width=33):
        tk.Label(parent, text=label, bg="white").grid(row=row, column=column, sticky="w", pady=5, padx=5)
        combo = ttk.Combobox(parent, values=values, width=width)
        combo.grid(row=row, column=column+1, pady=5, padx=5)
        self.entries[key] = combo

    def create_checkbox(self, parent, key, text, row, column, var_name=None):
        if var_name is None:
            var = tk.StringVar(value="no")
            self.entries[key] = var
        else:
            var = var_name
        cb = tk.Checkbutton(parent, text=text, variable=var, 
                           onvalue="si", offvalue="no", bg="white")
        cb.grid(row=row, column=column, pady=2, padx=10, sticky="w")
        return var

    # =========================
    # SECCIONES COMPLETAS
    # =========================
    def build_personales(self, parent):
        # 1. DATOS PERSONALES DEL COLABORADOR
        s = self.section(parent, "1. DATOS PERSONALES DEL COLABORADOR")
        
        self.create_entry(s, "Fecha de Visita", "visit_date", 0)
        self.create_entry(s, "Nombre", "name", 1)
        self.create_entry(s, "Cédula", "ci", 2)
        self.create_entry(s, "Edad", "age", 3)
        self.create_entry(s, "Lugar y Fecha de Nacimiento", "birth_date_place", 4)
        
        # Género
        tk.Label(s, text="Género", bg="white").grid(row=5, column=0, sticky="w", pady=5, padx=5)
        gender_frame = tk.Frame(s, bg="white")
        gender_frame.grid(row=5, column=1, sticky="w", pady=5, padx=5)
        self.entries["gender"] = tk.StringVar(value="masculino")
        tk.Radiobutton(gender_frame, text="Masculino", variable=self.entries["gender"], 
                      value="masculino", bg="white").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(gender_frame, text="Femenino", variable=self.entries["gender"], 
                      value="femenino", bg="white").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(gender_frame, text="Otro", variable=self.entries["gender"], 
                      value="otro", bg="white").pack(side=tk.LEFT, padx=5)
        
        # Estado Civil
        tk.Label(s, text="Estado Civil", bg="white").grid(row=6, column=0, sticky="w", pady=5, padx=5)
        civil_frame = tk.Frame(s, bg="white")
        civil_frame.grid(row=6, column=1, sticky="w", pady=5, padx=5)
        self.entries["marital_status"] = tk.StringVar(value="soltero")
        opciones = ["soltero", "casado", "separado", "divorciado", "unión libre", "viudo"]
        for i, opt in enumerate(opciones):
            tk.Radiobutton(civil_frame, text=opt.capitalize(), variable=self.entries["marital_status"], 
                          value=opt, bg="white").pack(side=tk.LEFT, padx=5)
        
        self.create_entry(s, "Email", "email", 7)
        self.create_entry(s, "Área de trabajo", "work_area", 8)
        self.create_entry(s, "Puesto que desempeña", "job_title", 9)
        self.create_entry(s, "Fecha de Ingreso", "entry_date", 10)
        
        # Discapacidad
        tk.Label(s, text="Tiene alguna discapacidad", bg="white").grid(row=11, column=0, sticky="w", pady=5, padx=5)
        self.entries["disability"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["disability"], value="si", bg="white").grid(row=11, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["disability"], value="no", bg="white").grid(row=11, column=2, sticky="w")
        
        self.create_entry(s, "Tipo de discapacidad", "disability_type", 12)
        self.create_entry(s, "Porcentaje de discapacidad", "disability_percentage", 13)
        
        # Nivel de Educación
        tk.Label(s, text="Nivel de Educación", bg="white").grid(row=14, column=0, sticky="w", pady=5, padx=5)
        edu_frame = tk.Frame(s, bg="white")
        edu_frame.grid(row=14, column=1, columnspan=3, sticky="w", pady=5, padx=5)
        self.entries["educational_level"] = tk.StringVar(value="bachiller")
        niveles = ["profesional", "bachiller", "basica", "primaria", "nuloestudio"]
        for i, niv in enumerate(niveles):
            tk.Radiobutton(edu_frame, text=niv.capitalize(), variable=self.entries["educational_level"], 
                          value=niv, bg="white").pack(side=tk.LEFT, padx=5)
        
        self.create_entry(s, "Observaciones de educación", "educational_observations", 15)
        
        # Actualmente estudia
        tk.Label(s, text="¿Actualmente estudia?", bg="white").grid(row=16, column=0, sticky="w", pady=5, padx=5)
        self.entries["currently_studying"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["currently_studying"], value="si", bg="white").grid(row=16, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["currently_studying"], value="no", bg="white").grid(row=16, column=2, sticky="w")
        
        self.create_entry(s, "Carrera que estudia", "study_career", 17)
        self.create_entry(s, "Nivel de estudio", "study_level", 18)
        
        self.create_entry(s, "Dirección Domiciliaria", "address", 19)
        self.create_entry(s, "Teléfono", "phone", 20)
        self.create_entry(s, "Contacto de emergencia", "emergency_contact", 21)
        
        # Foto
        s2 = self.section(parent, "Foto del Colaborador")
        tk.Entry(s2, textvariable=self.photo_path, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(s2, text="Buscar Foto", command=self.load_photo).pack(side=tk.LEFT, padx=5)

    def build_familia(self, parent):
        # 2. MIEMBROS DEL HOGAR
        s = self.section(parent, "2. MIEMBROS DEL HOGAR")
        
        tk.Label(s, text="Composición Familiar (viven dentro del hogar)", 
                font=("Segoe UI", 10, "bold"), bg="white").grid(row=0, column=0, columnspan=6, pady=10)
        
        # Headers
        headers = ["Nombre", "Edad", "Parentesco", "Ocupación", "Ingreso"]
        for i, h in enumerate(headers):
            tk.Label(s, text=h, font=("Segoe UI", 9, "bold"), bg="white", 
                    relief="ridge", width=12).grid(row=1, column=i, padx=1, pady=2)
        
        self.family_frame = tk.Frame(s, bg="white")
        self.family_frame.grid(row=2, column=0, columnspan=6, pady=5)
        self.family_rows = []
        
        tk.Button(s, text="+ Agregar Familiar", command=self.add_family,
                 bg="#3498db", fg="white").grid(row=3, column=0, pady=10)
        
        # Información de la Familia
        s2 = self.section(parent, "Información de la Familia")
        
        # Tipo de Familia
        tk.Label(s2, text="Tipo de Familia", bg="white", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        family_frame = tk.Frame(s2, bg="white")
        family_frame.grid(row=0, column=1, sticky="w", pady=5)
        self.entries["family_type"] = tk.StringVar(value="nuclear")
        tipos = ["nuclear", "cosanguinea", "afines", "ampliada", "materna", "paterna", "unipersonal"]
        for i, tip in enumerate(tipos):
            tk.Radiobutton(family_frame, text=tip.capitalize(), variable=self.entries["family_type"], 
                          value=tip, bg="white").pack(side=tk.LEFT, padx=5)
        
        self.create_entry(s2, "N° Hijos del colaborador", "children_count", 1)
        
        # Hijos fuera del matrimonio
        tk.Label(s2, text="Existen hijos fuera del matrimonio", bg="white").grid(row=2, column=0, sticky="w", pady=5)
        self.entries["children_outside_marriage"] = tk.StringVar(value="no")
        tk.Radiobutton(s2, text="Sí", variable=self.entries["children_outside_marriage"], value="si", bg="white").grid(row=2, column=1, sticky="w")
        tk.Radiobutton(s2, text="No", variable=self.entries["children_outside_marriage"], value="no", bg="white").grid(row=2, column=2, sticky="w")
        
        self.create_entry(s2, "Paga pensión alimenticia", "pays_alimony", 3)
        self.create_entry(s2, "N° de Matrimonio o Relación Actual", "marriage_number", 4)
        self.create_entry(s2, "Comparte tiempo de descanso con su familia", "family_time", 5)
        self.create_entry(s2, "Frecuencia", "family_time_frequency", 6)
        self.create_entry(s2, "Actividades familiares", "family_activities", 7)
        self.create_entry(s2, "Actividades fuera de jornada laboral", "other_activities", 8)
        
        # Otras actividades en domicilio
        tk.Label(s2, text="En su domicilio realiza otras actividades", bg="white").grid(row=9, column=0, sticky="w", pady=5)
        self.entries["other_activities"] = tk.StringVar(value="no")
        tk.Radiobutton(s2, text="Sí", variable=self.entries["other_activities"], value="si", bg="white").grid(row=9, column=1, sticky="w")
        tk.Radiobutton(s2, text="No", variable=self.entries["other_activities"], value="no", bg="white").grid(row=9, column=2, sticky="w")
        
        self.create_entry(s2, "Especificar otras actividades", "other_activities_specify", 10)
        self.create_entry(s2, "Hobbies", "hobbies", 11)
        self.create_entry(s2, "Tiempo dedicado a hobbies", "hobbies_time", 12)
        
        # Relación de pareja
        tk.Label(s2, text="Cómo califica la relación de pareja", bg="white").grid(row=13, column=0, sticky="w", pady=5)
        self.entries["partner_relationship"] = tk.StringVar(value="buena")
        rel_frame = tk.Frame(s2, bg="white")
        rel_frame.grid(row=13, column=1, columnspan=3, sticky="w")
        for i, val in enumerate(["muy buena", "buena", "regular", "mala"]):
            tk.Radiobutton(rel_frame, text=val.capitalize(), variable=self.entries["partner_relationship"], 
                          value=val, bg="white").pack(side=tk.LEFT, padx=5)
        self.create_entry(s2, "¿Por qué?", "partner_relationship_reason", 14)
        
        # Relación con hijos
        tk.Label(s2, text="Cómo califica la relación con los hijos", bg="white").grid(row=15, column=0, sticky="w", pady=5)
        self.entries["children_relationship"] = tk.StringVar(value="buena")
        rel2_frame = tk.Frame(s2, bg="white")
        rel2_frame.grid(row=15, column=1, columnspan=3, sticky="w")
        for i, val in enumerate(["muy buena", "buena", "regular", "mala"]):
            tk.Radiobutton(rel2_frame, text=val.capitalize(), variable=self.entries["children_relationship"], 
                          value=val, bg="white").pack(side=tk.LEFT, padx=5)
        self.create_entry(s2, "¿Por qué?", "children_relationship_reason", 16)
        
        # Problemas familiares
        tk.Label(s2, text="Problemas familiares", bg="white").grid(row=17, column=0, sticky="w", pady=5)
        prob_frame = tk.Frame(s2, bg="white")
        prob_frame.grid(row=17, column=1, columnspan=3, sticky="w")
        self.entries["family_problems"] = tk.StringVar(value="ningunos")
        problemas = ["divorcios", "separaciones", "resentimientos", "otros", "ningunos"]
        for i, prob in enumerate(problemas):
            tk.Radiobutton(prob_frame, text=prob.capitalize(), variable=self.entries["family_problems"], 
                          value=prob, bg="white").pack(side=tk.LEFT, padx=5)
        
        self.create_entry(s2, "Recibió ayuda", "family_problems_help", 18)
        
        # Migración
        tk.Label(s2, text="Miembros del hogar que salieron del país", bg="white").grid(row=19, column=0, sticky="w", pady=5)
        self.entries["family_migration"] = tk.StringVar(value="no")
        tk.Radiobutton(s2, text="Sí", variable=self.entries["family_migration"], value="si", bg="white").grid(row=19, column=1, sticky="w")
        tk.Radiobutton(s2, text="No", variable=self.entries["family_migration"], value="no", bg="white").grid(row=19, column=2, sticky="w")
        
        tk.Label(s2, text="Recibió dinero del exterior", bg="white").grid(row=20, column=0, sticky="w", pady=5)
        self.entries["family_migration_received"] = tk.StringVar(value="no")
        tk.Radiobutton(s2, text="Sí", variable=self.entries["family_migration_received"], value="si", bg="white").grid(row=20, column=1, sticky="w")
        tk.Radiobutton(s2, text="No", variable=self.entries["family_migration_received"], value="no", bg="white").grid(row=20, column=2, sticky="w")

    def build_vivienda(self, parent):
        s = self.section(parent, "3. VIVIENDA")
        
        # Sector
        tk.Label(s, text="Sector", bg="white").grid(row=0, column=0, sticky="w", pady=5)
        self.entries["sector"] = tk.StringVar(value="urbano")
        tk.Radiobutton(s, text="Urbano", variable=self.entries["sector"], value="urbano", bg="white").grid(row=0, column=1, sticky="w")
        tk.Radiobutton(s, text="Rural", variable=self.entries["sector"], value="rural", bg="white").grid(row=0, column=2, sticky="w")
        
        # Tenencia
        tk.Label(s, text="Tenencia de la vivienda", bg="white").grid(row=1, column=0, sticky="w", pady=5)
        self.entries["house_ownership"] = tk.StringVar(value="propia")
        tenencia_frame = tk.Frame(s, bg="white")
        tenencia_frame.grid(row=1, column=1, columnspan=4, sticky="w")
        tenencias = ["propia", "arrendada", "hipotecada", "prestada", "otros"]
        for i, ten in enumerate(tenencias):
            tk.Radiobutton(tenencia_frame, text=ten.capitalize(), variable=self.entries["house_ownership"], 
                          value=ten, bg="white").pack(side=tk.LEFT, padx=5)
        
        self.create_entry(s, "N° habitantes", "household_size", 2)
        self.create_entry(s, "Tiempo que vive en el sector", "time_living_sector", 3)
        
        tk.Label(s, text="Se considera seguro", bg="white").grid(row=4, column=0, sticky="w", pady=5)
        self.entries["is_safe"] = tk.StringVar(value="si")
        tk.Radiobutton(s, text="Sí", variable=self.entries["is_safe"], value="si", bg="white").grid(row=4, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["is_safe"], value="no", bg="white").grid(row=4, column=2, sticky="w")
        
        # Tipo de vivienda
        tk.Label(s, text="Tipo de vivienda", bg="white").grid(row=5, column=0, sticky="w", pady=5)
        self.entries["house_type"] = tk.StringVar(value="casa")
        tipo_frame = tk.Frame(s, bg="white")
        tipo_frame.grid(row=5, column=1, columnspan=3, sticky="w")
        tipos = ["casa", "departamento", "cuarto", "mediagua"]
        for i, tip in enumerate(tipos):
            tk.Radiobutton(tipo_frame, text=tip.capitalize(), variable=self.entries["house_type"], 
                          value=tip, bg="white").pack(side=tk.LEFT, padx=5)
        
        # Clase de vivienda
        tk.Label(s, text="Clase de vivienda", bg="white").grid(row=6, column=0, sticky="w", pady=5)
        self.entries["house_class"] = tk.StringVar(value="completa")
        clase_frame = tk.Frame(s, bg="white")
        clase_frame.grid(row=6, column=1, columnspan=3, sticky="w")
        clases = ["precaria", "elemental", "completa"]
        for i, cls in enumerate(clases):
            tk.Radiobutton(clase_frame, text=cls.capitalize(), variable=self.entries["house_class"], 
                          value=cls, bg="white").pack(side=tk.LEFT, padx=5)
        
        self.create_entry(s, "Avalúo de casa", "house_valuation", 7)
        self.create_entry(s, "Observaciones de vivienda", "house_observations", 8)
        
        # Distribución
        tk.Label(s, text="Distribución de la vivienda", bg="white", font=("Segoe UI", 9, "bold")).grid(row=9, column=0, columnspan=6, pady=10)
        dist_frame = tk.Frame(s, bg="white")
        dist_frame.grid(row=10, column=0, columnspan=6)
        self.entries["house_distribution"] = []
        distribuciones = ["dormitorio", "camas", "cocina", "comedor", "sala", "bano", "patio", "jardin", "terraza", "garaje"]
        for i, dist in enumerate(distribuciones):
            var = tk.StringVar(value="no")
            self.entries[f"house_distribution_{dist}"] = var
            cb = tk.Checkbutton(dist_frame, text=dist.capitalize(), variable=var, 
                               onvalue=dist, offvalue="no", bg="white")
            cb.grid(row=i//5, column=i%5, sticky="w", padx=10, pady=2)
        
        # Techo
        tk.Label(s, text="Techo o cubierta", bg="white", font=("Segoe UI", 9, "bold")).grid(row=11, column=0, columnspan=6, pady=10)
        techo_frame = tk.Frame(s, bg="white")
        techo_frame.grid(row=12, column=0, columnspan=6)
        self.entries["roof_type"] = tk.StringVar(value="eternit")
        techos = ["loza", "eternit", "zinc", "teja", "plastico"]
        for i, tech in enumerate(techos):
            tk.Radiobutton(techo_frame, text=tech.capitalize(), variable=self.entries["roof_type"], 
                          value=tech, bg="white").grid(row=0, column=i, padx=10)
        self.create_entry(s, "Otros materiales de techo", "roof_type_other", 13)
        
        tk.Label(s, text="Estado del techo", bg="white").grid(row=14, column=0, sticky="w", pady=5)
        self.entries["roof_status"] = tk.StringVar(value="bueno")
        tk.Radiobutton(s, text="Bueno", variable=self.entries["roof_status"], value="bueno", bg="white").grid(row=14, column=1, sticky="w")
        tk.Radiobutton(s, text="Regular", variable=self.entries["roof_status"], value="regular", bg="white").grid(row=14, column=2, sticky="w")
        tk.Radiobutton(s, text="Malo", variable=self.entries["roof_status"], value="malo", bg="white").grid(row=14, column=3, sticky="w")
        
        # Pared
        tk.Label(s, text="Tipo de pared", bg="white", font=("Segoe UI", 9, "bold")).grid(row=15, column=0, columnspan=6, pady=10)
        pared_frame = tk.Frame(s, bg="white")
        pared_frame.grid(row=16, column=0, columnspan=6)
        self.entries["wall_type"] = tk.StringVar(value="bloque")
        paredes = ["hormigon", "bloque", "adobe", "madera"]
        for i, par in enumerate(paredes):
            tk.Radiobutton(pared_frame, text=par.capitalize(), variable=self.entries["wall_type"], 
                          value=par, bg="white").grid(row=0, column=i, padx=10)
        self.create_entry(s, "Otros materiales de pared", "wall_type_other", 17)
        
        tk.Label(s, text="Estado de la pared", bg="white").grid(row=18, column=0, sticky="w", pady=5)
        self.entries["wall_status"] = tk.StringVar(value="bueno")
        tk.Radiobutton(s, text="Bueno", variable=self.entries["wall_status"], value="bueno", bg="white").grid(row=18, column=1, sticky="w")
        tk.Radiobutton(s, text="Regular", variable=self.entries["wall_status"], value="regular", bg="white").grid(row=18, column=2, sticky="w")
        tk.Radiobutton(s, text="Malo", variable=self.entries["wall_status"], value="malo", bg="white").grid(row=18, column=3, sticky="w")
        
        # Piso
        tk.Label(s, text="Tipo de piso", bg="white", font=("Segoe UI", 9, "bold")).grid(row=19, column=0, columnspan=6, pady=10)
        piso_frame = tk.Frame(s, bg="white")
        piso_frame.grid(row=20, column=0, columnspan=6)
        self.entries["floor_type"] = tk.StringVar(value="cemento")
        pisos = ["entablado", "baldosa", "cemento", "tierra"]
        for i, pis in enumerate(pisos):
            tk.Radiobutton(piso_frame, text=pis.capitalize(), variable=self.entries["floor_type"], 
                          value=pis, bg="white").grid(row=0, column=i, padx=10)
        self.create_entry(s, "Otros materiales de piso", "floor_type_other", 21)
        
        tk.Label(s, text="Estado del piso", bg="white").grid(row=22, column=0, sticky="w", pady=5)
        self.entries["floor_status"] = tk.StringVar(value="bueno")
        tk.Radiobutton(s, text="Bueno", variable=self.entries["floor_status"], value="bueno", bg="white").grid(row=22, column=1, sticky="w")
        tk.Radiobutton(s, text="Regular", variable=self.entries["floor_status"], value="regular", bg="white").grid(row=22, column=2, sticky="w")
        tk.Radiobutton(s, text="Malo", variable=self.entries["floor_status"], value="malo", bg="white").grid(row=22, column=3, sticky="w")
        
        # Estructura
        tk.Label(s, text="Tipo de estructura", bg="white", font=("Segoe UI", 9, "bold")).grid(row=23, column=0, columnspan=6, pady=10)
        struct_frame = tk.Frame(s, bg="white")
        struct_frame.grid(row=24, column=0, columnspan=6)
        self.entries["structure_type"] = tk.StringVar(value="hormigon")
        estructuras = ["hormigon", "hierro", "madera", "bloque"]
        for i, est in enumerate(estructuras):
            tk.Radiobutton(struct_frame, text=est.capitalize(), variable=self.entries["structure_type"], 
                          value=est, bg="white").grid(row=0, column=i, padx=10)
        self.create_entry(s, "Otros materiales de estructura", "structure_type_other", 25)
        
        tk.Label(s, text="Estado de la estructura", bg="white").grid(row=26, column=0, sticky="w", pady=5)
        self.entries["structure_status"] = tk.StringVar(value="bueno")
        tk.Radiobutton(s, text="Bueno", variable=self.entries["structure_status"], value="bueno", bg="white").grid(row=26, column=1, sticky="w")
        tk.Radiobutton(s, text="Regular", variable=self.entries["structure_status"], value="regular", bg="white").grid(row=26, column=2, sticky="w")
        tk.Radiobutton(s, text="Malo", variable=self.entries["structure_status"], value="malo", bg="white").grid(row=26, column=3, sticky="w")
        
        # Servicios Básicos
        tk.Label(s, text="Servicios Básicos", bg="white", font=("Segoe UI", 11, "bold")).grid(row=27, column=0, columnspan=6, pady=10)
        
        # Agua
        tk.Label(s, text="Agua", bg="white", font=("Segoe UI", 9, "bold")).grid(row=28, column=0, sticky="w", pady=5)
        self.entries["water_type"] = tk.StringVar(value="potable")
        agua_frame = tk.Frame(s, bg="white")
        agua_frame.grid(row=28, column=1, columnspan=5, sticky="w")
        aguas = ["potable", "cisterna", "vertiente", "repartidor", "sequia"]
        for i, ag in enumerate(aguas):
            tk.Radiobutton(agua_frame, text=ag.capitalize(), variable=self.entries["water_type"], 
                          value=ag, bg="white").pack(side=tk.LEFT, padx=5)
        
        # Luz
        tk.Label(s, text="Luz", bg="white", font=("Segoe UI", 9, "bold")).grid(row=29, column=0, sticky="w", pady=5)
        self.entries["light_type"] = tk.StringVar(value="permanente")
        luz_frame = tk.Frame(s, bg="white")
        luz_frame.grid(row=29, column=1, columnspan=5, sticky="w")
        luces = ["permanente", "temporal", "noenergia"]
        for i, luz in enumerate(luces):
            tk.Radiobutton(luz_frame, text=luz.capitalize(), variable=self.entries["light_type"], 
                          value=luz, bg="white").pack(side=tk.LEFT, padx=5)
        
        # SSHH
        tk.Label(s, text="SSHH (Baño)", bg="white", font=("Segoe UI", 9, "bold")).grid(row=30, column=0, sticky="w", pady=5)
        self.entries["sshh_type"] = tk.StringVar(value="propio")
        sshh_frame = tk.Frame(s, bg="white")
        sshh_frame.grid(row=30, column=1, columnspan=5, sticky="w")
        sshhs = ["propio", "compartido", "pozo", "libre"]
        for i, ssh in enumerate(sshhs):
            tk.Radiobutton(sshh_frame, text=ssh.capitalize(), variable=self.entries["sshh_type"], 
                          value=ssh, bg="white").pack(side=tk.LEFT, padx=5)
        
        self.create_entry(s, "Observaciones servicios básicos", "basic_services_other", 31)
        
        # Hacinamiento
        tk.Label(s, text="Hacinamiento", bg="white").grid(row=32, column=0, sticky="w", pady=5)
        self.entries["hacinamiento"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["hacinamiento"], value="si", bg="white").grid(row=32, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["hacinamiento"], value="no", bg="white").grid(row=32, column=2, sticky="w")
        
        self.create_entry(s, "Manejo de desechos", "waste_management", 33)
        
        # Transporte
        tk.Label(s, text="Transporte", bg="white").grid(row=34, column=0, sticky="w", pady=5)
        self.entries["transport_type"] = tk.StringVar(value="publico")
        trans_frame = tk.Frame(s, bg="white")
        trans_frame.grid(row=34, column=1, columnspan=5, sticky="w")
        transports = ["publico", "empresa", "privado"]
        for i, tr in enumerate(transports):
            tk.Radiobutton(trans_frame, text=tr.capitalize(), variable=self.entries["transport_type"], 
                          value=tr, bg="white").pack(side=tk.LEFT, padx=5)
        self.create_entry(s, "Otro transporte", "transport_type_other", 35)
        
        self.create_entry(s, "Electrodomésticos que posee", "appliances", 36)
        
        # Internet
        tk.Label(s, text="Dispone de servicio de Internet", bg="white").grid(row=37, column=0, sticky="w", pady=5)
        self.entries["internet"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["internet"], value="si", bg="white").grid(row=37, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["internet"], value="no", bg="white").grid(row=37, column=2, sticky="w")
        
        tk.Label(s, text="Tipo de Internet", bg="white").grid(row=38, column=0, sticky="w", pady=5)
        self.entries["internet_type"] = tk.StringVar(value="recarga")
        net_frame = tk.Frame(s, bg="white")
        net_frame.grid(row=38, column=1, columnspan=5, sticky="w")
        nets = ["recarga", "mensual", "satelital", "fibraoptica"]
        for i, net in enumerate(nets):
            tk.Radiobutton(net_frame, text=net.capitalize(), variable=self.entries["internet_type"], 
                          value=net, bg="white").pack(side=tk.LEFT, padx=5)
        
        # Animales
        tk.Label(s, text="Tiene animales", bg="white").grid(row=39, column=0, sticky="w", pady=5)
        self.entries["animals"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["animals"], value="si", bg="white").grid(row=39, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["animals"], value="no", bg="white").grid(row=39, column=2, sticky="w")
        
        self.create_entry(s, "Tipo de animales", "animal_type", 40)
        self.create_entry(s, "Cantidad de animales", "animal_quantity", 41)
        
        tk.Label(s, text="Zona de peste", bg="white").grid(row=42, column=0, sticky="w", pady=5)
        self.entries["peste"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["peste"], value="si", bg="white").grid(row=42, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["peste"], value="no", bg="white").grid(row=42, column=2, sticky="w")
        
        self.create_entry(s, "Lugar de tenencia de animales", "animal_location", 43)
        self.create_entry(s, "Observaciones de animales", "animal_observations", 44)

    def build_economia(self, parent):
        s = self.section(parent, "4. SITUACIÓN ECONÓMICA")
        
        # Gastos compartidos
        tk.Label(s, text="Los gastos en el hogar son compartidos", bg="white").grid(row=0, column=0, sticky="w", pady=5)
        self.entries["shared_expenses"] = tk.StringVar(value="si")
        tk.Radiobutton(s, text="Sí", variable=self.entries["shared_expenses"], value="si", bg="white").grid(row=0, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["shared_expenses"], value="no", bg="white").grid(row=0, column=2, sticky="w")
        
        # Aportes económicos
        tk.Label(s, text="Aportes económicos mensuales", bg="white", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, columnspan=3, pady=10)
        
        aportes = [
            ("Padre", "father_contribution"),
            ("Madre", "mother_contribution"),
            ("Hermanos", "siblings_contribution"),
            ("Colaborador", "collaborators_contribution"),
            ("Cónyuge", "spouse_contribution"),
            ("Hijos", "children_contribution"),
            ("Otros", "other_contribution")
        ]
        
        for i, (label, key) in enumerate(aportes):
            tk.Label(s, text=label, bg="white").grid(row=2+i, column=0, sticky="w", pady=2, padx=5)
            entry = tk.Entry(s, width=15)
            entry.grid(row=2+i, column=1, pady=2, padx=5)
            self.entries[key] = entry
        
        self.create_entry(s, "Total aportes", "total_contribution", 9)
        self.create_entry(s, "Monto de deudas", "debt_amount", 10)
        
        # Préstamos formales
        tk.Label(s, text="Préstamos formales", bg="white").grid(row=11, column=0, sticky="w", pady=5)
        self.entries["formal_loans"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["formal_loans"], value="si", bg="white").grid(row=11, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["formal_loans"], value="no", bg="white").grid(row=11, column=2, sticky="w")
        self.create_entry(s, "Monto préstamos formales", "formal_loans_amount", 12)
        
        # Préstamos informales
        tk.Label(s, text="Préstamos informales", bg="white").grid(row=13, column=0, sticky="w", pady=5)
        self.entries["informal_loans"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["informal_loans"], value="si", bg="white").grid(row=13, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["informal_loans"], value="no", bg="white").grid(row=13, column=2, sticky="w")
        self.create_entry(s, "Monto préstamos informales", "informal_loans_amount", 14)
        self.create_entry(s, "Préstamos familiares", "informal_loans_family_amount", 15)
        self.create_entry(s, "Préstamos chulqueros", "informal_loans_moneylender_amount", 16)
        self.create_entry(s, "Otros préstamos informales", "informal_loans_other_amount", 17)
        
        # Tarjetas de crédito
        tk.Label(s, text="Posee tarjetas de crédito", bg="white").grid(row=18, column=0, sticky="w", pady=5)
        self.entries["credit_cards"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["credit_cards"], value="si", bg="white").grid(row=18, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["credit_cards"], value="no", bg="white").grid(row=18, column=2, sticky="w")
        
        # Gastos
        tk.Label(s, text="Gastos mensuales", bg="white", font=("Segoe UI", 9, "bold")).grid(row=19, column=0, columnspan=3, pady=10)
        
        gastos = [
            ("Alimentación", "food_support"),
            ("Educación", "education_support"),
            ("Vivienda", "housing_support"),
            ("Vestimenta", "clothing_support"),
            ("Salud", "health_support"),
            ("Transporte", "transport_support"),
            ("Servicios básicos", "basic_services_support"),
            ("Internet", "internet_support"),
            ("TV Cable", "cable_tv_support"),
            ("Plan Celular", "cell_plan_support"),
            ("Préstamos", "loans_support"),
            ("Préstamos Quirografarios", "unsecured_loans_support"),
            ("Tarjetas de crédito", "credit_cards_support"),
            ("Pensión de Alimentos", "alimony_support"),
            ("Locales Comerciales", "commercial_properties_support"),
            ("Apoyo económico a terceros", "financial_support_others"),
            ("Otros Gastos", "other_expenses_support")
        ]
        
        for i, (label, key) in enumerate(gastos):
            tk.Label(s, text=label, bg="white").grid(row=20+i, column=0, sticky="w", pady=2, padx=5)
            entry = tk.Entry(s, width=15)
            entry.grid(row=20+i, column=1, pady=2, padx=5)
            self.entries[key] = entry
        
        self.create_entry(s, "TOTAL GASTOS", "total_expenses", 37)
        self.create_entry(s, "Total Ingresos", "total_income", 38)
        self.create_entry(s, "Saldo", "balance", 39)
        
        self.entries["total_income"].bind("<KeyRelease>", self.calculate_balance)
        self.entries["total_expenses"].bind("<KeyRelease>", self.calculate_balance)
        
        # Transporte propio
        tk.Label(s, text="Posee vehículo o medio de transporte", bg="white").grid(row=40, column=0, sticky="w", pady=5)
        self.entries["transportation"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["transportation"], value="si", bg="white").grid(row=40, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["transportation"], value="no", bg="white").grid(row=40, column=2, sticky="w")
        self.create_entry(s, "Descripción del transporte", "transportation_description", 41)
        
        self.create_entry(s, "Actividad económica adicional", "additional_economic_activity", 42)
        
        # Crianza de animales
        tk.Label(s, text="Se dedica a la crianza de animales", bg="white").grid(row=43, column=0, sticky="w", pady=5)
        self.entries["animal_breeding"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["animal_breeding"], value="si", bg="white").grid(row=43, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["animal_breeding"], value="no", bg="white").grid(row=43, column=2, sticky="w")

    def build_salud(self, parent):
        s = self.section(parent, "5. SALUD")
        
        # Deporte
        tk.Label(s, text="¿Practica algún deporte o actividad física?", bg="white").grid(row=0, column=0, sticky="w", pady=5)
        self.entries["sports"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["sports"], value="si", bg="white").grid(row=0, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["sports"], value="no", bg="white").grid(row=0, column=2, sticky="w")
        self.create_entry(s, "¿Cuál deporte?", "sports_description", 1)
        self.create_entry(s, "Frecuencia", "sports_frequency", 2)
        
        self.create_entry(s, "¿Tiene alguna enfermedad?", "disease", 3)
        
        # Problemas de salud familiares
        tk.Label(s, text="¿La familia tiene problemas de salud?", bg="white").grid(row=4, column=0, sticky="w", pady=5)
        self.entries["family_health_problems"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["family_health_problems"], value="si", bg="white").grid(row=4, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["family_health_problems"], value="no", bg="white").grid(row=4, column=2, sticky="w")
        self.create_entry(s, "¿Cuáles problemas de salud?", "family_health_problems_description", 5)
        
        # Discapacidad familiar
        tk.Label(s, text="¿La familia tiene algún tipo de discapacidad?", bg="white").grid(row=6, column=0, sticky="w", pady=5)
        self.entries["family_disability"] = tk.StringVar(value="no")
        tk.Radiobutton(s, text="Sí", variable=self.entries["family_disability"], value="si", bg="white").grid(row=6, column=1, sticky="w")
        tk.Radiobutton(s, text="No", variable=self.entries["family_disability"], value="no", bg="white").grid(row=6, column=2, sticky="w")
        
        self.create_entry(s, "Tipo de discapacidad familiar", "family_disability_type", 7)
        self.create_entry(s, "Porcentaje de discapacidad", "family_disability_percentage", 8)
        self.create_entry(s, "Parentezco", "family_disability_relationship", 9)

    def build_laboral(self, parent):
        s = self.section(parent, "6. SITUACIÓN LABORAL")
        
        self.create_entry(s, "Antes de ingresar, ¿a qué se dedicaba?", "previous_occupation", 0, width=60)
        self.create_entry(s, "Funciones que desempeña actualmente", "current_functions", 1, width=60)
        self.create_entry(s, "¿Su trabajo se relaciona con su formación?", "job_relation", 2, width=60)
        self.create_entry(s, "Relación con compañeros", "colleague_relationship", 3, width=60)
        self.create_entry(s, "¿Qué podría mejorar?", "improvement_suggestions", 4, width=60)
        self.create_entry(s, "¿Los conflictos se resuelven eficazmente?", "conflict_resolution", 5, width=60)
        self.create_entry(s, "¿Su trabajo es desgastante?", "job_exhaustion", 6, width=60)
        self.create_entry(s, "¿Siente presión laboral?", "job_pressure", 7, width=60)
        self.create_entry(s, "¿Le genera estrés?", "job_stress", 8, width=60)
        self.create_entry(s, "¿Le alcanza el tiempo para su trabajo?", "job_time_management", 9, width=60)
        self.create_entry(s, "¿Su trabajo es reconocido por su jefe?", "job_manager_recognition", 10, width=60)
        self.create_entry(s, "¿Se siente reconocido por la empresa?", "job_recognition", 11, width=60)
        self.create_entry(s, "Proyección y aspiraciones", "job_projection", 12, width=60)
        self.create_entry(s, "¿Se ha sentido discriminado?", "job_discrimination", 13, width=60)
        self.create_entry(s, "¿Qué haría para mejorar como trabajador?", "job_improvement", 14, width=60)
        self.create_entry(s, "Beneficios que podría implementar la empresa", "job_benefits", 15, width=60)

    def add_family(self):
        frame = tk.Frame(self.family_frame, bg="white")
        frame.pack(pady=2, fill="x")

        entries = {}
        
        fields = ["name", "age", "relation", "job", "income"]
        widths = [15, 8, 15, 15, 12]
        
        for i, (field, width) in enumerate(zip(fields, widths)):
            e = tk.Entry(frame, width=width)
            e.pack(side=tk.LEFT, padx=2)
            entries[field] = e
        
        tk.Button(frame, text="✖", command=lambda: frame.destroy(),
                 bg="#e74c3c", fg="white", width=2).pack(side=tk.LEFT, padx=5)

        self.family_rows.append(entries)

    def calculate_balance(self, event=None):
        try:
            income = float(self.entries.get("total_income", tk.Entry()).get() or 0)
            expenses = float(self.entries.get("total_expenses", tk.Entry()).get() or 0)
            balance = income - expenses
            if "balance" in self.entries:
                self.entries["balance"].delete(0, tk.END)
                self.entries["balance"].insert(0, f"{balance:.2f}")
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
                data[key] = entry.get()
            else:
                data[key] = entry
        
        # Procesar distribución de vivienda
        distribution = []
        for key in self.entries:
            if key.startswith("house_distribution_") and self.entries[key].get() != "no":
                distribution.append(self.entries[key].get())
        data["house_distribution"] = distribution if distribution else ["ninguno"]
        
        data["date"] = datetime.datetime.now().strftime("%d/%m/%Y")
        data["photo"] = self.photo_path.get()
        
        # Familia
        family = []
        for member in self.family_rows:
            family.append({k: e.get() for k, e in member.items()})
        data["family_members"] = family
        
        return data

    def save_all(self):
        try:
            data = self.get_data()
            
            pdf.create_pdf(data)
            xlsx.create_excel_file(data)
            
            messagebox.showinfo("Éxito", "Datos guardados correctamente en PDF y Excel")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()