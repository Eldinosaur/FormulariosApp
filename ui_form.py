import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime

from services import pdf_service as pdf, excel_service as xlsx


class App:

    def __init__(self, root):

        self.root = root
        self.root.title("Ficha Socioeconómica")

        self.entries = {}
        self.family_rows = []

        main = tk.Frame(root)
        main.pack(padx=10, pady=10)

        # -------------------------
        # DATOS PERSONALES
        # -------------------------
        tk.Label(main, text="DATOS PERSONALES", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=4)

        fields = [
            ("ci", "Cédula"),
            ("name", "Nombre"),
            ("age", "Edad"),
            ("email", "Email"),
            ("address", "Dirección"),
            ("phone", "Teléfono"),
        ]

        for i, (key, label) in enumerate(fields, start=1):
            tk.Label(main, text=label).grid(row=i, column=0, sticky="w")
            entry = tk.Entry(main, width=30)
            entry.grid(row=i, column=1)
            self.entries[key] = entry

        # FOTO
        tk.Label(main, text="Foto").grid(row=1, column=2)
        self.photo_path = tk.StringVar()

        tk.Entry(main, textvariable=self.photo_path, width=25).grid(row=1, column=3)
        tk.Button(main, text="Buscar", command=self.load_photo).grid(row=1, column=4)

        # -------------------------
        # VIVIENDA
        # -------------------------
        tk.Label(main, text="VIVIENDA", font=("Arial", 12, "bold")).grid(row=8, column=0, columnspan=4)

        self.entries["house_type"] = tk.Entry(main)
        self.entries["ownership"] = tk.Entry(main)
        self.entries["people_count"] = tk.Entry(main)
        self.entries["time_sector"] = tk.Entry(main)

        tk.Label(main, text="Tipo").grid(row=9, column=0)
        self.entries["house_type"].grid(row=9, column=1)

        tk.Label(main, text="Tenencia").grid(row=9, column=2)
        self.entries["ownership"].grid(row=9, column=3)

        tk.Label(main, text="Personas").grid(row=10, column=0)
        self.entries["people_count"].grid(row=10, column=1)

        tk.Label(main, text="Tiempo sector").grid(row=10, column=2)
        self.entries["time_sector"].grid(row=10, column=3)

        # SECTOR
        self.sector = tk.StringVar(value="urbano")
        ttk.Combobox(main, textvariable=self.sector, values=["urbano", "rural"]).grid(row=11, column=1)

        # -------------------------
        # SERVICIOS
        # -------------------------
        tk.Label(main, text="SERVICIOS", font=("Arial", 12, "bold")).grid(row=12, column=0, columnspan=4)

        self.water = tk.BooleanVar()
        self.light = tk.BooleanVar()
        self.internet = tk.BooleanVar()

        tk.Checkbutton(main, text="Agua", variable=self.water).grid(row=13, column=0)
        tk.Checkbutton(main, text="Luz", variable=self.light).grid(row=13, column=1)
        tk.Checkbutton(main, text="Internet", variable=self.internet).grid(row=13, column=2)

        # -------------------------
        # FAMILIA DINÁMICA
        # -------------------------
        tk.Label(main, text="FAMILIARES", font=("Arial", 12, "bold")).grid(row=14, column=0, columnspan=4)

        self.family_frame = tk.Frame(main)
        self.family_frame.grid(row=15, column=0, columnspan=5)

        tk.Button(main, text="Agregar Familiar", command=self.add_family).grid(row=16, column=0)

        # -------------------------
        # OBSERVACIONES
        # -------------------------
        tk.Label(main, text="Observaciones").grid(row=17, column=0)
        self.observations = tk.Text(main, width=60, height=4)
        self.observations.grid(row=18, column=0, columnspan=4)

        # -------------------------
        # BOTÓN
        # -------------------------
        tk.Button(main, text="GUARDAR TODO", command=self.save_all, bg="green", fg="white").grid(row=20, column=1)

    # -------------------------
    # FOTO
    # -------------------------
    def load_photo(self):
        path = filedialog.askopenfilename(filetypes=[("Imagen", "*.jpg *.png")])
        if path:
            self.photo_path.set(path)

    # -------------------------
    # FAMILIA
    # -------------------------
    def add_family(self):

        row_frame = tk.Frame(self.family_frame)
        row_frame.pack(pady=2)

        entries = {}

        for field in ["name", "age", "relation", "job", "income"]:
            e = tk.Entry(row_frame, width=12)
            e.pack(side=tk.LEFT)
            entries[field] = e

        tk.Button(row_frame, text="X", command=lambda: self.remove_family(row_frame)).pack(side=tk.LEFT)

        self.family_rows.append((row_frame, entries))

    def remove_family(self, frame):
        for item in self.family_rows:
            if item[0] == frame:
                self.family_rows.remove(item)
        frame.destroy()

    # -------------------------
    # DATA
    # -------------------------
    def get_data(self):

        data = {}

        for key, entry in self.entries.items():
            data[key] = entry.get()

        data["sector"] = self.sector.get()
        data["date"] = datetime.datetime.now().strftime("%d/%m/%Y")

        data["water"] = self.water.get()
        data["electricity"] = self.light.get()
        data["internet"] = self.internet.get()

        data["shared_expenses"] = True
        data["contributors"] = ""
        data["expenses"] = ""
        data["debts"] = ""

        data["sports"] = False
        data["disease"] = ""
        data["functions"] = ""
        data["life_plan"] = ""

        data["observations"] = self.observations.get("1.0", tk.END)

        data["photo"] = self.photo_path.get()

        # FAMILIA
        family = []

        for _, entries in self.family_rows:
            member = {k: e.get() for k, e in entries.items()}
            if any(member.values()):
                family.append(member)

        data["family_members"] = family

        return data

    # -------------------------
    # SAVE
    # -------------------------
    def save_all(self):

        try:
            data = self.get_data()

            #xlsx.create_excel_file(data)
            pdf.create_pdf(data)

            messagebox.showinfo("OK", "PDF y Excel generados")

        except Exception as e:
            messagebox.showerror("Error", str(e))
            print(e)