import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from services import excel_service as xlsx, pdf_service as pdf

# Función para manejar el evento de envío del formulario
def submit_form():
    # Obtener los datos del formulario
    id = id_entry.get()
    name = name_entry.get()
    email = email_entry.get()
    photo = photo_entry.get()
    
    # Validar que los campos no estén vacíos
    if not id or not name or not email:
        messagebox.showerror("Error", "Por favor, complete todos los campos.")
        return
    
    date = datetime.now().strftime('%Y-%m-%d')

    # Crear un diccionario con los datos del formulario
    data = {
        'id': id,
        'name': name,
        'email': email,
        'date': date,
        'photo': photo
    }

    # Guardar los datos en un archivo Excel
    xlsx.create_excel_file(data)

    # Crear un archivo PDF con los datos del formulario
    pdf.create_pdf(data)

    # Mostrar un mensaje de éxito
    messagebox.showinfo("Éxito", "El formulario ha sido enviado correctamente.")

    # Limpiar los campos del formulario
    id_entry.delete(0, tk.END)
    name_entry.delete(0, tk.END)
    email_entry.delete(0, tk.END)
    photo_entry.delete(0, tk.END)
    
# Crear la ventana principal
root = tk.Tk()
root.title("Formulario de Entrevista Social")

# Crear los campos del formulario
id_label = tk.Label(root, text="ID:")
id_label.pack()
id_entry = tk.Entry(root)
id_entry.pack()

name_label = tk.Label(root, text="Nombre:")
name_label.pack()
name_entry = tk.Entry(root)
name_entry.pack()

email_label = tk.Label(root, text="Email:")
email_label.pack()
email_entry = tk.Entry(root)
email_entry.pack()

photo_label = tk.Label(root, text="Foto:")
photo_label.pack()
photo_entry = tk.Entry(root)
photo_entry.pack()

# Crear un botón para enviar el formulario
submit_button = tk.Button(root, text="Enviar", command=submit_form)
submit_button.pack()

# Iniciar el bucle principal de la interfaz
root.mainloop()