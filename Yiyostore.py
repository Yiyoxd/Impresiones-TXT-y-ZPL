import os
import win32print
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from win10toast import ToastNotifier
from tkinterdnd2 import DND_FILES, TkinterDnD
import time

# Inicializar notificador
toaster = ToastNotifier()

# Obtener el directorio de datos de la aplicación del usuario
appdata_dir = os.path.join(os.getenv('APPDATA'), 'Yiyostore')
if not os.path.exists(appdata_dir):
    os.makedirs(appdata_dir)

config_file = os.path.join(appdata_dir, "config.txt")

def leer_configuracion():
    if os.path.exists(config_file):
        with open(config_file, 'r') as file:
            lines = file.readlines()
            impresora = lines[0].strip() if lines else ''
            carpeta = lines[1].strip() if len(lines) > 1 else ''
            return impresora, carpeta
    return '', ''

def guardar_configuracion(impresora, carpeta):
    with open(config_file, 'w') as file:
        file.write(f"{impresora}\n")
        file.write(f"{carpeta}\n")

def enviar_bloque_a_impresora(bloque_zpl, impresora):
    try:
        printer = win32print.OpenPrinter(impresora)
        job = win32print.StartDocPrinter(printer, 1, ("Etiqueta", None, "RAW"))
        win32print.StartPagePrinter(printer)
        win32print.WritePrinter(printer, bloque_zpl.encode())
        win32print.EndPagePrinter(printer)
        win32print.EndDocPrinter(printer)
        win32print.ClosePrinter(printer)
        
        toaster.show_toast(
            "Yiyostore",
            "Bloque enviado a la impresora correctamente.",
            duration=10,  # Duración de la notificación
            threaded=True
        )
    except Exception as e:
        messagebox.showerror("Error", f"Error al enviar el bloque a la impresora: {e}")

def enviar_a_impresora(archivo_zpl, impresora):
    try:
        with open(archivo_zpl, 'r') as file:
            zpl_data = file.read().split('^XZ')

        for bloque in zpl_data:
            if bloque.strip():
                bloque = bloque + '^XZ'
                enviar_bloque_a_impresora(bloque, impresora)
                time.sleep(0.01)  # Pequeño retraso entre envíos
    except FileNotFoundError:
        messagebox.showerror("Error", "El archivo no se encontró.")
    except ValueError as e:
        messagebox.showerror("Error", str(e))
    except Exception as e:
        messagebox.showerror("Error", f"Error al enviar el archivo a la impresora: {e}")

def seleccionar_archivo():
    archivos_zpl = filedialog.askopenfilenames(title="Seleccionar archivos de etiquetas", filetypes=[("Archivos ZPL", "*.txt *.zpl")])
    if archivos_zpl:
        impresora = impresora_combobox.get()
        if impresora:
            for archivo_zpl in archivos_zpl:
                enviar_a_impresora(archivo_zpl, impresora)
        else:
            messagebox.showerror("Error", "Seleccione la impresora.")

def obtener_lista_impresoras():
    printers = []
    for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
        printers.append(printer[2])
    return printers

def on_file_drop(event):
    archivos_zpl = event.data.split()  # Separar por espacios
    impresora = impresora_combobox.get()
    if impresora:
        for archivo_zpl in archivos_zpl:
            archivo_zpl = archivo_zpl.strip('{}')  # Eliminar llaves y espacios en blanco
            try:
                enviar_a_impresora(archivo_zpl, impresora)
            except Exception as e:
                messagebox.showerror("Error", f"Error al procesar el archivo: {archivo_zpl}\n\n{e}")
    else:
        messagebox.showerror("Error", "Seleccione la impresora.")

def seleccionar_carpeta():
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta para monitorear")
    if carpeta:
        carpeta_entry.delete(0, tk.END)
        carpeta_entry.insert(0, carpeta)
        guardar_configuracion(impresora_combobox.get(), carpeta)

def monitorear_carpeta():
    carpeta = carpeta_entry.get()
    archivos_en_proceso = set()
    if carpeta:
        for archivo in os.listdir(carpeta):
            if archivo.endswith(".txt") or archivo.endswith(".zpl"):
                archivo_zpl = os.path.join(carpeta, archivo)
                if archivo_zpl not in archivos_en_proceso:
                    archivos_en_proceso.add(archivo_zpl)
                    enviar_a_impresora(archivo_zpl, impresora_combobox.get())
                    os.remove(archivo_zpl)
                    archivos_en_proceso.remove(archivo_zpl)
    app.after(2000, monitorear_carpeta)

def on_impresora_selected(event):
    guardar_configuracion(impresora_combobox.get(), carpeta_entry.get())

app = TkinterDnD.Tk()
app.title("Impresiones Yiyostore")

frame = tk.Frame(app)
frame.pack(padx=20, pady=20)

impresora_label = tk.Label(frame, text="Seleccione la impresora:")
impresora_label.grid(row=0, column=0, padx=5, pady=5)

impresoras_disponibles = obtener_lista_impresoras()
impresora_seleccionada, carpeta_seleccionada = leer_configuracion()

impresora_combobox = ttk.Combobox(frame, values=impresoras_disponibles, width=30)
impresora_combobox.grid(row=0, column=1, padx=5, pady=5)
impresora_combobox.set(impresora_seleccionada if impresora_seleccionada else impresoras_disponibles[0])
impresora_combobox.bind("<<ComboboxSelected>>", on_impresora_selected)

carpeta_label = tk.Label(frame, text="Seleccione la carpeta:")
carpeta_label.grid(row=1, column=0, padx=5, pady=5)

carpeta_entry = tk.Entry(frame, width=30)
carpeta_entry.grid(row=1, column=1, padx=5, pady=5)
carpeta_entry.insert(0, carpeta_seleccionada)

carpeta_button = tk.Button(frame, text="Seleccionar", command=seleccionar_carpeta)
carpeta_button.grid(row=1, column=2, padx=5, pady=5)

seleccionar_btn = tk.Button(frame, text="Seleccione o arrastre los archivos a imprimir", command=seleccionar_archivo)
seleccionar_btn.grid(row=2, columnspan=3, pady=10)

app.drop_target_register(DND_FILES)
app.dnd_bind('<<Drop>>', on_file_drop)

app.after(2000, monitorear_carpeta)
app.mainloop()