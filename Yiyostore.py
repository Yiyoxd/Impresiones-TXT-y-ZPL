"""
Aplicación para gestionar la impresión de archivos ZPL en impresoras configuradas.
Permite seleccionar archivos manualmente, arrastrar y soltar archivos, y monitorear
una carpeta específica para imprimir automáticamente los archivos añadidos.

Requisitos:
- Python 3.x
- Módulos:
    - os
    - win32print
    - tkinter
    - win10toast
    - tkinterdnd2
    - time
"""

import os
import win32print
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from win10toast import ToastNotifier
from tkinterdnd2 import DND_FILES, TkinterDnD
import time

# Inicializa el sistema de notificaciones de Windows
toaster = ToastNotifier()

# Configuración del directorio de la aplicación para almacenar archivos de configuración
APPDATA_DIR = os.path.join(os.getenv('APPDATA'), 'Yiyostore')
os.makedirs(APPDATA_DIR, exist_ok=True)

CONFIG_FILE = os.path.join(APPDATA_DIR, "config.txt")


def leer_configuracion():
    """
    Lee la configuración guardada desde el archivo de configuración.

    Returns:
        tuple: (impresora, carpeta) donde cada uno es una cadena de texto.
    """
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as file:
            lines = file.readlines()
            impresora = lines[0].strip() if len(lines) > 0 else ''
            carpeta = lines[1].strip() if len(lines) > 1 else ''
            return impresora, carpeta
    return '', ''


def guardar_configuracion(impresora, carpeta):
    """
    Guarda la configuración actual en el archivo de configuración.

    Args:
        impresora (str): Nombre de la impresora seleccionada.
        carpeta (str): Ruta de la carpeta seleccionada.
    """
    with open(CONFIG_FILE, 'w') as file:
        file.write(f"{impresora}\n")
        file.write(f"{carpeta}\n")


def obtener_lista_impresoras():
    """
    Obtiene la lista de impresoras disponibles en el sistema.

    Returns:
        list: Lista de nombres de impresoras disponibles.
    """
    printers = [printer[2] for printer in win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    return printers


def enviar_bloque_a_impresora(bloque_zpl, impresora):
    """
    Envía un bloque de código ZPL a la impresora especificada.

    Args:
        bloque_zpl (str): Código ZPL a imprimir.
        impresora (str): Nombre de la impresora de destino.
    """
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
            duration=5,
            threaded=True
        )
    except Exception as e:
        messagebox.showerror("Error", f"Error al enviar el bloque a la impresora:\n{e}")


def enviar_a_impresora(archivo_zpl, impresora):
    """
    Lee un archivo ZPL y envía cada bloque de código a la impresora especificada.

    Args:
        archivo_zpl (str): Ruta del archivo ZPL.
        impresora (str): Nombre de la impresora de destino.
    """
    try:
        with open(archivo_zpl, 'r') as file:
            zpl_data = file.read()

        bloques = [bloque.strip() + '^XZ' for bloque in zpl_data.split('^XZ') if bloque.strip()]

        for bloque in bloques:
            enviar_bloque_a_impresora(bloque, impresora)
            time.sleep(0.05)  # Pequeño retraso para evitar saturar la impresora

    except FileNotFoundError:
        messagebox.showerror("Error", f"El archivo {archivo_zpl} no se encontró.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al enviar el archivo a la impresora:\n{e}")


def seleccionar_archivos():
    """
    Abre un diálogo para seleccionar uno o más archivos ZPL y los envía a la impresora seleccionada.
    """
    archivos_zpl = filedialog.askopenfilenames(
        title="Seleccionar archivos de etiquetas",
        filetypes=[("Archivos ZPL", "*.txt *.zpl")]
    )
    if archivos_zpl:
        impresora = impresora_combobox.get()
        if impresora:
            for archivo in archivos_zpl:
                enviar_a_impresora(archivo, impresora)
        else:
            messagebox.showwarning("Advertencia", "Por favor, seleccione una impresora.")


def on_file_drop(event):
    """
    Maneja el evento de arrastrar y soltar archivos en la ventana de la aplicación.

    Args:
        event: Evento de Tkinter que contiene la información de los archivos arrastrados.
    """
    archivos = app.splitlist(event.data)
    impresora = impresora_combobox.get()
    if impresora:
        for archivo in archivos:
            if archivo.lower().endswith(('.txt', '.zpl')):
                enviar_a_impresora(archivo, impresora)
            else:
                messagebox.showwarning("Advertencia", f"El archivo {archivo} no es un archivo ZPL válido.")
    else:
        messagebox.showwarning("Advertencia", "Por favor, seleccione una impresora.")


def seleccionar_carpeta():
    """
    Abre un diálogo para seleccionar una carpeta y actualiza la configuración.
    """
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta para monitorear")
    if carpeta:
        carpeta_entry.delete(0, tk.END)
        carpeta_entry.insert(0, carpeta)
        guardar_configuracion(impresora_combobox.get(), carpeta)


def monitorear_carpeta():
    """
    Monitorea la carpeta seleccionada y envía automáticamente a imprimir cualquier archivo ZPL nuevo que se agregue.

    Esta función se ejecuta periódicamente cada cierto intervalo de tiempo.
    """
    carpeta = carpeta_entry.get()
    impresora = impresora_combobox.get()
    if carpeta and impresora:
        try:
            archivos = [os.path.join(carpeta, f) for f in os.listdir(carpeta) if f.lower().endswith(('.txt', '.zpl'))]
            for archivo in archivos:
                enviar_a_impresora(archivo, impresora)
                os.remove(archivo)
        except Exception as e:
            messagebox.showerror("Error", f"Error al monitorear la carpeta:\n{e}")
    app.after(3000, monitorear_carpeta)  # Revisa la carpeta cada 3 segundos


def on_impresora_selected(event):
    """
    Actualiza la configuración cuando se selecciona una impresora diferente.

    Args:
        event: Evento de Tkinter que indica que se ha seleccionado una nueva opción en el Combobox.
    """
    guardar_configuracion(impresora_combobox.get(), carpeta_entry.get())


# Inicialización de la aplicación Tkinter con soporte para arrastrar y soltar
app = TkinterDnD.Tk()
app.title("Impresiones Yiyostore")
app.resizable(False, False)

# Configuración del marco principal
frame = tk.Frame(app, padding="10")
frame.pack(fill=tk.BOTH, expand=True)

# Etiqueta y Combobox para la selección de impresora
impresora_label = tk.Label(frame, text="Seleccione la impresora:")
impresora_label.grid(row=0, column=0, sticky=tk.W, pady=5)

impresoras_disponibles = obtener_lista_impresoras()
impresora_seleccionada, carpeta_seleccionada = leer_configuracion()

impresora_combobox = ttk.Combobox(frame, values=impresoras_disponibles, state="readonly")
impresora_combobox.grid(row=0, column=1, sticky=tk.EW, pady=5)
impresora_combobox.bind("<<ComboboxSelected>>", on_impresora_selected)

if impresora_seleccionada in impresoras_disponibles:
    impresora_combobox.set(impresora_seleccionada)
elif impresoras_disponibles:
    impresora_combobox.set(impresoras_disponibles[0])

# Etiqueta, Entry y Botón para la selección de carpeta
carpeta_label = tk.Label(frame, text="Seleccione la carpeta:")
carpeta_label.grid(row=1, column=0, sticky=tk.W, pady=5)

carpeta_entry = tk.Entry(frame)
carpeta_entry.grid(row=1, column=1, sticky=tk.EW, pady=5)
carpeta_entry.insert(0, carpeta_seleccionada)

carpeta_button = tk.Button(frame, text="Seleccionar", command=seleccionar_carpeta)
carpeta_button.grid(row=1, column=2, sticky=tk.E, pady=5)

# Botón para seleccionar archivos manualmente
seleccionar_archivos_button = tk.Button(
    frame,
    text="Seleccionar archivos a imprimir",
    command=seleccionar_archivos,
    width=30
)
seleccionar_archivos_button.grid(row=2, column=0, columnspan=3, pady=10)

# Configuración de arrastrar y soltar en la ventana principal
app.drop_target_register(DND_FILES)
app.dnd_bind('<<Drop>>', on_file_drop)

# Ajuste de las columnas para que se expandan correctamente
frame.columnconfigure(1, weight=1)

# Inicia el monitoreo de la carpeta seleccionada
monitorear_carpeta()

# Ejecuta el bucle principal de la aplicación
app.mainloop()
