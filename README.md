# Planilla-de-presentacion
Este código es una aplicación de generación de archivos Excel a partir de un archivo PDF llamado "CONTROL.pdf". La aplicación utiliza la biblioteca PyPDF2 para extraer el texto del PDF y luego procesa el texto extraído para extraer información relevante.
import os
import PyPDF2
import pandas as pd
import re
import tkinter as tk
from tkinter import messagebox

def generar_excel():
    pdf_file = open('CONTROL.pdf', 'rb')
    pdf_reader = PyPDF2.PdfReader(pdf_file)

    text = ""

    for page in pdf_reader.pages:
        text += page.extract_text()

    pdf_file.close()

    ventas = []
    try:
        for line in text.split("\n"):
            if line.strip().startswith("Venta"):
                venta = line.strip().split()[-1]
            elif line.strip().startswith("SKU"):
                sku = line.strip().split()[-1]
            elif line.strip().startswith("Cantidad"):
                cantidad_text = line.strip().split()[-1]
                if re.match(r'\d+', cantidad_text):
                    cantidad = int(cantidad_text)
                else:
                    # extract numeric value from PDF
                    # assuming the numeric value is the last number in the line
                    cantidad = int(re.findall(r'\d+', line)[-1])
                ventas.append([venta, sku, cantidad, "-"])
            elif "Color" in line.strip():
                color = line.strip().split(":")[-1].strip()
                ventas[-1][-1] = color
    except Exception as e:
        print(f"An error occurred: {e}")

    carpeta = os.getcwd()

    try:
        id_inicial = int(entry.get())
        nombre_xlsx = f"{id_inicial}_CONTROL.xlsx"

        id_actual = id_inicial
        id_venta = {}
        for venta in ventas:
            if venta[0] not in id_venta:
                id_venta[venta[0]] = id_actual
                id_actual += 1
            venta.insert(0, id_venta[venta[0]])

        df_ventas = pd.DataFrame(ventas, columns=["ID", "Venta", "SKU", "Cantidad", "Color"])
        df_ventas.to_excel(os.path.join(carpeta, nombre_xlsx), index=False, sheet_name="Ventas")

        df_relacionado = df_ventas
        df_relacionado.to_excel(os.path.join(carpeta, nombre_xlsx), index=False, sheet_name="Relacionado")

        messagebox.showinfo("Excel generado", f"Archivo {nombre_xlsx} creado en {carpeta}.")
        ventana.destroy()

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

ventana = tk.Tk()
ventana.geometry("880x440")
ventana.title("Generador de archivo Excel")

label = tk.Label(ventana, text="Ingrese el ID inicial:")
label.config(font=("Arial", 24))
label.pack(pady=50)

entry = tk.Entry(ventana)
entry.pack()

boton = tk.Button(ventana, text="Generar archivo Excel", command=generar_excel)
boton.pack(pady=20)

ventana.mainloop()
