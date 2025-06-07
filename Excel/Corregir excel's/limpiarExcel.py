#Mejoras:
#*convertir todos los datos como mayusculas

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

#Ruta de salida donde se guardará automáticamente el archivo limpio
RUTA_SALIDA = r"C:\Users\usuario\OneDrive\Escritorio\ApiConsumo\archivosExcel\excelLimpio.xlsx"

def limpiar_espacios_excel():
    try:
        #Abrir cuadro de dialogo para seleccionar el archivo de entrada
        ruta_entrada = filedialog.askopenfilename(
            title = "Seleccione el archivo Excel",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
        )

        if not ruta_entrada: # Si el usuario cancela la selección
            messagebox.showwarning("Advertencia", "No se seleccionó ningún archivo.")
            return
        
        #Cargar el archivo Excel
        xls = pd.ExcelFile(ruta_entrada)
        writer = pd.ExcelWriter(RUTA_SALIDA, engine='xlsxwriter')

        #Iterar sobre todas las hojas del archivo
        for hoja in xls.sheet_names:
            #Cargar la hoja en un DataFrame (convertir todo a string para evitar errores)
            df = pd.read_excel(xls, sheet_name=hoja, dtype=str)

            #Limpiar espacios en blanco en todas las celdas
            df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

            #Guardar la hoja en el nuevo archivo
            df.to_excel(writer, sheet_name=hoja, index=False)
        
        writer.close(); #Guardar cambios
        messagebox.showinfo("Éxito", f"Archivo Excel limpio guardado en: \n{RUTA_SALIDA}")

    except Exception as e:
        print(f"Ocurrió un error: {str(e)}")     

        messagebox.showerror("error", f"Ocurrió un error:\n{str(e)}")

#Crear la ventana de la aplicación
root = tk.Tk()
root.withdraw() #Ocultar la ventana principal de tkinder

#Ejecutar la función para limpiar el Excel
limpiar_espacios_excel()