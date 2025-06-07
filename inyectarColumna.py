import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

# Inicializar la ventana de selección de archivos
root = tk.Tk()
root.withdraw() # Ocultar la ventana de Tkinter

# Función para seleccionar archivo Excel
def seleccionar_archivo(mensaje):
    return filedialog.askopenfilename(title=mensaje, filetypes=[("Archivos Excel", "*.xlsx;*.xlsx")])

# Cargar archivos mediante selección del usuario
inyectable_path = seleccionar_archivo("Seleccione el archivo inyectable")
inyeccion_path = seleccionar_archivo("Seleccione el archivo de inyección")

# Verificar si se seleccionaron archivos
if not inyectable_path or not inyeccion_path:
    print("Error: Debe seleccionar ambos archivos.")
    exit()

# Cargar los archivos Excel
inyectable_df = pd.read_excel(inyectable_path, dtype=str)
inyeccion_df = pd.read_excel(inyeccion_path, dtype=str)

# Mostrar nombres de las columnas en ambos archivos
print("\nColumnas en el archivo inyectable:")
print(inyectable_df.columns.tolist())
print("\nColumnas en el archivo de inyección:")
print(inyeccion_df.columns.tolist())

# Capturar campos comparativos y columnas a inyectar
campo_inyectable = input("Ingrese el nombre del campo comparativo en el inyectable: ")
campo_inyeccion = input("Ingrese el nombre del campo comparativo en la inyección: ")
columnas_inyectar = input("Ingrese los nombres de las columnas a inyectar, separados por comas: ").split(',')

# Realizar la inyección de datos
inyectable_df = inyectable_df.merge(
    inyeccion_df[[campo_inyeccion] + columnas_inyectar],
    left_on = campo_inyectable,
    right_on = campo_inyeccion,
    how = "left"
).drop(columns=[campo_inyeccion]) # Eliminar duplicado del campo comparativo

# Crear directorio de salida si no existe
output_dir = "archivosExcel"
os.makedirs(output_dir, exist_ok=True)

# Guardar el resultado
output_path = os.path.join(output_dir, "columnasInyectadas.xlsx")
inyectable_df.to_excel(output_path, index=False)

print(f"Proceso completado. Archivo guardado en: {output_path}")
