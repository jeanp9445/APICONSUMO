#Añadir un print("Error") en la linea 100
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Ruta de salida donde se guardará el archivo modificado
RUTA_SALIDA = r"C:\Users\usuario\OneDrive\Escritorio\ApiConsumo\archivosExcel\datosInyectados.xlsx"

def cargar_excel(titulo="Selecciona un archivo Excel"):
    """Función para cargar un archivo Excel usando el explorador de archivos."""
    ruta = filedialog.askopenfilename(
        title=titulo,
        filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
    )
    return ruta

def mostrar_columnas(df, nombre_archivo):
    """Muestra las columnas del DataFrame con su respectivo nombre de archivo."""
    print(f"\nColumnas del archivo {nombre_archivo}:")
    for i, col in enumerate(df.columns):
        print(f"{i+1}. {col}")

def inyectar_datos():
    try:
        # Selección de archivos
        print("Seleccionar el archivo INYECTABLE")
        ruta_inyectable = cargar_excel("Selecciona el archivo INYECTABLE")
        if not ruta_inyectable:
            messagebox.showwarning("Advertencia", "No se seleccionó el archivo INYECTABLE")
            return

        print("Seleccionar el archivo INYECCIÓN")
        ruta_inyeccion = cargar_excel("Selecciona el archivo INYECCIÓN")
        if not ruta_inyeccion:
            messagebox.showwarning("Advertencia", "No se seleccionó el archivo INYECCIÓN")
            return

        # Cargar archivos en DataFrames, asegurándose de tratar los DNI como texto
        df_inyectable = pd.read_excel(ruta_inyectable, dtype=str)  # Leer todos los datos como texto
        df_inyeccion = pd.read_excel(ruta_inyeccion, dtype=str)    # Leer todos los datos como texto

        # Mostrar columnas
        mostrar_columnas(df_inyectable, "INYECTABLE")
        mostrar_columnas(df_inyeccion, "INYECCIÓN")

        # Selección del campo de comparación
        campo_comparacion_inyectable = input("\nEspecifica el campo de comparación en el archivo INYECTABLE: ").strip()
        campo_comparacion_inyeccion = input("Especifica el campo de comparación en el archivo INYECCIÓN: ").strip()

        # Validación de la existencia del campo de comparación
        if campo_comparacion_inyectable not in df_inyectable.columns or campo_comparacion_inyeccion not in df_inyeccion.columns:
            messagebox.showerror("Error", "El campo de comparación no existe en uno de los archivos")
            return

        # Selección de columnas a actualizar
        columnas_mapeo = {}
        while True:
            col_inyectable = input("\nEspecifica una columna del archivo INYECTABLE a actualizar (o presiona ENTER para finalizar): ").strip()
            if not col_inyectable:
                break
            col_inyeccion = input(f"Especifica la columna correspondiente en el archivo INYECCIÓN para {col_inyectable}: ").strip()

            if col_inyectable not in df_inyectable.columns or col_inyeccion not in df_inyeccion.columns:
                print("❌ Error: Alguna de las columnas especificadas no existe. Inténtalo nuevamente.")
                continue

            columnas_mapeo[col_inyectable] = col_inyeccion

        if not columnas_mapeo:
            messagebox.showwarning("Advertencia", "No se seleccionaron columnas para inyectar datos")
            return

        # Crear un diccionario con los valores de actualización usando el campo de comparación
        datos_actualizacion = df_inyeccion.set_index(campo_comparacion_inyeccion)[list(columnas_mapeo.values())].to_dict(orient='index')

        # Actualizar los valores en el DataFrame inyectable
        def actualizar_fila(fila):
            clave = fila[campo_comparacion_inyectable]
            
            # Mantener los ceros a la izquierda en los DNI (si la columna tiene que ver con DNI)
            if clave in datos_actualizacion:
                for col_inyectable, col_inyeccion in columnas_mapeo.items():
                    # Obtener el valor de la celda de la columna correspondiente
                    valor_inyeccion = datos_actualizacion[clave].get(col_inyeccion, None)

                    # Verificar que el valor no esté vacío, ni sea NaN
                    if pd.notna(valor_inyeccion) and valor_inyeccion != "":
                        if 'dni' in col_inyectable.lower() or 'dni' in col_inyeccion.lower():
                            # Asegurar que los DNI son tratados como cadenas, manteniendo ceros
                            fila[col_inyectable] = str(fila[col_inyectable]).zfill(8)  # Rellenar con ceros si es necesario
                        else:
                            fila[col_inyectable] = valor_inyeccion
            return fila

        df_inyectable = df_inyectable.apply(actualizar_fila, axis=1)

        # Guardar el archivo actualizado
        df_inyectable.to_excel(RUTA_SALIDA, index=False)
        print("Éxito", f"Archivo inyectado y guardado en: {RUTA_SALIDA}")
        messagebox.showinfo("Éxito", f"Archivo inyectado y guardado en: \n{RUTA_SALIDA}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")

# Crear la ventana de la aplicación
root = tk.Tk()
root.withdraw()  # Ocultar la ventana principal de tkinter

# Ejecutar la función para inyectar los datos
inyectar_datos()
