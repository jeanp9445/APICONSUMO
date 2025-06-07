#LOS ARCHIVOS EXCEL CARGABLES SON LOS QUE, EN EL, FUERON INSCRITOS LOS NOMBRES MEDIANTE EL FORMATO Apellidos-Nombres
#ES REQUERIBLE MENCIONAR CUAL ES EL NOMBRE DEL CAMPO ASOCIADO CUANDO EL SISTEMA LO PIDA

#MEJORAR EL MAPEO DE NOMBRES (QUE MAPEE POR 4 ESPACIOS)

import os
import shutil
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Función para limpiar texto en celdas
def limpiar_texto(texto):
    if pd.isna(texto):  # Si el valor es NaN, se convierte en cadena vacía
        return ""
    return str(texto).strip()  # Elimina espacios antes y después

# Función mejorada para separar apellidos y nombres
def separar_nombres(nombre_completo):
    nombre_completo = limpiar_texto(nombre_completo)  # Limpiar el texto primero

    # Si no hay nada en la celda, devolvemos valores vacíos
    if not nombre_completo:
        return "", ""

    # Separar el nombre completo en partes por espacios
    partes = nombre_completo.split(' ')

    # Caso 1: Si hay solo un apellido y un nombre
    if len(partes) == 2:
        apellidos = partes[0]
        nombres = partes[1]
        return apellidos, nombres
    # Caso 2: Si hay dos apellidos y un nombre
    elif len(partes) == 3:
        apellidos = f"{partes[0]} {partes[1]}"
        nombres = partes[2]
        return apellidos, nombres
    # Caso 3: Si hay dos apellidos y dos nombres
    elif len(partes) >= 4:
        apellidos = f"{partes[0]} {partes[1]}"
        nombres = ' '.join(partes[2:])
        return apellidos, nombres
    else:
        # Si solo hay un término, lo tratamos como apellidos y dejamos nombres vacíos
        return partes[0], ""

# Función para cargar el archivo Excel mediante el administrador de archivos
def cargar_archivo():
    # Inicializa la ventana de Tkinter (sin mostrarla)
    Tk().withdraw()

    # Abre el cuadro de diálogo para seleccionar el archivo
    archivo_entrada = askopenfilename(title="Selecciona el archivo Excel", filetypes=[("Archivos Excel", "*.xlsx;*.xls")])
    
    if archivo_entrada == "":
        print("Error: No se seleccionó ningún archivo.")
        return None
    
    # Verifica si el archivo es de Excel y lo guarda en la carpeta 'archivosExcel'
    if archivo_entrada.lower().endswith(('.xlsx', '.xls')): 
        # Crear carpeta si no existe
        carpeta_destino = os.path.join(os.getcwd(), "archivosExcel")
        if not os.path.exists(carpeta_destino):
            os.makedirs(carpeta_destino)
        
        # Guardar el archivo en la carpeta 'archivosExcel'
        archivo_destino = os.path.join(carpeta_destino, os.path.basename(archivo_entrada))
        
        # Verificar si el archivo ya existe y renombrarlo si es necesario
        if os.path.exists(archivo_destino):
            nombre, extension = os.path.splitext(os.path.basename(archivo_entrada))
            contador = 1
            while os.path.exists(archivo_destino):
                archivo_destino = os.path.join(carpeta_destino, f"{nombre}_{contador}{extension}")
                contador += 1
        
        shutil.copy(archivo_entrada, archivo_destino)
        print(f"Archivo cargado y guardado como: {archivo_destino}")
        return pd.read_excel(archivo_destino)
    
    else:
        print("Error: El archivo seleccionado no es válido. Solo se permiten archivos Excel.")
        return None

# Función principal
def procesar_excel():
    # Cargar el archivo Excel
    df = cargar_archivo()

    if df is None:
        return

    # Limpiar los nombres de las columnas eliminando espacios extra
    df.columns = [col.strip() for col in df.columns]

    # Limpiar todas las celdas del DataFrame
    df = df.applymap(limpiar_texto)

    # Mostrar las columnas corregidas
    print("Columnas disponibles en el archivo:", df.columns.tolist())

    # Solicitar el nombre de la columna con los nombres completos
    columna_nombres = input("Ingrese el nombre de la columna que contiene los nombres completos: ").strip()

    if columna_nombres not in df.columns:
        print(f"Error: La columna '{columna_nombres}' no existe en el archivo.")
    else:
        # Aplicar la separación de nombres y apellidos
        df[['Apellidos', 'Nombres']] = df[columna_nombres].apply(lambda x: pd.Series(separar_nombres(x)))

        # Verificar y agregar la extensión si es necesario
        archivo_salida = input("Ingrese el nombre del archivo de salida (por ejemplo, nombres_separados.xlsx): ").strip()

        # Asegurar que la extensión '.xlsx' esté presente
        if not archivo_salida.endswith('.xlsx'):
            archivo_salida += '.xlsx'

        # Guardar el archivo con los cambios
        df.to_excel(archivo_salida, index=False)
        print(f"Archivo procesado correctamente. Guardado como '{archivo_salida}'")

# Ejecutar el procesamiento
procesar_excel()
