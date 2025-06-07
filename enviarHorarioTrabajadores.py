import pandas as pd
import requests
import os
import re
import json
import tkinter as tk
from tkinter import filedialog

# Inicializar la vetana para seleccionar archivos
root = tk.Tk()
root.withdraw()

# Función para seleccionar un archivo Excel
def seleccionar_archivo():
    return filedialog.askopenfilename(title="Seleccione el archivo Excel", filetypes=[("Archivo Excel", "*.xlsx;*.xls")])

# Mapeo de días válidos en la API
DIAS_VALIDOS = {
    "lu": "Lunes",
    "ma": "Martes",
    "mi": "Miércoles",
    "ju": "Jueves",
    "vi": "Viernes",
    "sa": "Sábado",
    "do": "Domingo"
}

# Expresión regular para extraer los dos primeros caracteres de cada palabra (ignorando tildes y mayúsculas)
def extraer_dias(texto):
    texto = texto.lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    
    dias_extraidos = []
    
    # Si la cadena contiene " a " (con espacios), significa que es un rango
    if " a " in texto:
        partes = texto.split(" a ")
        if len(partes) == 2:
            inicio, fin = partes
            inicio = inicio.strip()[:2]  # Extraer dos primeras letras
            fin = fin.strip()[:2]
            
            # Verificar que ambos días existan en el diccionario
            if inicio in DIAS_VALIDOS and fin in DIAS_VALIDOS:
                dias_lista = list(DIAS_VALIDOS.keys())  # ["lu", "ma", "mi", "ju", "vi", "sa", "do"]
                start_index = dias_lista.index(inicio)
                end_index = dias_lista.index(fin)
                
                if start_index <= end_index:
                    dias_extraidos = [DIAS_VALIDOS[d] for d in dias_lista[start_index:end_index + 1]]

    # Si no es un rango, extraer días individuales
    if not dias_extraidos:
        coincidencias = re.findall(r'\b\w{2}', texto)
        dias_extraidos = [DIAS_VALIDOS[d] for d in coincidencias if d in DIAS_VALIDOS]

    return dias_extraidos


# Seleccionar archivo Excel
archivo_path = seleccionar_archivo()
if not archivo_path:
    print("Error: No se seleccionó ningún archivo.")
    exit()

# Cargar el archivo Excel
df = pd.read_excel(archivo_path, dtype=str)

#Mostrar nombres de las columnas
print("\nColumnas disponibles en el archivo:")
print(df.columns.tolist())

# Capturar la columna que contiene los IDs
columna_id = input("\nIngrese el nombre de la columna que contiene los ID de los trabajadores: ").strip()
if not columna_id in df.columns:
    print("Erro: La columna de ID no existe en el archivo.")
    exit()

# Capturar las columnas que contienen los días y sus horarios asociados
columnas_dias = input("Ingrese las columnas que contienen los días (separados por comas ): ").split(',')
columnas_horas_ingreso = input("Ingrese las columnas que contiene las horas de ingreso (en el mismo orden, separadas por comas): ").split(',')
columnas_horas_salida = input("Ingrese las columnas que continen las horas de salida (en el mismo orden, separadas por comas): ").split(',')

# Limpiar espacios en los nombres de columnas
columnas_dias = [col.strip() for col in columnas_dias]
columnas_horas_ingreso = [col.strip() for col in columnas_horas_ingreso]
columnas_horas_salida = [col.strip() for col in columnas_horas_salida]

# Validar que todas las columnas ingresadas existan en el archivo
if not all(col in df.columns for col in columnas_dias + columnas_horas_ingreso + columnas_horas_salida):
    print("Erro: Una o más columnas ingresadas no existen en el archivo.")
    exit()

# Iterar sobre cada fila del archivo
for _, fila in df.iterrows():
    worker_id = fila[columna_id]

    # Diccionario para agrupar los días según su horario
    horarios = {}

    #Iterar sobre las columnas de días y sus respectivas horas
    for i in range(len(columnas_dias)): 
        col_dia = columnas_dias[i] 
        col_hora_ingreso = columnas_horas_ingreso[i]
        col_hora_salida = columnas_horas_salida[i]

        if pd.notna(fila[col_dia]) and pd.notna(fila[col_hora_ingreso]) and pd.notna(fila[col_hora_salida]): # ***Posible conflicto por igualdad de nombres asociados a los campos
            dias_extraidos = extraer_dias(fila[col_dia])
            horario = (fila[col_hora_ingreso][:5], fila[col_hora_salida][:5])

            if horario in horarios:
                horarios[horario].extend(dias_extraidos)
            else: 
                horarios[horario] = dias_extraidos 

    # Enviar una solicitud Post por cada grupo de días con el mismo horario
    for (hora_ingreso, hora_salida), dias in horarios.items():
        data = {
            "scheduleWorker": {
                "horaIngreso": hora_ingreso,
                "horaSalida": hora_salida
            },
            "days": [{"nombreDia": dia} for dia in dias]
        }
        
        # Hacer la petición POST a la API
        
        url = f"http://192.168.0.102:8080/api/v1/workers/{worker_id}/scheduleWorkers" # url = f"http://spx-enterprise.com.pe/api/api/v1/workers/{worker_id}/scheduleWorkers"
        headers = {"Content-Type": "application/json"}

        print("\nPayload enviado: ", json.dumps(data, indent=2))

        try:
            response = requests.post(url, data=json.dumps(data), headers=headers)
            response.raise_for_status()
            print(f"✔ Datos enviados correctamente para el ID {worker_id} con horario {hora_ingreso} - {hora_salida}, Días: {dias}.")
        except requests.exceptions.RequestException as e:
            print(f"❌ Error al enviar datos para el ID {worker_id} con horario {hora_ingreso} - {hora_salida}: {e}")
      
print("\nProceso finalizado.")