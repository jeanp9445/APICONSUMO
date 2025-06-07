import pandas as pd
import tkinter as tk
from tkinter import filedialog
import requests
import base64
import io

def cargar_excel(titulo="Selecciona un archivo Excel"):
    """
    Función para cargar un archivo Excel usando el explorador de archivos.
    """
    ruta = filedialog.askopenfilename(
        title=titulo,
        filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
    )
    return ruta

def procesar_excel(ruta_excel):
    """
    Procesa el archivo Excel y extrae los datos para enviarlos como multipart/form-data.
    """
    df = pd.read_excel(ruta_excel)
    
    # Convertir las columnas de fechas a tipo datetime (si no están ya)
    df['fechaNacimiento'] = pd.to_datetime(df['fechaNacimiento'], errors='coerce')
    df['fechaInicioContrato'] = pd.to_datetime(df['fechaInicioContrato'], errors='coerce')
    df['fechaInicioLaboral'] = pd.to_datetime(df['fechaInicioLaboral'], errors='coerce')
    df['fechaFinContrato'] = pd.to_datetime(df['fechaFinContrato'], errors='coerce')
    df['fechaInicioPerComputable'] = pd.to_datetime(df['fechaInicioPerComputable'], errors='coerce')

    # Mostrar el tamaño del dataframe
    print(f"Total de registros en el archivo: {len(df)}")
    
    # Preparar la lista de datos a enviar
    data_list = []
    for _, row in df.iterrows():
        # Asegurarse de que el DNI y el celular son tratados como cadenas de texto
        dni = str(row['dni']).zfill(8)  # Añadir ceros a la izquierda si es necesario
        celular = str(row['celular']).zfill(9)  # Añadir ceros a la izquierda si es necesario

        # Mostrar los valores de DNI y Celular procesados
        print(f"DNI procesado: {dni}")
        print(f"Celular procesado: {celular}")

        # Preparar el registro de datos para este trabajador
        data = {
            'id': int(row['id']),
            'nombres': row['nombres'],
            'apellidos': row['apellidos'],
            'dni': dni,
            'estadoCivil': row['estadoCivil'],
            'fechaNacimiento': pd.Timestamp(row['fechaNacimiento']).strftime('%Y-%m-%d') if pd.notnull(row['fechaNacimiento']) else None,
            'cargo': row['cargo'],
            'tipoTrabajador': row['tipoTrabajador'],
            'direccion': row['direccion'],
            'distrito': row['distrito'],
            'celular': celular,
            'correoCorporativo': row['correoCorporativo'],
            'correoPersonal': row['correoPersonal'],
            'fechaInicioContrato': pd.Timestamp(row['fechaInicioContrato']).strftime('%Y-%m-%d') if pd.notnull(row['fechaInicioContrato']) else None,
            'fechaInicioLaboral': pd.Timestamp(row['fechaInicioLaboral']).strftime('%Y-%m-%d') if pd.notnull(row['fechaInicioLaboral']) else None,
            'fechaFinContrato': pd.Timestamp(row['fechaFinContrato']).strftime('%Y-%m-%d') if pd.notnull(row['fechaFinContrato']) else None,
            'fechaInicioPerComputable': pd.Timestamp(row['fechaInicioPerComputable']).strftime('%Y-%m-%d') if pd.notnull(row['fechaInicioPerComputable']) else None,
            'sueldo': float(row['sueldo']),
            'movilidad': float(row['movilidad']),
            'asignacionFamiliar': bool(row['asignacionFamiliar']),
            'numeroHijos': int(row['numeroHijos']),
            'foto': row['foto']  # Suponemos que la columna "foto" tiene el base64
        }
        data_list.append(data)
    
    return data_list

def enviar_put(data):
    """
    Envía una solicitud PUT con datos multipart/form-data.
    """
    url = f"http://spx-enterprise.com.pe/api/api/v1/workers/{data['id']}" #http://192.168.0.112:8080/api/v1/workers/{data['id']}
    
    # Preparar los datos como multipart/form-data
    files = {
        'nombres': (None, data['nombres']),
        'apellidos': (None, data['apellidos']),
        'dni': (None, data['dni']),
        'estadoCivil': (None, data['estadoCivil']),
        'fechaNacimiento': (None, data['fechaNacimiento']),
        'cargo': (None, data['cargo']),
        'tipoTrabajador': (None, data['tipoTrabajador']),
        'direccion': (None, data['direccion']),
        'distrito': (None, data['distrito']),
        'celular': (None, data['celular']),
        'correoCorporativo': (None, data['correoCorporativo']),
        'correoPersonal': (None, data['correoPersonal']),
        'fechaInicioContrato': (None, data['fechaInicioContrato']),
        'fechaInicioLaboral': (None, data['fechaInicioLaboral']),
        'fechaFinContrato': (None, data['fechaFinContrato']),
        'fechaInicioPerComputable': (None, data['fechaInicioPerComputable']),
        'sueldo': (None, str(data['sueldo'])),
        'movilidad': (None, str(data['movilidad'])),
        'asignacionFamiliar': (None, str(data['asignacionFamiliar'])),
        'numeroHijos': (None, str(data['numeroHijos'])),
    }
    
    # Si la foto está en formato base64, la convertimos y la añadimos como archivo
    if data['foto']:
        # Convertir la cadena base64 en un archivo binario
        foto_binaria = base64.b64decode(data['foto'])
        foto_io = io.BytesIO(foto_binaria)
        files['foto'] = ('foto.jpg', foto_io, 'image/jpeg')
    
    # Realizar el PUT
    response = requests.put(url, files=files)
    
    # Verificación de respuesta
    if response.status_code == 200:
        print(f"Datos enviados correctamente para ID: {data['id']}")
    else:
        print(f"Error al enviar datos para ID: {data['id']} - {response.status_code} - {response.text}")

# Crear la ventana de la aplicación
root = tk.Tk()
root.withdraw()  # Ocultar la ventana principal de tkinter

# Cargar el archivo Excel
ruta_excel = cargar_excel("Selecciona el archivo Excel")

if ruta_excel:
    # Procesar el archivo Excel
    data_list = procesar_excel(ruta_excel)
    
    # Enviar los datos usando PUT para cada registro
    for data in data_list:
        print("Enviando datos a la API:")
        print(f"DNI: {data['dni']}, Celular: {data['celular']}")
        enviar_put(data)
