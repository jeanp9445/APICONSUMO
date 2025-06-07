import os
import pandas as pd
import requests
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def cargar_archivo(tipo):
    Tk().withdraw()  # Oculta la ventana principal de tkinter
    archivo = askopenfilename(title=f"Seleccione el archivo {tipo}", filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
    
    if not archivo:
        print(f"Error: No se seleccionó ningún archivo {tipo}.")
        return None

    try:
        df = pd.read_excel(archivo)
        print(f"Archivo {tipo} cargado exitosamente: {archivo}")
        return df
    except Exception as e:
        print(f"Error al cargar el archivo {tipo}: {e}")
        return None

def leer_numeraciones():
    """Lee el archivo de numeraciones y retorna un diccionario con las numeraciones actuales."""
    numeraciones = {"celular": 900000000, "correoCorporativo": 1, "correoPersonal": 1}
    if os.path.exists("archivosTXT/numeraciones.txt"):
        with open("archivosTXT/numeraciones.txt", "r") as file:
            for line in file:
                campo, valor = line.strip().split("=")
                numeraciones[campo] = int(valor)
    return numeraciones

def actualizar_numeraciones(numeraciones):
    """Escribe las nuevas numeraciones en el archivo de texto."""
    with open("archivosTXT/numeraciones.txt", "w") as file:
        for campo, valor in numeraciones.items():
            file.write(f"{campo}={valor}\n")

def generar_archivo_producto(df_modelo, df_inyectable):
    columnas_modelo = df_modelo.columns.tolist()
    df_producto = pd.DataFrame(columns=columnas_modelo)
    
    print("Columnas en el archivo modelo:", columnas_modelo)
    print("Columnas disponibles en el archivo inyectable:", df_inyectable.columns.tolist())
    
    columnas_mapeo = {}
    for col_modelo in columnas_modelo:
        col_inyectable = input(f"¿Qué columna del inyectable asignas a '{col_modelo}'? (Déjalo vacío si no aplica): ").strip()
        if col_inyectable and col_inyectable in df_inyectable.columns:
            columnas_mapeo[col_modelo] = col_inyectable
        else:
            columnas_mapeo[col_modelo] = None
    
    # Cargar numeraciones desde el archivo
    numeraciones = leer_numeraciones()
    
    datos_ficticios = {
        "nombres": "sin nombre",
        "apellidos": "sin apellido",
        "dni": "sin dni",
        "estadoCivil": "Soltero",
        "fechaNacimiento": "2000-12-12",
        "cargo": "sin cargo",
        "tipoTrabajador": "Planilla",
        "direccion": "SIN DIRECCIÓN",
        "distrito": "SIN DISTRITO",
        "celular": numeraciones["celular"],
        "correoCorporativo": f"sincorreo{numeraciones['correoCorporativo']}@sanpiox.edu.pe",
        "correoPersonal": f"sincorreo{numeraciones['correoPersonal']}@gmail.com",
        "fechaInicioContrato": "2020-12-12",
        "fechaInicioLaboral": None,
        "fechaFinContrato": None,
        "fechaInicioPerComputable": None,
        "sueldo": 525,
        "movilidad": 500,
        "asignacionFamiliar": False,
        "numeroHijos": 0,
        "foto": os.path.join(os.getcwd(), "assets/perfil.png")
    }
    
    filas_nuevas = []
    incremento = 0
    for _, row in df_inyectable.iterrows():
        nueva_fila = {}
        for col in columnas_modelo:
            if columnas_mapeo[col]:
                valor = row.get(columnas_mapeo[col], datos_ficticios.get(col, "sin dato"))
                if col == "dni" and isinstance(valor, (int, str)):
                    valor = str(valor).zfill(8)
                nueva_fila[col] = valor
            else:
                if col == "celular":
                    nueva_fila[col] = numeraciones["celular"] + incremento
                elif col == "correoCorporativo":
                    nueva_fila[col] = datos_ficticios[col].replace("1", str(numeraciones['correoCorporativo'] + incremento))
                elif col == "correoPersonal":
                    nueva_fila[col] = datos_ficticios[col].replace("1", str(numeraciones['correoPersonal'] + incremento))
                else:
                    nueva_fila[col] = datos_ficticios.get(col, "sin dato")
        incremento += 1
        filas_nuevas.append(nueva_fila)
    
    df_producto = pd.concat([df_producto, pd.DataFrame(filas_nuevas)], ignore_index=True)
    
    archivo_producto = "archivosExcel/producto.xlsx"
    df_producto.to_excel(archivo_producto, index=False)
    print(f"Archivo producto generado: {archivo_producto}")
    
    # Actualizar las numeraciones
    numeraciones["celular"] += incremento
    numeraciones["correoCorporativo"] += incremento
    numeraciones["correoPersonal"] += incremento
    actualizar_numeraciones(numeraciones)
    
    return df_producto

def enviar_datos_a_api(df_producto):
    url = 'http://192.168.2.142:8080/api/v1/workers' #url = http://spx-enterprise.com.pe/api/api/v1/workers
    
    for _, row in df_producto.iterrows():
        archivos = {}
        if 'foto' in row and isinstance(row['foto'], str) and os.path.exists(row['foto']):
            archivos['foto'] = open(row['foto'], 'rb')

        datos = {col: row[col] for col in df_producto.columns if col != 'foto'}

        print(f"Enviando datos: {datos}")
        
        try:
            response = requests.post(url, data=datos, files=archivos)
            if response.status_code == 200:
                print(f"Datos enviados correctamente para {row['nombres']} {row['apellidos']}")
            else:
                print(f"Error al enviar los datos para {row['nombres']} {row['apellidos']}: {response.text}")
        except Exception as e:
            print(f"Error al enviar los datos: {e}")

def procesar_archivos():
    print("Cargando archivo modelo...")
    df_modelo = cargar_archivo("modelo")
    if df_modelo is None:
        return

    print("Cargando archivo inyectable...")
    df_inyectable = cargar_archivo("inyectable")
    if df_inyectable is None:
        return

    df_producto = generar_archivo_producto(df_modelo, df_inyectable)
    enviar_datos_a_api(df_producto)

if __name__ == '__main__':
    procesar_archivos()
