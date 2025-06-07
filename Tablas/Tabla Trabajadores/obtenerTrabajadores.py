import requests
import pandas as pd

import os

# ------------------------------------------------------------------
# CONFIGURACIÓN (ajusta si tu ruta o modelo cambian)
# ------------------------------------------------------------------
SWAGGER_URL = "http://192.168.2.142:8080/swagger.json"
MODEL_NAME  = "Worker"          # nombre exacto en swagger.json > definitions
OUT_FILE    = "archivosExcel/workerstableget.xlsx"

# ------------------------------------------------------------------
# AUXILIAR · Leer columnas desde swagger.json
# ------------------------------------------------------------------
def get_campos_desde_swagger() -> list[str]:
    """
    Devuelve una lista con los nombres de las propiedades del modelo `Worker`
    descrito en swagger.json. Si no puede encontrarlas, retorna [].
    """
    try:
        r = requests.get(SWAGGER_URL, timeout=10)
        r.raise_for_status()
        spec = r.json()

        # 1) Buscar en definitions
        defs = spec.get("definitions", {})
        modelo = defs.get(MODEL_NAME, {})
        if "properties" in modelo:
            return list(modelo["properties"].keys())

        # 2) Fallback: inspeccionar la respuesta 200 de GET /api/v1/workers
        path_item = spec.get("paths", {}).get("/api/v1/workers", {})
        get_op = path_item.get("get", {})
        resp200 = get_op.get("responses", {}).get("200", {})
        schema = resp200.get("schema", {})
        if schema.get("type") == "array":
            ref = schema.get("items", {}).get("$ref")
            if ref:
                ref_name = ref.split("/")[-1]
                modelo = defs.get(ref_name, {})
                if "properties" in modelo:
                    return list(modelo["properties"].keys())

    except Exception as e:
        print("No se pudo leer swagger.json:", e)

    return []   # Si todo falla

# ------------------------------------------------------------------
# AUXILIAR · Generar Excel vacío con encabezados
# ------------------------------------------------------------------
def generar_excel_encabezado() -> None:
    columnas = get_campos_desde_swagger()
    if columnas:
        print("Campos detectados:", ", ".join(columnas))
    else:
        print("⚠️  No se pudo determinar la estructura.")
    df = pd.DataFrame(columns=columnas)          # DataFrame sin filas
    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    df.to_excel(OUT_FILE, index=False)
    print("Archivo Excel (encabezados) generado:", OUT_FILE)


def obtener_registros():
    url = "http://192.168.2.142:8080/api/v1/workers" #

    try:
        #Realizamos la solicitud para obtener todos los registros
        print(f"Solicitando todos los registros...")
        response = requests.get(url, timeout=10)

        #Verificamos si la solicitud fue exitosa
        if response.status_code == 200:
            registros = response.json()
            
            #Si no hay registros en la respuesta
            if not registros:
                print("No se encontraron registros.")
                generar_excel_encabezado()  
                return
            
            #Convertimos los registros obtenidos a un DataFrame de pandas
            df_registros = pd.DataFrame(registros)
            
            #Convertimos los registros obtenidos a un DataFrame de Pandas
            archivo_excel = "archivosExcel/workerstableget.xlsx"
            df_registros.to_excel(archivo_excel, index=False)

            print(f"Archivo Excel generado: {archivo_excel}")
        else:
            print(f"Erro al obtener los registros. Código de estado: {response.status_code}")
            print(f"Detalles: {response.text}")

    except Exception as e:
        print(f"Ocurrió un error: {e}")
    
# Llamamos a la función para obtener todos los registros y guardarlos en un archivo Excel
obtener_registros()


