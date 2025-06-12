import pandas as pd
import requests
import logging
import sys
import tkinter as tk
from tkinter import filedialog
from datetime import datetime 
import json
import math
import os

# Configuración del logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s', handlers=[logging.StreamHandler(sys.stdout)])
# RUTA DE LA FOTO PERFIL POR DEFECTO
FOTO_DEFAULT_PATH = r"C:\Users\usuario\OneDrive\Escritorio\ApiConsumo\assets\perfil.png"

def seleccionar_archivo() -> str:

    """
    Abre una ventana de diálogo para seleccionar un archivo Excel.
    :return: ruta al fichero seleccionado
    """
    root = tk.Tk()
    root.withdraw() #Oculta la ventana principal
    ruta = filedialog.askopenfilename(
        title="Selecciona el archivo Excel de trabajadores",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not ruta:
        logging.error("No se ha seleccionado ningún archivo.")
        sys.exit(1)
    logging.info(f"Archivo seleccionado: {ruta}")
    return ruta

def cargar_excel(ruta_excel: str, hoja: str = None) -> pd.DataFrame:
    """
    Carga el archivo Excel en un DataFrame de pandas.
    :param ruta_excel: ruta al fichero .xlsx
    :param hoja: nombre o índice de la hoja (por defecto la primera)
    :return: DataFrame con los datos
    """
    try:
        df = pd.read_excel(ruta_excel, engine='openpyxl')
        logging.info(f"Leídas {len(df)} filas desde '{ruta_excel}'")
        return df
    except Exception as e:
        logging.error(f"Error al leer el archivo Excel: {e}")
        sys.exit(1)

def dic_a_multipart(dic: dict) -> dict:
    """
    Convierte {'campo':'valor'} -> {'campo': (None, 'valor')}
    para que requests envíe multipart/form-data.
    """
    return {k: (None, str(v)) for k, v in dic.items()}
    
def enviar_post(url: str, payload: dict, foto_file=None, timeout: int = 10) -> requests.Response:
    """
    Envía un POST con el payload JSON a la URL indicada.
    :param url: endpoint al que hacer POST
    :param payload: diccionario con los datos a enviar
    :param timeout: segundos antes de timeout
    :return: objeto Response de requests
    """

    try:
        files = dic_a_multipart(payload)
        if foto_file:
            files["foto"] = foto_file  # Añadir la foto si se proporciona

        resp = requests.post(url, files=files, timeout=timeout)
        resp.raise_for_status() # Lanza un error si la respuesta no es 200
        return resp
    except requests.RequestException as e:
        logging.error(f"Error en POST a {url}: {e}")
        raise

def crear_payload(fila: pd.Series) -> dict:
    def safe_str(val, max_len=None):
        if pd.isna(val):
            return ""
        val = str(val).strip()
        return val[:max_len] if max_len else val
    
    def safe_str_required(val, min_len=2, max_len=None):
        if pd.isna(val):
            return None
        val = str(val).strip() 
        if len(val) < min_len:
            return None
        return val[:max_len] if max_len else val
    
    def safe_date(val, default="1900-01-01"):
        if pd.isna(val) or val == "":
            return default
        if isinstance(val, pd.Timestamp):
            return val.date().isoformat()
        if isinstance(val, str):
            try:
                return pd.to_datetime(val).date().isoformat()
            except:
                return default
        
        return default
    
    def safe_float(val):
        try:
            f = float(val)
            return f if not math.isnan(f) else 0.0
        except:
            return 0.0
        
    def safe_float_min(val, min_val=0.0):
        try:
            f = float(val)
            if math.isnan(f) or f < min_val:
                return None
            return f
        except:
            return None
        
    def safe_bool(val):
        if isinstance(val, bool):
            return val
        if isinstance(val, str):
            return val.strip().lower() in ['true', '1', 'yes', 'si', 'sí', 'false', 'verdadero', 'falso', 'no']
        if isinstance(val, (int, float)):
            return val != 0
        return False

    def safe_int(val):
        try:
            return int(val)
        except:
            return 0
        
    def validar_campo_sueldo(campo:str) -> float:
        sueldo_val = safe_float(fila.get("sueldo"))
        if sueldo_val is None or sueldo_val < 525:
            logging.warning(f"[Fila inválida] Sueldo inválido: {sueldo_val} — se omite esta fila")
            sueldo_val = 50000  # Valor por defecto si no es válido
        return sueldo_val
    
    def validar_sexo(val) -> str:
        if pd.isna(val):
            return None
        val = str(val).strip().lower()
        if val in ['m', 'masculino', 'hombre']:
            return 'M'
        if val in ['f', 'femenino', 'mujer']:
            return 'F'
        else:
            logging.warning(f"[Campo inválido] Valor de 'sexo' no reconocido: {val}")
            return None
        
    def validar_estado_civil(val) -> str:
        val = str(val).strip().lower()
        match val:
            case 'soltero' | 'SOLTERO':
                return "Soltero"
            case 'casado' | 'CASADO':
                return "Casado"
            case 'divorciado' | 'DIVORCIADO':
                return "Divorciado"
            case 'conviviente' | 'CONVIVIENTE':
                return "Conviviente"
            case _:
                return "Soltero"  # Valor por defecto si no es válido

    def validar_tipo_trabajador(tipoTrabajador: str) -> str:
        tipoTrabajador = str(tipoTrabajador).strip().lower()
        match tipoTrabajador:
            case 'empleado' | 'EMPLEADO':
                return "Empleado"
            case 'ejecutivo' | 'EJECUTIVO':
                return "Ejecutivo"
            case _:
                return "Empleado"  # Valor por defecto si no es válido

    def cargar_foto(path: str):
        if not path or pd.isna(path):
            path = os.path.join(os.getcwd(), "assets/perfil.png")
        try:
            with open(path, 'rb') as f:
                filename = os.path.basename(path)
                return (filename, f.read(), "image/png")
        except Exception as e:
            logging.warning(f"[Foto inválida] No se pudo abrir la foto '{path}': {e}")
            return None
        
    def validar_correo_corporativo(val: str) -> str:
        if pd.isna(val):
            return None
        return val.strip().lower() if isinstance(val, str) else None
        
    payload = {
        "nombres": safe_str(fila.get("nombres"), 40),
        "apellidos": safe_str(fila.get("apellidos"), 40),
        "dni": safe_str(fila.get("dni"), 8),
        "sexo": validar_sexo(fila.get("sexo")),
        "area": safe_str_required(fila.get("area"), min_len=2, max_len=30),
        "status": safe_str(fila.get("status"), 1) or "v", #v=celda vacía
        "referencia": safe_str(fila.get("referencia"), 100) or "SIN REFERENCIA",
        "estadoCivil": validar_estado_civil(safe_str(fila.get("estadoCivil"), 20)),
        "fechaNacimiento": safe_date(fila.get("fechaNacimiento")),
        "cargo": safe_str(fila.get("cargo"), 80),
        "tipoTrabajador": validar_tipo_trabajador(safe_str(fila.get("tipoTrabajador"), 30)),
        "direccion": safe_str(fila.get("direccion"), 200),
        "distrito": safe_str(fila.get("distrito"), 30),
        "celular": safe_str(fila.get("celular"), 9),
        "correoCorporativo": validar_correo_corporativo(safe_str(fila.get("correoCorporativo"), 70)),
        "correoPersonal": safe_str(fila.get("correoPersonal"), 100),
        "fechaInicioContrato": safe_date(fila.get("fechaInicioContrato")),
        "fechaInicioLaboral": safe_date(fila.get("fechaInicioLaboral")),
        "fechaFinContrato": safe_date(fila.get("fechaFinContrato")),
        "fechaInicioPerComputable": safe_date(fila.get("fechaInicioPerComputable")),
        "sueldo": validar_campo_sueldo(fila.get("sueldo")),
        "movilidad": safe_float(fila.get("movilidad")),
        "asignacionFamiliar": safe_bool(fila.get("asignacionFamiliar")),
        "urlDireccion": safe_str(fila.get("urlDireccion"), 255) or "https://longitudlargadecaracteres.com",
        "numeroHijos": safe_int(fila.get("numeroHijos")),
    }

    foto_path = fila.get("foto")
    foto_file = cargar_foto(foto_path)

    return payload, foto_file

def main():
    # --- Configuración fija ---
    ENDPOINT = "http://192.168.2.142:8080/api/v1/workers" # Ajusta la URL de tu API

    # 1) Seleccionar archivo Excel mediante diálogo
    ruta_excel = seleccionar_archivo()

    # 2) Cargar datos desde Excel
    df = cargar_excel(ruta_excel)

    # 3) Recorrer cada fila y hacer POST
    for id, fila in df.iterrows():

        try:
            payload, foto_file = crear_payload(fila)
            if payload is None:
                continue

            logging.info(f"Enviado trabajador: {payload['nombres']} {payload['apellidos']}")
            resp = enviar_post(ENDPOINT, payload, foto_file=foto_file)
            logging.info(f"✔️ {resp.status_code} - {resp.text}")

        except requests.exceptions.HTTPError as e:

            if e.response is not None:

                try:
                    error_json = e.response.json()
                    mensaje = error_json.get("message", "")

                    # 🔇 Omitir mensaje si ya existe el trabajador
                    if "ya existe con el mismo nombre y apellido" in mensaje:
                        logging.info(f"Procesando fila ID: {id}")
                        pass
                    else:
                        error_pretty = json.dumps(error_json, indent=2, ensure_ascii=False)
                        logging.error(f"❌ Error de API - Código {e.response.status_code}:\n{error_pretty}")
                        logging.info(f"Procesando fila ID: {id}")



                    """""
                    error_json = e.response.json()
                    error_pretty = json.dumps(error_json, indent=2, ensure_ascii=False)
                    logging.error(f"❌ Error de API - Código {e.response.status_code}:\n{error_pretty}")"""
                except ValueError:
                    logging.error(f"❌ Error de API - Cód {e.response.status_code} - {e.response.text}")
            else:
                logging.error(f"❌ Error HTTP inesperado: {e}")

            print("\n" + "-" * 80 + "\n")
            continue

if __name__ == '__main__':
    main()