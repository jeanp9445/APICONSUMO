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
import re
from requests_toolbelt.multipart.encoder import MultipartEncoder

# Configuración del logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s', handlers=[logging.StreamHandler(sys.stdout)])

#Memoria temporal de correos_corporativos
contador_correo_ficticio = 1
correos_vistos_excel = set()
dnis_vistos_excel = set()
dni_ficticio_actual = 99999999

def correo_ficticio() -> str:
    global contador_correo_ficticio
    backend_correos = obtener_correos_existentes()

    while True:
        correo = f"ficticio{contador_correo_ficticio}@sanpiox.edu.pe"
        contador_correo_ficticio += 1

        # Verificar que no se repita en Excel cargado ni en BD
        if correo not in correos_vistos_excel and correo not in backend_correos:
            correos_vistos_excel.add(correo)
            return correo
        
def obtener_correos_existentes() -> set:
    url = "http://192.168.100.5:8080/api/v1/users"
    try:
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        correos = {w["correoCorporativo"].strip().lower() for w in data if w.get("correoCorporativo")}
        return correos
    except requests.RequestException as e:
        logging.error(f"❌ Error al obtener correos existentes: {e}")
        return set()

def obtener_dnis_existentes() -> set:
    url = "http://192.168.100.5:8080/api/v1/users"
    try:
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        dnis = {w["dni"] for w in data if w.get("dni")}
        return dnis
    except requests.RequestException as e:
        logging.error(f"❌ Error al obtener DNIs existentes: {e}")
        return set()

dnis_existentes_backend = obtener_dnis_existentes()

'''Limpiar el dni considerando: añadir un cero incial si son 7 caracteres, validar que solo sean numeros y 
    si no lo son eliminar los carteres especiales'''
def correct_dni(dni: str) -> str:
    global dni_ficticio_actual

    if pd.isna(dni) or not dni:
        return generar_dni_ficticio()

    dni_limpio = ''.join(filter(str.isdigit, str(dni).strip()))

    if len(dni_limpio) == 7:
        dni_limpio = '0' + dni_limpio

    if len(dni_limpio) != 8:
        return generar_dni_ficticio()

    if dni_limpio in dnis_vistos_excel or dni_limpio in dnis_existentes_backend:
        return generar_dni_ficticio()

    dnis_vistos_excel.add(dni_limpio)
    return dni_limpio

def generar_dni_ficticio() -> str:
    global dni_ficticio_actual
    
    while (
        str(dni_ficticio_actual) in dnis_vistos_excel 
        or str(dni_ficticio_actual) in dnis_existentes_backend
    ):
        dni_ficticio_actual -= 1
        if dni_ficticio_actual < 90000000:
            raise ValueError("❌ Se agotaron los DNIs ficticios disponibles")
        
    dnis_vistos_excel.add(str(dni_ficticio_actual))
    
    return str(dni_ficticio_actual)

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

def enviar_post(url: str, payload: dict, timeout: int = 10) -> requests.Response:
    """
    Envía un POST con el payload en formato JSON (application/json).
    """
    try:
        headers = {'Content-Type': 'application/json'}
        resp = requests.post(url, json=payload, headers=headers, timeout=timeout)
        resp.raise_for_status()
        return resp
    except requests.RequestException as e:
        logging.error(f"Error en POST a {url}: {e}")
        raise

        raise

def repeated_corporate_mail(val: str) -> bool:
    if not val:
        return False
    val_normalizado = val.strip().lower()
    if val_normalizado in correos_vistos_excel:
        return True
    correos_vistos_excel.add(val_normalizado)

    return False

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
        
    def validar_correo_corporativo(val: str) -> str:
        if repeated_corporate_mail(val):
            return correo_ficticio()
        else:
            if pd.isna(val) or val == "":
                return correo_ficticio()

            return val

    def validar_dni(dni: str) -> str:
        if pd.isna(dni):
            return None
        
        clean_dni = str(dni).strip().lower()
        valid_dni = correct_dni(clean_dni)

        return valid_dni

    payload = {
        "nombres": safe_str(fila.get("nombres"), 40),
        "apellidos": safe_str(fila.get("apellidos"), 40),
        "dni": validar_dni(fila.get("dni")),
        "sexo": validar_sexo(fila.get("sexo")),
        "fechaNacimiento": safe_date(fila.get("fechaNacimiento")),
        "direccion": safe_str(fila.get("direccion"), 200),
        "telefono": safe_str(fila.get("celular"), 9),
        "correo": safe_str(fila.get("correoPersonal"),30),
    }

    dni_final = payload["dni"] if payload["dni"] else "00000000"
    ultimos_4 = dni_final[-4:] if len(dni_final) >= 4 else "0000"
    primer_nombre = payload["nombres"].split()[0].lower() if payload["nombres"] else "usuario"

    payload["username"] = f"{primer_nombre}.{ultimos_4}"
    payload["password"] = f"SanPioX{ultimos_4}"
    payload["roles"] = []
    payload["sedes"] = []

    return payload

def main():
    # --- Configuración fija ---
    ENDPOINT = "http://192.168.100.5:8080/api/v1/users" # Ajusta la URL de tu API

    # 1) Seleccionar archivo Excel mediante diálogo
    ruta_excel = seleccionar_archivo()

    # 2) Cargar datos desde Excel
    df = cargar_excel(ruta_excel)

    # 3) Recorrer cada fila y hacer POST
    for id, fila in df.iterrows():

        try:
            payload = crear_payload(fila)
            if payload is None:
                continue

            logging.info(f"Enviado trabajador: {payload['nombres']} {payload['apellidos']}")
            resp = enviar_post(ENDPOINT, payload)
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