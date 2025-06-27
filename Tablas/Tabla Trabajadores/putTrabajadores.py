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
# RUTA DE LA FOTO PERFIL POR DEFECTO
FOTO_DEFAULT_PATH = r"C:\Users\jeanm\Downloads\apiconsumo25062025\assets\perfil.png"

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
    url = "http://192.168.2.142:8080/api/v1/workers" # "http://spx-enterprise.com.pe/api/api/v1/workers" "http://192.168.2.142:8080/api/v1/workers"
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
    url = "http://192.168.2.142:8080/api/v1/workers" # "http://spx-enterprise.com.pe/api/api/v1/workers" "http://192.168.2.142:8080/api/v1/workers"
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
def correct_dni(dni: str) -> str | None:
    if pd.isna(dni) or not dni:
        return None

    dni_limpio = ''.join(filter(str.isdigit, str(dni).strip()))

    if len(dni_limpio) == 7:
        dni_limpio = '0' + dni_limpio

    if len(dni_limpio) != 8:
        return None

    if dni_limpio in dnis_vistos_excel or dni_limpio in dnis_existentes_backend:
        return dni_limpio

    dnis_vistos_excel.add(dni_limpio)
    return dni_limpio

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

def enviar_put(url: str, payload: dict, foto_file, timeout: int = 10) -> requests.Response:
    """
    Envía un PUT con payload y foto usando multipart/form-data.
    payload se debe serializar como JSON y enviarse bajo el campo 'worker'.
    """
    fields = {
        "worker": ("worker.json", json.dumps(payload), "application/json")
    }

    if foto_file:
        filename, content, mime_type = foto_file
        fields["foto"] = (filename, content, mime_type)

    m = MultipartEncoder(fields=fields)

    headers = {'Content-Type': m.content_type}

    resp = requests.put(url, data=m, headers=headers, timeout=timeout)
    resp.raise_for_status()
    return resp

def repeated_corporate_mail(val: str) -> bool:
    if not val:
        return False
    val_normalizado = val.strip().lower()
    if val_normalizado in correos_vistos_excel:
        return True
    correos_vistos_excel.add(val_normalizado)

    return False

def obtener_id_por_dni(dni: str) -> int | None:
    """
    Consulta al backend por el trabajador con ese DNI y retorna su ID si existe.
    """
    url = "http://192.168.2.142:8080/api/v1/workers" # "http://spx-enterprise.com.pe/api/api/v1/workers" "http://192.168.2.142:8080/api/v1/workers"
    try:
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        trabajadores = resp.json()
        for t in trabajadores:
            if t.get("dni") == dni:
                return t.get("id")
        return None
    except requests.RequestException as e:
        logging.error(f"❌ Error al obtener ID por DNI: {e}")
        return None

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
        if not path or pd.isna(path) or not os.path.exists(path):
            path = FOTO_DEFAULT_PATH

        try:
            with open(path, 'rb') as f:
                filename = os.path.basename(path)
                return (filename, f.read(), "image/png")
        except Exception as e:
            logging.warning(f"[Foto inválida] No se pudo abrir la foto '{path}': {e}")
            # Devolver un archivo vacío para cumplir con el requerimiento del backend
            return ("vacio.png", b"", "image/png")

        
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
    
    def campo_opcional(val):
        if pd.isna(val) or str(val).strip() == "":
            return None
        return str(val).strip()

        
    payload = {
        "nombres": safe_str_required(fila.get("nombres"), min_len=2, max_len=100),
        "apellidos": safe_str_required(fila.get("apellidos"), min_len=2, max_len=100),
        "dni": validar_dni(fila.get("dni")),
        "fechaNacimiento": safe_date(fila.get("fechaNacimiento"), default="1990-01-01"),
        "correo": safe_str(fila.get("correo"), 70) or "correo@ficticio.com",
        "direccion": safe_str(fila.get("direccion"), 100) or "No especificada",
        "telefono": safe_str(fila.get("telefono"), 15) or "000000000",

        "sexo": validar_sexo(fila.get("sexo")),
        "area": safe_str_required(fila.get("area"), min_len=2, max_len=30),
        "status": safe_str(fila.get("status"), 1) or "v", #v=celda vacía
        "referencia": safe_str(fila.get("referencia"), 100) or "SIN REFERENCIA",
        "estadoCivil": validar_estado_civil(safe_str(fila.get("estadoCivil"), 20)),
        "cargo": safe_str(fila.get("cargo"), 80),
        "tipoTrabajador": validar_tipo_trabajador(safe_str(fila.get("tipoTrabajador"), 30)),
        "distrito": safe_str(fila.get("distrito"), 30),
        "correoCorporativo": validar_correo_corporativo(safe_str(fila.get("correoCorporativo"), 70)),
        "fechaInicioContrato": safe_date(fila.get("fechaInicioContrato")),
        "fechaInicioLaboral": safe_date(fila.get("fechaInicioLaboral")),
        "fechaFinContrato": safe_date(fila.get("fechaFinContrato")),
        "fechaInicioPerComputable": safe_date(fila.get("fechaInicioPerComputable")),
        "sueldo": validar_campo_sueldo(fila.get("sueldo")),
        "movilidad": safe_float(fila.get("movilidad")),
        "asignacionFamiliar": safe_bool(fila.get("asignacionFamiliar")),
        "urlDireccion": safe_str(fila.get("urlDireccion"), 255) or "https://longitudlargadecaracteres.com",
        "numeroHijos": safe_int(fila.get("numeroHijos")),

        "correoPersonal": campo_opcional(fila.get("correoPersonal")),
        "celular": campo_opcional(fila.get("celular")),
        "vacation": None,
        "sedes": [],
        "horarios": []
    }

    foto_path = fila.get("foto")
    foto_file = cargar_foto(foto_path)

    # Añadir campo requerido 'vacation' con valores por defecto
    payload["vacation"] = {
        "id": 0,  # Valor genérico para evitar errores si se espera un ID
        "fechaInicio": "2025-01-01",  # Fecha ficticia válida
        "fechaFin": "2025-01-15",     # Fecha ficticia válida
        "estado": True,               # Valor booleano requerido
        "workerId": 0                 # Este se sobreescribirá luego con el ID real
    }

    return payload, foto_file

def strict_dni(dni: str) -> str | None:
    if pd.isna(dni) or not str(dni).strip().isdigit():
        return None
    dni_str = str(dni).strip()
    if len(dni_str) == 7:
        dni_str = '0' + dni_str
    return dni_str if len(dni_str) == 8 else None


def main():
    # --- Configuración fija ---
    ENDPOINT = "http://192.168.2.142:8080/api/v1/workers" # "http://spx-enterprise.com.pe/api/api/v1/workers" "http://192.168.2.142:8080/api/v1/workers"

    # 1) Seleccionar archivo Excel mediante diálogo
    ruta_excel = seleccionar_archivo()

    # 2) Cargar datos desde Excel
    df = cargar_excel(ruta_excel)

    # 3) Recorrer cada fila y hacer POST
    for id, fila in df.iterrows():

        try:
            dni_raw = fila.get("dni")
            dni = strict_dni(dni_raw)
            if not dni:
                logging.warning(f"[Fila {id}] DNI no disponible. Se omite.")
                continue

            worker_id = obtener_id_por_dni(dni)
            if worker_id is None:
                logging.warning(f"Fila {id} No se encontró ID para el DNI {dni}. Se omite.")
                continue

            payload, foto_file = crear_payload(fila)
            if payload is None:
                continue

            payload["id"] = worker_id

            if "vacation" in payload:
                payload["vacation"]["workerId"] = worker_id

            endpoint = f"{ENDPOINT}/{worker_id}"
            logging.info(f"Actualizando ID={worker_id}")
            resp = enviar_put(endpoint, payload, foto_file)
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