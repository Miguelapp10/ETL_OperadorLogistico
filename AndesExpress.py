import time
import datetime as dt  # Para fechas (opcional)
import glob  # Para jalar todo los archivos en una carpeta
import os  # Para trabajar con rutas
import locale
from io import StringIO  # Usado para definir función nueva
from pathlib import Path
import numpy as np  # Para operaciones matemáticas
import xlsxwriter  # Funcionalidad para trabajar con archivos Excel
from pandas import ExcelWriter  # Para exportar tabla a Excel
from xlsx2csv import Xlsx2csv  # Usado para definir función nueva
from decimal import Decimal
import re
import unidecode
import requests
import json
import pandas as pd
from datetime import datetime, timedelta
import geopandas as gpd
from shapely.geometry import Point, LineString
from requests.adapters import HTTPAdapter
from urllib3.exceptions import InsecureRequestWarning
from urllib3.util.retry import Retry
from io import BytesIO
from google.cloud import bigquery
import pytz
import warnings
from pandas_gbq import to_gbq 
from concurrent.futures import ThreadPoolExecutor, as_completed 

# Suprimir advertencias SSL
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
# URL de inicio de sesión y consulta
login_url = "https://andesexpress.pe/vAuthenticate_2.php"
consulta_url_datos = "https://andesexpress.pe/api_consulta_guia_export.php"

# Credenciales de usuario
payload_login = {
    "username": "",  # Usuario
    "password": "",    # Contraseña
}
# Simula un navegador con encabezados
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
}
# Configuración del proyecto de Google Cloud
project_id = ""
dataset_id = ""
table_id_2 = "AndesExpress"
# Inicia una sesión para mantener las cookies
session = requests.Session()
# Paso 1: Realizar inicio de sesión
login_response = session.post(login_url, data=payload_login, headers=headers, verify=False)
if login_response.status_code == 200 and "location.replace" in login_response.text:
    print("Inicio de sesión exitoso.")

# Configurar fechas inicial y final
fecha_fin = datetime.now()
fecha_inicio = fecha_fin - timedelta(days=30)
fecha_limite = datetime(2024,11, 1) # Fecha de inicio de los datos
# Bucle para descargar por tramos de 90 días
consolidated_df = pd.DataFrame()
 
while fecha_inicio >= fecha_limite:  # Comparación correcta (2023, 7, 1)
    print(f"Procesando rango: {fecha_inicio.strftime('%Y-%m-%d')} a {fecha_fin.strftime('%Y-%m-%d')}")
    consulta_payload = {
        "c_cli": "2663078",
        "f_ini": fecha_inicio.strftime("%Y-%m-%d"),
        "f_fin": fecha_fin.strftime("%Y-%m-%d"),
        "c_serv": "",
        "estado": "",
        "doccliente": "",
        "pedido": "",
    }

    # Paso 2: Realizar consulta a la API
    consulta_response = session.get(consulta_url_datos, params=consulta_payload, headers=headers, verify=False)
    if consulta_response.status_code == 200:
        print("Consulta exitosa.")
        # Si el contenido de la respuesta es un archivo Excel:
        if 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in consulta_response.headers.get('Content-Type', ''):
            try:
        # Leer el archivo Excel con pandas especificando el motor 'openpyxl'
                df = pd.read_excel(BytesIO(consulta_response.content), engine='openpyxl')
            # Renombrar la columna 'id' a 'vehicle'
                df = df.rename(columns={'Nro de guia':'Nro_de_guia',
                'Fecha':'Fecha',
                'Estado':'Estado',
                'Servicio':'Servicio',
                'Peso (Kg)':'Peso_kg',
                'Origen':'Origen',
                'Destino':'Destino',
                'Remitente':'Remitente',
                'Direccion Origen':'Direccion_Origen',
                'Consignatario':'Consignatario',
                'Direccion Destino':'Direccion_Destino',
                'Fecha de Entrega':'Fecha_de_Entrega',
                'Observacion Estado Actual':'Observacion_Estado_Actual',
                'Doc. Cliente':'Doc_Cliente',
                'Trasnporte':'Trasnporte',
                'Obs Adicional':'Obs_Adicional',
                'Nro Pedido':'Nro_Pedido',
                'Nro Consultora':'Nro_Consultora',
                'Ultima Ocurrencia':'Ultima_Ocurrencia',
                'Descripcion':'Descripcion',
                'Fecha Visita1':'Fecha_Visita1',
                'Observacion Visita1':'Observacion_Visita1',
                'Fecha Visita2':'Fecha_Visita2',
                'Observacion Visita2':'Observacion_Visita2',
                'Fecha Visita3':'Fecha_Visita3',
                'Observacion Visita3':'Observacion_Visita3',
                'Fecha Visita4':'Fecha_Visita4',
                'Observacion Visita4':'Observacion_Visita4',
                'Contador Visitas':'Contador_Visitas',
                'Nro Documento Destinatario':'Nro_Documento_Destinatario',
                'Departamento Destinatario':'Departamento_Destinatario',
                'Provincia Destinatario':'Provincia_Destinatario',
                'Distrito Destinatario':'Distrito_Destinatario',
                'Fecha Recojo':'Fecha_Recojo',
                'Coordenadas Entrega':'Coordenadas_Entrega',
                'Telefono':'Telefono',
                'Bultos':'Bultos',
                'Ubigeo Origen':'Ubigeo_Origen',
                'Ubigeo Destino':'Ubigeo_Destino'})
                # Asegurarse de que la columna 'Nro_Documento_Destinatario' sea de tipo 'string'
               
                # Selección de columnas y conversión a string
                df[['Nro_de_guia','Estado','Servicio','Peso_kg','Origen','Destino','Remitente','Direccion_Origen','Consignatario',
                    'Direccion_Destino', 'Observacion_Estado_Actual','Doc_Cliente','Trasnporte','Obs_Adicional','Nro_Pedido', 'Nro_Consultora',
                    'Ultima_Ocurrencia','Descripcion', 'Observacion_Visita1','Observacion_Visita2','Observacion_Visita3','Observacion_Visita4',
                    'Nro_Documento_Destinatario', 'Departamento_Destinatario','Provincia_Destinatario','Distrito_Destinatario', 
                    'Coordenadas_Entrega','Telefono','Bultos','Ubigeo_Origen','Ubigeo_Destino' ]] = \
                 df[['Nro_de_guia','Estado','Servicio','Peso_kg','Origen','Destino','Remitente','Direccion_Origen','Consignatario',
                    'Direccion_Destino', 'Observacion_Estado_Actual','Doc_Cliente','Trasnporte','Obs_Adicional','Nro_Pedido', 'Nro_Consultora',
                    'Ultima_Ocurrencia','Descripcion', 'Observacion_Visita1','Observacion_Visita2','Observacion_Visita3','Observacion_Visita4',
                    'Nro_Documento_Destinatario', 'Departamento_Destinatario','Provincia_Destinatario','Distrito_Destinatario', 
                    'Coordenadas_Entrega','Telefono','Bultos','Ubigeo_Origen','Ubigeo_Destino' ]].astype(str)

                # Concatenar datos al DataFrame consolidado
                if not df.empty:
                    consolidated_df = pd.concat([consolidated_df, df.dropna(how='all', axis=1)], ignore_index=True)

                print(f"Datos procesados para rango {fecha_inicio.strftime('%Y-%m-%d')} a {fecha_fin.strftime('%Y-%m-%d')}")
            except Exception as e:
                print(f"Error al procesar el archivo Excel para el rango: {e}")
        else:
            print("El formato de la respuesta no es un archivo Excel.")
    else:
        print(f"Error en la consulta. Código de estado: {consulta_response.status_code}")

    # Actualizar fechas para el siguiente tramo
    fecha_fin = fecha_inicio - timedelta(days=1)
    fecha_inicio = fecha_fin - timedelta(days=30)

# Cargar datos consolidados a BigQuery
if not consolidated_df.empty:
    to_gbq(consolidated_df,
           destination_table=f"{dataset_id}.{table_id_2}",
            project_id=project_id,
            if_exists="replace")
    print("Los datos consolidados se han cargado exitosamente en BigQuery.")
else:
    print("No se encontraron datos para cargar en BigQuery.")
