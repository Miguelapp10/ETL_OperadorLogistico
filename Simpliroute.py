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
# Definir las zonas horarias
utc = pytz.UTC
lima_tz = pytz.timezone("America/Lima")
# Suprimir sólo las advertencias de InsecureRequestWarning
warnings.filterwarnings('ignore', category=InsecureRequestWarning)

# Información de autenticación
token = ''

base_url_vehicles = 'https://api.simpliroute.com/v1/routes/vehicles/'
base_url_routes = 'https://api.simpliroute.com/v1/routes/routes/'
base_url_visits = 'https://api.simpliroute.com/v1/routes/visits/'

# Encabezados de la solicitud
headers = {
    'Content-Type': 'application/json',
    'Authorization': f'Token {token}'
}

# Función para generar un rango de fechas
def date_range(start_date, end_date):
    for n in range(int((end_date - start_date).days) + 1):
        yield start_date + timedelta(n)

# Definir el rango de fechas
start_date = datetime.strptime('2025-01-01', '%Y-%m-%d') ##'2024-04-01'
end_date = datetime.today()  # Usar la fecha actual para end_date

# Función para obtener datos de una URL en una fecha específica
def fetch_data_for_date(base_url, single_date):
    formatted_date = single_date.strftime('%Y-%m-%d')
    url = f'{base_url}?planned_date={formatted_date}'

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        try:
            response_data = response.json()
            if response_data:
                return response_data  # Devolver los datos obtenidos
        except json.JSONDecodeError:
            return []
    return []

# Función para obtener datos de una URL en un rango de fechas usando ThreadPoolExecutor
def get_data_parallel(base_url, start_date, end_date):
    all_data = []
    dates = list(date_range(start_date, end_date))

    with ThreadPoolExecutor(max_workers=20) as executor:
        futures = [executor.submit(fetch_data_for_date, base_url, date) for date in dates]
        for future in futures:
            result = future.result()
            if result:
                all_data.extend(result)

    return all_data


# Obtener datos de ambos endpoints en paralelo
all_vehicles_data = get_data_parallel(base_url_vehicles, start_date, end_date)
all_routes_data = get_data_parallel(base_url_routes, start_date, end_date)
all_visits_data = get_data_parallel(base_url_visits, start_date, end_date)

# Convertir todos los datos a DataFrames de pandas
df_vehicles = pd.json_normalize(all_vehicles_data)
df_routes = pd.json_normalize(all_routes_data)
df_visits = pd.json_normalize(all_visits_data)

# Seleccionar las columnas deseadas de cada DataFrame
df_vehicles_selected = df_vehicles[['id','name',
                                   #'shift_start','shift_end','capacity','capacity_2','capacity_3','default_driver','location_start_address',
                                   #'location_start_latitude','location_start_longitude','location_end_address','location_end_latitude',
                                   #'location_end_longitude','skills','created','modified',
                                   'color',
                                   #'reference_id','license_plate','cost','min_load','min_load_2','min_load_3','max_visit',
                                   #'rest_time_start','rest_time_end','rest_time_duration','deleted','max_time','codrivers','status',
                                   ]].drop_duplicates().rename(columns={'id': 'vehicle'})

df_routes_selected = df_routes[['id','vehicle','driver','plan','status','planned_date','estimated_time_start',
                                'estimated_time_end','total_duration','total_distance','total_load',
                                'total_load_percentage','total_load_2','total_load_2_percentage','total_load_3',
                                'total_load_3_percentage','location_start_address','location_start_latitude',
                                'location_start_longitude','location_end_address','location_end_latitude',
                                'location_end_longitude','comment','start_time','end_time','created','modified',
                                'kilometers','total_visits','latitude_init','longitude_init','latitude_finish',
                                'longitude_finish','is_revised',
                                'reference']].drop_duplicates().rename(columns={'id': 'route',
                                'status':'status_route','created':'created_route','modified':'modified_route',
                                'reference':'reference_x'})

df_routes_selected.loc[:, 'comment'] = df_routes_selected['comment'].fillna('')
df_routes_selected.loc[:, 'reference_x'] = df_routes_selected['reference_x'].fillna('')

# Seleccionar las columnas deseadas
selected_columns_v1 = ['id', 'order', 'tracking_id', 'status', 'title', 'address',  'latitude', 'longitude',
                    #'load', 'load_2', 'load_3', 'window_start', 'window_end', 'window_start_2', 'window_end_2', 'duration',
                    'contact_name', 'contact_phone', 'contact_email','reference', 'notes', #'skills_required','skills_optional','tags',
                    'planned_date', 'programmed_date', 'route',
                    'route_estimated_time_start', 'estimated_time_arrival', 'estimated_time_departure', 'checkin_time',
                    'checkout_time', 'checkout_latitude', 'checkout_longitude', 'checkout_comment',
                    'checkout_observation',# 'signature','pictures',
                    'created', 'modified', 'eta_predicted','eta_current', 'driver', 'vehicle', 'priority','has_alert',
                    'priority_level','geocode_alert','visit_type','current_eta','fleet','on_its_way','seller',
                    'extra_field_values.visitado',
                    'extra_field_values',
                        ]
#df_visits_selected = df_visits[selected_columns_v1] 

df_visits_selected = df_visits[selected_columns_v1].copy()

df_visits_selected.loc[:, 'programmed_date'] = df_visits_selected['programmed_date'].fillna('')
df_visits_selected.loc[:, 'geocode_alert'] = df_visits_selected['geocode_alert'].fillna('')
df_visits_selected.loc[:, 'visit_type'] = df_visits_selected['visit_type'].fillna('')
df_visits_selected.loc[:, 'fleet'] = df_visits_selected['fleet'].fillna('')
df_visits_selected.loc[:, 'seller'] = df_visits_selected['seller'].fillna('')

# Renombrar las columnas con nombres no válidos
df_visits_selected.columns = [
    col.replace('.', '_') for col in df_visits_selected.columns
]
# Renombrar la columna 'id' a 'vehicle'
df_visits_selected = df_visits_selected.rename(columns={'order': 'order_visit'})

# Función para mapear valores de checkout_observation a una nueva columna
def map_observation_to_new_column(observation):
    mapping = {
        '1a1d65aa-d355-45b6-8c3f-3f2295ee4c5a':'Producto no corresponde',
        '56a04e5b-2fc5-42df-b4dc-6ef75d97f63c':'Cliente no quiere identificarse',
        '6084c66c-c720-4136-b20f-e01c80a73378':'Fallas mecánicas',
        '830808ee-2ef6-4c96-973c-751d530ba0f9':'Entregado a conserje',
        '9dc634d3-865f-470b-92c0-14fb63a40637':'Fuera De Horario',
        'e4d21dd5-1107-4c99-a9b7-fae1bac83882':'Entregado a familiar',
        'f97966aa-47f5-4c4d-8d42-1b6df9729157':'Entregado a titular',
        'c505bc38-1215-48bf-9e6b-d78bce3dc2f2':'Cliente Solicito Reprogramación',
        '8c15fdad-2f27-49eb-9967-f4fa94a366b1':'Entregado en tienda',
        '04902a7e-116c-4083-82bb-2849e1928db3':'Cliente Ausente',
        'fe8ac91f-0ead-4d84-b33c-aa094173d32b':'Domicilio Sin Acceso',
        '2fee0a67-e6d8-4742-ac52-f8c4f0382543':'Rechazado Por Cliente',
        '412375a8-cd64-4e8a-81be-5572ca883018':'Producto Dañado',
        '15d8ba69-4dff-40c8-9c27-28f8c3d09779':'Distrito errado',
        'a5da6d6e-964e-4a87-b4b8-3ba3a970938c':'Dirección Incorrecta',
        '11feca76-3da6-443e-88cd-5becc64821c8':'Entregado a titular',
        'd2be558b-1910-47fc-ae6b-1ae0c32e60bf':'Fuera de Ruta',
        'bc17ba7d-f735-4aec-ab8f-2baac8772c00':'Siniestro OL',
        '0a782b98-7998-48bf-89d7-7be410a98a55':'Dirección de Agencia',
        'c8be72af-c364-4169-9742-1346519b6a86':'Zona peligrosa',
        'c5ed62de-19fa-4c47-8e7c-193f3fd8573e':'Sin Información',
        '086f1611-2ee6-48c7-984a-5106acc8b631':'Fuera De Horario',
        '362584e0-8fdf-4a59-82b5-f6283745c4f8':'Rechazado Por Cliente',
        '34641928-f1b8-4f19-af63-2077649a9d5e':'Cliente Ausente'
        }
    return mapping.get(observation, 'Otro')

# Aplicar la función al DataFrame
df_visits_selected['Observaciones'] = df_visits_selected['checkout_observation'].apply(map_observation_to_new_column)

# Unir los DataFrames en el campo común 'vehicle'
df_routes_vehicles = pd.merge(df_routes_selected, df_vehicles_selected, on='vehicle', how='inner')

# Unir los DataFrames en el campo común 'vehicle'
df_visits_routes_vehicles = pd.merge(df_visits_selected, df_routes_vehicles, on=('vehicle','driver' ,'route', 'planned_date'), how='left')

# Convertir 'contact_name' a mayúsculas
df_visits_routes_vehicles['tienda_name'] = df_visits_routes_vehicles['contact_name'].str.upper()

# Función para extraer la parte derecha de 'title' después del primer '-' o dejarlo igual si no hay '-'
def extract_tienda_name(row):
    if pd.isna(row['tienda_name']) or row['tienda_name'].strip() == '':
        if '-' in row['title']:  # Verifica si hay un guion antes de hacer split
            return row['title'].split('-', 1)[1].strip()
        else:
            return row['title'].strip()  # Si no hay guion, usa el título completo
    return row['tienda_name']

# Aplicar la función
df_visits_routes_vehicles['tienda_name'] = df_visits_routes_vehicles.apply(extract_tienda_name, axis=1)

# Función que intenta convertir las fechas considerando microsegundos o no
def convert_to_datetime(date_str):
    try:
        return pd.to_datetime(date_str, format="%Y-%m-%dT%H:%M:%S.%fZ", errors='raise')
    except ValueError:
        return pd.to_datetime(date_str, format="%Y-%m-%dT%H:%M:%SZ", errors='raise')

# Convertir las columnas de fecha
df_visits_routes_vehicles["current_eta"] = df_visits_routes_vehicles["current_eta"].apply(convert_to_datetime)
df_visits_routes_vehicles["checkin_time"] = df_visits_routes_vehicles["checkin_time"].apply(convert_to_datetime)
df_visits_routes_vehicles["checkout_time"] = df_visits_routes_vehicles["checkout_time"].apply(convert_to_datetime)
df_visits_routes_vehicles["created"] = df_visits_routes_vehicles["created"].apply(convert_to_datetime)
df_visits_routes_vehicles["modified"] = df_visits_routes_vehicles["modified"].apply(convert_to_datetime)
df_visits_routes_vehicles["start_time"] = df_visits_routes_vehicles["start_time"].apply(convert_to_datetime)
df_visits_routes_vehicles["created_route"] = df_visits_routes_vehicles["created_route"].apply(convert_to_datetime)
df_visits_routes_vehicles["modified_route"] = df_visits_routes_vehicles["modified_route"].apply(convert_to_datetime)

# Restar 5 horas
df_visits_routes_vehicles["current_eta"] = df_visits_routes_vehicles["current_eta"] - timedelta(hours=5)
df_visits_routes_vehicles["checkin_time"] = df_visits_routes_vehicles["checkin_time"] - timedelta(hours=5)
df_visits_routes_vehicles["checkout_time"] = df_visits_routes_vehicles["checkout_time"] - timedelta(hours=5)
df_visits_routes_vehicles["created"] = df_visits_routes_vehicles["created"] - timedelta(hours=5)
df_visits_routes_vehicles["modified"] = df_visits_routes_vehicles["modified"] - timedelta(hours=5)
df_visits_routes_vehicles["start_time"] = df_visits_routes_vehicles["start_time"] - timedelta(hours=5)
df_visits_routes_vehicles["created_route"] = df_visits_routes_vehicles["created_route"] - timedelta(hours=5)
df_visits_routes_vehicles["modified_route"] = df_visits_routes_vehicles["modified_route"] - timedelta(hours=5)

project_id = ''
# Configuración del dataset y la tabla en BigQuery
dataset_id = ''
table_id = ''

# Carga el DataFrame a BigQuery
to_gbq(df_visits_routes_vehicles,destination_table=f"{dataset_id}.{table_id}",
                   project_id=project_id,
                   if_exists="replace")  # Opciones: "append", "replace", "fail"
print("Los datos se han cargado exitosamente en BigQuery.")

#################################################################################################################################################
#################################################################################################################################################

# Información de autenticación
token_qolqas = '42898f92de610efb5b6455f6562d9b61a324e08d'
base_url_vehicles_qolqas = 'https://api.simpliroute.com/v1/routes/vehicles/'
base_url_routes_qolqas = 'https://api.simpliroute.com/v1/routes/routes/'
base_url_visits_qolqas = 'https://api.simpliroute.com/v1/routes/visits/'

# Encabezados de la solicitud
headers_qolqas = {
    'Content-Type': 'application/json',
    'Authorization': f'Token {token_qolqas}'
}

# Función para generar un rango de fechas
def date_range(start_date_qolqas, end_date_qolqas):
    for n in range(int((end_date_qolqas - start_date_qolqas).days) + 1):
        yield start_date_qolqas + timedelta(n)

# Definir el rango de fechas
start_date_qolqas = datetime.strptime('2025-01-01', '%Y-%m-%d') ## '2024-04-01'
end_date_qolqas  = datetime.today()  # Usar la fecha actual para end_date

# Función para obtener datos de una URL en una fecha específica
def fetch_data_for_date(base_url_qolqas, single_date_qolqas):
    formatted_date_qolqas = single_date_qolqas.strftime('%Y-%m-%d')
    url_qolqas = f'{base_url_qolqas}?planned_date={formatted_date_qolqas}'

    response = requests.get(url_qolqas, headers=headers_qolqas)
    if response.status_code == 200:
        try:
            response_data = response.json()
            if response_data:
                return response_data  # Devolver los datos obtenidos
        except json.JSONDecodeError:
            return []
    return []

# Función para obtener datos de una URL en un rango de fechas usando ThreadPoolExecutor
def get_data_parallel(base_url_qolqas, start_date_qolqas, end_date_qolqas):
    all_data_qolqas = []
    dates_qolqas = list(date_range(start_date_qolqas, end_date_qolqas))

    with ThreadPoolExecutor(max_workers=20) as executor:
        futures_qolqas = [executor.submit(fetch_data_for_date, base_url_qolqas, date) for date in dates_qolqas]
        for future in futures_qolqas:
            result = future.result()
            if result:
                all_data_qolqas.extend(result)

    return all_data_qolqas


# Obtener datos de ambos endpoints en paralelo
all_vehicles_data_qolqas = get_data_parallel(base_url_vehicles_qolqas, start_date_qolqas, end_date_qolqas)
all_routes_data_qolqas = get_data_parallel(base_url_routes_qolqas, start_date_qolqas, end_date_qolqas)
all_visits_data_qolqas = get_data_parallel(base_url_visits_qolqas, start_date_qolqas, end_date_qolqas)

# Convertir todos los datos a DataFrames de pandas
df_vehicles_qolqas = pd.json_normalize(all_vehicles_data_qolqas)
df_routes_qolqas = pd.json_normalize(all_routes_data_qolqas)
df_visits_qolqas = pd.json_normalize(all_visits_data_qolqas)

# Seleccionar las columnas deseadas de cada DataFrame
df_vehicles_qolqas_selected = df_vehicles_qolqas[['id','name',
                                    #'capacity','capacity_2','default_driver','location_start_address',
                                    #'location_start_latitude','location_start_longitude','location_end_address',
                                    #'location_end_latitude','location_end_longitude','created','modified',
                                    'color']].drop_duplicates()
# Renombrar la columna 'id' a 'vehicle'
df_vehicles_qolqas_selected = df_vehicles_qolqas_selected.rename(columns={'id': 'vehicle'})

df_routes_qolqas_selected = df_routes_qolqas[['id','vehicle','driver','plan','status','planned_date','estimated_time_start',
                                'estimated_time_end','total_duration','total_distance','total_load','total_load_percentage',
                                'location_start_address','location_start_latitude','location_start_longitude',
                                'location_end_address','location_end_latitude','location_end_longitude','start_time',
                                'end_time','created','modified','kilometers','total_visits','latitude_init','longitude_init',
                                'latitude_finish','longitude_finish'
]].drop_duplicates()  # Reemplaza 'other_columns' con las columnas deseadas de df_routes
# Renombrar la columna 'id' a 'vehicle'
df_routes_qolqas_selected = df_routes_qolqas_selected.rename(columns={'id': 'route','status':'status_route','created':'created_route','modified':'modified_route'})

# Seleccionar las columnas deseadas
selected_columns = ['id', 'order', 'tracking_id', 'status', 'title', 'address',  'latitude', 'longitude',
                    #'load', 'load_2', 'load_3', 'window_start', 'window_end', 'window_start_2', 'window_end_2', 'duration',
                        'contact_name', 'contact_phone', #'contact_email',
                        'reference', 'notes', #'skills_required','skills_optional','tags',
                        'planned_date',#'programmed_date',
                        'route','route_estimated_time_start',
                        'estimated_time_arrival','estimated_time_departure','checkin_time','checkout_time','checkout_latitude',
                        'checkout_longitude',#'checkout_comment',
                        'checkout_observation',#'signature','pictures',
                        'created','modified','eta_predicted','eta_current','driver','vehicle',#'priority','has_alert',
                        'priority_level',
                        'geocode_alert',#'visit_type','current_eta','fleet',
                        'on_its_way',#'seller','extra_field_values',
                        ]
df_visits_qolqas_selected = df_visits_qolqas[selected_columns]

# Renombrar la columna 'id' a 'vehicle'
df_visits_qolqas_selected = df_visits_qolqas_selected.rename(columns={'order': 'order_visit'})

# Aplicar la función al DataFrame
df_visits_qolqas_selected['Observaciones'] = df_visits_qolqas_selected['checkout_observation'].apply(map_observation_to_new_column)

# Unir los DataFrames en el campo común 'vehicle'
df_routes_vehicles_qolqas = pd.merge(df_routes_qolqas_selected, df_vehicles_qolqas_selected, on='vehicle', how='inner')

# Unir los DataFrames en el campo común 'vehicle'
df_visits_routes_vehicles_qolqas = pd.merge(df_visits_qolqas_selected,df_routes_vehicles_qolqas , on=('vehicle','driver' ,'route', 'planned_date'), how='left')

# Configuración del dataset y la tabla en BigQuery
table_id_1= ''

# Carga el DataFrame a BigQuery
to_gbq(df_visits_routes_vehicles_qolqas,
       destination_table=f"{dataset_id}.{table_id_1}",
        project_id=project_id,
        if_exists="replace")  # Opciones: "append", "replace", "fail"

print("Los datos se han cargado exitosamente en BigQuery.")
