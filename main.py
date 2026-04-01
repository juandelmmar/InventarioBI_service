import requests
import pandas as pd
import time
from azure.identity import InteractiveBrowserCredential
from datetime import datetime, timezone
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# 1. Autenticacion
print("Iniciando sesion interactiva...")
credential = InteractiveBrowserCredential()
token_obj = credential.get_token("https://analysis.windows.net/powerbi/api/.default")
headers = {'Authorization': f'Bearer {token_obj.token}'}

# ── Sesion con reintentos automaticos ──
session = requests.Session()
retry_strategy = Retry(
    total=4,
    backoff_factor=2,          # espera 2, 4, 8, 16 segundos entre reintentos
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=["GET"]
)
adapter = HTTPAdapter(max_retries=retry_strategy)
session.mount("https://", adapter)
session.headers.update(headers)

def get_safe(url, pausa=0.4):
    """GET con reintentos y pausa entre llamadas."""
    try:
        time.sleep(pausa)
        res = session.get(url, timeout=30)
        res.raise_for_status()
        return res.json()
    except requests.exceptions.ConnectionError as e:
        print(f"  ⚠ Conexion interrumpida, reintentando en 10s... {url[-60:]}")
        time.sleep(10)
        try:
            res = session.get(url, timeout=30)
            return res.json()
        except Exception:
            print(f"  ✗ Fallo definitivo: {url[-60:]}")
            return {}
    except Exception as e:
        print(f"  ✗ Error en GET: {str(e)[:80]}")
        return {}

def obtener_origenes(ws_id, ds_id):
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{ws_id}/datasets/{ds_id}/datasources"
    res = get_safe(url).get('value', [])
    origenes = []
    for ds in res:
        tipo = ds.get('datasourceType', 'Desconocido')
        conn = ds.get('connectionDetails', {})
        if tipo == 'Sql':
            origenes.append(f"SQL Server: {conn.get('server','')}/{conn.get('database','')}")
        elif tipo == 'AzureDataLakeStorage':
            origenes.append(f"ADLS: {conn.get('url','')}")
        elif tipo == 'SharePoint':
            origenes.append(f"SharePoint: {conn.get('url','')}")
        elif tipo == 'Web':
            origenes.append(f"Web: {conn.get('url','')}")
        elif tipo == 'OData':
            origenes.append(f"OData: {conn.get('url','')}")
        elif tipo in ('File', 'Folder', 'Excel'):
            origenes.append(f"{tipo}: {conn.get('path','')}")
        elif tipo == 'AzureSqlDatabase':
            origenes.append(f"Azure SQL: {conn.get('server','')}/{conn.get('database','')}")
        elif tipo == 'AzureSqlDW':
            origenes.append(f"Synapse: {conn.get('server','')}/{conn.get('database','')}")
        elif tipo == 'Databricks':
            origenes.append(f"Databricks: {conn.get('server','')}")
        elif tipo == 'GoogleBigQuery':
            origenes.append(f"BigQuery: {conn.get('billingProjectId','')}")
        elif tipo == 'PowerBIDataset':
            origenes.append("Dataset PBI compartido")
        elif tipo == 'AnalysisServices':
            origenes.append(f"Analysis Services: {conn.get('server','')}/{conn.get('database','')}")
        else:
            origenes.append(f"{tipo}: {str(conn)[:80]}")
    return " | ".join(origenes) if origenes else "Sin origen detectado"

print("Buscando todas las areas de trabajo...")
workspaces = get_safe("https://api.powerbi.com/v1.0/myorg/groups").get('value', [])

all_data = []
ahora = datetime.now(timezone.utc)

# 2. Iterar por Workspaces
for i, ws in enumerate(workspaces):
    ws_name = ws.get('name')
    ws_id = ws.get('id')
    print(f"[{i+1}/{len(workspaces)}] Procesando: {ws_name}...")

    reports = get_safe(
        f"https://api.powerbi.com/v1.0/myorg/groups/{ws_id}/reports"
    ).get('value', [])

    for r in reports:
        ds_id = r.get('datasetId')
        report_info = {
            'Nombre Workspace':   ws_name,
            'Nombre Tablero':     r.get('name'),
            'Report ID':          r.get('id'),
            'Dataset ID':         ds_id,
            'Origen de Datos':    'Sin datos',
            'Ultima Actualizacion': 'Sin datos',
            'Estado':             'Desconocido'
        }

        if ds_id:
            # 3. Origen
            report_info['Origen de Datos'] = obtener_origenes(ws_id, ds_id)

            # 4. Ultima actualizacion
            refresh = get_safe(
                f"https://api.powerbi.com/v1.0/myorg/groups/{ws_id}/datasets/{ds_id}/refreshes?$top=1"
            ).get('value', [])

            if refresh:
                last_date_str = refresh[0].get('endTime')
                if last_date_str:
                    try:
                        last_date = pd.to_datetime(
                            last_date_str.replace('Z', '+00:00')
                        ).to_pydatetime()
                        report_info['Ultima Actualizacion'] = last_date.strftime('%Y-%m-%d %H:%M')
                        dias_dif = (ahora - last_date).days
                        report_info['Estado'] = 'Inactivo' if dias_dif > 30 else 'Activo'
                    except:
                        report_info['Ultima Actualizacion'] = 'Error formato'
            else:
                report_info['Ultima Actualizacion'] = 'Nunca actualizado'
                report_info['Estado'] = 'Inactivo'

        all_data.append(report_info)

# 5. Exportar
df_final = pd.DataFrame(all_data)
if not df_final.empty:
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nombre_archivo = f"Inventario_PBI_{timestamp}.xlsx"
    try:
        df_final.to_excel(nombre_archivo, index=False)
        print(f"\n✓ Finalizado — {len(df_final)} registros → {nombre_archivo}")
    except PermissionError:
        nombre_csv = nombre_archivo.replace('.xlsx', '.csv')
        df_final.to_csv(nombre_csv, index=False, encoding='utf-8-sig')
        print(f"\n✓ Guardado como CSV: {nombre_csv}")
else:
    print("No se extrajeron datos.")