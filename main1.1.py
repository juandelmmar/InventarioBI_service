import requests
import pandas as pd
from azure.identity import InteractiveBrowserCredential
from datetime import datetime, timezone

# 1. Autenticacion
print("Iniciando sesion interactiva...")
credential = InteractiveBrowserCredential()
token_obj = credential.get_token("https://analysis.windows.net/powerbi/api/.default")
headers = {'Authorization': f'Bearer {token_obj.token}'}

print("Buscando todas las areas de trabajo...")
url_groups = "https://api.powerbi.com/v1.0/myorg/groups"
workspaces = requests.get(url_groups, headers=headers).json().get('value', [])

all_data = []
ahora = datetime.now(timezone.utc)

# 2. Iterar por Workspaces
for ws in workspaces:
    ws_name = ws.get('name')
    ws_id = ws.get('id')
    print(f"Procesando: {ws_name}...")
    
    url_reports = f"https://api.powerbi.com/v1.0/myorg/groups/{ws_id}/reports"
    reports_res = requests.get(url_reports, headers=headers).json()
    reports = reports_res.get('value', [])
    
    for r in reports:
        report_info = {
            'Nombre Workspace': ws_name,
            'Nombre Tablero': r.get('name'),
            'Report ID': r.get('id'),
            'Dataset ID': r.get('datasetId'),
            'Ultima Actualizacion': 'Sin datos',
            'Estado': 'Desconocido'
        }
        
        # 3. Consultar la ultima actualizacion del Dataset asociado
        ds_id = r.get('datasetId')
        if ds_id:
            url_refresh = f"https://api.powerbi.com/v1.0/myorg/groups/{ws_id}/datasets/{ds_id}/refreshes?$top=1"
            res_refresh = requests.get(url_refresh, headers=headers).json().get('value', [])
            
            if res_refresh:
                last_date_str = res_refresh[0].get('endTime')
                if last_date_str:
                    try:
                        # Limpieza robusta de la fecha (manejando Z y microsegundos)
                        clean_date_str = last_date_str.replace('Z', '+00:00')
                        # Usamos pd.to_datetime que es mas inteligente para detectar formatos locos de fechas
                        last_date = pd.to_datetime(clean_date_str).to_pydatetime()
                        
                        report_info['Ultima Actualizacion'] = last_date.strftime('%Y-%m-%d %H:%M')
                        
                        # 4. Logica de Activo/Inactivo (30 dias)
                        dias_dif = (ahora - last_date).days
                        report_info['Estado'] = 'Inactivo' if dias_dif > 30 else 'Activo'
                    except:
                        report_info['Ultima Actualizacion'] = 'Error formato'
            else:
                report_info['Ultima Actualizacion'] = 'Nunca actualizado'
                report_info['Estado'] = 'Inactivo'

        all_data.append(report_info)

# 5. Exportar a Excel
df_final = pd.DataFrame(all_data)
if not df_final.empty:
    df_final.to_excel("Inventario_PBI_Con_Estado.xlsx", index=False)
    print(f"\n¡Finalizado! Se procesaron {len(df_final)} registros.")
    print("El archivo 'Inventario_PBI_Con_Estado.xlsx' ha sido generado.")
else:
    print("No se extrajeron datos.")