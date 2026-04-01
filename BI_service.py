import requests
import pandas as pd
from azure.identity import InteractiveBrowserCredential

# 1. Autenticacion interactiva (abrira tu navegador para iniciar sesion)
print("Abriendo navegador para iniciar sesion...")
credential = InteractiveBrowserCredential()

# El scope (alcance) por defecto para la API de Power BI
scope = "https://analysis.windows.net/powerbi/api/.default"
token_obj = credential.get_token(scope)
token = token_obj.token

headers = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {token}'
}

# 2. Obtener Workspaces (Areas de trabajo)
print("Descargando Workspaces...")
url_groups = "https://api.powerbi.com/v1.0/myorg/groups"
res_groups = requests.get(url_groups, headers=headers)

if res_groups.status_code == 200:
    workspaces = res_groups.json().get('value', [])
    df_workspaces = pd.DataFrame(workspaces)
else:
    print(f"Error al obtener workspaces: {res_groups.status_code}")
    df_workspaces = pd.DataFrame()

# 3. Obtener Reportes (Tableros)
print("Descargando Reportes...")
url_reports = "https://api.powerbi.com/v1.0/myorg/reports"
res_reports = requests.get(url_reports, headers=headers)

if res_reports.status_code == 200:
    reports = res_reports.json().get('value', [])
    df_reports = pd.DataFrame(reports)
else:
    print(f"Error al obtener reportes: {res_reports.status_code}")
    df_reports = pd.DataFrame()

# 4. Exportar a Excel
print("Guardando en Excel...")
with pd.ExcelWriter("Inventario_PowerBI.xlsx") as writer:
    if not df_workspaces.empty:
        df_workspaces.to_excel(writer, sheet_name="Workspaces", index=False)
    if not df_reports.empty:
        df_reports.to_excel(writer, sheet_name="Reportes", index=False)

print("Proceso finalizado. Revisa el archivo Inventario_PowerBI.xlsx")