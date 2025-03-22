import gspread
import pandas as pd
import re
from google.oauth2.service_account import Credentials

def simplify_address(address):
    """
    Busca el primer dígito en la dirección y devuelve 15 caracteres a partir de ahí.
    Si no encuentra ningún dígito, devuelve los primeros 15 caracteres.
    Luego, elimina caracteres especiales, convierte a minúsculas y quita espacios extra.
    """
    address = address.strip()
    match = re.search(r'\d', address)
    if match:
        start_index = match.start()
        substring = address[start_index:start_index+15]
    else:
        substring = address[:15]
    simplified = re.sub(r'[^0-9a-zA-Z\s]', '', substring).lower().strip()
    return simplified

# ----------------------------------------------------------------------------
# 1. Autenticación con Google Sheets usando el archivo JSON de credenciales
# ----------------------------------------------------------------------------
json_path = r'C:\Users\jhonn\Downloads\Telegram Desktop\winged-memory-454119-q3-48d5c4851974.json'
credentials = Credentials.from_service_account_file(
    json_path,
    scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets"]
)
client = gspread.authorize(credentials)

# ----------------------------------------------------------------------------
# 2. Abrir la hoja "Key Register" y obtener sus valores
# ----------------------------------------------------------------------------
sheet = client.open_by_key('1AEX3jKwAdO5cROqTe6k4uNv7BCy7lPOKHrGQjZA3om0').worksheet('Key Register')
data = sheet.get_all_values()

# Imprimir algunas filas para ver la estructura (opcional)
print("Fila 0:", data[0])
print("Fila 1:", data[1])
print("Fila 2:", data[2])

# ----------------------------------------------------------------------------
# 3. Crear DataFrame usando la fila 1 como encabezado y las filas 2+ como datos
# ----------------------------------------------------------------------------
df_keys = pd.DataFrame(data[2:], columns=data[1])
df_keys = df_keys.drop(columns='', errors='ignore')
print("Columnas en df_keys:", df_keys.columns.tolist())

# ----------------------------------------------------------------------------
# 4. Filtrar llaves disponibles (donde "Observation" esté vacío o sea NaN)
df_keys_available = df_keys[(df_keys["Observation"].isna()) | (df_keys["Observation"].str.strip() == "")]
print("Número de llaves disponibles:", df_keys_available.shape[0])

# ----------------------------------------------------------------------------
# 5. Cargar el archivo CSV de tareas IGMS
# ----------------------------------------------------------------------------
df_igms = pd.read_csv(r'C:\Users\jhonn\Downloads\Telegram Desktop\igms_tasks.csv')
print("Columnas en df_igms:", df_igms.columns.tolist())

# ----------------------------------------------------------------------------
# 6. Crear la columna "Simplified Address" en ambos DataFrames usando la función
# Para IGMS, se usa la parte de "Property Nickname" antes del guión
df_igms["Simplified Address"] = df_igms["Property Nickname"].apply(
    lambda x: simplify_address(x.split('-')[0])
)
# Para Key Register, se usa la columna "Property Address"
df_keys_available["Simplified Address"] = df_keys_available["Property Address"].apply(
    lambda x: simplify_address(x)
)

# ----------------------------------------------------------------------------
# 7. Depurar: Imprimir algunas filas de "Simplified Address" en ambos DataFrames
print("\n--- IGMS Simplified Addresses ---")
print(df_igms[["Property Nickname", "Simplified Address"]].head(10))

print("\n--- Key Register Simplified Addresses ---")
print(df_keys_available[["Property Address", "Simplified Address"]].head(10))

# ----------------------------------------------------------------------------
# 8. Merge entre IGMS y llaves disponibles usando "Simplified Address"
df_merged = pd.merge(
    df_igms,
    df_keys_available,
    on="Simplified Address",
    how="left",
    suffixes=('_IGMS', '_Key')
)
print("\nColumnas en df_merged:", df_merged.columns.tolist())

# ----------------------------------------------------------------------------
# 9. Crear columna "M_Key": si "Tag" existe y empieza con "M", se asigna; de lo contrario, cadena vacía
df_merged["M_Key"] = df_merged.apply(
    lambda row: row["Tag"] if pd.notna(row["Tag"]) and row["Tag"].strip().startswith("M") else "",
    axis=1
)

# ----------------------------------------------------------------------------
# 10. Filtrar solo las propiedades que tienen llave M disponible (M_Key no vacía)
df_filtered = df_merged[df_merged["M_Key"].str.startswith("M")]

# ----------------------------------------------------------------------------
# 11. Seleccionar columnas de interés para el reporte final
df_result = df_filtered[[
    "Property Nickname",   # Desde IGMS
    "Cleaner",             # Desde IGMS (ajusta el nombre si es necesario)
    "M_Key",               # Llave M disponible
    "Simplified Address"   # Para referencia
]]

# ----------------------------------------------------------------------------
# 12. Ordenar el resultado final por la columna "Cleaner"
df_result_sorted = df_result.sort_values(by="Cleaner")

# ----------------------------------------------------------------------------
# 13. Mostrar y exportar el resultado final
print("\n--- Resultado final ordenado por Cleaner ---")
print(df_result_sorted)

df_result_sorted.to_csv('resultado_llaves_m_sorted.csv', index=False)
