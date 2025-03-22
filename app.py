import streamlit as st
import gspread
import pandas as pd
import re
from google.oauth2.service_account import Credentials

# ========================
# Función para autorizar gspread usando st.secrets
# ========================
def authorize_gspread():
    # Obtiene las credenciales desde st.secrets
    creds_dict = st.secrets["gcp_service_account"]
    credentials = Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    )
    client = gspread.authorize(credentials)
    return client

# ========================
# Funciones para simplificar la dirección
# ========================
def simplify_address_15chars(address):
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

def simplify_address_basic(address):
    """
    Versión básica: elimina espacios, caracteres especiales y pasa a minúsculas.
    """
    address = address.strip()
    simplified = re.sub(r'[^0-9a-zA-Z\s]', '', address).lower().strip()
    return simplified

# ========================
# Función principal de procesamiento
# ========================
def process_files(igms_csv_file, apply_15chars, order_by_cleaner):
    try:
        # Autorizar gspread usando st.secrets
        client = authorize_gspread()
        # Usar el spreadsheet_id definido en st.secrets
        spreadsheet_id = st.secrets["gcp_service_account"].get("spreadsheet_id", "YOUR_SPREADSHEET_ID")
        sheet = client.open_by_key(spreadsheet_id).worksheet("Key Register")
        data = sheet.get_all_values()
        
        # Crear DataFrame de llaves (la fila 1 tiene encabezados, filas 2+ datos)
        df_keys = pd.DataFrame(data[2:], columns=data[1])
        df_keys = df_keys.drop(columns='', errors='ignore')
        # Filtrar llaves disponibles (Observation vacío o NaN)
        df_keys_available = df_keys[(df_keys["Observation"].isna()) | (df_keys["Observation"].str.strip() == "")]
        
        # Leer el CSV de IGMS (archivo subido)
        df_igms = pd.read_csv(igms_csv_file)
        
        # Crear la columna "Simplified Address" en ambos DataFrames
        if apply_15chars:
            df_igms["Simplified Address"] = df_igms["Property Nickname"].apply(
                lambda x: simplify_address_15chars(x.split('-')[0])
            )
            df_keys_available["Simplified Address"] = df_keys_available["Property Address"].apply(
                lambda x: simplify_address_15chars(x)
            )
        else:
            df_igms["Simplified Address"] = df_igms["Property Nickname"].apply(
                lambda x: simplify_address_basic(x.split('-')[0])
            )
            df_keys_available["Simplified Address"] = df_keys_available["Property Address"].apply(
                lambda x: simplify_address_basic(x)
            )
        
        # Merge entre IGMS y llaves disponibles
        df_merged = pd.merge(
            df_igms,
            df_keys_available,
            on="Simplified Address",
            how="left",
            suffixes=('_IGMS', '_Key')
        )
        
        # Crear la columna "M_Key": asigna la llave (columna "Tag") si empieza con "M"
        df_merged["M_Key"] = df_merged.apply(
            lambda row: row["Tag"] if pd.notna(row["Tag"]) and row["Tag"].strip().startswith("M") else "",
            axis=1
        )
        
        # Seleccionar columnas de interés
        df_result = df_merged[[ 
            "Property Nickname",   # IGMS
            "Cleaner",             # IGMS
            "M_Key",               # Llave M
            "Simplified Address"   # Referencia
        ]]
        df_result["Cleaner"] = df_result["Cleaner"].fillna("")
        
        # Ordenar según la opción del usuario
        sort_column = "Cleaner" if order_by_cleaner else "Property Nickname"
        df_result_sorted = df_result.sort_values(by=sort_column, na_position="first")
        df_result_sorted = df_result_sorted.rename(columns={
            "Property Nickname": "Propiedad",
            "Cleaner": "Responsable",
            "M_Key": "Llave M",
            "Simplified Address": "Direccion Simplificada"
        })
        
        csv_bytes = df_result_sorted.to_csv(index=False).encode('utf-8')
        result_str = df_result_sorted.to_string(index=False, max_rows=None)
        return df_result_sorted, csv_bytes, result_str, "Proceso completado."
    except Exception as e:
        return None, None, "", f"Error al procesar: {e}"

# ========================
# Función para generar reporte agrupado
# ========================
def generate_grouped_report(df_result_sorted):
    """
    Agrupa el DataFrame final por Propiedad y concatena todas las llaves M disponibles
    en una sola celda, ordenado por Responsable y Propiedad. Se aplica un formato atractivo.
    """
    df_grouped = df_result_sorted.groupby("Propiedad", as_index=False).agg({
        "Responsable": "first",
        "Direccion Simplificada": "first",
        "Llave M": lambda x: ", ".join(sorted(set(x.dropna().astype(str).str.strip()).difference({''})))
    })
    df_grouped = df_grouped.sort_values(by=["Responsable", "Propiedad"], na_position="first")
    
    output_xlsx = "resultado_llaves_m_grouped.xlsx"
    with pd.ExcelWriter(output_xlsx, engine="xlsxwriter") as writer:
        df_grouped.to_excel(writer, sheet_name="Reporte", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Reporte"]
        
        # Formato para encabezados
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4F81BD',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        for col_num, value in enumerate(df_grouped.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Formato para celdas de datos
        cell_format = workbook.add_format({
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        worksheet.set_column(0, 0, 30, cell_format)  # Propiedad
        worksheet.set_column(1, 1, 20, cell_format)  # Responsable
        worksheet.set_column(2, 2, 40, cell_format)  # Direccion Simplificada
        worksheet.set_column(3, 3, 30, cell_format)  # Llave M
    with open(output_xlsx, "rb") as f:
        xlsx_data = f.read()
    grouped_str = df_grouped.to_string(index=False, max_rows=None)
    return output_xlsx, xlsx_data, grouped_str

# ========================
# Interfaz con Streamlit
# ========================
st.title("Procesador de Llaves M")

st.markdown("### Subir archivo CSV IGMS")
csv_file = st.file_uploader("Selecciona el archivo CSV", type=["csv"])

apply_15chars = st.checkbox("Aplicar 15 caracteres desde el primer dígito", value=True)
order_option = st.radio("Ordenar por:", ("Cleaner", "Nombre Propiedad"))
order_by_cleaner = True if order_option == "Cleaner" else False
grouped = st.checkbox("Generar reporte agrupado (una fila por unidad)", value=True)

if st.button("Procesar"):
    if csv_file is None:
        st.error("Por favor, sube el archivo CSV IGMS.")
    else:
        df_result_sorted, csv_bytes, result_str, msg = process_files(csv_file, apply_15chars, order_by_cleaner)
        st.success(msg)
        st.text_area("Resultado Plano", result_str, height=300)
        st.download_button(label="Descargar CSV", data=csv_bytes, file_name="resultado_llaves_m_sorted.csv", mime="text/csv")
        if grouped and df_result_sorted is not None:
            output_xlsx, xlsx_bytes, grouped_str = generate_grouped_report(df_result_sorted)
            st.text_area("Reporte Agrupado", grouped_str, height=300)
            st.download_button(label="Descargar Excel Agrupado", data=xlsx_bytes, file_name="resultado_llaves_m_grouped.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
