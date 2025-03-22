import PySimpleGUI as sg
import gspread
import pandas as pd
import re
from google.oauth2.service_account import Credentials

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
def process_files(json_path, igms_csv_path, apply_15chars, order_by_cleaner):
    """
    Procesa los archivos:
      - Se autentica en Google Sheets y lee la hoja "Key Register".
      - Filtra las llaves disponibles (donde "Observation" esté vacío).
      - Lee el CSV de IGMS.
      - Crea la columna 'Simplified Address' usando la regla elegida.
      - Realiza un merge para emparejar propiedades con llaves.
      - Crea la columna 'M_Key' asignando la llave si empieza con "M".
      - NO filtra las propiedades; se muestran todas.
      - Ordena el resultado según la opción del usuario.
      - Renombra columnas para un Excel más amigable.
      - Exporta el resultado final a CSV.
    Devuelve el DataFrame final, su representación en texto y un mensaje.
    """
    try:
        # 1. Autenticación en Google Sheets
        credentials = Credentials.from_service_account_file(
            json_path,
            scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets"]
        )
        client = gspread.authorize(credentials)
        
        # 2. Cargar la hoja "Key Register"
        sheet = client.open_by_key('1AEX3jKwAdO5cROqTe6k4uNv7BCy7lPOKHrGQjZA3om0').worksheet('Key Register')
        data = sheet.get_all_values()
        
        # 3. Crear DataFrame de llaves (fila 1 = encabezados, filas 2+ = datos)
        df_keys = pd.DataFrame(data[2:], columns=data[1])
        df_keys = df_keys.drop(columns='', errors='ignore')
        
        # 4. Filtrar llaves disponibles (donde "Observation" esté vacío o sea NaN)
        df_keys_available = df_keys[(df_keys["Observation"].isna()) | (df_keys["Observation"].str.strip() == "")]
        
        # 5. Cargar CSV IGMS
        df_igms = pd.read_csv(igms_csv_path)
        
        # 6. Crear la columna "Simplified Address" en ambos DataFrames
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
        
        # 7. Merge entre IGMS y llaves disponibles
        df_merged = pd.merge(
            df_igms,
            df_keys_available,
            on="Simplified Address",
            how="left",
            suffixes=('_IGMS', '_Key')
        )
        
        # 8. Crear la columna "M_Key": asigna la llave (columna "Tag") si empieza con "M"
        df_merged["M_Key"] = df_merged.apply(
            lambda row: row["Tag"] if pd.notna(row["Tag"]) and row["Tag"].strip().startswith("M") else "",
            axis=1
        )
        
        # 9. Seleccionar columnas de interés para el reporte final
        df_result = df_merged[[ 
            "Property Nickname",   # Desde IGMS
            "Cleaner",             # Desde IGMS (puede estar vacío)
            "M_Key",               # Llave M (vacío si no hay)
            "Simplified Address"   # Para referencia
        ]]
        
        # Asegurarse de que "Cleaner" no tenga NaN para ordenar
        df_result["Cleaner"] = df_result["Cleaner"].fillna("")
        
        # 10. Ordenar según la opción del usuario
        sort_column = "Cleaner" if order_by_cleaner else "Property Nickname"
        df_result_sorted = df_result.sort_values(by=sort_column, na_position="first")
        
        # 11. Renombrar columnas para un Excel más amigable
        df_result_sorted = df_result_sorted.rename(columns={
            "Property Nickname": "Propiedad",
            "Cleaner": "Responsable",
            "M_Key": "Llave M",
            "Simplified Address": "Direccion Simplificada"
        })
        
        # 12. Exportar el resultado a un archivo CSV
        output_csv = r'resultado_llaves_m_sorted.csv'
        df_result_sorted.to_csv(output_csv, index=False)
        
        # Convertir a string para mostrar en GUI (sin truncar)
        result_str = df_result_sorted.to_string(index=False, max_rows=None)
        
        return df_result_sorted, result_str, f"Proceso completado. Resultado exportado a '{output_csv}'."
    
    except Exception as e:
        return None, "", f"Error al procesar: {str(e)}"

# ========================
# Función para generar reporte agrupado (una fila por unidad)
# ========================
def generate_grouped_report(df_result_sorted):
    """
    Agrupa el DataFrame final (ya renombrado) por Propiedad y concatena todas las llaves M 
    disponibles en una sola celda separadas por comas. Además, se ordena por Responsable y Propiedad.
    Se aplican formatos para un resultado más llamativo.
    """
    # Agrupar y concatenar llaves M para cada Propiedad
    df_grouped = df_result_sorted.groupby("Propiedad", as_index=False).agg({
        "Responsable": "first",
        "Direccion Simplificada": "first",
        "Llave M": lambda x: ", ".join(sorted(set(x.dropna().astype(str).str.strip()).difference({''})))
    })
    # Ordenar por Responsable y luego por Propiedad
    df_grouped = df_grouped.sort_values(by=["Responsable", "Propiedad"], na_position="first")
    
    output_xlsx = r'resultado_llaves_m_grouped.xlsx'
    with pd.ExcelWriter(output_xlsx, engine="xlsxwriter") as writer:
        df_grouped.to_excel(writer, sheet_name="Reporte", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Reporte"]
        
        # Formato para encabezados: fondo azul, texto blanco, negrita y bordes
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4F81BD',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        # Escribir encabezados
        for col_num, value in enumerate(df_grouped.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Formato para celdas de datos
        cell_format = workbook.add_format({
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        # Ajustar columnas y aplicar formato
        worksheet.set_column(0, 0, 30, cell_format)  # Propiedad
        worksheet.set_column(1, 1, 20, cell_format)  # Responsable
        worksheet.set_column(2, 2, 40, cell_format)  # Direccion Simplificada
        worksheet.set_column(3, 3, 30, cell_format)  # Llave M
    return output_xlsx, df_grouped.to_string(index=False, max_rows=None)

# ========================
# Interfaz gráfica con PySimpleGUI
# ========================
def main():
    sg.theme("LightBlue2")
    layout = [
        [sg.Text("Archivo JSON de credenciales:"), sg.Input(key="-JSON-"), sg.FileBrowse(file_types=(("JSON Files", "*.json"),))],
        [sg.Text("Archivo CSV IGMS:"), sg.Input(key="-CSV-"), sg.FileBrowse(file_types=(("CSV Files", "*.csv"),))],
        [sg.Frame("Opciones", [
            [sg.Checkbox("Aplicar 15 caracteres desde primer dígito", default=True, key="-APPLY15-")],
            [sg.Text("Ordenar por:"), sg.Radio("Cleaner", "ORDER", default=True, key="-ORDER_CLEANER-"), sg.Radio("Nombre Propiedad", "ORDER", key="-ORDER_NAME-")],
            [sg.Checkbox("Generar reporte agrupado (una fila por unidad)", default=True, key="-GROUPED-")]
        ])],
        [sg.Button("Procesar"), sg.Exit()],
        [sg.Text("Resultado:")],
        [sg.Multiline(size=(80,20), key="-OUTPUT-")]
    ]
    
    window = sg.Window("Procesador de Llaves M", layout)
    
    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "Exit"):
            break
        if event == "Procesar":
            json_path = values["-JSON-"]
            csv_path = values["-CSV-"]
            apply_15chars = values["-APPLY15-"]
            order_by_cleaner = values["-ORDER_CLEANER-"]
            grouped = values["-GROUPED-"]
            
            if not json_path or not csv_path:
                sg.popup("Por favor, selecciona ambos archivos (JSON y CSV).")
            else:
                df_result_sorted, result_str, msg = process_files(json_path, csv_path, apply_15chars, order_by_cleaner)
                output_text = f"{msg}\n\n{result_str}"
                
                if grouped and df_result_sorted is not None:
                    output_xlsx, grouped_str = generate_grouped_report(df_result_sorted)
                    output_text += f"\n\nReporte agrupado exportado a '{output_xlsx}':\n{grouped_str}"
                
                window["-OUTPUT-"].update(output_text)
    
    window.close()

if __name__ == "__main__":
    main()
