import streamlit as st
import gspread
import pandas as pd
import re
from io import BytesIO
from google.oauth2.service_account import Credentials

# -------------------------------------------------
# 1. Autenticación con Google Sheets
# -------------------------------------------------
def authorize_gspread():
    creds_dict = st.secrets["gcp_service_account"]
    credentials = Credentials.from_service_account_info(
        creds_dict,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    return gspread.authorize(credentials)

# -------------------------------------------------
# 2. Utilidades para simplificar direcciones
# -------------------------------------------------
def simplify_address_15chars(address: str) -> str:
    address = address.strip()
    match = re.search(r"\d", address)
    substring = address[match.start(): match.start()+15] if match else address[:15]
    return re.sub(r"[^0-9a-zA-Z\s]", "", substring).lower().strip()

# -------------------------------------------------
# 3. Procesamiento principal
# -------------------------------------------------
def create_grouped_excel(igms_csv_file):
    try:
        # --- Leer Google Sheet ---
        client = authorize_gspread()
        spreadsheet_id = st.secrets["gcp_service_account"]["spreadsheet_id"]
        sheet = client.open_by_key(spreadsheet_id).worksheet("Key Register")
        data = sheet.get_all_values()

        df_keys = pd.DataFrame(data[2:], columns=data[1]).drop(columns="", errors="ignore")
        df_keys_available = df_keys[df_keys["Observation"].str.strip().fillna("") == ""]

        # --- Leer CSV subido ---
        df_igms = pd.read_csv(igms_csv_file)

        # --- Simplificación de direcciones (15 caracteres) ---
        df_igms["Simplified"] = df_igms["Property Nickname"].apply(
            lambda x: simplify_address_15chars(x.split("-")[0])
        )
        df_keys_available["Simplified"] = df_keys_available["Property Address"].apply(
            simplify_address_15chars
        )

        # --- Merge y filtrado de llaves M ---
        merged = pd.merge(
            df_igms,
            df_keys_available,
            on="Simplified",
            how="left",
            suffixes=("_IGMS", "_Key"),
        )
        merged["Llave M"] = merged.apply(
            lambda r: r["Tag"] if pd.notna(r["Tag"]) and r["Tag"].strip().startswith("M") else "",
            axis=1,
        )

        result = merged[["Property Nickname", "Cleaner", "Llave M"]]
        result = result.rename(columns={"Property Nickname": "Apartamento", "Cleaner": "Asignado"})
        result = result.sort_values(by="Asignado", na_position="first")

        # --- Agrupar por Apartamento y concatenar Llaves ---
        grouped = (
            result.groupby("Apartamento", as_index=False)
            .agg({"Asignado": "first", "Llave M": lambda x: ", ".join(sorted({v.strip() for v in x if v.strip()}))})
        )

        # --- Crear Excel en memoria ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            grouped.to_excel(writer, sheet_name="Reporte", index=False)
            wb = writer.book
            ws = writer.sheets["Reporte"]

            header_fmt = wb.add_format({
                "bold": True, "bg_color": "#305496", "font_color": "white",
                "border": 1, "align": "center", "valign": "vcenter"
            })
            for col, col_name in enumerate(grouped.columns):
                ws.write(0, col, col_name, header_fmt)

            cell_fmt = wb.add_format({"border": 1, "align": "left", "valign": "vcenter"})
            ws.set_column("A:A", 30, cell_fmt)  # Apartamento
            ws.set_column("B:B", 25, cell_fmt)  # Asignado
            ws.set_column("C:C", 40, cell_fmt)  # Llaves

        output.seek(0)
        return grouped, output.read(), "Proceso completado."
    except Exception as e:
        return None, None, f"Error: {e}"

# -------------------------------------------------
# 4. Interfaz Streamlit
# -------------------------------------------------
st.title("Generador de Reporte de Llaves M")

csv_file = st.file_uploader("Sube el archivo CSV de IGMS", type=["csv"])

if st.button("Procesar"):
    if csv_file is None:
        st.error("Por favor, sube un archivo CSV.")
    else:
        df_grouped, xlsx_bytes, msg = create_grouped_excel(csv_file)
        if df_grouped is not None:
            st.success(msg)
            st.dataframe(df_grouped)
            st.download_button(
                label="Descargar Excel Agrupado",
                data=xlsx_bytes,
                file_name="reporte_llaves_m.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error(msg)
