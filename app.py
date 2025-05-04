# ==========================================================
# EduSpot – Generador de Reporte de Llaves M
#   · Agrupa y ordena SIEMPRE por “Cleaner”
#   · Excel más legible (bandas + ancho fijo)
#   · Muestra solo: Cleaner | Apartamento | Llaves
# ==========================================================
import streamlit as st
import gspread
import pandas as pd
import re
from io import BytesIO
from google.oauth2.service_account import Credentials

# ----------------------------------------------------------
# 1. Autenticación Google Sheets
# ----------------------------------------------------------
def authorize_gspread():
    creds = st.secrets["gcp_service_account"]
    credentials = Credentials.from_service_account_info(
        creds,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    return gspread.authorize(credentials)

# ----------------------------------------------------------
# 2. Simplificación de dirección (15 caracteres)
# ----------------------------------------------------------
def simplify_address_15chars(address: str) -> str:
    address = address.strip()
    match = re.search(r"\d", address)
    substr = address[match.start(): match.start() + 15] if match else address[:15]
    return re.sub(r"[^0-9a-zA-Z\s]", "", substr).lower().strip()

# ----------------------------------------------------------
# 3. Procesamiento principal
# ----------------------------------------------------------
def create_grouped_excel(igms_csv_file):
    try:
        # --- Leer Google Sheet ---
        client = authorize_gspread()
        sheet_id = st.secrets["gcp_service_account"]["spreadsheet_id"]
        sheet = client.open_by_key(sheet_id).worksheet("Key Register")
        data = sheet.get_all_values()

        df_keys = (
            pd.DataFrame(data[2:], columns=data[1])
            .drop(columns="", errors="ignore")
        )
        df_keys = df_keys[df_keys["Observation"].str.strip().fillna("") == ""]

        # --- Leer CSV IGMS ---
        df_igms = pd.read_csv(igms_csv_file)

        # --- Simplificar direcciones ---
        df_igms["Simplified"] = df_igms["Property Nickname"].apply(
            lambda x: simplify_address_15chars(x.split("-")[0])
        )
        df_keys["Simplified"] = df_keys["Property Address"].apply(simplify_address_15chars)

        # --- Merge IGMS + llaves ---
        merged = pd.merge(
            df_igms,
            df_keys,
            on="Simplified",
            how="left",
            suffixes=("_IGMS", "_Key"),
        )

        merged["Llave M"] = merged.apply(
            lambda r: r["Tag"] if pd.notna(r["Tag"]) and r["Tag"].strip().startswith("M") else "",
            axis=1,
        )

        result = merged[["Cleaner", "Property Nickname", "Llave M"]].rename(columns={
            "Cleaner": "Cleaner",
            "Property Nickname": "Apartamento",
        })

        # --- FILTRAR: solo filas con Cleaner no vacío ---
        result = result[result["Cleaner"].fillna("").str.strip() != ""]

        # --- Agrupar por Cleaner / Apartamento ---
        grouped = (
            result.groupby(["Cleaner", "Apartamento"], as_index=False)
                  .agg({"Llave M": lambda x: ", ".join(sorted({v.strip() for v in x if v.strip()}))})
                  .sort_values(["Cleaner", "Apartamento"])
        )

        # --- Crear Excel en memoria ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            grouped.to_excel(writer, sheet_name="Reporte", index=False)
            wb, ws = writer.book, writer.sheets["Reporte"]

            header_fmt = wb.add_format({
                "bold": True, "bg_color": "#305496", "font_color": "white",
                "border": 1, "align": "center", "valign": "vcenter"})
            cell_fmt   = wb.add_format({"border": 1, "align": "left", "valign": "vcenter"})
            band_fmt   = wb.add_format({"border": 1, "bg_color": "#F2F2F2",
                                        "align": "left", "valign": "vcenter"})

            for col, col_name in enumerate(grouped.columns):
                ws.write(0, col, col_name, header_fmt)

            ws.set_column("A:A", 25, cell_fmt)  # Cleaner
            ws.set_column("B:B", 30, cell_fmt)  # Apartamento
            ws.set_column("C:C", 40, cell_fmt)  # Llaves

            for row in range(1, len(grouped) + 1):
                ws.set_row(row, None, band_fmt if row % 2 == 0 else cell_fmt)

        output.seek(0)
        return grouped, output.read(), "Reporte agrupado por Cleaner generado correctamente."
    except Exception as e:
        return None, None, f"Error: {e}"

# ----------------------------------------------------------
# 4. Interfaz Streamlit
# ----------------------------------------------------------
st.title("Reporte de Llaves M agrupado por Cleaner")

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
