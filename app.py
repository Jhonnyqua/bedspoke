import streamlit as st
import gspread
import pandas as pd
import re
from io import BytesIO
from google.oauth2.service_account import Credentials

# ----------------------------------------------------------
# 1. Autenticación con Google Sheets (st.secrets)
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
# 2. Simplificar dirección (15 caracteres tras primer dígito)
# ----------------------------------------------------------
def simplify_address_15chars(address: str) -> str:
    address = address.strip()
    match = re.search(r"\d", address)
    if match:
        start = match.start()
        substr = address[start : start + 15]
    else:
        substr = address[:15]
    return re.sub(r"[^0-9A-Za-z\s]", "", substr).lower().strip()

# ----------------------------------------------------------
# 3. Generar reporte agrupado siempre por Cleaner
# ----------------------------------------------------------
def create_grouped_excel(igms_csv):
    # 3.1 Leer llaves disponibles desde Google Sheet
    client = authorize_gspread()
    sheet_id = st.secrets["gcp_service_account"]["spreadsheet_id"]
    sheet = client.open_by_key(sheet_id).worksheet("Key Register")
    data = sheet.get_all_values()

    df_keys = pd.DataFrame(data[2:], columns=data[1]).drop(columns="", errors="ignore")
    df_keys = df_keys[df_keys["Observation"].str.strip().fillna("") == ""]

    # 3.2 Leer CSV IGMS
    df_igms = pd.read_csv(igms_csv)

    # 3.3 Simplificar direcciones en ambos
    df_igms["Simplified"] = df_igms["Property Nickname"].apply(
        lambda x: simplify_address_15chars(x.split("-")[0])
    )
    df_keys["Simplified"] = df_keys["Property Address"].apply(simplify_address_15chars)

    # 3.4 Merge y extraer solo llaves "M"
    merged = pd.merge(
        df_igms, df_keys, on="Simplified", how="left", suffixes=("_IGMS", "_Key")
    )
    merged["Llave M"] = merged.apply(
        lambda r: r["Tag"] if pd.notna(r["Tag"]) and r["Tag"].startswith("M") else "",
        axis=1,
    )

    # 3.5 Seleccionar columnas y filtrar filas sin Cleaner
    df = merged[["Cleaner", "Property Nickname", "Llave M"]].rename(
        columns={"Property Nickname": "Apartamento"}
    )
    df = df[df["Cleaner"].str.strip().fillna("") != ""]

    # 3.6 Agrupar por Cleaner → Apartamento
    grouped = (
        df.groupby(["Cleaner", "Apartamento"], as_index=False)
        .agg({"Llave M": lambda x: ", ".join(sorted({v.strip() for v in x if v.strip()}))})
        .sort_values(["Cleaner", "Apartamento"])
    )

    # 3.7 Crear Excel en memoria con formato
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        grouped.to_excel(writer, sheet_name="Reporte", index=False)
        wb = writer.book
        ws = writer.sheets["Reporte"]

        # Cabecera
        header_fmt = wb.add_format({
            "bold": True, "bg_color": "#305496", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter"
        })
        # Celdas normales y bandas
        cell_fmt = wb.add_format({"border": 1, "align": "left", "valign": "vcenter"})
        band_fmt = wb.add_format({"border": 1, "bg_color": "#F2F2F2",
                                  "align": "left", "valign": "vcenter"})

        # Escribir cabeceras
        for col, name in enumerate(grouped.columns):
            ws.write(0, col, name, header_fmt)

        # Anchos fijos
        ws.set_column("A:A", 25, cell_fmt)  # Cleaner
        ws.set_column("B:B", 30, cell_fmt)  # Apartamento
        ws.set_column("C:C", 40, cell_fmt)  # Llaves

        # Bandas alternas
        for row in range(1, len(grouped) + 1):
            ws.set_row(row, None, band_fmt if row % 2 == 0 else cell_fmt)

    output.seek(0)
    return grouped, output.read()

# ----------------------------------------------------------
# 4. Interfaz Streamlit
# ----------------------------------------------------------
st.title("🗝️ Reporte de Llaves M (agrupado por Cleaner)")

csv_file = st.file_uploader("📥 Sube tu CSV de IGMS", type="csv")
if csv_file:
    df_grp, excel_data = create_grouped_excel(csv_file)
    st.success("✅ Reporte creado correctamente")
    st.dataframe(df_grp, use_container_width=True)
    st.download_button(
        "⬇️ Descargar Reporte (Excel)",
        data=excel_data,
        file_name="reporte_llaves_m.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("📄 Esperando que subas un archivo CSV de IGMS")
