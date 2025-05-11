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
        substr = address[match.start() : match.start() + 15]
    else:
        substr = address[:15]
    return re.sub(r"[^0-9A-Za-z\s]", "", substr).lower().strip()

# ----------------------------------------------------------
# 3. Crear y formatear el reporte Excel agrupado
#    (siempre ordenado por Cleaner → Dirección → Llave M)
# ----------------------------------------------------------
def create_grouped_excel(igms_csv):
    # 3.1 Leer Google Sheet con llaves
    client = authorize_gspread()
    sheet_id = st.secrets["gcp_service_account"]["spreadsheet_id"]
    ws = client.open_by_key(sheet_id).worksheet("Key Register")
    data = ws.get_all_values()

    # 3.2 DataFrame de llaves y limpieza de encabezados
    df_keys = pd.DataFrame(data[2:], columns=data[1]).drop(columns="", errors="ignore")
    df_keys.columns = df_keys.columns.str.strip()  # quitar espacios invisibles
    df_keys = df_keys[df_keys["Observation"].str.strip().fillna("") == ""]

    # 3.3 Leer CSV IGMS subido
    df_igms = pd.read_csv(igms_csv)
    df_igms["Simplified"] = df_igms["Property Nickname"].apply(
        lambda x: simplify_address_15chars(x.split("-")[0])
    )

    # 3.4 Simplified en df_keys
    df_keys["Simplified"] = df_keys["Property Address"].apply(simplify_address_15chars)

    # 3.5 Merge
    merged = pd.merge(
        df_igms,
        df_keys,
        on="Simplified",
        how="left",
        suffixes=("_IGMS", "_Key"),
    )

    # 3.6 Extraer llaves que empiecen por "M"
    def extract_m_key(row):
        if "Tag" in row.index:
            tag = row["Tag"]
            if isinstance(tag, str) and tag.strip().startswith("M"):
                return tag.strip()
        return ""
    merged["Llave M"] = merged.apply(extract_m_key, axis=1)

    # 3.7 Selección y renombrado
    df = merged[["Cleaner", "Property Nickname", "Llave M"]].rename(columns={
        "Cleaner": "Encargado",
        "Property Nickname": "Dirección"
    })
    df = df[df["Encargado"].str.strip().fillna("") != ""]

    # 3.8 Agrupar y ordenar por Encargado → Dirección
    grouped = (
        df.groupby(["Encargado", "Dirección"], as_index=False)
          .agg({"Llave M": lambda x: ", ".join(sorted({v.strip() for v in x if v.strip()}))})
          .sort_values(["Encargado", "Dirección"])
    )

    # 3.9 Reordenar columnas: Dirección, Encargado, Llave M
    grouped = grouped[["Dirección", "Encargado", "Llave M"]]

    # 3.10 Generar Excel en memoria con formato
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        grouped.to_excel(writer, sheet_name="Reporte", index=False)
        wb, ws = writer.book, writer.sheets["Reporte"]

        # Formato de encabezado
        header_fmt = wb.add_format({
            "bold": True,
            "bg_color": "#305496",
            "font_color": "white",
