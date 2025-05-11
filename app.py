import streamlit as st
import gspread
import pandas as pd
import re
from io import BytesIO
from google.oauth2.service_account import Credentials

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

def simplify_address_15chars(address: str) -> str:
    address = address.strip()
    match = re.search(r"\d", address)
    if match:
        substr = address[match.start() : match.start() + 15]
    else:
        substr = address[:15]
    return re.sub(r"[^0-9A-Za-z\s]", "", substr).lower().strip()

def create_grouped_excel(igms_csv):
    client = authorize_gspread()
    sheet_id = st.secrets["gcp_service_account"]["spreadsheet_id"]
    ws = client.open_by_key(sheet_id).worksheet("Key Register")
    data = ws.get_all_values()

    df_keys = pd.DataFrame(data[2:], columns=data[1]).drop(columns="", errors="ignore")
    df_keys.columns = df_keys.columns.str.strip()
    df_keys = df_keys[df_keys["Observation"].str.strip().fillna("") == ""]

    df_igms = pd.read_csv(igms_csv)
    df_igms["Simplified"] = df_igms["Property Nickname"].apply(
        lambda x: simplify_address_15chars(x.split("-")[0])
    )

    df_keys["Simplified"] = df_keys["Property Address"].apply(simplify_address_15chars)

    merged = pd.merge(
        df_igms,
        df_keys,
        on="Simplified",
        how="left",
        suffixes=("_IGMS", "_Key"),
    )

    def extract_m_key(row):
        if "Tag" in row.index:
            tag = row["Tag"]
            if isinstance(tag, str) and tag.strip().startswith("M"):
                return tag.strip()
        return ""
    merged["Llave M"] = merged.apply(extract_m_key, axis=1)

    df = merged[["Cleaner", "Property Nickname", "Llave M"]].rename(columns={
        "Cleaner": "Encargado",
        "Property Nickname": "Dirección"
    })
    df = df[df["Encargado"].str.strip().fillna("") != ""]

    grouped = (
        df.groupby(["Encargado", "Dirección"], as_index=False)
          .agg({"Llave M": lambda x: ", ".join(sorted({v.strip() for v in x if v.strip()}))})
          .sort_values(["Encargado", "Dirección"])
    )

    grouped = grouped[["Dirección", "Encargado", "Llave M"]]

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        grouped.to_excel(writer, sheet_name="Reporte", index=False)
        wb, ws = writer.book, writer.sheets["Reporte"]

        header_fmt = wb.add_format({
            "bold": True,
            "bg_color": "#305496",
            "font_color": "white",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })
        cell_fmt = wb.add_format({"border": 1, "align": "left", "valign": "vcenter"})
        band_fmt = wb.add_format({
            "border": 1,
            "bg_color": "#F2F2F2",
            "align": "left",
            "valign": "vcenter",
        })

        for col, name in enumerate(grouped.columns):
            ws.write(0, col, name, header_fmt)
        ws.set_column("A:A", 35, cell_fmt)
        ws.set_column("B:B", 25, cell_fmt)
        ws.set_column("C:C", 40, cell_fmt)

        for row in range(1, len(grouped) + 1):
            ws.set_row(row, None, band_fmt if row % 2 == 0 else cell_fmt)

    output.seek(0)
    return grouped, output.read()

st.set_page_config(page_title="Reporte Llaves M", layout="wide")
st.title("🗝️ Reporte de Llaves M (ordenado por Cleaner)")
csv_file = st.file_uploader("📥 Sube tu CSV de IGMS", type="csv")
if csv_file:
    df_grp, excel_bytes = create_grouped_excel(csv_file)
    st.success("✅ Reporte generado correctamente")
    st.dataframe(df_grp, use_container_width=True)
    st.download_button(
        label="⬇️ Descargar Reporte (Excel)",
        data=excel_bytes,
        file_name="reporte_llaves_m.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("📄 Esperando que subas un archivo CSV de IGMS")
