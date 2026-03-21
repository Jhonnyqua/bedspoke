import base64
import json
import re
from io import BytesIO

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

# ----------------------------------------------------------
# GOOGLE SHEETS
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

def load_key_register() -> pd.DataFrame:
    client = authorize_gspread()
    sheet_id = st.secrets["gcp_service_account"]["spreadsheet_id"]
    sheet = client.open_by_key(sheet_id).worksheet("Key Register")
    data = sheet.get_all_values()
    if len(data) < 2:
        raise ValueError("La hoja 'Key Register' no tiene suficiente información.")
    df_keys = pd.DataFrame(data[2:], columns=data[1]).drop(columns="", errors="ignore")
    for col in ["Property Address", "Tag"]:
        if col not in df_keys.columns:
            raise ValueError(f"No encontré la columna '{col}' en 'Key Register'.")
    if "Observation" in df_keys.columns:
        df_keys = df_keys[df_keys["Observation"].fillna("").str.strip() == ""]
    df_keys["Property Address"] = df_keys["Property Address"].fillna("").astype(str).str.strip()
    df_keys["Tag"] = df_keys["Tag"].fillna("").astype(str).str.strip()
    return df_keys

# ----------------------------------------------------------
# NORMALIZAR DIRECCIÓN
# ----------------------------------------------------------
def simplify_address_15chars(address: str) -> str:
    if not isinstance(address, str):
        return ""
    address = address.strip()
    m = re.search(r"\d", address)
    substr = address[m.start():m.start() + 15] if m else address[:15]
    return re.sub(r"[^0-9A-Za-z\s]", "", substr).lower().strip()

# ----------------------------------------------------------
# PDF → IMÁGENES BASE64 (una por página)
# ----------------------------------------------------------
def pdf_to_base64_images(pdf_bytes: bytes) -> list:
    try:
        import fitz
    except ImportError:
        raise ImportError("Falta PyMuPDF. Agregá 'pymupdf' a requirements.txt")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page in doc:
        mat = fitz.Matrix(1.5, 1.5)  # zoom moderado para no exceder límites
        pix = page.get_pixmap(matrix=mat)
        images.append(base64.b64encode(pix.tobytes("png")).decode())
    doc.close()
    return images

# ----------------------------------------------------------
# LLAMAR CLAUDE API: UNA PÁGINA A LA VEZ
# ----------------------------------------------------------
SYSTEM_PROMPT = """Eres un extractor de datos de reportes de housekeeping.
Se te muestra UNA página de un "Housekeeping Daily Summary".
Extraé cada propiedad de la página con su dirección y cleaner asignado.

REGLAS:
- La columna IZQUIERDA tiene la dirección de la propiedad.
- La columna DERECHA (encabezado "Assigned To") tiene el nombre del cleaner.
- Si el cleaner dice "Unassigned", usá "Unassigned".
- Una dirección puede cortarse al final de la página — incluila igual, completa lo que puedas.
- Ignorá encabezados, pies de página, fechas, datos de huéspedes.
- Respondé SOLO con JSON válido, sin texto extra, sin markdown:
[{"address": "...", "cleaner": "..."}]
- Si la página no tiene propiedades (ej. es solo encabezado), respondé: []"""

def call_claude_page(img_b64: str, page_num: int) -> list:
    import requests

    response = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"Content-Type": "application/json"},
        json={
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 1024,
            "system": SYSTEM_PROMPT,
            "messages": [{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": img_b64
                        }
                    },
                    {
                        "type": "text",
                        "text": f"Extraé las propiedades de esta página ({page_num})."
                    }
                ]
            }]
        },
        timeout=60
    )

    if response.status_code != 200:
        raise ValueError(f"Error API página {page_num}: {response.status_code}")

    raw = ""
    for block in response.json().get("content", []):
        if block.get("type") == "text":
            raw += block["text"]

    raw = re.sub(r"```json|```", "", raw).strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # Intentar extraer JSON del texto
        m = re.search(r"\[.*\]", raw, re.DOTALL)
        if m:
            return json.loads(m.group())
        st.warning(f"⚠️ Página {page_num}: no pude parsear JSON. Respuesta: {raw[:200]}")
        return []

# ----------------------------------------------------------
# PROCESAR TODAS LAS PÁGINAS
# ----------------------------------------------------------
def extract_all_pages(images: list, progress_bar) -> list:
    """Llama Claude una vez por página, muestra progreso."""
    all_records = []
    n = len(images)

    # Saltear primera página si es solo portada/encabezado (igual la procesamos)
    for i, img in enumerate(images):
        page_num = i + 1
        pct = 0.1 + (0.7 * (i / n))
        progress_bar.progress(pct, f"Leyendo página {page_num} de {n}...")

        records = call_claude_page(img, page_num)
        all_records.extend(records)

    return all_records

# ----------------------------------------------------------
# GENERAR EXCEL
# ----------------------------------------------------------
def create_report_excel(pdf_bytes: bytes, progress_bar):
    progress_bar.progress(0.05, "Convirtiendo PDF a imágenes...")
    images = pdf_to_base64_images(pdf_bytes)

    records = extract_all_pages(images, progress_bar)
    progress_bar.progress(0.82, "Cruzando con Key Register...")

    if not records:
        raise ValueError("Claude no encontró propiedades en el PDF.")

    df_pdf = pd.DataFrame(records)
    df_pdf = df_pdf.rename(columns={"address": "Property Nickname", "cleaner": "Cleaner"})
    df_pdf["Cleaner"] = df_pdf["Cleaner"].fillna("Unassigned").astype(str).str.strip()
    df_pdf["Property Nickname"] = df_pdf["Property Nickname"].fillna("").astype(str).str.strip()
    df_pdf = df_pdf[df_pdf["Property Nickname"] != ""].reset_index(drop=True)

    df_keys = load_key_register()
    df_pdf["Simplified"] = df_pdf["Property Nickname"].apply(simplify_address_15chars)
    df_keys["Simplified"] = df_keys["Property Address"].apply(simplify_address_15chars)

    merged = pd.merge(df_pdf, df_keys, on="Simplified", how="left", suffixes=("_PDF", "_Key"))
    merged["Llave M"] = merged["Tag"].apply(
        lambda x: x if pd.notna(x) and str(x).strip().upper().startswith("M") else ""
    )

    df_report = merged[["Cleaner", "Property Nickname", "Llave M"]].rename(
        columns={"Cleaner": "Encargado", "Property Nickname": "Dirección"}
    )
    df_report = df_report.fillna("").astype(str)
    df_report["Encargado"] = df_report["Encargado"].str.strip().replace("", "Unassigned")
    df_report = df_report[df_report["Dirección"].str.strip() != ""]

    grouped = (
        df_report.groupby(["Encargado", "Dirección"], as_index=False)
        .agg({"Llave M": lambda x: ", ".join(sorted({v.strip() for v in x if v.strip()}))})
        .sort_values(["Encargado", "Dirección"])
    )[["Dirección", "Encargado", "Llave M"]]

    progress_bar.progress(0.93, "Generando Excel...")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        grouped.to_excel(writer, sheet_name="Reporte", index=False)
        df_pdf.to_excel(writer, sheet_name="Extraido_PDF", index=False)
        merged.to_excel(writer, sheet_name="Merge_Debug", index=False)

        wb = writer.book
        hdr = wb.add_format({"bold": True, "bg_color": "#305496", "font_color": "white",
                              "border": 1, "align": "center", "valign": "vcenter"})
        cel = wb.add_format({"border": 1, "align": "left", "valign": "vcenter"})
        alt = wb.add_format({"border": 1, "bg_color": "#F2F2F2",
                              "align": "left", "valign": "vcenter"})

        ws = writer.sheets["Reporte"]
        for col, name in enumerate(grouped.columns):
            ws.write(0, col, name, hdr)
        ws.set_column("A:A", 45, cel)
        ws.set_column("B:B", 32, cel)
        ws.set_column("C:C", 40, cel)
        for row in range(1, len(grouped) + 1):
            ws.set_row(row, None, alt if row % 2 == 0 else cel)

        ws2 = writer.sheets["Extraido_PDF"]
        for col, name in enumerate(df_pdf.columns):
            ws2.write(0, col, name, hdr)
        ws2.set_column("A:A", 32)
        ws2.set_column("B:B", 55)

        ws3 = writer.sheets["Merge_Debug"]
        for col, name in enumerate(merged.columns):
            ws3.write(0, col, name, hdr)
        ws3.set_column(0, len(merged.columns) - 1, 22)

    output.seek(0)
    progress_bar.progress(1.0, "¡Listo!")
    return grouped, output.read(), df_pdf, merged

# ----------------------------------------------------------
# UI
# ----------------------------------------------------------
st.set_page_config(page_title="Reporte de Llaves M", layout="wide")
st.title("🗝️ Reporte de Llaves M desde PDF")
st.caption("Claude lee cada página visualmente — sin errores de parseo ni columnas mezcladas.")

pdf_file = st.file_uploader("📥 Sube tu PDF", type=["pdf"])

if pdf_file:
    if st.button("🚀 Generar Reporte", type="primary"):
        progress = st.progress(0, "Iniciando...")
        try:
            pdf_bytes = pdf_file.read()
            grouped_df, excel_data, extracted_df, merged_df = create_report_excel(pdf_bytes, progress)

            st.success("✅ Reporte generado correctamente")

            c1, c2, c3 = st.columns(3)
            c1.metric("Trabajos detectados", len(extracted_df))
            c2.metric("Filas reporte final", len(grouped_df))
            c3.metric("Cruces totales", len(merged_df))

            st.subheader("Vista previa: extraído del PDF")
            st.dataframe(extracted_df, use_container_width=True)

            st.subheader("Reporte final")
            st.dataframe(grouped_df, use_container_width=True)

            st.download_button(
                label="⬇️ Descargar Excel",
                data=excel_data,
                file_name="reporte_llaves_m_desde_pdf.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"❌ Error: {e}")
            import traceback
            st.code(traceback.format_exc())