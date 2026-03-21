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
# EXTRAER PDF COMO IMÁGENES BASE64 (página por página)
# ----------------------------------------------------------
def pdf_to_base64_images(pdf_bytes: bytes) -> list[str]:
    """Convierte cada página del PDF a imagen PNG base64."""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise ImportError("Instalá PyMuPDF: pip install pymupdf")

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page in doc:
        mat = fitz.Matrix(2, 2)  # 2x zoom = ~150dpi
        pix = page.get_pixmap(matrix=mat)
        images.append(base64.b64encode(pix.tobytes("png")).decode())
    doc.close()
    return images

# ----------------------------------------------------------
# LLAMAR A CLAUDE API CON LAS IMÁGENES
# ----------------------------------------------------------
def extract_with_claude_api(page_images: list[str], progress_bar) -> list[dict]:
    """
    Manda todas las páginas a Claude en un solo llamado.
    Pide JSON con lista de {address, cleaner}.
    """
    import requests

    # Construir content con todas las imágenes
    content = []
    for i, img_b64 in enumerate(page_images):
        content.append({
            "type": "text",
            "text": f"=== PÁGINA {i+1} ==="
        })
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/png",
                "data": img_b64
            }
        })

    content.append({
        "type": "text",
        "text": """Este es un reporte "Housekeeping Daily Summary" de múltiples páginas.

Para CADA propiedad listada, extrae:
1. La dirección completa de la propiedad (columna izquierda)
2. El nombre del cleaner/housekeeper asignado (columna derecha, fila "Assigned To")

IMPORTANTE:
- Si una dirección se corta al final de la página y continúa en la siguiente, unilas.
- Si el cleaner dice "Unassigned", ponlo como "Unassigned".
- Ignorá encabezados, pies de página, y datos de reservas/huéspedes.
- La columna derecha tiene: Assigned To (nombre del cleaner), luego datos de reserva (ignorar).

Respondé ÚNICAMENTE con un JSON válido, sin texto antes ni después, sin markdown:
[
  {"address": "dirección completa", "cleaner": "Nombre Cleaner"},
  ...
]"""
    })

    progress_bar.progress(0.3, "Enviando páginas a Claude...")

    response = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"Content-Type": "application/json"},
        json={
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 4096,
            "messages": [{"role": "user", "content": content}]
        },
        timeout=120
    )

    progress_bar.progress(0.7, "Procesando respuesta...")

    if response.status_code != 200:
        raise ValueError(f"Error API Claude: {response.status_code} - {response.text}")

    data = response.json()
    raw = ""
    for block in data.get("content", []):
        if block.get("type") == "text":
            raw += block["text"]

    # Limpiar posibles backticks
    raw = re.sub(r"```json|```", "", raw).strip()

    try:
        records = json.loads(raw)
    except json.JSONDecodeError as e:
        raise ValueError(f"No pude parsear el JSON de Claude:\n{raw[:500]}\nError: {e}")

    return records

# ----------------------------------------------------------
# GENERAR EXCEL
# ----------------------------------------------------------
def create_report_excel(pdf_bytes: bytes, progress_bar):
    progress_bar.progress(0.1, "Convirtiendo PDF a imágenes...")
    images = pdf_to_base64_images(pdf_bytes)

    records = extract_with_claude_api(images, progress_bar)
    progress_bar.progress(0.8, "Cruzando con Key Register...")

    df_pdf = pd.DataFrame(records)
    if df_pdf.empty or "address" not in df_pdf.columns:
        raise ValueError("Claude no retornó datos válidos.")

    df_pdf = df_pdf.rename(columns={"address": "Property Nickname", "cleaner": "Cleaner"})
    df_pdf["Cleaner"] = df_pdf["Cleaner"].fillna("Unassigned").astype(str).str.strip()
    df_pdf["Property Nickname"] = df_pdf["Property Nickname"].fillna("").astype(str).str.strip()
    df_pdf = df_pdf[df_pdf["Property Nickname"] != ""]

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

    progress_bar.progress(0.95, "Generando Excel...")

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
st.write("Sube el PDF de Housekeeping Daily Summary. Claude lo lee visualmente — sin errores de parseo.")

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