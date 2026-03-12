import re
from io import BytesIO

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextBox, LAParams

st.set_page_config(page_title="Reporte de Llaves M", layout="wide")

# ----------------------------------------------------------
# CONSTANTES
# ----------------------------------------------------------
STREET_KEYWORDS = [
    "Street", "St", "Road", "Rd", "Terrace", "Tce", "Lane", "Way",
    "Quay", "Avenue", "Ave", "Grove", "Court", "Ct", "Boulevard",
    "Bvd", "Drive", "Dr", "Place", "Pl", "Close", "Cl",
]

# Palabras que indican que una caja de la col derecha NO es un nombre de cleaner
BANNED_NAME_WORDS = [
    "reservation", "arrival", "arriving", "depart", "departure",
    "check", "guest", "housekeeping", "printed", "resly", "welcome",
    "bedspoke", "clean", "done", "scheduled", "unassigned",
    "due", "eta", "please", "bring", "important", "feedback",
    "property", "task", "back", "deep", "return", "assigned",
]

# Cajas de col izquierda que siempre deben ignorarse
SKIP_LEFT_EXACT = {
    "Brisbane", "Bedspoke Pty Ltd", "Property", "Housekeeping Tasks",
}
SKIP_LEFT_STARTSWITH = (
    "Address:", "Email:", "Phone:", "ABN:", "ACN:", "Licence",
    "Licensee:", "12 Mar", "13 Mar", "14 Mar",
)

# Cajas de col derecha que siempre deben ignorarse
SKIP_RIGHT_EXACT = {
    "Assigned To", "63 Depart",
}
SKIP_RIGHT_STARTSWITH = ("Page ",)


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


# ----------------------------------------------------------
# NORMALIZAR DIRECCIÓN
# ----------------------------------------------------------
def simplify_address_15chars(address: str) -> str:
    if not isinstance(address, str):
        return ""
    address = address.strip()
    m = re.search(r"\d", address)
    if m:
        substr = address[m.start():m.start() + 15]
    else:
        substr = address[:15]
    return re.sub(r"[^0-9A-Za-z\s]", "", substr).lower().strip()


# ----------------------------------------------------------
# EXTRAER CAJAS EN ORDEN DE DOCUMENTO
# ----------------------------------------------------------
def extract_boxes_ordered(pdf_file) -> list:
    """
    Retorna todas las cajas de texto, ordenadas por posición en el documento:
    página ascendente, luego y1 descendente (de arriba a abajo).
    """
    laparams = LAParams(line_margin=0.5, char_margin=3.0, word_margin=0.1)
    boxes = []
    for page_num, page_layout in enumerate(extract_pages(pdf_file, laparams=laparams)):
        page_width = page_layout.width
        page_height = page_layout.height
        for element in page_layout:
            if isinstance(element, LTTextBox):
                raw = element.get_text().strip()
                if not raw:
                    continue
                text = re.sub(r"\s*\n\s*", " | ", raw).strip()
                boxes.append({
                    "page": page_num,           # 0-indexed
                    "x0": element.x0,
                    "y1": element.y1,
                    "x0_pct": element.x0 / page_width * 100,
                    "page_height": page_height,
                    "text": text,
                })
    # Orden documento: página ASC, y1 DESC (arriba primero)
    boxes.sort(key=lambda b: (b["page"], -b["y1"]))
    return boxes


# ----------------------------------------------------------
# CLASIFICAR CAJAS
# ----------------------------------------------------------
def is_left_col(box) -> bool:
    return box["x0_pct"] < 15

def is_right_col(box) -> bool:
    return box["x0_pct"] > 75

def should_skip_left(text: str) -> bool:
    t = text.strip()
    if t in SKIP_LEFT_EXACT:
        return True
    for prefix in SKIP_LEFT_STARTSWITH:
        if t.startswith(prefix):
            return True
    return False

def should_skip_right(text: str) -> bool:
    t = text.strip()
    if t in SKIP_RIGHT_EXACT:
        return True
    for prefix in SKIP_RIGHT_STARTSWITH:
        if t.startswith(prefix):
            return True
    return False

def has_street_keyword(text: str) -> bool:
    return any(k.lower() in text.lower() for k in STREET_KEYWORDS)

def starts_with_number(text: str) -> bool:
    return bool(re.match(r"^[A-Za-z]?\d", text.strip()))

def is_valid_address(text: str) -> bool:
    """Una dirección válida empieza con (letra+)?número y tiene keyword de calle."""
    return starts_with_number(text) and has_street_keyword(text)

def clean_address(text: str) -> str:
    """Reemplaza separadores de línea, corta en postcode si existe."""
    text = re.sub(r"\s*\|\s*", " ", text).strip()
    text = re.sub(r"\s+", " ", text)
    m = re.search(r"\b([2-9]\d{3})\b", text)
    if m:
        text = text[:m.start() + 4].strip()
    return text

def extract_cleaner_name(text: str) -> str:
    """
    Extrae nombre de cleaner de una caja de la col derecha.
    El nombre viene ANTES de los datos de reserva (XAxC, fechas, comas, palabras baneadas).
    """
    parts = re.split(r"\s*\|\s*|\n", text)
    name_words = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        # Stop signals
        if re.search(r"\([0-9]+A[0-9]+C\)", part):
            break
        if re.search(r"\d{2}/\d{2}/\d{4}", part):
            break
        if re.search(r"\d", part):
            break
        if "," in part:
            break
        low = part.lower()
        if any(re.search(r"\b" + re.escape(w) + r"\b", low) for w in BANNED_NAME_WORDS):
            break
        words = part.split()
        if all(re.fullmatch(r"[A-Za-z\xc0-\xff''\-]+", w) for w in words) and 1 <= len(words) <= 5:
            name_words.extend(words)
        else:
            break
    return " ".join(name_words) if name_words else ""


# ----------------------------------------------------------
# PARSER PRINCIPAL: secuencia de bloques
# ----------------------------------------------------------
def parse_housekeeping_pdf(pdf_file) -> pd.DataFrame:
    """
    Estrategia: procesa las cajas en orden de documento.
    - Col izquierda (x0_pct < 15%): identifica inicio de nueva propiedad o continuación.
    - Col derecha (x0_pct > 75%): busca nombre de cleaner.
    - Cada vez que encontramos una dirección válida, iniciamos un nuevo record.
    - Una caja de la izq sin número pero con suburb (ej. "Spring Hill QLD 4000")
      se ADHIERE al record anterior si este no tiene postcode aún.
    - Una caja de la derecha con nombre válido se asigna al record activo.
    """
    boxes = extract_boxes_ordered(pdf_file)

    records = []          # lista de {"address": str, "cleaner": str}
    current = None        # record en construcción

    # Para detectar "pending right box" — una caja derecha sin record izq aún
    pending_right = None  # nombre de cleaner pendiente de asignar

    for box in boxes:
        text = box["text"].strip()
        if not text:
            continue

        if is_left_col(box):
            if should_skip_left(text):
                continue

            cleaned = clean_address(text)

            if is_valid_address(cleaned):
                # Nueva propiedad
                if current is not None:
                    records.append(current)
                current = {"address": cleaned, "cleaner": ""}

                # ¿Había un cleaner pendiente del salto de página anterior?
                if pending_right:
                    current["cleaner"] = pending_right
                    pending_right = None

            else:
                # Posible continuación de dirección (ej. "Spring Hill QLD 4000")
                # o texto irrelevante (ej. "Back To Back", "Fortitude Valley")
                # Solo adhirimos si el record actual no tiene postcode y el texto
                # parece un suburb (letras, sin número al inicio, corto)
                if (current is not None
                        and not re.search(r"\b[2-9]\d{3}\b", current["address"])
                        and not starts_with_number(cleaned)
                        and len(cleaned) < 50
                        and not re.search(r"\d", cleaned)):
                    # Verificar que no es "Back To Back" u otro ruido
                    low = cleaned.lower()
                    if not any(w in low for w in ["back", "printed", "resly", "due", "housekeeping"]):
                        # Adhirir
                        current["address"] = current["address"].rstrip(",").strip() + " " + cleaned
                        # Re-limpiar al postcode
                        current["address"] = clean_address(current["address"])

        elif is_right_col(box):
            if should_skip_right(text):
                continue

            name = extract_cleaner_name(text)
            if not name:
                continue

            if current is not None:
                if not current["cleaner"]:
                    current["cleaner"] = name
                # Si ya tiene cleaner, ignorar (es de otro bloque)
            else:
                # Todavía no hay record activo — guardar como pending
                # (caso: cleaner aparece antes que la dirección en el stream)
                pending_right = name

    # Guardar el último record
    if current is not None:
        records.append(current)

    df = pd.DataFrame(records)
    if df.empty:
        return df

    df["cleaner"] = df["cleaner"].replace("", "Unassigned")
    df = df.rename(columns={"address": "Property Nickname", "cleaner": "Cleaner"})
    df = df.drop_duplicates(subset=["Property Nickname"])
    df = df[df["Property Nickname"].str.strip() != ""]
    return df[["Cleaner", "Property Nickname"]]


# ----------------------------------------------------------
# CARGAR KEY REGISTER
# ----------------------------------------------------------
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
# GENERAR EXCEL
# ----------------------------------------------------------
def create_report_excel(pdf_file):
    df_pdf = parse_housekeeping_pdf(pdf_file)
    if df_pdf.empty:
        raise ValueError("No pude extraer propiedades del PDF.")

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
    return grouped, output.read(), df_pdf, merged


# ----------------------------------------------------------
# UI
# ----------------------------------------------------------
st.title("🗝️ Reporte de Llaves M desde PDF")
st.write("Sube el PDF de Housekeeping Daily Summary para cruzarlo con el Key Register.")

pdf_file = st.file_uploader("📥 Sube tu PDF", type=["pdf"])

if pdf_file:
    try:
        grouped_df, excel_data, extracted_df, merged_df = create_report_excel(pdf_file)

        st.success("✅ Reporte generado correctamente")

        c1, c2, c3 = st.columns(3)
        c1.metric("Trabajos detectados en PDF", len(extracted_df))
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
else:
    st.info("📄 Esperando que subas un PDF")