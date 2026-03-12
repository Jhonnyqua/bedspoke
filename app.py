import re
from io import BytesIO

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextBox, LAParams

st.set_page_config(page_title="Reporte de Llaves M", layout="wide")

STREET_KEYWORDS = [
    "Street", "St", "Road", "Rd", "Terrace", "Tce", "Lane", "Way",
    "Quay", "Avenue", "Ave", "Grove", "Court", "Ct", "Boulevard",
    "Bvd", "Drive", "Dr", "Place", "Pl", "Close", "Cl"
]

BANNED_NAME_WORDS = [
    "reservation", "arrival", "arriving", "depart", "departure",
    "check", "guest", "housekeeping", "printed", "resly", "welcome",
    "bedspoke", "clean", "done", "scheduled", "unassigned",
    "due", "eta", "please", "bring", "important", "feedback",
    "property", "task",
]
BANNED_NAME_PHRASES = ["back to back", "deep clean", "return the", "pls return"]


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
    match = re.search(r"\d", address)
    if match:
        substr = address[match.start():match.start() + 15]
    else:
        substr = address[:15]
    return re.sub(r"[^0-9A-Za-z\s]", "", substr).lower().strip()


# ----------------------------------------------------------
# EXTRAER CAJAS DE TEXTO CON COORDENADAS
# ----------------------------------------------------------
def extract_boxes(pdf_file) -> list:
    """
    Returns all text boxes with page, x0, y1, x0_pct, text.
    Text within each box has newlines replaced with spaces.
    """
    laparams = LAParams(line_margin=0.5, char_margin=3.0, word_margin=0.1)
    boxes = []
    for page_num, page_layout in enumerate(extract_pages(pdf_file, laparams=laparams)):
        page_width = page_layout.width
        for element in page_layout:
            if isinstance(element, LTTextBox):
                text = element.get_text().strip()
                if not text:
                    continue
                # Normalize internal newlines to spaces
                text = re.sub(r"\s*\n\s*", " ", text).strip()
                boxes.append({
                    "page": page_num,
                    "x0": element.x0,
                    "y1": element.y1,   # top of box (higher = higher on page)
                    "x0_pct": element.x0 / page_width * 100,
                    "text": text,
                })
    return boxes


# ----------------------------------------------------------
# LIMPIAR TEXTO DE DIRECCIÓN
# ----------------------------------------------------------
def clean_address_text(text: str) -> str:
    """
    Address boxes contain only the address lines separated by spaces.
    Remove any trailing operational text just in case.
    """
    text = re.sub(r"\s+", " ", text).strip()
    # Cut at postcode if present
    m = re.search(r"\b([2-9]\d{3})\b", text)
    if m:
        text = text[:m.start() + 4].strip()
    return text


def is_valid_address(text: str) -> bool:
    """Check if text looks like a property address."""
    if not re.match(r"^[A-Za-z]?\d+[A-Za-z]?(?:/\d+[A-Za-z]?)?", text):
        return False
    return any(k.lower() in text.lower() for k in STREET_KEYWORDS)


# ----------------------------------------------------------
# LIMPIAR TEXTO DE CLEANER
# ----------------------------------------------------------
def is_cleaner_name_word(word: str) -> bool:
    """Check if a single word could be part of a cleaner name."""
    if not word:
        return False
    if re.search(r"\d", word):
        return False
    if not re.fullmatch(r"[A-Za-z\xc0-\xff''\-]+", word):
        return False
    return True


def extract_cleaner_from_text(text: str) -> str:
    """
    From the right-column text box, extract the cleaner name.
    The box may contain: 'Thomas Clarisse' or 'Brenda Anahi Bedon'
    It may also contain noise from the Reservation column bleeding in.
    We take only words that look like name words (letters only, no digits).
    Stop at first line that looks like a guest entry or date.
    """
    # Split on pipe separators or newlines that pdfminer may produce
    parts = re.split(r"\s*\|\s*|\n", text)
    
    name_words = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        # Stop if this looks like a guest line (SURNAME, Firstname format)
        if re.search(r"\([0-9]+A[0-9]+C\)", part):
            break
        if re.search(r"\d{2}/\d{2}/\d{4}", part):
            break
        if re.search(r"\d", part):
            break
        if "," in part:
            break
        # Check banned words
        low = part.lower()
        banned = False
        for word in BANNED_NAME_WORDS:
            if re.search(r"\b" + re.escape(word) + r"\b", low):
                banned = True
                break
        for phrase in BANNED_NAME_PHRASES:
            if phrase in low:
                banned = True
                break
        if banned:
            break
        # Only keep if all words are name-like
        words = part.split()
        if all(is_cleaner_name_word(w) for w in words) and 1 <= len(words) <= 4:
            name_words.extend(words)
            if len(name_words) >= 5:
                break

    if not name_words:
        return "Unassigned"
    return " ".join(name_words)


# ----------------------------------------------------------
# PARSER PRINCIPAL
# ----------------------------------------------------------
def parse_housekeeping_pdf(pdf_file) -> pd.DataFrame:
    boxes = extract_boxes(pdf_file)

    # Separate into address boxes (x0_pct < 15%) and cleaner boxes (x0_pct > 75%)
    address_boxes = []
    cleaner_boxes = []

    for box in boxes:
        pct = box["x0_pct"]
        text = box["text"]

        if pct < 15:
            # Left column — check if it's a valid address
            cleaned = clean_address_text(text)
            if is_valid_address(cleaned):
                address_boxes.append({
                    "page": box["page"],
                    "y1": box["y1"],
                    "address": cleaned,
                })

        elif pct > 75:
            # Right column — potential cleaner name
            cleaner = extract_cleaner_from_text(text)
            if cleaner != "Unassigned":
                cleaner_boxes.append({
                    "page": box["page"],
                    "y1": box["y1"],
                    "cleaner": cleaner,
                })

    # Match each address to the nearest cleaner box by (page, y1)
    rows = []
    for addr in address_boxes:
        best_cleaner = "Unassigned"
        best_dist = float("inf")

        for c in cleaner_boxes:
            if c["page"] != addr["page"]:
                continue
            dist = abs(c["y1"] - addr["y1"])
            if dist < best_dist:
                best_dist = dist
                best_cleaner = c["cleaner"]

        # Only accept match if within ~60 points vertically
        if best_dist > 60:
            best_cleaner = "Unassigned"

        rows.append({
            "Cleaner": best_cleaner,
            "Property Nickname": addr["address"],
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df = df.drop_duplicates()
    df["Cleaner"] = df["Cleaner"].fillna("Unassigned").astype(str).str.strip()
    df["Property Nickname"] = df["Property Nickname"].fillna("").astype(str).str.strip()
    return df[df["Property Nickname"] != ""]


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
# GENERAR REPORTE
# ----------------------------------------------------------
def create_grouped_excel_from_pdf(pdf_file):
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
        bnd = wb.add_format({"border": 1, "bg_color": "#F2F2F2",
                              "align": "left", "valign": "vcenter"})

        ws = writer.sheets["Reporte"]
        for col, name in enumerate(grouped.columns):
            ws.write(0, col, name, hdr)
        ws.set_column("A:A", 42, cel)
        ws.set_column("B:B", 30, cel)
        ws.set_column("C:C", 40, cel)
        for row in range(1, len(grouped) + 1):
            ws.set_row(row, None, bnd if row % 2 == 0 else cel)

        ws2 = writer.sheets["Extraido_PDF"]
        for col, name in enumerate(df_pdf.columns):
            ws2.write(0, col, name, hdr)
        ws2.set_column("A:A", 30)
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
st.write("Sube el PDF de Housekeeping Daily Summary para cruzarlo con tu Key Register.")

pdf_file = st.file_uploader("📥 Sube tu PDF", type=["pdf"])

if pdf_file:
    try:
        grouped_df, excel_data, extracted_df, merged_df = create_grouped_excel_from_pdf(pdf_file)

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
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
        import traceback
        st.code(traceback.format_exc())
else:
    st.info("📄 Esperando que subas un PDF")