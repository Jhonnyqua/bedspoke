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
    laparams = LAParams(line_margin=0.5, char_margin=3.0, word_margin=0.1)
    boxes = []
    for page_num, page_layout in enumerate(extract_pages(pdf_file, laparams=laparams)):
        page_width = page_layout.width
        page_height = page_layout.height
        for element in page_layout:
            if isinstance(element, LTTextBox):
                text = element.get_text().strip()
                if not text:
                    continue
                text_clean = re.sub(r"\s*\n\s*", " ", text).strip()
                boxes.append({
                    "page": page_num,
                    "page_height": page_height,
                    "x0": element.x0,
                    "y1": element.y1,
                    "x0_pct": element.x0 / page_width * 100,
                    "text": text_clean,
                })
    return boxes


# ----------------------------------------------------------
# CONVERTIR (page, y1) A POSICIÓN ABSOLUTA EN EL DOCUMENTO
# Para comparar posiciones entre páginas distintas
# ----------------------------------------------------------
def abs_position(page: int, y1: float, page_height: float) -> float:
    """Higher value = earlier in document (page 0 top > page 1 top)."""
    return page * 10000 + (page_height - y1)


# ----------------------------------------------------------
# LIMPIAR TEXTO DE DIRECCIÓN
# ----------------------------------------------------------
def clean_address_text(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    m = re.search(r"\b([2-9]\d{3})\b", text)
    if m:
        text = text[:m.start() + 4].strip()
    return text


def is_valid_address(text: str) -> bool:
    if not re.match(r"^[A-Za-z]?\d+[A-Za-z]?(?:/\d+[A-Za-z]?)?", text):
        return False
    return any(k.lower() in text.lower() for k in STREET_KEYWORDS)


# ----------------------------------------------------------
# LIMPIAR TEXTO DE CLEANER
# ----------------------------------------------------------
def is_cleaner_name_word(word: str) -> bool:
    if not word:
        return False
    if re.search(r"\d", word):
        return False
    if not re.fullmatch(r"[A-Za-z\xc0-\xff''\-]+", word):
        return False
    return True


def extract_cleaner_from_text(text: str) -> str:
    parts = re.split(r"\s*\|\s*|\n", text)
    name_words = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if re.search(r"\([0-9]+A[0-9]+C\)", part):
            break
        if re.search(r"\d{2}/\d{2}/\d{4}", part):
            break
        if re.search(r"\d", part):
            break
        if "," in part:
            break
        low = part.lower()
        banned = False
        for word in BANNED_NAME_WORDS:
            if re.search(r"\b" + re.escape(word) + r"\b", low):
                banned = True
                break
        if not banned:
            for phrase in BANNED_NAME_PHRASES:
                if phrase in low:
                    banned = True
                    break
        if banned:
            break
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

    address_list = []  # {abs_pos, address}
    cleaner_list = []  # {abs_pos, cleaner}

    for box in boxes:
        pct = box["x0_pct"]
        text = box["text"]
        apos = abs_position(box["page"], box["y1"], box["page_height"])

        if pct < 15:
            cleaned = clean_address_text(text)
            if is_valid_address(cleaned):
                address_list.append({"abs_pos": apos, "address": cleaned})

        elif pct > 70:
            cleaner = extract_cleaner_from_text(text)
            if cleaner != "Unassigned":
                cleaner_list.append({"abs_pos": apos, "cleaner": cleaner})

    # Sort both lists by document position
    address_list.sort(key=lambda x: x["abs_pos"])
    cleaner_list.sort(key=lambda x: x["abs_pos"])

    # One-to-one greedy matching: for each address, find the closest unused cleaner
    used_cleaner_indices = set()
    rows = []

    for addr in address_list:
        best_idx = None
        best_dist = float("inf")

        for i, c in enumerate(cleaner_list):
            if i in used_cleaner_indices:
                continue
            dist = abs(c["abs_pos"] - addr["abs_pos"])
            if dist < best_dist:
                best_dist = dist
                best_idx = i

        # Accept match if within ~300 abs units (generous tolerance)
        if best_idx is not None and best_dist <= 300:
            cleaner = cleaner_list[best_idx]["cleaner"]
            used_cleaner_indices.add(best_idx)
        else:
            cleaner = "Unassigned"

        rows.append({
            "Cleaner": cleaner,
            "Property Nickname": addr["address"],
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df = df.drop_duplicates(subset=["Property Nickname"])
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