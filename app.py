import re
from io import BytesIO

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextBox, LTTextLine, LTChar, LAParams

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
# PDF EXTRACTION — LEFT COLUMN ONLY BY X COORDINATE
# ----------------------------------------------------------
def extract_left_column_lines(pdf_file) -> list:
    """
    Extract text elements from the LEFT column only, using X coordinates.
    The PDF has 3 columns:
      Left (~x < 40% of page width):  Property address + cleaner name  ← we want this
      Middle:                          Task info (Scheduled, Due, etc.)
      Right:                           Reservation/guest info
    We collect text boxes whose left edge (x0) is in the left portion of the page,
    sort them top-to-bottom, and return as lines.
    """
    laparams = LAParams(line_margin=0.5, char_margin=3.0, word_margin=0.1)

    all_elements = []  # list of (page_num, y_top, x0, text)

    for page_num, page_layout in enumerate(extract_pages(pdf_file, laparams=laparams)):
        page_width = page_layout.width
        # Left column threshold: anything starting in the left ~38% of the page
        left_threshold = page_width * 0.38

        for element in page_layout:
            if isinstance(element, LTTextBox):
                x0 = element.x0
                # Only keep elements from the left column
                if x0 < left_threshold:
                    text = element.get_text().strip()
                    if text:
                        y_top = element.y1  # top of the element (higher = higher on page)
                        all_elements.append((page_num, y_top, x0, text))

    # Sort by page, then by y descending (top of page first)
    all_elements.sort(key=lambda e: (e[0], -e[1]))

    # Split each text box into individual lines
    lines = []
    for _, _, _, text in all_elements:
        for line in text.splitlines():
            line = line.strip()
            if line:
                lines.append(line)

    return lines


# ----------------------------------------------------------
# LIMPIEZA DE LÍNEAS
# ----------------------------------------------------------
def clean_line(line: str) -> str:
    if not isinstance(line, str):
        return ""
    line = line.replace("\u00a0", " ").replace("￾", " ").replace("\t", " ")
    return re.sub(r"\s+", " ", line).strip()


def is_footer_or_header(line: str) -> bool:
    line = clean_line(line)
    if not line:
        return True
    patterns = [
        r"^Housekeeping Daily Summary$",
        r"^\d{2}/\d{2}/\d{4}$",
        r"^Housekeeping Tasks\b",
        r"^Property Housekeeping$",
        r"^Task$",
        r"^Reservation Assigned To$",
        r"^Brisbane$",
        r"^Bedspoke",
        r"^Address:",
        r"^Email:",
        r"^Phone:",
        r"^Printed by Resly$",
        r"^Page \d+ of \d+$",
        r"^\d{1,2} \w{3} \d{4}.*Printed by Resly",
        r"^\d{1,2} \w{3} \d{4} \d{2}:\d{2}:\d{2}",
    ]
    return any(re.search(p, line, flags=re.IGNORECASE) for p in patterns)


def is_date_stay_line(line: str) -> bool:
    return bool(re.search(r"\d{2}/\d{2}/\d{4}\s+to\s+\d{2}/\d{2}/\d{4}", line))


def is_guest_line(line: str) -> bool:
    return bool(re.search(r"\([0-9]+A[0-9]+C\)", line))


def is_operation_line(line: str) -> bool:
    bad_exact = {
        "Scheduled", "In House", "Done", "[D] Depart Clean", "[D] Depart",
        "Clean", "Departing", "Departing Today", "Departure", "Arriving",
        "Arriving Today", "Arrival", "Back To Back", "Due:", "Unassigned",
        "Property Housekeeping", "Task", "Reservation Assigned To",
    }
    if line in bad_exact:
        return True
    bad_patterns = [
        r"^Due:", r"^\d{2}/\d{2}/\d{4}$",
        r"^\d{2}/\d{2}/\d{4}\s+to\s+\d{2}/\d{2}/\d{4}",
        r"^\d+N$", r"^ETA:", r"^Welcome", r"^Please ",
        r"^Departure Time:", r"^There was a long stay",
        r"^No Arrival", r"^IMPORTANT", r"^return the", r"^pls return",
        r"^DEEP CLEAN", r"^Guest requested", r"^G feedback",
        r"^11am ", r"^10:30am ", r"^12nn$", r"^3pm$",
        r"^8am$", r"^3PM$", r"^2:30pm$",
    ]
    return any(re.search(p, line, flags=re.IGNORECASE) for p in bad_patterns)


# ----------------------------------------------------------
# DETECCIÓN DE DIRECCIONES
# ----------------------------------------------------------
def looks_like_address_start(line: str, next_line: str = "") -> bool:
    line = clean_line(line)
    if not line or is_footer_or_header(line) or is_date_stay_line(line):
        return False
    if not re.match(r"^[A-Za-z]?\d+[A-Za-z]?(?:/\d+[A-Za-z]?)?", line):
        return False
    combined = line + " " + clean_line(next_line)
    return any(k.lower() in combined.lower() for k in STREET_KEYWORDS)


def build_address(lines, start_idx):
    parts = [lines[start_idx]]
    i = start_idx + 1
    while i < len(lines):
        line = lines[i]
        if not line or is_footer_or_header(line):
            break
        next_line = lines[i + 1] if i + 1 < len(lines) else ""
        if looks_like_address_start(line, next_line):
            break
        if is_operation_line(line) or is_guest_line(line) or is_date_stay_line(line):
            break
        has_postcode = bool(re.search(
            r"\bQLD\b|\bNSW\b|\bVIC\b|\bACT\b|\bWA\b|\bSA\b|\bTAS\b|\bNT\b|\b\d{4}\b",
            line, re.IGNORECASE
        ))
        is_safe = bool(re.fullmatch(r"[A-Za-z0-9,.\- ]+", line))
        if has_postcode or is_safe:
            parts.append(line)
            i += 1
            if re.search(r"\b\d{4}\b", " ".join(parts)):
                break
            continue
        break

    address = re.sub(r"\s+,", ",", " ".join(parts))
    address = re.sub(r"\s{2,}", " ", address).strip()
    return address, i


# ----------------------------------------------------------
# DETECCIÓN DE NOMBRE DEL CLEANER
# ----------------------------------------------------------
def is_cleaner_name_line(line: str) -> bool:
    line = clean_line(line)
    if not line:
        return False
    if is_footer_or_header(line) or is_operation_line(line):
        return False
    if is_guest_line(line) or is_date_stay_line(line):
        return False
    if re.search(r"\d", line):
        return False
    if "," in line or "(" in line or ")" in line:
        return False

    low = line.lower()
    for word in BANNED_NAME_WORDS:
        if re.search(r"\b" + re.escape(word) + r"\b", low):
            return False
    for phrase in BANNED_NAME_PHRASES:
        if phrase in low:
            return False

    if not re.fullmatch(r"[A-Za-z\xc0-\xff''\- ]+", line):
        return False

    return 1 <= len(line.split()) <= 6


def extract_cleaner_from_block(block_lines: list) -> str:
    name_parts = [clean_line(l) for l in block_lines if is_cleaner_name_line(l)]
    if not name_parts:
        return "Unassigned"
    return re.sub(r"\s{2,}", " ", " ".join(name_parts)).strip() or "Unassigned"


# ----------------------------------------------------------
# PARSER PRINCIPAL
# ----------------------------------------------------------
def parse_housekeeping_pdf(pdf_file) -> pd.DataFrame:
    # Extract ONLY the left column lines (address + cleaner name)
    raw_lines = extract_left_column_lines(pdf_file)
    all_lines = [clean_line(l) for l in raw_lines if clean_line(l)]

    # Pass 1: locate all addresses
    address_entries = []
    i = 0
    while i < len(all_lines):
        next_line = all_lines[i + 1] if i + 1 < len(all_lines) else ""
        if looks_like_address_start(all_lines[i], next_line):
            address, next_i = build_address(all_lines, i)
            address_entries.append({
                "address": address,
                "addr_start": i,
                "addr_end": next_i,
            })
            i = next_i
        else:
            i += 1

    # Pass 2: cleaner = name lines in block AFTER address, BEFORE next address
    rows = []
    for idx, entry in enumerate(address_entries):
        block_start = entry["addr_end"]
        block_end = (
            address_entries[idx + 1]["addr_start"]
            if idx + 1 < len(address_entries)
            else len(all_lines)
        )
        block_lines = [
            l for l in all_lines[block_start:block_end]
            if not is_footer_or_header(l)
        ]
        cleaner = extract_cleaner_from_block(block_lines)
        rows.append({"Cleaner": cleaner, "Property Nickname": entry["address"]})

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