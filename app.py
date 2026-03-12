import re
from io import BytesIO

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from pypdf import PdfReader

# ----------------------------------------------------------
# CONFIG
# ----------------------------------------------------------
st.set_page_config(page_title="Reporte de Llaves M", layout="wide")

STREET_KEYWORDS = [
    "Street", "St", "Road", "Rd", "Terrace", "Tce", "Lane", "Way",
    "Quay", "Avenue", "Ave", "Grove", "Court", "Ct", "Boulevard",
    "Bvd", "Drive", "Dr", "Place", "Pl", "Close", "Cl"
]


# ----------------------------------------------------------
# 1. GOOGLE SHEETS
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
# 2. NORMALIZAR DIRECCIÓN
# ----------------------------------------------------------
def simplify_address_15chars(address: str) -> str:
    if not isinstance(address, str):
        return ""
    address = address.strip()
    match = re.search(r"\d", address)
    if match:
        start = match.start()
        substr = address[start:start + 15]
    else:
        substr = address[:15]
    return re.sub(r"[^0-9A-Za-z\s]", "", substr).lower().strip()


# ----------------------------------------------------------
# 3. PDF A TEXTO
# ----------------------------------------------------------
def extract_text_from_pdf(pdf_file) -> str:
    reader = PdfReader(pdf_file)
    text_parts = []
    for page in reader.pages:
        txt = page.extract_text()
        if txt:
            text_parts.append(txt)
    return "\n".join(text_parts)


# ----------------------------------------------------------
# 4. LIMPIEZA DE LÍNEAS
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
        r"^Bedspoke$",
        r"^Bedspoke Pty Ltd",
        r"^Address:",
        r"^Email:",
        r"^Phone:",
        r"^Printed by Resly$",
        r"^Page \d+ of \d+$",
        r"^\d{1,2} \w{3} \d{4} .* Printed by Resly Page \d+ of \d+$",
        r"^\d{1,2} \w{3} \d{4} \d{2}:\d{2}:\d{2}$",
        r"^\d{1,2} \w{3} \d{4} \d{2}:\d{2}:\d{2} \| Printed by Resly Page \d+ of \d+.*$",
    ]
    return any(re.search(p, line, flags=re.IGNORECASE) for p in patterns)


def is_date_stay_line(line: str) -> bool:
    return bool(re.search(r"\d{2}/\d{2}/\d{4}\s+to\s+\d{2}/\d{2}/\d{4}", clean_line(line)))


def is_guest_line(line: str) -> bool:
    return bool(re.search(r"\([0-9]+A[0-9]+C\)", clean_line(line)))


def is_operation_line(line: str) -> bool:
    line = clean_line(line)
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
        r"^\d+N$", r"^ETA:", r"^Welcome and enjoy your stay",
        r"^Welcome and Happy birthday", r"^Please ", r"^Departure Time:",
        r"^There was a long stay", r"^No Arrival Reservation", r"^IMPORTANT",
        r"^return the", r"^pls return", r"^DEEP CLEAN", r"^Guest requested",
        r"^G feedback", r"^11am ", r"^10:30am ", r"^12nn$", r"^3pm$",
        r"^8am$", r"^3PM$", r"^2:30pm$",
    ]
    return any(re.search(p, line, flags=re.IGNORECASE) for p in bad_patterns)


def _has_street_keyword(text: str) -> bool:
    return any(k.lower() in text.lower() for k in STREET_KEYWORDS)


def _starts_like_property(text: str) -> bool:
    return bool(re.match(r"^[A-Za-z]?\d+[A-Za-z]?(?:/\d+[A-Za-z]?)?", text))


# ----------------------------------------------------------
# 5. DETECCIÓN DE DIRECCIONES
# ----------------------------------------------------------
def looks_like_address_start(line: str, next_line: str = "") -> bool:
    """
    KEY FIX: pypdf splits 'Thomas Clarisse' and '1/56 Dobson Street'
    into separate lines. A line like '1/56 Dobson' has the number pattern
    but the street keyword is on the next line 'Street, Ascot,'.
    We check BOTH the current line and the next line for the street keyword.
    """
    line = clean_line(line)
    if not line or is_footer_or_header(line) or is_date_stay_line(line):
        return False
    if not _starts_like_property(line):
        return False
    combined = line + " " + clean_line(next_line)
    return _has_street_keyword(combined)


def build_address(lines, start_idx):
    parts = [clean_line(lines[start_idx])]
    i = start_idx + 1

    while i < len(lines):
        line = clean_line(lines[i])
        if not line or is_footer_or_header(line):
            break
        next_line = clean_line(lines[i + 1]) if i + 1 < len(lines) else ""
        if looks_like_address_start(line, next_line):
            break
        if is_operation_line(line) or is_guest_line(line) or is_date_stay_line(line):
            break
        has_postcode = bool(re.search(
            r"\bQLD\b|\bNSW\b|\bVIC\b|\bACT\b|\bWA\b|\bSA\b|\bTAS\b|\bNT\b|\b\d{4}\b",
            line, flags=re.IGNORECASE
        ))
        is_safe_chars = bool(re.fullmatch(r"[A-Za-z0-9,.\- ]+", line))
        if has_postcode or is_safe_chars:
            parts.append(line)
            i += 1
            if re.search(r"\b\d{4}\b", " ".join(parts)):
                break
            continue
        break

    address = " ".join(parts)
    address = re.sub(r"\s+,", ",", address)
    address = re.sub(r"\s{2,}", " ", address).strip()
    return address, i


# ----------------------------------------------------------
# 6. DETECCIÓN DE NOMBRE DEL CLEANER
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
    banned_fragments = [
        "reservation", "arrival", "arriving", "depart", "departure",
        "check", "guest", "housekeeping", "printed", "resly", "welcome",
        "bedspoke", "clean", "done", "scheduled", "unassigned",
        "back to back", "due", "eta", "please", "return", "bring",
        "important", "deep clean", "feedback",
    ]
    if any(x in line.lower() for x in banned_fragments):
        return False
    if not re.fullmatch(r"[A-Za-z\xc0-\xff''\- ]+", line):
        return False
    words = line.split()
    return 1 <= len(words) <= 6


def extract_cleaner_from_block(block_lines: list) -> str:
    """
    In the PDF, the cleaner name appears at the TAIL of each property block
    (just before the next address). It may span 1-3 lines e.g.:
      'Brenda Anahi'
      'Bedon'
    or:
      'Thomas'
      'Clarisse'
    We collect ALL lines that look like name lines in the block.
    """
    name_parts = []
    for line in block_lines:
        line = clean_line(line)
        if is_cleaner_name_line(line):
            name_parts.append(line)

    if not name_parts:
        return "Unassigned"

    cleaner = " ".join(name_parts)
    return re.sub(r"\s{2,}", " ", cleaner).strip() or "Unassigned"


# ----------------------------------------------------------
# 7. PARSER PRINCIPAL DEL PDF
# ----------------------------------------------------------
def parse_housekeeping_pdf(pdf_file) -> pd.DataFrame:
    text = extract_text_from_pdf(pdf_file)
    all_lines = [clean_line(l) for l in text.splitlines()]
    all_lines = [l for l in all_lines if l]

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

    # Pass 2: for each address, the cleaner is in the block AFTER the address
    # and BEFORE the next address starts
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

        rows.append({
            "Cleaner": cleaner,
            "Property Nickname": entry["address"],
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df = df.drop_duplicates()
    df["Cleaner"] = df["Cleaner"].fillna("Unassigned").astype(str).str.strip()
    df["Property Nickname"] = df["Property Nickname"].fillna("").astype(str).str.strip()
    df = df[df["Property Nickname"] != ""]
    return df


# ----------------------------------------------------------
# 8. CARGAR KEY REGISTER
# ----------------------------------------------------------
def load_key_register() -> pd.DataFrame:
    client = authorize_gspread()
    sheet_id = st.secrets["gcp_service_account"]["spreadsheet_id"]
    sheet = client.open_by_key(sheet_id).worksheet("Key Register")
    data = sheet.get_all_values()

    if len(data) < 2:
        raise ValueError("La hoja 'Key Register' no tiene suficiente información.")

    df_keys = pd.DataFrame(data[2:], columns=data[1]).drop(columns="", errors="ignore")

    required_cols = ["Property Address", "Tag"]
    for col in required_cols:
        if col not in df_keys.columns:
            raise ValueError(f"No encontré la columna '{col}' en la hoja 'Key Register'.")

    if "Observation" in df_keys.columns:
        df_keys = df_keys[df_keys["Observation"].fillna("").str.strip() == ""]

    df_keys["Property Address"] = df_keys["Property Address"].fillna("").astype(str).str.strip()
    df_keys["Tag"] = df_keys["Tag"].fillna("").astype(str).str.strip()
    return df_keys


# ----------------------------------------------------------
# 9. GENERAR REPORTE
# ----------------------------------------------------------
def create_grouped_excel_from_pdf(pdf_file):
    df_pdf = parse_housekeeping_pdf(pdf_file)

    if df_pdf.empty:
        raise ValueError("No pude extraer propiedades del PDF. Revisa si cambió el formato.")

    df_keys = load_key_register()

    df_pdf["Simplified"] = df_pdf["Property Nickname"].apply(simplify_address_15chars)
    df_keys["Simplified"] = df_keys["Property Address"].apply(simplify_address_15chars)

    merged = pd.merge(df_pdf, df_keys, on="Simplified", how="left", suffixes=("_PDF", "_Key"))

    merged["Llave M"] = merged["Tag"].apply(
        lambda x: x if pd.notna(x) and str(x).strip().upper().startswith("M") else ""
    )

    df_report = merged[["Cleaner", "Property Nickname", "Llave M"]].rename(columns={
        "Cleaner": "Encargado",
        "Property Nickname": "Dirección"
    })

    df_report["Encargado"] = df_report["Encargado"].fillna("Unassigned").astype(str).str.strip()
    df_report["Dirección"] = df_report["Dirección"].fillna("").astype(str).str.strip()
    df_report["Llave M"] = df_report["Llave M"].fillna("").astype(str).str.strip()
    df_report = df_report[df_report["Dirección"] != ""]

    grouped = (
        df_report.groupby(["Encargado", "Dirección"], as_index=False)
        .agg({"Llave M": lambda x: ", ".join(sorted({v.strip() for v in x if v.strip()}))})
        .sort_values(["Encargado", "Dirección"])
    )
    grouped = grouped[["Dirección", "Encargado", "Llave M"]]

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        grouped.to_excel(writer, sheet_name="Reporte", index=False)
        df_pdf.to_excel(writer, sheet_name="Extraido_PDF", index=False)
        merged.to_excel(writer, sheet_name="Merge_Debug", index=False)

        workbook = writer.book
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": "#305496", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter"
        })
        cell_fmt = workbook.add_format({"border": 1, "align": "left", "valign": "vcenter"})
        band_fmt = workbook.add_format({
            "border": 1, "bg_color": "#F2F2F2", "align": "left", "valign": "vcenter"
        })

        ws_report = writer.sheets["Reporte"]
        ws_pdf = writer.sheets["Extraido_PDF"]
        ws_merge = writer.sheets["Merge_Debug"]

        for col, name in enumerate(grouped.columns):
            ws_report.write(0, col, name, header_fmt)
        ws_report.set_column("A:A", 42, cell_fmt)
        ws_report.set_column("B:B", 30, cell_fmt)
        ws_report.set_column("C:C", 40, cell_fmt)
        for row in range(1, len(grouped) + 1):
            ws_report.set_row(row, None, band_fmt if row % 2 == 0 else cell_fmt)

        for col, name in enumerate(df_pdf.columns):
            ws_pdf.write(0, col, name, header_fmt)
        ws_pdf.set_column("A:A", 30)
        ws_pdf.set_column("B:B", 55)

        for col, name in enumerate(merged.columns):
            ws_merge.write(0, col, name, header_fmt)
        ws_merge.set_column(0, len(merged.columns) - 1, 22)

    output.seek(0)
    return grouped, output.read(), df_pdf, merged


# ----------------------------------------------------------
# 10. UI
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