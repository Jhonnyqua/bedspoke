import streamlit as st
import gspread
import pandas as pd
import re
from io import BytesIO
from google.oauth2.service_account import Credentials
from pypdf import PdfReader

# ----------------------------------------------------------
# CONFIG
# ----------------------------------------------------------
st.set_page_config(page_title="Reporte de Llaves M", layout="wide")


# ----------------------------------------------------------
# 1. AUTENTICACIÓN GOOGLE SHEETS
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
# 3. EXTRAER TEXTO DE PDF
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

    line = line.replace("\u00a0", " ")
    line = line.replace("￾", " ")
    line = line.replace("\t", " ")
    line = re.sub(r"\s+", " ", line).strip()
    return line


def is_footer_or_header(line: str) -> bool:
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
        r"^Bedspoke Pty Ltd$",
        r"^Address:$",
        r"^Email:$",
        r"^Phone:$",
        r"^Printed by Resly$",
        r"^Page \d+ of \d+$",
        r"^\d{1,2} \w{3} \d{4} .* Printed by Resly Page \d+ of \d+$",
        r"^\d{1,2} \w{3} \d{4} \d{2}:\d{2}:\d{2}$",
        r"^\d{1,2} \w{3} \d{4} \d{2}:\d{2}:\d{2} Printed by Resly Page \d+ of \d+$",
    ]

    return any(re.search(p, line, flags=re.IGNORECASE) for p in patterns)


# ----------------------------------------------------------
# 5. DETECCIÓN DE DIRECCIONES
# ----------------------------------------------------------
def looks_like_address_start(line: str) -> bool:
    """
    Detecta inicio de propiedad.
    Ej:
    1/484 Upper Edward Street
    92 Cricket Street
    A71/41 Gotha Street
    """
    line = clean_line(line)

    if not line:
        return False

    # Excluir fechas
    if re.match(r"^\d{2}/\d{2}/\d{4}$", line):
        return False

    # Excluir rango de fechas de reservas
    if re.search(r"\d{2}/\d{2}/\d{4}\s+to\s+\d{2}/\d{2}/\d{4}", line):
        return False

    # Debe comenzar parecido a número/unidad
    starts_like_property = bool(
        re.match(r"^[A-Za-z]?\d+[A-Za-z]?(?:/\d+[A-Za-z]?)?", line)
    )

    street_keywords = [
        "Street", "St", "Road", "Rd", "Terrace", "Tce", "Lane", "Way",
        "Quay", "Avenue", "Ave", "Grove", "Court", "Ct", "Boulevard",
        "Bvd", "Drive", "Dr", "Place", "Pl", "Close", "Cl"
    ]

    has_street_word = any(k.lower() in line.lower() for k in street_keywords)

    return starts_like_property and has_street_word


def is_address_continuation(line: str) -> bool:
    """
    Segunda línea de dirección, por ejemplo:
    Spring Hill, QLD 4000
    Brisbane City, QLD 4000
    """
    line = clean_line(line)

    if not line:
        return False

    if looks_like_address_start(line):
        return False

    # No debe parecer línea operativa
    forbidden_starts = [
        "Scheduled",
        "[D]",
        "Due:",
        "Departing",
        "Arriving",
        "Arrival",
        "Back To Back",
        "Unassigned",
        "No Arrival Reservation",
        "Welcome",
        "Please",
        "IMPORTANT",
        "ETA:",
    ]
    if any(line.startswith(x) for x in forbidden_starts):
        return False

    # No debe ser guest stay date
    if re.search(r"\d{2}/\d{2}/\d{4}\s+to\s+\d{2}/\d{2}/\d{4}", line):
        return False

    # Línea con suburbio / estado / postcode
    if re.search(r"\bQLD\b|\bNSW\b|\bVIC\b|\bACT\b|\bWA\b|\bSA\b|\bTAS\b|\bNT\b|\b\d{4}\b", line, flags=re.IGNORECASE):
        return True

    # Texto tipo continuación
    return bool(re.fullmatch(r"[A-Za-z0-9,.\- ]+", line))


def build_address(lines, start_idx):
    address_parts = [clean_line(lines[start_idx])]
    i = start_idx + 1

    while i < len(lines):
        line = clean_line(lines[i])

        if not line:
            break

        if is_footer_or_header(line):
            break

        if looks_like_address_start(line):
            break

        if is_address_continuation(line):
            address_parts.append(line)
            i += 1

            # Si ya tiene postcode, normalmente está completa
            if re.search(r"\b\d{4}\b", " ".join(address_parts)):
                break
            continue

        break

    address = " ".join(address_parts)
    address = re.sub(r"\s+,", ",", address)
    address = re.sub(r"\s{2,}", " ", address).strip()

    return address, i


# ----------------------------------------------------------
# 6. DETECCIÓN DE CLEANER
# ----------------------------------------------------------
def is_guest_line(line: str) -> bool:
    line = clean_line(line)
    return bool(re.search(r"\([0-9]+A[0-9]+C\)", line))


def is_operation_line(line: str) -> bool:
    line = clean_line(line)

    bad_exact = {
        "Scheduled",
        "[D] Depart Clean",
        "[D] Depart",
        "Clean",
        "Departing",
        "Arriving",
        "Arrival",
        "Back To Back",
        "Due:",
        "Unassigned",
        "Property Housekeeping",
        "Task",
        "Reservation Assigned To",
    }
    if line in bad_exact:
        return True

    bad_patterns = [
        r"^Due:",
        r"^\d{2}/\d{2}/\d{4}$",
        r"^\d{2}/\d{2}/\d{4}\s+to\s+\d{2}/\d{2}/\d{4}",
        r"^\d+N$",
        r"^ETA:",
        r"^Welcome and enjoy your stay",
        r"^Please ",
        r"^Departure Time:",
        r"^There was a long stay",
        r"^No Arrival Reservation",
        r"^IMPORTANT",
        r"^return the",
        r"^pls return",
        r"^DEEP CLEAN",
        r"^Printed by Resly",
        r"^Housekeeping Tasks",
    ]

    return any(re.search(p, line, flags=re.IGNORECASE) for p in bad_patterns)


def is_cleaner_name_line(line: str) -> bool:
    line = clean_line(line)

    if not line:
        return False

    if is_footer_or_header(line):
        return False

    if is_operation_line(line):
        return False

    if is_guest_line(line):
        return False

    # Excluir fechas o números
    if re.search(r"\d", line):
        return False

    # Excluir formatos de huésped
    if "," in line or "(" in line or ")" in line:
        return False

    # Excluir frases largas operativas disfrazadas
    banned_fragments = [
        "reservation",
        "arrival",
        "depart",
        "check",
        "guest",
        "housekeeping",
        "printed",
        "resly",
        "welcome",
        "bedspoke",
    ]
    low = line.lower()
    if any(x in low for x in banned_fragments):
        return False

    # Solo letras/espacios/apóstrofes/guiones
    if not re.fullmatch(r"[A-Za-zÀ-ÿ'’\- ]+", line):
        return False

    words = line.split()
    return 1 <= len(words) <= 6


def extract_cleaner_from_block(block_lines):
    """
    Busca el cleaner desde el final del bloque hacia arriba.
    Permite nombres partidos en varias líneas.
    """
    cleaned = [clean_line(x) for x in block_lines if clean_line(x)]

    # Limpiar basura del final
    while cleaned and (is_footer_or_header(cleaned[-1]) or is_operation_line(cleaned[-1])):
        cleaned.pop()

    if not cleaned:
        return "Unassigned"

    name_parts = []
    i = len(cleaned) - 1

    while i >= 0:
        line = cleaned[i]

        if is_cleaner_name_line(line):
            name_parts.insert(0, line)
            i -= 1
        else:
            if name_parts:
                break
            i -= 1

    cleaner = " ".join(name_parts)
    cleaner = re.sub(r"\s{2,}", " ", cleaner).strip()

    if not cleaner:
        return "Unassigned"

    bad_cleaners = {
        "Back To Back",
        "Departing",
        "Arriving",
        "Scheduled",
        "Due",
        "Due:",
        "Property Housekeeping",
        "Task",
        "Reservation Assigned To",
    }

    if cleaner in bad_cleaners:
        return "Unassigned"

    return cleaner


# ----------------------------------------------------------
# 7. PARSEAR PDF HOUSEKEEPING
# ----------------------------------------------------------
def parse_housekeeping_pdf(pdf_file) -> pd.DataFrame:
    text = extract_text_from_pdf(pdf_file)
    raw_lines = text.splitlines()

    lines = []
    for line in raw_lines:
        line = clean_line(line)
        if not line:
            continue
        if is_footer_or_header(line):
            continue
        lines.append(line)

    rows = []
    i = 0

    while i < len(lines):
        line = lines[i]

        if looks_like_address_start(line):
            address, next_i = build_address(lines, i)

            block = []
            j = next_i

            while j < len(lines):
                if looks_like_address_start(lines[j]):
                    break
                block.append(lines[j])
                j += 1

            cleaner = extract_cleaner_from_block(block)

            rows.append({
                "Cleaner": cleaner,
                "Property Nickname": address
            })

            i = j
        else:
            i += 1

    df = pd.DataFrame(rows)

    if df.empty:
        return df

    df = df.drop_duplicates()

    # Limpieza extra
    df["Cleaner"] = df["Cleaner"].fillna("Unassigned").astype(str).str.strip()
    df["Property Nickname"] = df["Property Nickname"].fillna("").astype(str).str.strip()

    # Quitar filas vacías
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
# 9. GENERAR REPORTE EXCEL
# ----------------------------------------------------------
def create_grouped_excel_from_pdf(pdf_file):
    df_pdf = parse_housekeeping_pdf(pdf_file)

    if df_pdf.empty:
        raise ValueError(
            "No pude extraer propiedades del PDF. Revisa si el formato cambió."
        )

    df_keys = load_key_register()

    # Simplificar direcciones
    df_pdf["Simplified"] = df_pdf["Property Nickname"].apply(simplify_address_15chars)
    df_keys["Simplified"] = df_keys["Property Address"].apply(simplify_address_15chars)

    # Merge
    merged = pd.merge(
        df_pdf,
        df_keys,
        on="Simplified",
        how="left",
        suffixes=("_PDF", "_Key")
    )

    # Solo llaves M
    merged["Llave M"] = merged["Tag"].apply(
        lambda x: x if pd.notna(x) and str(x).strip().upper().startswith("M") else ""
    )

    # Reporte base
    df_report = merged[["Cleaner", "Property Nickname", "Llave M"]].rename(columns={
        "Cleaner": "Encargado",
        "Property Nickname": "Dirección"
    })

    df_report["Encargado"] = df_report["Encargado"].fillna("Unassigned").astype(str).str.strip()
    df_report["Dirección"] = df_report["Dirección"].fillna("").astype(str).str.strip()
    df_report["Llave M"] = df_report["Llave M"].fillna("").astype(str).str.strip()

    df_report = df_report[df_report["Dirección"] != ""]

    # Agrupar por encargado + dirección
    grouped = (
        df_report.groupby(["Encargado", "Dirección"], as_index=False)
        .agg({
            "Llave M": lambda x: ", ".join(sorted({v.strip() for v in x if v.strip()}))
        })
        .sort_values(["Encargado", "Dirección"])
    )

    grouped = grouped[["Dirección", "Encargado", "Llave M"]]

    # Crear Excel
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        grouped.to_excel(writer, sheet_name="Reporte", index=False)
        df_pdf.to_excel(writer, sheet_name="Extraido_PDF", index=False)
        merged.to_excel(writer, sheet_name="Merge_Debug", index=False)

        workbook = writer.book
        ws_report = writer.sheets["Reporte"]
        ws_pdf = writer.sheets["Extraido_PDF"]
        ws_merge = writer.sheets["Merge_Debug"]

        header_fmt = workbook.add_format({
            "bold": True,
            "bg_color": "#305496",
            "font_color": "white",
            "border": 1,
            "align": "center",
            "valign": "vcenter"
        })

        cell_fmt = workbook.add_format({
            "border": 1,
            "align": "left",
            "valign": "vcenter"
        })

        band_fmt = workbook.add_format({
            "border": 1,
            "bg_color": "#F2F2F2",
            "align": "left",
            "valign": "vcenter"
        })

        # Encabezados Reporte
        for col, name in enumerate(grouped.columns):
            ws_report.write(0, col, name, header_fmt)

        ws_report.set_column("A:A", 42, cell_fmt)
        ws_report.set_column("B:B", 30, cell_fmt)
        ws_report.set_column("C:C", 40, cell_fmt)

        for row in range(1, len(grouped) + 1):
            ws_report.set_row(row, None, band_fmt if row % 2 == 0 else cell_fmt)

        # Extraido_PDF
        for col, name in enumerate(df_pdf.columns):
            ws_pdf.write(0, col, name, header_fmt)
        ws_pdf.set_column("A:A", 30)
        ws_pdf.set_column("B:B", 50)

        # Merge_Debug
        for col, name in enumerate(merged.columns):
            ws_merge.write(0, col, name, header_fmt)
        ws_merge.set_column(0, len(merged.columns) - 1, 22)

    output.seek(0)
    return grouped, output.read(), df_pdf, merged


# ----------------------------------------------------------
# 10. INTERFAZ STREAMLIT
# ----------------------------------------------------------
st.title("🗝️ Reporte de Llaves M desde PDF")
st.write("Sube el PDF de Housekeeping Daily Summary para cruzarlo con tu Key Register.")

pdf_file = st.file_uploader("📥 Sube tu PDF", type=["pdf"])

if pdf_file:
    try:
        grouped_df, excel_data, extracted_df, merged_df = create_grouped_excel_from_pdf(pdf_file)

        st.success("✅ Reporte generado correctamente")

        col1, col2, col3 = st.columns(3)
        col1.metric("Trabajos detectados en PDF", len(extracted_df))
        col2.metric("Filas reporte final", len(grouped_df))
        col3.metric("Cruces totales", len(merged_df))

        st.subheader("Vista previa: propiedades y cleaners detectados desde el PDF")
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
else:
    st.info("📄 Esperando que subas un PDF")