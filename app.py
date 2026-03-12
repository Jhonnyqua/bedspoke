import streamlit as st
import gspread
import pandas as pd
import re
from io import BytesIO
from google.oauth2.service_account import Credentials
from pypdf import PdfReader

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
# 3. Leer texto del PDF
# ----------------------------------------------------------
def extract_text_from_pdf(pdf_file) -> str:
    reader = PdfReader(pdf_file)
    pages_text = []
    for page in reader.pages:
        txt = page.extract_text()
        if txt:
            pages_text.append(txt)
    return "\n".join(pages_text)

# ----------------------------------------------------------
# 4. Utilidades para detectar direcciones y cleaners
# ----------------------------------------------------------
def clean_line(line: str) -> str:
    line = line.replace("\u00a0", " ")
    line = line.replace("￾", " ")
    return re.sub(r"\s+", " ", line).strip()

def is_noise_line(line: str) -> bool:
    if not line:
        return True

    noise_patterns = [
        r"^Housekeeping Daily Summary$",
        r"^\d{2}/\d{2}/\d{4}$",
        r"^Housekeeping Tasks",
        r"^Property Housekeeping$",
        r"^Task$",
        r"^Reservation Assigned To$",
        r"^Bedspoke$",
        r"^Bedspoke Pty Ltd",
        r"^Address:",
        r"^Email:",
        r"^Phone:",
        r"^Printed by Resly",
        r"^Page \d+ of \d+$",
    ]

    for p in noise_patterns:
        if re.search(p, line, flags=re.IGNORECASE):
            return True

    return False

def looks_like_address_start(lines, i) -> bool:
    """
    Detecta inicio de dirección mirando línea actual + siguientes.
    """
    line = clean_line(lines[i])

    if not line:
        return False

    # Debe empezar con algo tipo 34/15, 1208/35, A71/41, 2A/6, etc.
    if not re.match(r"^[A-Za-z]?\d+[A-Za-z]?(?:[-/][A-Za-z0-9]+)*", line):
        return False

    # Excluir fechas
    if re.match(r"^\d{2}/\d{2}/\d{4}$", line):
        return False

    # Mirar una ventana de hasta 3 líneas para confirmar que parece dirección
    window = " ".join(clean_line(lines[j]) for j in range(i, min(i + 3, len(lines))))
    street_words = [
        "Street", "St", "Avenue", "Ave", "Road", "Rd", "Court", "Ct",
        "Terrace", "Tce", "Way", "Quay", "Lane", "Grove", "Hope",
        "Norfolk", "Margaret", "Brunswick", "Cordelia", "Merivale",
        "Alfred", "Manning", "Hercules", "Masters", "Connor", "Gotha",
        "Vulture", "Exford", "Britannia", "Serisier", "South Sea Islander"
    ]

    has_street_word = any(word.lower() in window.lower() for word in street_words)
    has_postcode = bool(re.search(r"\b\d{4}\b", window))

    return has_street_word or has_postcode

def build_address(lines, start_idx):
    """
    Junta varias líneas hasta completar la dirección.
    """
    address_parts = []
    i = start_idx

    while i < len(lines):
        line = clean_line(lines[i])

        if not line or is_noise_line(line):
            break

        # Si ya empezó y aparece otra dirección, detener
        if i > start_idx and looks_like_address_start(lines, i):
            break

        address_parts.append(line)

        # Si ya apareció postcode, ya debe estar completa
        joined = " ".join(address_parts)
        if re.search(r"\b\d{4}\b", joined):
            i += 1
            break

        # Evitar capturar demasiado
        if len(address_parts) >= 4:
            i += 1
            break

        i += 1

    address = " ".join(address_parts)
    address = re.sub(r"\s+,", ",", address)
    address = re.sub(r"\s{2,}", " ", address).strip()

    return address, i

def is_name_like(line: str) -> bool:
    """
    Detecta líneas que parecen parte del nombre del cleaner.
    """
    line = clean_line(line)

    if not line:
        return False

    # Excluir líneas con números o demasiada puntuación
    if re.search(r"\d", line):
        return False

    excluded = {
        "In House", "Done", "Urgent", "Clean", "Ready", "Scheduled",
        "[D] Depart Clean", "Due:", "Back To Back", "Departing Today",
        "Arriving Today", "Arrival", "Departure", "Welcome and enjoy your stay.",
        "Please", "ETA:", "Late check out", "Late checkout"
    }

    if line in excluded:
        return False

    # Excluir reservas tipo "SMITH, John (2A0C)"
    if "," in line or "(" in line or ")" in line:
        return False

    # Debe ser texto estilo nombre
    if not re.fullmatch(r"[A-Za-zÀ-ÿ'’\- ]+", line):
        return False

    words = line.split()
    if len(words) == 0 or len(words) > 5:
        return False

    return True

def extract_cleaner_from_block(block_lines):
    """
    Toma las últimas líneas del bloque que parezcan nombre del cleaner.
    """
    if not block_lines:
        return ""

    cleaner_parts = []
    i = len(block_lines) - 1

    while i >= 0:
        line = clean_line(block_lines[i])

        if is_name_like(line):
            cleaner_parts.insert(0, line)
            i -= 1
        else:
            if cleaner_parts:
                break
            i -= 1

    cleaner = " ".join(cleaner_parts)
    cleaner = re.sub(r"\s{2,}", " ", cleaner).strip()
    return cleaner

# ----------------------------------------------------------
# 5. Parsear PDF de Housekeeping
# ----------------------------------------------------------
def parse_housekeeping_pdf(pdf_file) -> pd.DataFrame:
    text = extract_text_from_pdf(pdf_file)
    raw_lines = text.splitlines()
    lines = [clean_line(x) for x in raw_lines if clean_line(x)]

    rows = []
    i = 0

    while i < len(lines):
        line = lines[i]

        if is_noise_line(line):
            i += 1
            continue

        if looks_like_address_start(lines, i):
            address, next_i = build_address(lines, i)

            # Capturar bloque hasta la siguiente dirección
            block = []
            j = next_i
            while j < len(lines):
                if looks_like_address_start(lines, j):
                    break
                if not is_noise_line(lines[j]):
                    block.append(lines[j])
                j += 1

            cleaner = extract_cleaner_from_block(block)

            if address and cleaner:
                rows.append({
                    "Cleaner": cleaner,
                    "Property Nickname": address
                })

            i = j
        else:
            i += 1

    df_pdf = pd.DataFrame(rows).drop_duplicates()

    return df_pdf

# ----------------------------------------------------------
# 6. Generar reporte agrupado desde PDF
# ----------------------------------------------------------
def create_grouped_excel_from_pdf(pdf_file):
    # 6.1 Leer Google Sheet con llaves
    client = authorize_gspread()
    sheet_id = st.secrets["gcp_service_account"]["spreadsheet_id"]
    sheet = client.open_by_key(sheet_id).worksheet("Key Register")
    data = sheet.get_all_values()

    df_keys = pd.DataFrame(data[2:], columns=data[1]).drop(columns="", errors="ignore")

    # Filtrar solo filas sin observación
    if "Observation" in df_keys.columns:
        df_keys = df_keys[df_keys["Observation"].fillna("").str.strip() == ""]

    # Validación mínima
    required_key_cols = ["Property Address", "Tag"]
    for col in required_key_cols:
        if col not in df_keys.columns:
            raise ValueError(f"No encontré la columna '{col}' en la hoja Key Register.")

    # 6.2 Leer PDF de housekeeping
    df_pdf = parse_housekeeping_pdf(pdf_file)

    if df_pdf.empty:
        raise ValueError(
            "No pude extraer propiedades y cleaners del PDF. "
            "Puede que el formato haya cambiado y haya que ajustar el parser."
        )

    # 6.3 Simplificar direcciones
    df_pdf["Simplified"] = df_pdf["Property Nickname"].apply(simplify_address_15chars)
    df_keys["Simplified"] = df_keys["Property Address"].apply(simplify_address_15chars)

    # 6.4 Merge
    merged = pd.merge(
        df_pdf,
        df_keys,
        on="Simplified",
        how="left",
        suffixes=("_PDF", "_Key"),
    )

    # 6.5 Extraer solo llaves M
    merged["Llave M"] = merged["Tag"].apply(
        lambda x: x if pd.notna(x) and str(x).startswith("M") else ""
    )

    # 6.6 Selección de columnas
    df = merged[["Cleaner", "Property Nickname", "Llave M"]].rename(columns={
        "Cleaner": "Encargado",
        "Property Nickname": "Dirección"
    })

    df = df[df["Encargado"].fillna("").str.strip() != ""]

    # 6.7 Agrupar
    grouped = (
        df.groupby(["Encargado", "Dirección"], as_index=False)
          .agg({
              "Llave M": lambda x: ", ".join(sorted({str(v).strip() for v in x if str(v).strip()}))
          })
          .sort_values(["Encargado", "Dirección"])
    )

    grouped = grouped[["Dirección", "Encargado", "Llave M"]]

    # 6.8 Crear Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        grouped.to_excel(writer, sheet_name="Reporte", index=False)
        wb = writer.book
        ws = writer.sheets["Reporte"]

        header_fmt = wb.add_format({
            "bold": True,
            "bg_color": "#305496",
            "font_color": "white",
            "border": 1,
            "align": "center",
            "valign": "vcenter"
        })

        cell_fmt = wb.add_format({
            "border": 1,
            "align": "left",
            "valign": "vcenter"
        })

        band_fmt = wb.add_format({
            "border": 1,
            "bg_color": "#F2F2F2",
            "align": "left",
            "valign": "vcenter"
        })

        for col, name in enumerate(grouped.columns):
            ws.write(0, col, name, header_fmt)

        ws.set_column("A:A", 38, cell_fmt)  # Dirección
        ws.set_column("B:B", 28, cell_fmt)  # Encargado
        ws.set_column("C:C", 40, cell_fmt)  # Llave M

        for row in range(1, len(grouped) + 1):
            ws.set_row(row, None, band_fmt if row % 2 == 0 else cell_fmt)

    output.seek(0)
    return grouped, output.read(), df_pdf

# ----------------------------------------------------------
# 7. Interfaz Streamlit
# ----------------------------------------------------------
st.set_page_config(page_title="Reporte de Llaves M", layout="wide")

st.title("🗝️ Reporte de Llaves M desde PDF")
st.caption("Sube el PDF de Housekeeping Daily Summary y generaré el reporte cruzado con Key Register.")

pdf_file = st.file_uploader("📥 Sube tu PDF de Housekeeping", type=["pdf"])

if pdf_file:
    try:
        df_grp, excel_data, df_extracted = create_grouped_excel_from_pdf(pdf_file)

        st.success("✅ Reporte generado correctamente")

        with st.expander("Ver propiedades y cleaners detectados desde el PDF"):
            st.dataframe(df_extracted, use_container_width=True)

        st.subheader("Reporte final")
        st.dataframe(df_grp, use_container_width=True)

        st.download_button(
            "⬇️ Descargar Reporte (Excel)",
            data=excel_data,
            file_name="reporte_llaves_m_desde_pdf.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Ocurrió un error: {e}")
else:
    st.info("📄 Esperando que subas un archivo PDF")