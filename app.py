import re
from io import BytesIO

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextBox, LAParams

st.set_page_config(page_title="Reporte de Llaves M - DEBUG", layout="wide")


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
# DEBUG: show ALL text boxes with coordinates
# ----------------------------------------------------------
def extract_all_boxes(pdf_file) -> list:
    """Returns list of (page, x0, y0, x1, y1, text) for ALL text boxes."""
    laparams = LAParams(line_margin=0.5, char_margin=3.0, word_margin=0.1)
    boxes = []
    for page_num, page_layout in enumerate(extract_pages(pdf_file, laparams=laparams)):
        page_width = page_layout.width
        page_height = page_layout.height
        for element in page_layout:
            if isinstance(element, LTTextBox):
                text = element.get_text().strip()
                if text:
                    boxes.append({
                        "page": page_num + 1,
                        "x0": round(element.x0, 1),
                        "y0": round(element.y0, 1),
                        "x1": round(element.x1, 1),
                        "y1": round(element.y1, 1),
                        "x0_pct": round(element.x0 / page_width * 100, 1),
                        "page_width": round(page_width, 1),
                        "text": text.replace("\n", " | "),
                    })
    return boxes


st.title("🔍 DEBUG - Coordenadas del PDF")
st.write("Sube el PDF para ver exactamente qué texto extrae pdfminer y en qué posición X.")

pdf_file = st.file_uploader("📥 Sube tu PDF", type=["pdf"])

if pdf_file:
    try:
        boxes = extract_all_boxes(pdf_file)
        df = pd.DataFrame(boxes)

        st.success(f"✅ Extraídos {len(df)} bloques de texto")

        # Show page width info
        if not df.empty:
            st.info(f"Ancho de página: {df['page_width'].iloc[0]} pts. "
                    f"El 38% sería x < {df['page_width'].iloc[0] * 0.38:.0f} pts")

        st.subheader("Todos los bloques (página 1)")
        page1 = df[df["page"] == 1].sort_values("y1", ascending=False)
        st.dataframe(page1[["x0", "x0_pct", "y1", "text"]], use_container_width=True)

        st.subheader("Solo columna izquierda (x0_pct < 38)")
        left = df[df["x0_pct"] < 38].sort_values(["page", "y1"], ascending=[True, False])
        st.dataframe(left[["page", "x0", "x0_pct", "y1", "text"]], use_container_width=True)

        # Download full debug
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.sort_values(["page", "y1"], ascending=[True, False]).to_excel(
                writer, sheet_name="All_Boxes", index=False
            )
            left.to_excel(writer, sheet_name="Left_Column", index=False)
        output.seek(0)

        st.download_button(
            label="⬇️ Descargar debug Excel",
            data=output.read(),
            file_name="debug_pdf_boxes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
        import traceback
        st.code(traceback.format_exc())
else:
    st.info("📄 Esperando que subas un PDF")