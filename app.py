import re
from io import BytesIO

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextBox, LAParams

st.set_page_config(page_title="DEBUG v2", layout="wide")

STREET_KEYWORDS = [
    "Street", "St", "Road", "Rd", "Terrace", "Tce", "Lane", "Way",
    "Quay", "Avenue", "Ave", "Grove", "Court", "Ct", "Boulevard",
    "Bvd", "Drive", "Dr", "Place", "Pl", "Close", "Cl"
]

def extract_boxes(pdf_file):
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
                text_clean = re.sub(r"\s*\n\s*", " | ", text).strip()
                boxes.append({
                    "page": page_num + 1,
                    "x0": round(element.x0, 1),
                    "y1": round(element.y1, 1),
                    "x0_pct": round(element.x0 / page_width * 100, 1),
                    "page_width": round(page_width, 1),
                    "page_height": round(page_height, 1),
                    "text": text_clean,
                })
    return boxes

def clean_address_text(text):
    text = re.sub(r"\s+", " ", text).strip()
    m = re.search(r"\b([2-9]\d{3})\b", text)
    if m:
        text = text[:m.start() + 4].strip()
    return text

def is_valid_address(text):
    if not re.match(r"^[A-Za-z]?\d+[A-Za-z]?(?:/\d+[A-Za-z]?)?", text):
        return False
    return any(k.lower() in text.lower() for k in STREET_KEYWORDS)

st.title("🔍 DEBUG v2 - Detección de cajas")

pdf_file = st.file_uploader("📥 Sube tu PDF", type=["pdf"])

if pdf_file:
    boxes = extract_boxes(pdf_file)
    df = pd.DataFrame(boxes)

    st.info(f"Total cajas: {len(df)} | Página ancho: {df['page_width'].iloc[0]} pts")

    # Address boxes
    addr_df = df[df["x0_pct"] < 15].copy()
    addr_df["cleaned"] = addr_df["text"].apply(clean_address_text)
    addr_df["is_address"] = addr_df["cleaned"].apply(is_valid_address)
    valid_addr = addr_df[addr_df["is_address"]]

    st.subheader(f"📍 Cajas columna IZQUIERDA (x0_pct < 15%) — {len(addr_df)} total, {len(valid_addr)} direcciones válidas")
    st.dataframe(addr_df[["page", "x0_pct", "y1", "is_address", "cleaned"]].sort_values(["page","y1"], ascending=[True,False]), use_container_width=True)

    # Right column boxes
    right_df = df[df["x0_pct"] > 70].copy()
    st.subheader(f"👤 Cajas columna DERECHA (x0_pct > 70%) — {len(right_df)} cajas")
    st.dataframe(right_df[["page", "x0_pct", "y1", "text"]].sort_values(["page","y1"], ascending=[True,False]), use_container_width=True)

    # Middle column (x0_pct 15-70) - what's there?
    mid_df = df[(df["x0_pct"] >= 15) & (df["x0_pct"] <= 70)].copy()
    st.subheader(f"⚙️ Cajas MEDIO (x0_pct 15-70%) — {len(mid_df)} cajas")
    st.dataframe(mid_df[["page", "x0_pct", "y1", "text"]].sort_values(["page","y1"], ascending=[True,False]), use_container_width=True)

    # Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.sort_values(["page","y1"], ascending=[True,False]).to_excel(writer, sheet_name="All", index=False)
        addr_df.sort_values(["page","y1"], ascending=[True,False]).to_excel(writer, sheet_name="Left_col", index=False)
        right_df.sort_values(["page","y1"], ascending=[True,False]).to_excel(writer, sheet_name="Right_col", index=False)
        mid_df.sort_values(["page","y1"], ascending=[True,False]).to_excel(writer, sheet_name="Mid_col", index=False)
    output.seek(0)

    st.download_button("⬇️ Descargar debug Excel", data=output.read(),
                       file_name="debug_v2.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")