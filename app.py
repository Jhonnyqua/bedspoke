import base64
import json
import re
from io import BytesIO

import gspread
import pandas as pd
import requests
import streamlit as st
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Reporte de Llaves M", layout="wide")

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

    return df_keys.reset_index(drop=True)


# ----------------------------------------------------------
# NORMALIZACIÓN SIMPLE
# ----------------------------------------------------------
def simplify_address_15chars(address: str) -> str:
    if not isinstance(address, str):
        return ""
    address = address.strip()
    m = re.search(r"\d", address)
    substr = address[m.start():m.start() + 15] if m else address[:15]
    return re.sub(r"[^0-9A-Za-z\s]", "", substr).lower().strip()


# ----------------------------------------------------------
# NORMALIZACIÓN FUERTE DE DIRECCIONES
# ----------------------------------------------------------
def normalize_address(addr: str) -> str:
    if not isinstance(addr, str):
        return ""

    addr = addr.lower().strip()

    replacements = {
        " street": " st",
        " road": " rd",
        " terrace": " tce",
        " avenue": " ave",
        " boulevard": " bvd",
        " drive": " dr",
        " place": " pl",
        " court": " ct",
        " lane": " ln",
        " quay": " qy",
        " saint ": " st ",
        " queensland": " qld",
    }

    for old, new in replacements.items():
        addr = addr.replace(old, new)

    addr = addr.replace(",", " ")
    addr = re.sub(r"[^a-z0-9\s/]", " ", addr)
    addr = re.sub(r"\s+", " ", addr).strip()
    return addr


def extract_address_parts(addr: str) -> dict:
    a = normalize_address(addr)

    unit = ""
    street_number = ""
    postcode = ""
    street_type = ""
    suburb = ""

    unit_match = re.match(r"^([a-z]?\d+[a-z]?/\d+[a-z]?)\b", a)
    if unit_match:
        unit = unit_match.group(1)

    num_match = re.search(r"\b(\d+)\b", a)
    if num_match:
        street_number = num_match.group(1)

    pc_match = re.search(r"\b([2-9]\d{3})\b", a)
    if pc_match:
        postcode = pc_match.group(1)

    stype_match = re.search(r"\b(st|rd|tce|ave|bvd|dr|pl|ct|ln|qy)\b", a)
    if stype_match:
        street_type = stype_match.group(1)

    state_words = {"qld", "nsw", "vic", "act", "wa", "sa", "tas", "nt"}
    type_words = {"st", "rd", "tce", "ave", "bvd", "dr", "pl", "ct", "ln", "qy"}

    tokens = a.split()
    text_tokens = [t for t in tokens if not re.search(r"\d", t) and t not in state_words and t not in type_words]

    if len(text_tokens) >= 2:
        suburb = " ".join(text_tokens[-2:])
    elif len(text_tokens) == 1:
        suburb = text_tokens[-1]

    return {
        "normalized": a,
        "unit": unit,
        "street_number": street_number,
        "postcode": postcode,
        "street_type": street_type,
        "suburb": suburb,
        "tokens": set(text_tokens),
    }


def score_address_match(pdf_addr: str, key_addr: str) -> float:
    p = extract_address_parts(pdf_addr)
    k = extract_address_parts(key_addr)

    score = 0.0

    if p["unit"] and p["unit"] == k["unit"]:
        score += 35
    if p["street_number"] and p["street_number"] == k["street_number"]:
        score += 20
    if p["street_type"] and p["street_type"] == k["street_type"]:
        score += 10
    if p["postcode"] and p["postcode"] == k["postcode"]:
        score += 10
    if p["suburb"] and p["suburb"] == k["suburb"]:
        score += 10

    token_overlap = len(p["tokens"] & k["tokens"])
    score += min(token_overlap * 5, 20)

    # bonus por simplificación corta exacta
    if simplify_address_15chars(pdf_addr) == simplify_address_15chars(key_addr):
        score += 10

    return round(score, 2)


def find_best_match(pdf_addr: str, df_keys: pd.DataFrame):
    candidates = []

    for _, row in df_keys.iterrows():
        key_addr = str(row["Property Address"])
        score = score_address_match(pdf_addr, key_addr)

        if score > 0:
            candidates.append({
                "Property Address": key_addr,
                "Tag": row.get("Tag", ""),
                "score": score,
                "row_data": row.to_dict(),
            })

    if not candidates:
        return None, []

    candidates = sorted(candidates, key=lambda x: x["score"], reverse=True)
    best = candidates[0]
    return best, candidates[:3]


# ----------------------------------------------------------
# PDF → IMÁGENES BASE64
# ----------------------------------------------------------
def pdf_to_base64_images(pdf_bytes: bytes) -> list:
    try:
        import fitz
    except ImportError:
        raise ImportError("Falta PyMuPDF. Agregá 'pymupdf' a requirements.txt")

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []

    for page in doc:
        mat = fitz.Matrix(1.5, 1.5)
        pix = page.get_pixmap(matrix=mat)
        images.append(base64.b64encode(pix.tobytes("png")).decode())

    doc.close()
    return images


# ----------------------------------------------------------
# GPT EXTRACTION
# ----------------------------------------------------------
PROMPT = """Esta es una página de un reporte "Housekeeping Daily Summary".

Extraé CADA propiedad visible en la página.

Para cada propiedad devolvé:
- address: dirección completa lo mejor posible
- cleaner: nombre del cleaner asignado
- page: número de página
- address_confidence: número entre 0 y 1
- cleaner_confidence: número entre 0 y 1
- notes: texto corto si hubo ambigüedad, o ""

REGLAS:
- Si el cleaner figura como "Unassigned", devolvé exactamente "Unassigned".
- No confundas huéspedes con cleaners.
- Ignorá encabezados, pies de página, fechas, reservas, ETA, comentarios, notas internas y nombres de huéspedes.
- Si una dirección está fragmentada en varias líneas, unila.
- Si una dirección quedó incompleta, devolvela igual, no inventes.
- Respondé SOLO con JSON válido en esta forma:
{
  "records": [
    {
      "address": "...",
      "cleaner": "...",
      "page": 1,
      "address_confidence": 0.93,
      "cleaner_confidence": 0.88,
      "notes": ""
    }
  ]
}
Si no hay propiedades, respondé:
{"records": []}
"""


def call_gpt_page(img_b64: str, page_num: int) -> list:
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        raise ValueError("Falta 'OPENAI_API_KEY' en los secrets de Streamlit.")

    response = requests.post(
        "https://api.openai.com/v1/chat/completions",
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        json={
            "model": "gpt-4o",
            "temperature": 0,
            "max_tokens": 1800,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{img_b64}",
                                "detail": "high",
                            },
                        },
                        {
                            "type": "text",
                            "text": PROMPT,
                        },
                    ],
                }
            ],
        },
        timeout=90,
    )

    if response.status_code != 200:
        err = response.json().get("error", {}).get("message", response.text)
        raise ValueError(f"Error API GPT página {page_num}: {response.status_code} — {err}")

    raw = response.json()["choices"][0]["message"]["content"].strip()
    raw = re.sub(r"```json|```", "", raw).strip()

    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        st.warning(f"⚠️ Página {page_num}: no pude parsear JSON. Respuesta: {raw[:300]}")
        return []

    if isinstance(parsed, dict) and "records" in parsed:
        records = parsed["records"]
    elif isinstance(parsed, list):
        records = parsed
    else:
        records = []

    cleaned = []
    for r in records:
        try:
            page_value = int(r.get("page", page_num) or page_num)
        except Exception:
            page_value = page_num

        try:
            address_conf = float(r.get("address_confidence", 0) or 0)
        except Exception:
            address_conf = 0.0

        try:
            cleaner_conf = float(r.get("cleaner_confidence", 0) or 0)
        except Exception:
            cleaner_conf = 0.0

        cleaned.append({
            "address": str(r.get("address", "")).strip(),
            "cleaner": str(r.get("cleaner", "Unassigned")).strip() or "Unassigned",
            "page": page_value,
            "address_confidence": address_conf,
            "cleaner_confidence": cleaner_conf,
            "notes": str(r.get("notes", "")).strip(),
        })

    return cleaned


def extract_all_pages(images: list, progress_bar) -> list:
    all_records = []
    n = len(images)

    for i, img in enumerate(images):
        page_num = i + 1
        pct = 0.08 + (0.55 * ((i + 1) / max(n, 1)))
        progress_bar.progress(min(pct, 0.63), text=f"Leyendo página {page_num} de {n} con IA...")
        records = call_gpt_page(img, page_num)
        all_records.extend(records)

    return all_records


# ----------------------------------------------------------
# GPT TIEBREAKER
# ----------------------------------------------------------
def resolve_match_with_gpt(pdf_address: str, candidates: list) -> dict:
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        return {"selected_address": "", "confidence": 0, "reason": "No API key"}

    prompt = f"""
Debes elegir el mejor match entre una dirección extraída de un PDF y hasta 3 candidatos del Key Register.

Dirección PDF:
{pdf_address}

Candidatos:
{json.dumps(candidates, ensure_ascii=False, indent=2)}

Responde SOLO JSON válido:
{{
  "selected_address": "dirección elegida o vacío",
  "confidence": 0.0,
  "reason": "motivo breve"
}}
"""

    response = requests.post(
        "https://api.openai.com/v1/chat/completions",
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        json={
            "model": "gpt-4o",
            "temperature": 0,
            "max_tokens": 500,
            "messages": [{"role": "user", "content": prompt}],
        },
        timeout=60,
    )

    if response.status_code != 200:
        return {"selected_address": "", "confidence": 0, "reason": response.text[:200]}

    raw = response.json()["choices"][0]["message"]["content"].strip()
    raw = re.sub(r"```json|```", "", raw).strip()

    try:
        return json.loads(raw)
    except Exception:
        return {"selected_address": "", "confidence": 0, "reason": "JSON inválido"}


# ----------------------------------------------------------
# BUILD MATCHES
# ----------------------------------------------------------
def build_matches(df_pdf: pd.DataFrame, df_keys: pd.DataFrame):
    matched_rows = []
    review_rows = []

    for _, pdf_row in df_pdf.iterrows():
        pdf_addr = str(pdf_row["Property Nickname"])
        best, top3 = find_best_match(pdf_addr, df_keys)

        final_address = ""
        final_tag = ""
        match_score = 0
        match_method = "none"
        review_reason = ""

        if best is not None:
            match_score = float(best["score"])

            if match_score >= 75:
                final_address = best["Property Address"]
                final_tag = best["Tag"]
                match_method = "rule_auto"

            elif match_score >= 55:
                gpt_decision = resolve_match_with_gpt(
                    pdf_addr,
                    [{"address": c["Property Address"], "score": c["score"]} for c in top3],
                )
                selected = str(gpt_decision.get("selected_address", "")).strip()

                if selected:
                    chosen = next((c for c in top3 if c["Property Address"] == selected), None)
                    if chosen:
                        final_address = chosen["Property Address"]
                        final_tag = chosen["Tag"]
                        match_score = chosen["score"]
                        match_method = "gpt_tiebreak"
                    else:
                        review_reason = f"GPT eligió dirección fuera del top3: {selected}"
                else:
                    review_reason = str(gpt_decision.get("reason", "No selection"))

            else:
                review_reason = "Low score"

        if final_address:
            matched_rows.append({
                **pdf_row.to_dict(),
                "Matched Address": final_address,
                "Tag": final_tag,
                "Match Score": match_score,
                "Match Method": match_method,
                "Llave M": final_tag if str(final_tag).strip().upper().startswith("M") else "",
            })
        else:
            top_candidates_str = " | ".join(
                [f"{c['Property Address']} ({c['score']})" for c in top3]
            ) if top3 else ""

            review_rows.append({
                **pdf_row.to_dict(),
                "Top Candidates": top_candidates_str,
                "Review Reason": review_reason or "No match found",
            })

    matched_df = pd.DataFrame(matched_rows)
    review_df = pd.DataFrame(review_rows)

    return matched_df, review_df


# ----------------------------------------------------------
# GENERAR EXCEL
# ----------------------------------------------------------
def create_report_excel(pdf_bytes: bytes, progress_bar):
    progress_bar.progress(0.03, text="Convirtiendo PDF a imágenes...")
    images = pdf_to_base64_images(pdf_bytes)

    records = extract_all_pages(images, progress_bar)
    progress_bar.progress(0.66, text="Preparando extracción...")

    if not records:
        raise ValueError("GPT no encontró propiedades en el PDF.")

    df_pdf = pd.DataFrame(records)
    df_pdf = df_pdf.rename(columns={"address": "Property Nickname", "cleaner": "Cleaner"})

    for col in ["Cleaner", "Property Nickname", "notes"]:
        if col in df_pdf.columns:
            df_pdf[col] = df_pdf[col].fillna("").astype(str).str.strip()

    df_pdf["Cleaner"] = df_pdf["Cleaner"].replace("", "Unassigned")
    df_pdf = df_pdf[df_pdf["Property Nickname"] != ""].reset_index(drop=True)

    progress_bar.progress(0.72, text="Cargando Key Register...")
    df_keys = load_key_register()

    progress_bar.progress(0.80, text="Haciendo match inteligente...")
    matched_df, review_df = build_matches(df_pdf, df_keys)

    if matched_df.empty:
        raise ValueError("No se logró hacer ningún match. Revisa la extracción o el Key Register.")

    df_report = matched_df[["Cleaner", "Property Nickname", "Llave M"]].rename(
        columns={"Cleaner": "Encargado", "Property Nickname": "Dirección"}
    )

    df_report = df_report.fillna("").astype(str)
    df_report["Encargado"] = df_report["Encargado"].str.strip().replace("", "Unassigned")
    df_report["Dirección"] = df_report["Dirección"].str.strip()
    df_report["Llave M"] = df_report["Llave M"].str.strip()
    df_report = df_report[df_report["Dirección"] != ""]

    grouped = (
        df_report.groupby(["Encargado", "Dirección"], as_index=False)
        .agg({"Llave M": lambda x: ", ".join(sorted({v.strip() for v in x if v.strip()}))})
        .sort_values(["Encargado", "Dirección"])
    )[["Dirección", "Encargado", "Llave M"]]

    progress_bar.progress(0.92, text="Generando Excel...")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        grouped.to_excel(writer, sheet_name="Reporte", index=False)
        df_pdf.to_excel(writer, sheet_name="Extraido_PDF", index=False)
        matched_df.to_excel(writer, sheet_name="Matched_Debug", index=False)

        if review_df.empty:
            review_df = pd.DataFrame(columns=[
                "Cleaner", "Property Nickname", "page", "address_confidence",
                "cleaner_confidence", "notes", "Top Candidates", "Review Reason"
            ])
        review_df.to_excel(writer, sheet_name="Review_Needed", index=False)

        wb = writer.book
        hdr = wb.add_format({
            "bold": True,
            "bg_color": "#305496",
            "font_color": "white",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })
        cel = wb.add_format({
            "border": 1,
            "align": "left",
            "valign": "vcenter",
        })
        alt = wb.add_format({
            "border": 1,
            "bg_color": "#F2F2F2",
            "align": "left",
            "valign": "vcenter",
        })

        # Reporte
        ws = writer.sheets["Reporte"]
        for col, name in enumerate(grouped.columns):
            ws.write(0, col, name, hdr)
        ws.set_column("A:A", 48, cel)
        ws.set_column("B:B", 30, cel)
        ws.set_column("C:C", 40, cel)
        for row in range(1, len(grouped) + 1):
            ws.set_row(row, None, alt if row % 2 == 0 else cel)

        # Extraido_PDF
        ws2 = writer.sheets["Extraido_PDF"]
        for col, name in enumerate(df_pdf.columns):
            ws2.write(0, col, name, hdr)
        ws2.set_column("A:A", 55)
        ws2.set_column("B:B", 28)
        ws2.set_column("C:F", 18)

        # Matched_Debug
        ws3 = writer.sheets["Matched_Debug"]
        for col, name in enumerate(matched_df.columns):
            ws3.write(0, col, name, hdr)
        ws3.set_column(0, len(matched_df.columns) - 1, 24)

        # Review_Needed
        ws4 = writer.sheets["Review_Needed"]
        for col, name in enumerate(review_df.columns):
            ws4.write(0, col, name, hdr)
        ws4.set_column(0, len(review_df.columns) - 1, 28)

    output.seek(0)
    progress_bar.progress(1.0, text="¡Listo!")
    return grouped, output.read(), df_pdf, matched_df, review_df


# ----------------------------------------------------------
# UI
# ----------------------------------------------------------
st.title("🗝️ Reporte de Llaves M desde PDF")
st.caption("Lectura visual con IA + matching inteligente + revisión de casos dudosos.")

pdf_file = st.file_uploader("📥 Sube tu PDF", type=["pdf"])

if pdf_file:
    if st.button("🚀 Generar Reporte", type="primary"):
        progress = st.progress(0, text="Iniciando...")
        try:
            pdf_bytes = pdf_file.read()

            grouped_df, excel_data, extracted_df, matched_df, review_df = create_report_excel(
                pdf_bytes,
                progress
            )

            st.success("✅ Reporte generado correctamente")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Trabajos detectados", len(extracted_df))
            c2.metric("Matches automáticos/finales", len(matched_df))
            c3.metric("Pendientes de revisión", len(review_df))
            c4.metric("Filas reporte final", len(grouped_df))

            st.subheader("Vista previa: extraído del PDF")
            st.dataframe(extracted_df, use_container_width=True)

            st.subheader("Reporte final")
            st.dataframe(grouped_df, use_container_width=True)

            if not review_df.empty:
                st.subheader("Casos para revisión")
                st.dataframe(review_df, use_container_width=True)

            st.download_button(
                label="⬇️ Descargar Excel",
                data=excel_data,
                file_name="reporte_llaves_m_ai_hibrido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"❌ Error: {e}")
            import traceback
            st.code(traceback.format_exc())
else:
    st.info("📄 Esperando que subas un PDF")