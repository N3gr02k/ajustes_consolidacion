import streamlit as st
import fitz
import pandas as pd
import re
import os
from io import BytesIO

from pdf2image import convert_from_bytes
import pytesseract


# =====================================================
# CONFIGURACIÓN
# =====================================================

POPPLER_PATH = r"D:\poppler\poppler-25.12.0\Library\bin"

TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
TESSDATA_PATH = r"C:\Program Files\Tesseract-OCR\tessdata"

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
os.environ["TESSDATA_PREFIX"] = TESSDATA_PATH

OCR_LANG = "spa"   # usar español


DATE = re.compile(r"\b\d{2}-\d{2}\b")
AMOUNT = re.compile(r"\d[\d,]*\.\d{2}")


# =====================================================
# UTILIDADES
# =====================================================

def parse_amount(v):

    v = v.replace(",", "")

    if v.endswith("-"):
        v = "-" + v[:-1]

    try:
        return float(v)
    except:
        return None


# =====================================================
# MÉTODO 1 — TEXTO DIRECTO
# =====================================================

def extract_text_method(pdf_bytes):

    st.write("Método 1: extracción directa")

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    rows = []

    for page in doc:

        text = page.get_text()

        for line in text.split("\n"):

            if DATE.search(line) and AMOUNT.search(line):

                amounts = AMOUNT.findall(line)

                rows.append({
                    "Fecha": DATE.search(line).group(),
                    "Concepto": line,
                    "Monto": parse_amount(amounts[-1])
                })

    return pd.DataFrame(rows)


# =====================================================
# MÉTODO 2 — RECONSTRUCCIÓN POR LAYOUT
# =====================================================

def extract_layout_method(pdf_bytes):

    st.write("Método 2: reconstrucción layout")

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    rows = []

    for page in doc:

        words = page.get_text("words")

        lines = {}

        for w in words:

            x0,y0,x1,y1,text,_,_,_ = w

            key = round(y0)

            if key not in lines:
                lines[key] = []

            lines[key].append((x0,text))

        for y in sorted(lines):

            ordered = sorted(lines[y], key=lambda x: x[0])

            line = " ".join(w[1] for w in ordered)

            if DATE.search(line) and AMOUNT.search(line):

                amounts = AMOUNT.findall(line)

                rows.append({
                    "Fecha": DATE.search(line).group(),
                    "Concepto": line,
                    "Monto": parse_amount(amounts[-1])
                })

    return pd.DataFrame(rows)


# =====================================================
# MÉTODO 3 — OCR
# =====================================================

def extract_ocr_method(pdf_bytes):

    st.write("Método 3: OCR")

    images = convert_from_bytes(
        pdf_bytes,
        dpi=300,
        poppler_path=POPPLER_PATH
    )

    rows = []
    buffer = ""

    for i,img in enumerate(images):

        st.write(f"OCR página {i+1}")

        text = pytesseract.image_to_string(
            img,
            lang=OCR_LANG,
            config="--psm 6"
        )

        for line in text.split("\n"):

            line = line.strip()

            if DATE.search(line):

                buffer = line

            else:

                buffer += " " + line

            if AMOUNT.search(buffer):

                amounts = AMOUNT.findall(buffer)

                rows.append({
                    "Fecha": DATE.search(buffer).group(),
                    "Concepto": buffer,
                    "Monto": parse_amount(amounts[-1])
                })

                buffer = ""

    return pd.DataFrame(rows)


# =====================================================
# MOTOR HÍBRIDO
# =====================================================

def extract_movements(pdf_file):

    pdf_bytes = pdf_file.read()

    df = extract_text_method(pdf_bytes)

    if len(df) > 20:

        st.success("Extracción directa exitosa")

        return df


    df = extract_layout_method(pdf_bytes)

    if len(df) > 20:

        st.success("Extracción por layout exitosa")

        return df


    df = extract_ocr_method(pdf_bytes)

    if len(df) > 0:

        st.success("Extracción OCR exitosa")

        return df


    st.error("No se pudieron detectar movimientos")

    return df


# =====================================================
# EXPORTAR
# =====================================================

def export_excel(df):

    buffer = BytesIO()

    df.to_excel(buffer, index=False)

    buffer.seek(0)

    return buffer


# =====================================================
# STREAMLIT
# =====================================================

def main():

    st.set_page_config(layout="wide")

    st.title("STRATUM v57.1 — Extractor Bancario Universal")

    pdf = st.file_uploader("Subir estado bancario PDF", type="pdf")

    if pdf:

        if st.button("Extraer movimientos"):

            df = extract_movements(pdf)

            st.write("Movimientos detectados:", len(df))

            if len(df) > 0:

                st.dataframe(df)

                excel = export_excel(df)

                st.download_button(
                    "Descargar Excel",
                    excel,
                    "movimientos_extraidos.xlsx"
                )


if __name__ == "__main__":
    main()