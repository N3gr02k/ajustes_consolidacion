"""
Módulo de extracción de PDFs bancarios
Soporta múltiples bancos mediante estrategia híbrida
"""
import re
import streamlit as st
from typing import Optional, List, Dict
import pandas as pd
import fitz
from io import BytesIO

try:
    from pdf2image import convert_from_bytes
    import pytesseract
    OCR_DISPONIBLE = True
except ImportError:
    OCR_DISPONIBLE = False

from .config import (
    POPPLER_PATH,
    TESSERACT_PATH,
    TESSDATA_PATH,
    OCR_LANG,
    OCR_DPI,
    DATE_PATTERNS,
    AMOUNT_PATTERNS,
    UMBRAL_EXTRACCION_EXITO
)

# Configurar Tesseract
import os
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
os.environ["TESSDATA_PREFIX"] = TESSDATA_PATH


class ExtractorPDF:
    """Extractor de movimientos bancarios desde PDFs"""

    def __init__(self):
        self.date_patterns = [re.compile(p) for p in DATE_PATTERNS.values()]
        self.amount_patterns = [re.compile(p) for p in AMOUNT_PATTERNS.values()]

    def parse_amount(self, v: str) -> Optional[float]:
        """Convierte string de monto a float"""
        if not v:
            return None

        # Limpiar el valor
        v = v.strip()
        v = v.replace(",", "")

        # Manejar signos negativos
        if v.endswith("-"):
            v = "-" + v[:-1]

        # Manejar paréntesis para negativos
        if v.startswith("(") and v.endswith(")"):
            v = "-" + v[1:-1]

        try:
            return float(v)
        except (ValueError, TypeError):
            return None

    def find_date(self, text: str) -> Optional[str]:
        """Encuentra una fecha en el texto"""
        for pattern in self.date_patterns:
            match = pattern.search(text)
            if match:
                return match.group()
        return None

    def find_amount(self, text: str) -> Optional[str]:
        """Encuentra un monto en el texto"""
        for pattern in self.amount_patterns:
            matches = pattern.findall(text)
            if matches:
                return matches[-1]  # Retornar el último monto (通常是 el monto total)
        return None

    def extract_text_direct(self, pdf_bytes: bytes) -> pd.DataFrame:
        """Método 1: Extracción directa de texto"""
        st.write("  📄 Método 1: Extracción directa de texto")

        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        rows = []

        for page in doc:
            text = page.get_text()

            for line in text.split("\n"):
                line = line.strip()
                if not line:
                    continue

                fecha = self.find_date(line)
                monto_str = self.find_amount(line)

                if fecha and monto_str:
                    monto = self.parse_amount(monto_str)
                    if monto is not None:
                        rows.append({
                            "Fecha": fecha,
                            "Concepto": line,
                            "Monto": monto,
                            "MontoOriginal": monto_str
                        })

        doc.close()
        return pd.DataFrame(rows)

    def extract_layout(self, pdf_bytes: bytes) -> pd.DataFrame:
        """Método 2: Reconstrucción por layout (ordenamiento de palabras)"""
        st.write("  📐 Método 2: Reconstrucción por layout")

        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        rows = []

        for page in doc:
            words = page.get_text("words")
            lines = {}

            for w in words:
                x0, y0, x1, y1, text, _, _, _ = w
                key = round(y0)

                if key not in lines:
                    lines[key] = []
                lines[key].append((x0, text))

            for y in sorted(lines):
                ordered = sorted(lines[y], key=lambda x: x[0])
                line = " ".join(w[1] for w in ordered)

                fecha = self.find_date(line)
                monto_str = self.find_amount(line)

                if fecha and monto_str:
                    monto = self.parse_amount(monto_str)
                    if monto is not None:
                        rows.append({
                            "Fecha": fecha,
                            "Concepto": line,
                            "Monto": monto,
                            "MontoOriginal": monto_str
                        })

        doc.close()
        return pd.DataFrame(rows)

    def extract_ocr(self, pdf_bytes: bytes) -> pd.DataFrame:
        """Método 3: OCR para PDFs escaneados o con imágenes"""
        if not OCR_DISPONIBLE:
            st.warning("  ⚠️ OCR no disponible - faltabibliotecas")
            return pd.DataFrame()

        st.write("  🔍 Método 3: OCR (reconocimiento óptico)")

        try:
            images = convert_from_bytes(
                pdf_bytes,
                dpi=OCR_DPI,
                poppler_path=POPPLER_PATH
            )
        except Exception as e:
            st.warning(f"  ⚠️ Error en conversión: {e}")
            return pd.DataFrame()

        rows = []
        buffer = ""

        for i, img in enumerate(images):
            st.write(f"    Procesando página {i + 1}/{len(images)}")

            try:
                text = pytesseract.image_to_string(
                    img,
                    lang=OCR_LANG,
                    config="--psm 6"
                )
            except Exception as e:
                st.warning(f"  ⚠️ Error OCR página {i+1}: {e}")
                continue

            for line in text.split("\n"):
                line = line.strip()

                if not line:
                    continue

                # Si encuentra fecha, inicia nuevo registro
                if self.find_date(line):
                    buffer = line
                else:
                    buffer += " " + line

                # Buscar monto en el buffer
                monto_str = self.find_amount(buffer)
                if monto_str:
                    monto = self.parse_amount(monto_str)
                    if monto is not None:
                        fecha = self.find_date(buffer)
                        rows.append({
                            "Fecha": fecha or "",
                            "Concepto": buffer,
                            "Monto": monto,
                            "MontoOriginal": monto_str
                        })
                    buffer = ""

        return pd.DataFrame(rows)

    def extraer_movimientos(self, pdf_file) -> pd.DataFrame:
        """Método híbrido: intenta múltiples estrategias"""
        pdf_bytes = pdf_file.read()

        # Intentar método 1: texto directo
        df = self.extract_text_direct(pdf_bytes)
        if len(df) >= UMBRAL_EXTRACCION_EXITO:
            st.success(f"  ✅ Extracción directa exitosa: {len(df)} movimientos")
            return df

        # Intentar método 2: layout
        df = self.extract_layout(pdf_bytes)
        if len(df) >= UMBRAL_EXTRACCION_EXITO:
            st.success(f"  ✅ Extracción por layout exitosa: {len(df)} movimientos")
            return df

        # Intentar método 3: OCR
        if OCR_DISPONIBLE:
            df = self.extract_ocr(pdf_bytes)
            if len(df) > 0:
                st.success(f"  ✅ Extracción OCR exitosa: {len(df)} movimientos")
                return df
        else:
            st.warning("  ⚠️ OCR no disponible, no se puede intentar este método")

        st.error("  ❌ No se pudieron detectar movimientos")
        return pd.DataFrame()

    def normalizar_fechas(self, df: pd.DataFrame, year: int) -> pd.DataFrame:
        """Normaliza las fechas agregando el año"""
        if df.empty:
            return df

        df = df.copy()

        # Intentar diferentes formatos de fecha
        def parse_fecha(fecha_str):
            if not fecha_str:
                return None

            # Formato DD-MM
            if "-" in fecha_str:
                try:
                    return pd.to_datetime(f"{fecha_str}-{year}", format="%d-%m-%Y")
                except:
                    pass

            # Formato DD/MM/YYYY
            if "/" in fecha_str and len(fecha_str.split("/")[-1]) == 4:
                try:
                    return pd.to_datetime(fecha_str, format="%d/%m/%Y")
                except:
                    pass

            # Formato DD/MM/YY
            if "/" in fecha_str and len(fecha_str.split("/")[-1]) == 2:
                try:
                    return pd.to_datetime(f"{fecha_str}-{year}", format="%d/%m/%y-%Y")
                except:
                    pass

            return None

        df["Fecha"] = df["Fecha"].apply(parse_fecha)
        df = df.dropna(subset=["Fecha"])

        return df