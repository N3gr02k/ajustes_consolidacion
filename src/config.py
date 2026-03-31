"""
Configuracion del sistema de conciliacion bancaria.
"""
import os

# Rutas de herramientas externas
POPPLER_PATH = os.getenv("POPPLER_PATH", r"D:\poppler\poppler-25.12.0\Library\bin")
TESSERACT_PATH = os.getenv("TESSERACT_PATH", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
TESSDATA_PATH = os.getenv("TESSDATA_PREFIX", r"C:\Program Files\Tesseract-OCR\tessdata")

# Configuracion OCR
OCR_LANG = "spa"
OCR_DPI = 300

# Expresiones regulares para detectar fechas y montos
DATE_PATTERNS = {
    "dd-mm": r"\b\d{2}-\d{2}\b",
    "dd/mm/yyyy": r"\b\d{2}/\d{2}/\d{4}\b",
    "dd/mm/yy": r"\b\d{2}/\d{2}/\d{2}\b",
}

# El orden importa: primero negativos, luego montos sin signo.
AMOUNT_PATTERNS = {
    "negative_paren": r"\(\d{1,3}(?:,\d{3})*\.\d{2}\)",
    "with_sign": r"\d{1,3}(?:,\d{3})*\.\d{2}-",
    "standard": r"\d{1,3}(?:,\d{3})*\.\d{2}",
}

# Configuracion de conciliacion
TOLERANCIA_MONTO = 0.01
TOLERANCIA_DIAS = 2
MIN_MONTO_AGRUPACION = 50.0
MAX_ELEMENTOS_GRUPO = 4

# Configuracion de extraccion
UMBRAL_EXTRACCION_EXITO = 15
