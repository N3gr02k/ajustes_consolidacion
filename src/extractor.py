"""
Módulo deprecado - usar extractor_pdf.py
Este archivo se mantiene por compatibilidad pero no se usa.
"""
from .extractor_pdf import ExtractorPDF

# Mantener la función por compatibilidad
def parse_bcp_line(line):
    """Función de compatibilidad - usar ExtractorPDF"""
    extractor = ExtractorPDF()
    return extractor.parse_amount(line)