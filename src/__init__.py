"""
Sistema de Conciliación Bancaria STRATUM
"""
from .extractor_pdf import ExtractorPDF
from .lector_excel import LectorExcelGestion
from .conciliador import Conciliador
from .exportador import ExportadorExcel
from .config import (
    TOLERANCIA_MONTO,
    TOLERANCIA_DIAS,
    MIN_MONTO_AGRUPACION,
    MAX_ELEMENTOS_GRUPO
)

__all__ = [
    "ExtractorPDF",
    "LectorExcelGestion",
    "Conciliador",
    "ExportadorExcel",
    "TOLERANCIA_MONTO",
    "TOLERANCIA_DIAS",
    "MIN_MONTO_AGRUPACION",
    "MAX_ELEMENTOS_GRUPO"
]

__version__ = "2.0.0"