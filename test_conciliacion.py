"""Test de conciliacion completo"""
import sys

# Evita errores de impresion en consola Windows (cp1252)
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

sys.path.insert(0, "D:/cloud/ajustes_consolidacion")

import pandas as pd

from src import Conciliador, ExtractorPDF

print("=== PRUEBA COMPLETA DE CONCILIACION ===")
print()

# 1. Extraer PDF de Mayo
print("1. Extrayendo PDF Mayo...")
with open("D:/cloud/ajustes_consolidacion/data/pdf/Mayo.pdf", "rb") as f:
    pdf_bytes = f.read()

extractor = ExtractorPDF()
df_banco = extractor.extract_text_direct(pdf_bytes)
df_banco = extractor.normalizar_fechas(df_banco, 2024)
print(f"   Movimientos banco: {len(df_banco)}")
print()

# 2. Leer Excel
print("2. Leyendo Excel sistema...")
df_gestion = pd.read_excel("D:/cloud/ajustes_consolidacion/data/excel/excel mayo.xls", header=15)

# Normalizar fechas
df_gestion["Fecha"] = pd.to_datetime(df_gestion["Fecha"], dayfirst=True, errors="coerce")
df_gestion = df_gestion.dropna(subset=["Fecha"])

# Normalizar montos
df_gestion["Ingreso"] = pd.to_numeric(df_gestion["Ingreso"].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
df_gestion["Egreso"] = pd.to_numeric(df_gestion["Egreso"].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
df_gestion["Monto"] = df_gestion["Ingreso"] - df_gestion["Egreso"]

# Agregar columna de estado
df_gestion["Estado de Auditoria"] = "Pendiente"

print("Columnas disponibles:", list(df_gestion.columns))

cols_seleccionadas = [
    c
    for c in ["Fecha", "Concepto", "Referencia", "Ingreso", "Egreso", "Monto", "Estado de Auditoria"]
    if c in df_gestion.columns
]
df_gestion = df_gestion[cols_seleccionadas].copy()

df_gestion = df_gestion[df_gestion["Monto"] != 0]

print(f"   Movimientos sistema: {len(df_gestion)}")
print()

# 3. Conciliar
print("3. Ejecutando conciliacion...")
conciliador = Conciliador(df_banco, df_gestion)
resultado = conciliador.conciliar()
print()

# 4. Resultados
print("4. RESULTADOS:")
print(f"   Total registros: {len(resultado)}")

estado_col = "Estado de Auditoría" if "Estado de Auditoría" in resultado.columns else "Estado de Auditoria"
conciliados = int(resultado[estado_col].astype(str).str.contains("Conciliado", na=False).sum())
print(f"   Conciliados: {conciliados}")

hallazgos = len(resultado) - conciliados
print(f"   Hallazgos: {hallazgos}")

print()
print("5. Primeros 10 resultados:")
print(resultado.head(10).to_string())
