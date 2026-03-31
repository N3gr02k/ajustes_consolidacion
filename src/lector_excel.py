"""
Lector de Excel del sistema de gestion.
"""
from __future__ import annotations

import pandas as pd
import streamlit as st


class LectorExcelGestion:
    """Lector del Excel del sistema de gestion."""

    def __init__(self):
        self.columnas_encontradas = []
        self.metadata = {}

    def buscar_fila_header(self, df_raw: pd.DataFrame) -> int:
        """Busca automaticamente la fila donde estan los headers."""
        for idx, row in df_raw.iterrows():
            row_str = [str(x).lower() for x in row.values if pd.notna(x)]
            row_text = " ".join(row_str)
            if "fecha" in row_text and ("concepto" in row_text or "egreso" in row_text or "ingreso" in row_text):
                return idx
        return 15

    def extraer_metadata(self, df_raw: pd.DataFrame, fila_header: int) -> dict:
        """Extrae el encabezado institucional del Excel fuente."""
        metadata = {}
        limite = min(fila_header, len(df_raw))
        for idx in range(limite):
            clave = df_raw.iloc[idx, 0] if df_raw.shape[1] > 0 else None
            valor = df_raw.iloc[idx, 1] if df_raw.shape[1] > 1 else None
            if pd.isna(clave) or pd.isna(valor):
                continue
            clave_txt = str(clave).strip()
            valor_txt = str(valor).strip()
            if not clave_txt or not valor_txt:
                continue
            metadata[clave_txt] = valor_txt
        return metadata

    def detectar_columnas(self, df: pd.DataFrame) -> dict:
        """Detecta que columnas existen en el Excel."""
        columnas_disponibles = [str(c) for c in df.columns]
        mapa = {}

        for col in columnas_disponibles:
            col_lower = col.lower().strip()
            if "fecha" in col_lower:
                mapa["fecha"] = col
                break

        for col in columnas_disponibles:
            col_lower = col.lower().strip()
            if any(x in col_lower for x in ["concepto", "descripcion", "detalle", "glosa"]):
                mapa["concepto"] = col
                break

        for col in columnas_disponibles:
            col_lower = col.lower().strip()
            if any(x in col_lower for x in ["operacion", "nro", "numero"]):
                mapa["operacion"] = col
                break

        for col in columnas_disponibles:
            col_lower = col.lower().strip()
            if "ingreso" in col_lower and "egreso" not in col_lower:
                mapa["ingreso"] = col
            elif "egreso" in col_lower:
                mapa["egreso"] = col
            elif "monto" in col_lower:
                mapa["monto"] = col
            elif "saldo" in col_lower:
                mapa["saldo"] = col
            elif "referencia" in col_lower:
                mapa["referencia"] = col

        return mapa

    def normalizar_excel(self, archivo_excel) -> pd.DataFrame:
        """Normaliza el Excel del sistema de gestion."""
        st.info("  Normalizando Excel del sistema...")

        try:
            if hasattr(archivo_excel, "seek"):
                archivo_excel.seek(0)
            df_raw = pd.read_excel(archivo_excel, header=None)
            st.write(f"    Excel detectado: {df_raw.shape[0]} filas")

            fila_header = self.buscar_fila_header(df_raw)
            self.metadata = self.extraer_metadata(df_raw, fila_header)
            st.write(f"    Headers encontrados en fila: {fila_header + 1}")

            if hasattr(archivo_excel, "seek"):
                archivo_excel.seek(0)
            df = pd.read_excel(archivo_excel, header=fila_header)
        except Exception as e:
            st.error(f"    Error al leer Excel: {e}")
            return pd.DataFrame()

        df.columns = [str(c).strip() for c in df.columns]
        mapa = self.detectar_columnas(df)
        self.columnas_encontradas = list(df.columns)
        st.write(f"    Columnas mapeadas: {mapa}")

        if "fecha" in mapa:
            df[mapa["fecha"]] = pd.to_datetime(df[mapa["fecha"]], dayfirst=True, errors="coerce")
            df = df.dropna(subset=[mapa["fecha"]])

        tiene_monto_calculado = False
        if "ingreso" in mapa and "egreso" in mapa:
            df[mapa["ingreso"]] = pd.to_numeric(df[mapa["ingreso"]].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
            df[mapa["egreso"]] = pd.to_numeric(df[mapa["egreso"]].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
            df["Monto"] = df[mapa["ingreso"]] - df[mapa["egreso"]]
            tiene_monto_calculado = True
        elif "monto" in mapa:
            df["Monto"] = pd.to_numeric(df[mapa["monto"]].astype(str).str.replace(",", ""), errors="coerce")
            tiene_monto_calculado = True

        if not tiene_monto_calculado:
            st.error("    No se pudo detectar columna de monto")
            return pd.DataFrame()

        if "Estado de Auditoría" not in df.columns and "Estado de Auditoria" not in df.columns:
            df["Estado de Auditoría"] = "Pendiente"

        columnas_salida = []
        if "fecha" in mapa:
            columnas_salida.append(mapa["fecha"])
        if "concepto" in mapa:
            columnas_salida.append(mapa["concepto"])
        if "referencia" in mapa:
            columnas_salida.append(mapa["referencia"])
        if "operacion" in mapa:
            columnas_salida.append(mapa["operacion"])
        if "ingreso" in mapa:
            columnas_salida.append(mapa["ingreso"])
        if "egreso" in mapa:
            columnas_salida.append(mapa["egreso"])
        columnas_salida.append("Monto")
        if "saldo" in mapa:
            columnas_salida.append(mapa["saldo"])
        if "Estado de Auditoría" in df.columns:
            columnas_salida.append("Estado de Auditoría")
        if "Estado de Auditoria" in df.columns:
            columnas_salida.append("Estado de Auditoria")

        columnas_salida = [c for c in columnas_salida if c in df.columns]
        columnas_salida = list(dict.fromkeys(columnas_salida))
        df_normalizado = df[columnas_salida].copy()

        rename_map = {}
        if "fecha" in mapa:
            rename_map[mapa["fecha"]] = "Fecha"
        if "concepto" in mapa:
            rename_map[mapa["concepto"]] = "Concepto"
        if "referencia" in mapa:
            rename_map[mapa["referencia"]] = "Referencia"
        if "operacion" in mapa:
            rename_map[mapa["operacion"]] = "Número de operación"
        if "ingreso" in mapa:
            rename_map[mapa["ingreso"]] = "Ingreso"
        if "egreso" in mapa:
            rename_map[mapa["egreso"]] = "Egreso"
        if "saldo" in mapa:
            rename_map[mapa["saldo"]] = "Saldo"
        df_normalizado.rename(columns=rename_map, inplace=True)

        df_normalizado = df_normalizado[df_normalizado["Monto"].notna()]
        df_normalizado = df_normalizado[df_normalizado["Monto"] != 0]

        st.success(f"    Excel normalizado: {len(df_normalizado)} registros")
        return df_normalizado
