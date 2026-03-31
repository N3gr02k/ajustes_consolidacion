"""
Motor de conciliacion bancaria orientado a uso contable real.
"""
from __future__ import annotations

from datetime import timedelta
from itertools import combinations
import re
import unicodedata
from typing import List, Optional, Set

import pandas as pd
import streamlit as st

from .config import (
    MAX_ELEMENTOS_GRUPO,
    MIN_MONTO_AGRUPACION,
    TOLERANCIA_DIAS,
    TOLERANCIA_MONTO,
)


class Conciliador:
    """Motor de conciliacion entre banco y sistema."""

    ESTADO_COL = "Estado de Auditoría"
    ESTADO_COL_ALT = "Estado de Auditoria"
    STOPWORDS = {
        "ABONO",
        "CARGO",
        "EFECTIVO",
        "TRANSFERENCIA",
        "BPI",
        "VEN",
        "TLC",
        "INT",
        "CAJ",
        "POS",
        "PAGO",
        "BANCO",
        "DEPOSITO",
        "DEPOSITOS",
        "OTROS",
        "MANT",
        "MANTENIMIENTO",
        "INGRESO",
        "EGRESO",
    }

    def __init__(self, df_banco: pd.DataFrame, df_gestion: pd.DataFrame):
        self.df_banco = df_banco.copy()
        self.df_gestion = df_gestion.copy()
        self.registros_usados_gestion: Set[int] = set()
        self.registros_usados_banco: Set[int] = set()
        self._contador_progreso = 0

    def _normalizar_columna_estado(self, df: pd.DataFrame) -> pd.DataFrame:
        """Unifica la columna de estado para soportar variantes con/sin tilde."""
        df = df.copy()
        if self.ESTADO_COL not in df.columns:
            if self.ESTADO_COL_ALT in df.columns:
                df[self.ESTADO_COL] = df[self.ESTADO_COL_ALT]
            else:
                df[self.ESTADO_COL] = "Pendiente"

        if self.ESTADO_COL_ALT in df.columns:
            df.drop(columns=[self.ESTADO_COL_ALT], inplace=True)

        return df

    def _normalizar_texto(self, valor: object) -> str:
        texto = "" if pd.isna(valor) else str(valor)
        texto = unicodedata.normalize("NFKD", texto)
        texto = texto.encode("ascii", "ignore").decode("ascii")
        texto = re.sub(r"[^A-Z0-9]+", " ", texto.upper()).strip()
        return texto

    def _extraer_identificadores(self, fila: pd.Series) -> Set[str]:
        """Extrae tokens utiles para comparar descripciones operativas."""
        partes = [
            fila.get("Concepto", ""),
            fila.get("Referencia", ""),
            fila.get("Número de operación", ""),
        ]
        texto = self._normalizar_texto(" ".join("" if pd.isna(x) else str(x) for x in partes))
        tokens = set()

        for token in texto.split():
            if len(token) < 3:
                continue
            if token in self.STOPWORDS:
                continue
            tiene_letras = any(c.isalpha() for c in token)
            tiene_numeros = any(c.isdigit() for c in token)

            if tiene_letras and tiene_numeros:
                tokens.add(token)
                continue

            if tiene_letras and len(token) >= 5:
                tokens.add(token)

        return tokens

    def _preparar_movimientos(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
        df["Monto"] = pd.to_numeric(df["Monto"], errors="coerce")
        df = df.dropna(subset=["Fecha", "Monto"])
        df["MontoRound"] = df["Monto"].round(2)
        df["Identificadores"] = df.apply(self._extraer_identificadores, axis=1)
        return df

    def _candidatos_por_monto(
        self,
        df_candidatos: pd.DataFrame,
        monto_banco: float,
        indices_usados: Set[int],
    ) -> pd.DataFrame:
        return df_candidatos[
            (~df_candidatos.index.isin(indices_usados))
            & (abs(df_candidatos["Monto"] - monto_banco) <= TOLERANCIA_MONTO)
        ].copy()

    def buscar_coincidencia_directa(
        self,
        movimiento_banco: pd.Series,
        df_candidatos: pd.DataFrame,
        indices_usados: Set[int],
    ) -> Optional[tuple[int, str]]:
        """Prioriza monto igual y fecha mas cercana."""
        monto_banco = float(movimiento_banco["Monto"])
        fecha_banco = movimiento_banco["Fecha"]
        candidatos = self._candidatos_por_monto(df_candidatos, monto_banco, indices_usados)

        if len(candidatos) == 0:
            return None

        candidatos["diff_dias"] = abs((candidatos["Fecha"] - fecha_banco).dt.days)
        mismos_dia = candidatos[candidatos["diff_dias"] == 0]
        if len(mismos_dia) > 0:
            idx = int(mismos_dia.sort_index().index[0])
            return idx, "Conciliado"

        idx = int(candidatos.sort_values(["diff_dias", "Fecha"]).index[0])
        return idx, "Conciliado por fecha"

    def buscar_coincidencia_por_referencia(
        self,
        movimiento_banco: pd.Series,
        resultados_gestion: pd.DataFrame,
    ) -> Optional[tuple[int, str]]:
        """
        Busca un match fuera de la tolerancia de dias usando monto + identificadores.
        Esto ayuda cuando el administrador registra el movimiento en otra fecha.
        """
        monto_banco = float(movimiento_banco["Monto"])
        fecha_banco = movimiento_banco["Fecha"]
        ids_banco = movimiento_banco.get("Identificadores", set())

        if not ids_banco:
            return None

        candidatos = self._candidatos_por_monto(
            resultados_gestion,
            monto_banco,
            self.registros_usados_gestion,
        )

        if len(candidatos) == 0:
            return None

        def _score(fila: pd.Series) -> tuple[int, int]:
            ids_gestion = fila.get("Identificadores", set())
            overlap = len(ids_banco.intersection(ids_gestion))
            diff = abs((fila["Fecha"] - fecha_banco).days)
            return overlap, -diff

        candidatos["score_overlap"] = candidatos.apply(lambda fila: _score(fila)[0], axis=1)
        candidatos["score_fecha"] = candidatos.apply(lambda fila: _score(fila)[1], axis=1)
        candidatos = candidatos[candidatos["score_overlap"] > 0]

        if len(candidatos) == 0:
            return None

        idx = int(
            candidatos.sort_values(
                ["score_overlap", "score_fecha", "Fecha"],
                ascending=[False, False, True],
            ).index[0]
        )
        return idx, "Conciliado por referencia"

    def buscar_agrupacion(
        self,
        movimiento_banco: pd.Series,
        df_candidatos: pd.DataFrame,
        indices_usados: Set[int],
    ) -> Optional[List[int]]:
        """Busca combinaciones de registros que sumen el monto objetivo."""
        monto_banco = float(movimiento_banco["Monto"])
        fecha_banco = movimiento_banco["Fecha"]
        signo_objetivo = 1 if monto_banco > 0 else -1

        candidatos = df_candidatos[~df_candidatos.index.isin(indices_usados)].copy()
        candidatos = candidatos[candidatos["Monto"] * signo_objetivo > 0]

        if len(candidatos) < 2:
            return None

        if abs(monto_banco) < MIN_MONTO_AGRUPACION:
            return None

        candidatos["diff_dias"] = abs((candidatos["Fecha"] - fecha_banco).dt.days)
        candidatos = candidatos.sort_values(["diff_dias", "Fecha"]).head(18)

        idxs = list(candidatos.index)
        max_elementos = min(MAX_ELEMENTOS_GRUPO, len(idxs))

        for r in range(2, max_elementos + 1):
            for combo in combinations(idxs, r):
                suma = float(candidatos.loc[list(combo), "Monto"].sum())
                if abs(suma - monto_banco) <= TOLERANCIA_MONTO:
                    return list(combo)

        return None

    def _registrar_hallazgo_banco(
        self,
        resultados_banco: List[dict],
        movimiento: pd.Series,
        estado: str,
    ) -> None:
        monto = float(movimiento["Monto"])
        resultados_banco.append(
            {
                "Fecha": movimiento["Fecha"],
                "Concepto": f"HALLAZGO BANCO: {movimiento.get('Concepto', '')}",
                "Referencia": "",
                "Número de operación": "",
                "Ingreso": monto if monto > 0 else "",
                "Egreso": abs(monto) if monto < 0 else "",
                "Saldo": "",
                self.ESTADO_COL: estado,
            }
        )

    def conciliar(self) -> pd.DataFrame:
        """Ejecuta la conciliacion completa."""
        if self.df_banco.empty:
            st.warning("  No hay movimientos del banco")
            return self.df_gestion

        if self.df_gestion.empty:
            st.warning("  No hay movimientos en el sistema")
            return self.df_banco

        self.df_banco = self._preparar_movimientos(self.df_banco)
        self.df_gestion = self._preparar_movimientos(self._normalizar_columna_estado(self.df_gestion))

        resultados_gestion = self.df_gestion.copy()
        resultados_banco: List[dict] = []
        total_banco = len(self.df_banco)

        for idx_banco, movimiento in self.df_banco.sort_values(["Fecha", "Monto"]).iterrows():
            self._contador_progreso += 1
            if self._contador_progreso % 25 == 0:
                st.write(f"    Procesando... {self._contador_progreso}/{total_banco}")

            monto = float(movimiento["Monto"])
            fecha = movimiento["Fecha"]
            fecha_min = fecha - timedelta(days=TOLERANCIA_DIAS)
            fecha_max = fecha + timedelta(days=TOLERANCIA_DIAS)

            candidatos_cercanos = resultados_gestion[
                resultados_gestion["Fecha"].between(fecha_min, fecha_max)
            ]

            coincidencia = self.buscar_coincidencia_directa(
                movimiento,
                candidatos_cercanos,
                self.registros_usados_gestion,
            )
            if coincidencia is not None:
                idx_coincidencia, estado = coincidencia
                resultados_gestion.loc[idx_coincidencia, self.ESTADO_COL] = estado
                self.registros_usados_gestion.add(idx_coincidencia)
                self.registros_usados_banco.add(idx_banco)
                continue

            coincidencia_ref = self.buscar_coincidencia_por_referencia(
                movimiento,
                resultados_gestion,
            )
            if coincidencia_ref is not None:
                idx_coincidencia, estado = coincidencia_ref
                resultados_gestion.loc[idx_coincidencia, self.ESTADO_COL] = estado
                self.registros_usados_gestion.add(idx_coincidencia)
                self.registros_usados_banco.add(idx_banco)
                continue

            grupo = self.buscar_agrupacion(
                movimiento,
                candidatos_cercanos,
                self.registros_usados_gestion,
            )
            if grupo:
                for idx_g in grupo:
                    resultados_gestion.loc[idx_g, self.ESTADO_COL] = "Conciliado por agrupacion"
                    self.registros_usados_gestion.add(idx_g)

                self.registros_usados_banco.add(idx_banco)
                continue

            self.registros_usados_banco.add(idx_banco)

            candidatos_mismo_monto = self._candidatos_por_monto(
                resultados_gestion,
                monto,
                self.registros_usados_gestion,
            )
            if len(candidatos_mismo_monto) > 0:
                self._registrar_hallazgo_banco(
                    resultados_banco,
                    movimiento,
                    "Hallazgo - Monto coincide pero fecha/referencia no",
                )
            else:
                self._registrar_hallazgo_banco(
                    resultados_banco,
                    movimiento,
                    "Hallazgo - Falta en Sistema",
                )

        mask_sobran_gestion = ~resultados_gestion.index.isin(self.registros_usados_gestion)
        faltan_en_banco = int(mask_sobran_gestion.sum())
        if faltan_en_banco > 0:
            st.warning(f"  {faltan_en_banco} movimientos en sistema sin correspondencia en banco")
            resultados_gestion.loc[mask_sobran_gestion, self.ESTADO_COL] = "Hallazgo - Falta en Banco"

        columnas_aux = ["MontoRound", "Identificadores"]
        resultados_gestion = resultados_gestion.drop(
            columns=[col for col in columnas_aux if col in resultados_gestion.columns]
        )

        if len(resultados_banco) > 0:
            df_banco_resultado = pd.DataFrame(resultados_banco)
            resultado_final = pd.concat([resultados_gestion, df_banco_resultado], ignore_index=True)
        else:
            resultado_final = resultados_gestion.copy()

        resultado_final = resultado_final.sort_values("Fecha").reset_index(drop=True)

        total = len(resultado_final)
        conciliados = int(
            resultado_final[self.ESTADO_COL].astype(str).str.contains("Conciliado", na=False).sum()
        )
        hallazgos = total - conciliados

        st.success("  Conciliacion completada:")
        st.write(f"     - Total registros: {total}")
        st.write(f"     - Conciliados: {conciliados}")
        st.write(f"     - Hallazgos: {hallazgos}")

        return resultado_final

    def generar_reporte_diferencias(self) -> dict:
        """Genera un reporte detallado de diferencias."""
        banco_no_en_gestion = self.df_banco[
            ~self.df_banco.index.isin(self.registros_usados_banco)
        ]
        gestion_no_en_banco = self.df_gestion[
            ~self.df_gestion.index.isin(self.registros_usados_gestion)
        ]

        return {
            "solo_en_banco": banco_no_en_gestion,
            "solo_en_gestion": gestion_no_en_banco,
            "total_banco": len(self.df_banco),
            "total_gestion": len(self.df_gestion),
            "conciliados_banco": len(self.registros_usados_banco),
            "conciliados_gestion": len(self.registros_usados_gestion),
        }
