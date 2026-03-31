"""
STRATUM v2.0 - Sistema de conciliacion bancaria universal.
Aplicacion principal con Streamlit.
"""
from datetime import datetime

import streamlit as st

from src import Conciliador, ExportadorExcel, ExtractorPDF, LectorExcelGestion


st.set_page_config(
    page_title="STRATUM v2.0 - Conciliacion Bancaria",
    page_icon="SB",
    layout="wide",
)


def main():
    """Aplicacion principal."""
    st.title("STRATUM v2.0 - Conciliacion Bancaria Universal")
    st.markdown("---")

    with st.sidebar:
        st.header("Configuracion")

        col_mes, col_ano = st.columns(2)
        with col_mes:
            mes = st.selectbox(
                "Mes",
                [
                    "Enero",
                    "Febrero",
                    "Marzo",
                    "Abril",
                    "Mayo",
                    "Junio",
                    "Julio",
                    "Agosto",
                    "Septiembre",
                    "Octubre",
                    "Noviembre",
                    "Diciembre",
                ],
                index=datetime.now().month - 1,
            )
        with col_ano:
            year = st.number_input(
                "Ano",
                min_value=2020,
                max_value=2030,
                value=datetime.now().year,
            )

        meses_codigo = {
            "Enero": "01",
            "Febrero": "02",
            "Marzo": "03",
            "Abril": "04",
            "Mayo": "05",
            "Junio": "06",
            "Julio": "07",
            "Agosto": "08",
            "Septiembre": "09",
            "Octubre": "10",
            "Noviembre": "11",
            "Diciembre": "12",
        }
        mes_codigo = meses_codigo[mes]

        st.markdown("---")
        st.markdown("### Parametros de conciliacion")

        from src import MIN_MONTO_AGRUPACION, TOLERANCIA_DIAS, TOLERANCIA_MONTO

        st.info(
            f"""
        **Parametros activos:**
        - Tolerancia monto: +/-{TOLERANCIA_MONTO}
        - Dias de tolerancia: {TOLERANCIA_DIAS}
        - Min. monto para agrupar: {MIN_MONTO_AGRUPACION}
        """
        )

        st.markdown("---")
        st.markdown("### Instrucciones")
        st.markdown(
            """
        **Paso 1:** Subir el PDF del estado de cuenta bancario

        **Paso 2:** Subir el Excel del sistema de gestion

        **Paso 3:** Hacer clic en "Ejecutar conciliacion"

        El sistema detectara:
        - Movimientos conciliados
        - Movimientos en banco que faltan en sistema
        - Movimientos en sistema que faltan en banco
        - Grupos de transacciones
        """
        )

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Estado de Cuenta (PDF)")
        pdf_file = st.file_uploader(
            "Subir PDF del banco",
            type=["pdf"],
            help="Estado de cuenta en formato PDF",
        )

    with col2:
        st.subheader("Sistema de Gestion (Excel)")
        excel_file = st.file_uploader(
            "Subir Excel del sistema",
            type=["xls", "xlsx"],
            help="Archivo Excel con los movimientos registrados",
        )

    st.markdown("---")

    if pdf_file and excel_file:
        if st.button("Ejecutar conciliacion", type="primary", use_container_width=True):
            try:
                st.header("Etapa 1: Extraccion del PDF")

                extractor = ExtractorPDF()
                with st.spinner("Extrayendo movimientos del PDF..."):
                    df_banco = extractor.extraer_movimientos(pdf_file)

                if df_banco.empty:
                    st.error("No se pudieron extraer movimientos del PDF")
                    return

                df_banco = extractor.normalizar_fechas(df_banco, year)
                st.write(f"**Movimientos extraidos del banco:** {len(df_banco)}")
                st.dataframe(df_banco.head(10), use_container_width=True)

                st.markdown("---")
                st.header("Etapa 2: Lectura del sistema de gestion")

                lector = LectorExcelGestion()
                with st.spinner("Normalizando Excel del sistema..."):
                    df_gestion = lector.normalizar_excel(excel_file)

                if df_gestion.empty:
                    st.error("No se pudieron leer movimientos del Excel")
                    return

                st.write(f"**Movimientos del sistema:** {len(df_gestion)}")
                st.dataframe(df_gestion.head(10), use_container_width=True)

                st.markdown("---")
                st.header("Etapa 3: Conciliacion")

                conciliador = Conciliador(df_banco, df_gestion)
                with st.spinner("Realizando conciliacion..."):
                    resultado = conciliador.conciliar()

                st.markdown("---")
                st.header("Resultados de la conciliacion")

                total = len(resultado)
                estado_col = (
                    "Estado de Auditoría"
                    if "Estado de Auditoría" in resultado.columns
                    else "Estado de Auditoria"
                )
                conciliados = len(
                    resultado[resultado[estado_col].astype(str).str.contains("Conciliado", na=False)]
                )
                hallazgos = total - conciliados

                m1, m2, m3 = st.columns(3)
                m1.metric("Total registros", total)
                m2.metric("Conciliados", conciliados, delta=conciliados, delta_color="normal")
                m3.metric("Hallazgos", hallazgos, delta=-hallazgos, delta_color="inverse")

                st.dataframe(resultado, use_container_width=True)

                st.header("Exportar resultados")
                exportador = ExportadorExcel()

                excel_buffer = exportador.exportar(
                    resultado,
                    f"conciliacion_stratum_{mes_codigo}_{year}.xlsx",
                    metadata=lector.metadata,
                )

                st.download_button(
                    label="Descargar Excel de conciliacion",
                    data=excel_buffer,
                    file_name=f"conciliacion_stratum_{year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

                resumen_buffer = exportador.exportar_resumen(
                    df_banco,
                    df_gestion,
                    metadata=lector.metadata,
                )

                st.download_button(
                    label="Descargar resumen",
                    data=resumen_buffer,
                    file_name=f"resumen_conciliacion_{year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.markdown("---")
                st.header("Analisis de hallazgos")

                hallazgos_sistema = resultado[
                    resultado[estado_col].astype(str).str.contains("Falta en Sistema", na=False)
                ]
                hallazgos_banco = resultado[
                    resultado[estado_col].astype(str).str.contains("Falta en Banco", na=False)
                ]
                por_revisar = resultado[
                    resultado[estado_col]
                    .astype(str)
                    .str.contains("No conciliado|fecha/referencia", na=False)
                ]

                if len(hallazgos_sistema) > 0:
                    st.subheader(
                        f"Movimientos en banco que faltan en sistema ({len(hallazgos_sistema)})"
                    )
                    st.dataframe(hallazgos_sistema, use_container_width=True)

                if len(hallazgos_banco) > 0:
                    st.subheader(
                        f"Movimientos en sistema que faltan en banco ({len(hallazgos_banco)})"
                    )
                    st.dataframe(hallazgos_banco, use_container_width=True)

                if len(por_revisar) > 0:
                    st.subheader(f"Movimientos por revisar ({len(por_revisar)})")
                    st.dataframe(por_revisar, use_container_width=True)

                agrupados = resultado[
                    resultado[estado_col].astype(str).str.contains("agrupacion", case=False, na=False)
                ]
                if len(agrupados) > 0:
                    st.subheader(f"Transacciones agrupadas ({len(agrupados)})")
                    st.dataframe(agrupados, use_container_width=True)

            except Exception as e:
                st.error(f"Error durante la conciliacion: {str(e)}")
                import traceback

                st.code(traceback.format_exc())

    else:
        st.info("Sube ambos archivos para comenzar la conciliacion.")

        st.markdown("---")
        st.subheader("Formatos esperados")

        col_e1, col_e2 = st.columns(2)
        with col_e1:
            st.markdown("**PDF del banco:**")
            st.markdown(
                """
            - Estados de cuenta en formato PDF
            - Puede ser de cualquier banco
            - El sistema intentara multiples metodos de extraccion:
              1. Texto directo
              2. Reconstruccion por layout
              3. OCR como ultimo recurso
            """
            )

        with col_e2:
            st.markdown("**Excel del sistema:**")
            st.markdown(
                """
            - Debe contener columnas de:
              - Fecha
              - Concepto o descripcion
              - Ingreso o egreso, o monto
            - El sistema detectara automaticamente las columnas
            """
            )


if __name__ == "__main__":
    main()
