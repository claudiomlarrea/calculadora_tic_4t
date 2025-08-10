
import streamlit as st
import pandas as pd
from io import BytesIO
from analyzer import generar_analisis_tic_ampliado, generar_informe_narrativo_tic

st.set_page_config(page_title="Calculadora TIC – 4º Trimestre (2017–2024)", layout="wide")
st.title("📊 Calculadora TIC – 4º Trimestre (2017–2024)")
st.markdown("Subí las bases TIC (hogares e individuos). La app integra y genera Excel + Word con **conclusiones profundas**.")

anio = st.selectbox("📅 Seleccioná el año del 4º trimestre", list(range(2017, 2025)))

c1, c2 = st.columns(2)
with c1:
    hogares_tic_file = st.file_uploader("🏠 Base de hogares TIC (.xlsx)", type=["xlsx"])
with c2:
    individuos_tic_file = st.file_uploader("👤 Base de individuos TIC (.xlsx)", type=["xlsx"])

pdf_file = st.file_uploader("📘 Manual de códigos TIC (opcional, .pdf)", type=["pdf"])

st.caption("La app intenta detectar automáticamente edad (CH06/EDAD), sexo (CH04), educación (NIVEL_ED...), ingreso del hogar (ITF...) y pesos muestrales (PONDERA...).")

if hogares_tic_file and individuos_tic_file:
    st.success("Archivos cargados. Hacé clic en el botón para procesar.")

    if st.button("▶️ Generar informe"):
        with st.spinner("Leyendo y unificando bases..."):
            df_h = pd.read_excel(hogares_tic_file)
            df_i = pd.read_excel(individuos_tic_file)

            # Detectar claves disponibles para merge
            claves = [c for c in ["CODUSU","NRO_HOGAR","AGLOMERADO"] if c in df_i.columns and c in df_h.columns]
            if not claves:
                st.error("No se hallaron claves comunes para el merge (se esperan: CODUSU, NRO_HOGAR, AGLOMERADO).")
                st.stop()

            df = pd.merge(df_i, df_h, on=claves, how="left")

        with st.spinner("Calculando indicadores y brechas..."):
            try:
                tablas, df_enriquecido, resumen = generar_analisis_tic_ampliado(df)
            except Exception as e:
                st.exception(e)
                st.stop()

        # Exportar Excel con todas las tablas
        excel_io = BytesIO()
        with pd.ExcelWriter(excel_io, engine="openpyxl") as wr:
            for hoja, tabla in tablas.items():
                tabla.to_excel(wr, sheet_name=hoja[:31], index=False)
        excel_io.seek(0)
        st.download_button("📥 Descargar Excel del análisis TIC", data=excel_io.getvalue(), file_name=f"TIC_{anio}_analisis.xlsx")

        # Exportar Word con narrativa extendida
        with st.spinner("Redactando conclusiones extendidas..."):
            word_io = generar_informe_narrativo_tic(tablas, anio=str(anio), resumen=resumen)
        st.download_button("📄 Descargar Informe Word TIC", data=word_io.getvalue(), file_name=f"Informe_TIC_{anio}.docx")

        if pdf_file:
            st.download_button("📘 Descargar Manual PDF TIC", data=pdf_file.read(), file_name=f"Manual_TIC_{anio}.pdf")

        st.success("Listo. Abajo se muestran vistas previas.")
        for nombre, tabla in list(tablas.items())[:4]:
            st.subheader(nombre)
            st.dataframe(tabla, use_container_width=True)
else:
    st.info("📥 Subí las bases de hogares e individuos (y opcionalmente el instructivo PDF) para comenzar.")

