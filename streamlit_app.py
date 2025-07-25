import streamlit as st
import pandas as pd
from analyzer import generar_analisis_tic_ampliado, generar_informe_narrativo_tic
from io import BytesIO

st.set_page_config(page_title="Calculadora TIC 4ºT", layout="wide")
st.title("📊 Calculadora TIC – 4º Trimestre")

st.markdown("Subí las bases de hogares e individuos TIC para generar un análisis completo.")

col1, col2 = st.columns(2)
with col1:
    hogares_tic_file = st.file_uploader("🏠 Base de hogares TIC (.xlsx)", type="xlsx")
with col2:
    individuos_tic_file = st.file_uploader("👤 Base de individuos TIC (.xlsx)", type="xlsx")

anio = st.selectbox("🗓️ Seleccioná el año del trimestre", list(range(2017, 2025)), index=0)

if hogares_tic_file and individuos_tic_file:
    st.success("Archivos cargados correctamente.")

    if st.button("▶️ Generar análisis"):
        df_hogar = pd.read_excel(hogares_tic_file)
        df_ind = pd.read_excel(individuos_tic_file)
        df = pd.merge(df_ind, df_hogar, on=["CODUSU", "NRO_HOGAR", "AGLOMERADO"], how="left")

        resumen_dict, df_ampliado = generar_analisis_tic_ampliado(df)

        # Exportar Excel
        excel_io = BytesIO()
        with pd.ExcelWriter(excel_io, engine="openpyxl") as writer:
            for hoja, tabla in resumen_dict.items():
                tabla.to_excel(writer, sheet_name=hoja[:30], index=False)
        st.download_button("📥 Descargar análisis Excel", data=excel_io.getvalue(), file_name="Analisis_TIC.xlsx")

        # Exportar Word
        word_io = generar_informe_narrativo_tic(resumen_dict, anio=anio)
        st.download_button("📄 Descargar informe Word", data=word_io.getvalue(), file_name="Informe_TIC.docx")
