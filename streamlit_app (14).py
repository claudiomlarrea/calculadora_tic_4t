import streamlit as st
import pandas as pd
from analyzer import analizar_tic, generar_informe_tic
from io import BytesIO

st.set_page_config(page_title="Calculadora TIC 4º Trimestre", layout="wide")

st.title("📊 Calculadora TIC – 4º Trimestre (2017–2024)")
st.markdown("Subí las bases TIC y generá tu informe automáticamente.")

st.selectbox("📅 Seleccioná el año del 4º trimestre", options=list(range(2017, 2025)), index=0, key="anio_tic")

col1, col2 = st.columns(2)
with col1:
    hogares_tic_file = st.file_uploader("🏠 Base de hogares TIC (.xlsx)", type="xlsx")
with col2:
    individuos_tic_file = st.file_uploader("👤 Base de individuos TIC (.xlsx)", type="xlsx")

manual_pdf = st.file_uploader("📄 Manual de códigos TIC (.pdf)", type="pdf")

if hogares_tic_file and individuos_tic_file and manual_pdf:
    st.success("Archivos cargados correctamente. Hacé clic en el botón para procesar.")

    if st.button("▶️ Generar informe"):
        df_hogar = pd.read_excel(hogares_tic_file)
        df_ind = pd.read_excel(individuos_tic_file)

        resultados, resumen = analizar_tic(df_hogar, df_ind)
        output_word = generar_informe_tic(resultados, resumen, st.session_state.anio_tic)

        st.download_button("📥 Descargar informe Word", data=output_word, file_name="Informe_TIC.docx")
        st.dataframe(resumen)