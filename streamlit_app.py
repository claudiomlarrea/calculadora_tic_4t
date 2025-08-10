
import streamlit as st
import pandas as pd
from analyzer import generar_analisis_tic_ampliado, generar_informe_narrativo_tic
from io import BytesIO

st.set_page_config(page_title="Calculadora TIC – 4º Trimestre (2017–2024)", layout="wide")
st.title("📊 Calculadora TIC – 4º Trimestre (2017–2024)")
st.markdown("Subí las bases TIC (hogares e individuos), el sistema las integra y genera un Excel con tablas y un Word con conclusiones **profundas y robustas**.")

anio = st.selectbox("📅 Seleccioná el año del 4º trimestre", list(range(2017, 2025)))

col1, col2 = st.columns(2)
with col1:
    hogares_tic_file = st.file_uploader("🏠 Base de hogares TIC (.xlsx)", type=["xlsx"])
with col2:
    individuos_tic_file = st.file_uploader("👤 Base de individuos TIC (.xlsx)", type=["xlsx"])

pdf_file = st.file_uploader("📘 Manual de códigos TIC (opcional, .pdf)", type=["pdf"])

st.caption("**Nota:** si tu base tiene nombres de columnas alternativos (p. ej., EDAD en lugar de CH06, PONDERA como factor de expansión), la app los detecta automáticamente.")

if hogares_tic_file and individuos_tic_file:
    st.success("Archivos cargados correctamente. Hacé clic en el botón para procesar.")

    if st.button("▶️ Generar informe"):
        with st.spinner("Procesando y generando tablas..."):
            df_hogar = pd.read_excel(hogares_tic_file)
            df_ind = pd.read_excel(individuos_tic_file)
            # Merge estándar EPH
            on_cols = [c for c in ["CODUSU", "NRO_HOGAR", "AGLOMERADO"] if c in df_ind.columns and c in df_hogar.columns]
            if not on_cols:
                st.error("No se encontraron claves comunes para merge (esperadas: CODUSU, NRO_HOGAR, AGLOMERADO). Revisá tus archivos.")
                st.stop()
            df = pd.merge(df_ind, df_hogar, on=on_cols, how="left")

            try:
                tablas, df_ampliado, resumen = generar_analisis_tic_ampliado(df)
            except Exception as e:
                st.exception(e)
                st.stop()

        # Exportar Excel (todas las tablas)
        excel_io = BytesIO()
        with pd.ExcelWriter(excel_io, engine="openpyxl") as writer:
            for hoja, tabla in tablas.items():
                # Nombre de hoja <= 31 caracteres
                sheet = hoja[:31]
                tabla.to_excel(writer, sheet_name=sheet, index=False)
        excel_io.seek(0)
        st.download_button("📥 Descargar Excel del análisis TIC", data=excel_io.getvalue(), file_name=f"TIC_{anio}_analisis.xlsx")

        # Exportar Word con narrativa robusta
        with st.spinner("Armando informe Word..."):
            word_io = generar_informe_narrativo_tic(tablas, anio=str(anio), resumen=resumen)
        st.download_button("📄 Descargar Informe Word TIC", data=word_io.getvalue(), file_name=f"Informe_TIC_{anio}.docx")

        if pdf_file:
            st.download_button("📘 Descargar Manual PDF TIC", data=pdf_file.read(), file_name=f"Manual_TIC_{anio}.pdf")

        st.success("Listo. Abajo podés previsualizar algunas tablas clave.")
        # Preview
        for nombre, tabla in list(tablas.items())[:3]:
            st.subheader(nombre)
            st.dataframe(tabla, use_container_width=True)

else:
    st.info("📥 Subí las bases de hogares e individuos (y opcionalmente el instructivo PDF) para comenzar.")
