import pandas as pd
from docx import Document
from io import BytesIO

def analizar_tic(df_hogar, df_ind):
    df = pd.merge(df_ind, df_hogar, on=["CODUSU", "NRO_HOGAR", "AGLOMERADO"], how="left")
    df["excluido_binario"] = ((df["IP_III_04"] == "No") & (df["IP_III_06"] == "No")).astype(int)

    resumen = df.groupby("sexo")["excluido_binario"].mean().reset_index()
    resumen["excluido_binario"] = (resumen["excluido_binario"] * 100).round(2)
    resumen.columns = ["Sexo", "Porcentaje de Exclusión Digital"]
    return df, resumen

def generar_informe_tic(df, resumen, anio):
    doc = Document()
    doc.add_heading(f"Informe TIC – 4º Trimestre {anio}", 0)
    doc.add_paragraph("Este informe presenta los resultados del análisis del módulo TIC "
                      f"del 4º trimestre del año {anio}.")
    doc.add_heading("1. Exclusión Digital", level=1)
    doc.add_paragraph("Se define como exclusión digital la carencia de acceso o habilidades para utilizar tecnologías.")
    doc.add_paragraph("Se midió mediante la ausencia simultánea de uso de computadora e internet.")
    doc.add_heading("2. Resultados por sexo", level=1)
    for i, row in resumen.iterrows():
        doc.add_paragraph(f"{row['Sexo']}: {row['Porcentaje de Exclusión Digital']}%", style="List Bullet")
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
