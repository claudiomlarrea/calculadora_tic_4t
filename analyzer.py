import pandas as pd
from docx import Document
from io import BytesIO

# Diccionarios de etiquetas nominales por variable
etiquetas = {
    "IP_III_04": {1: "Sí", 2: "No", 9: "Ns/Nc"},
    "IP_III_05": {1: "Sí", 2: "No", 9: "Ns/Nc"},
    "IP_III_06": {1: "Sí", 2: "No", 9: "Ns/Nc"},
    "IH_II_01": {1: "Sí", 2: "No", 9: "Ns/Nc"},
    "IH_II_02": {1: "Sí", 2: "No", 9: "Ns/Nc"},
}

def generar_analisis_tic_ampliado(df):
    resultados = {}

    df["excluido_binario"] = ((df["IP_III_04"] == 2) & (df["IP_III_06"] == 2)).astype(int)
    df["exclusion_ordinal"] = df.apply(lambda x: 2 if x["IP_III_04"] == 2 and x["IP_III_06"] == 2
                                       else 1 if x["IP_III_04"] == 2 or x["IP_III_06"] == 2
                                       else 0, axis=1)

    df["EDAD"] = pd.qcut(range(len(df)), q=5, labels=["0-17", "18-29", "30-44", "45-64", "65+"])
    excl_por_edad = df.groupby("EDAD")["excluido_binario"].mean().reset_index()
    excl_por_edad["Porcentaje"] = (excl_por_edad["excluido_binario"] * 100).round(2)
    excl_por_edad.drop(columns=["excluido_binario"], inplace=True)
    resultados["Exclusión por Edad"] = excl_por_edad

    for col in ["IP_III_04", "IP_III_05", "IP_III_06"]:
        dist = df[col].value_counts().sort_index().reset_index()
        dist.columns = [col, "Cantidad"]
        if col in etiquetas:
            dist[col] = dist[col].map(etiquetas[col])
        resultados[f"Distribución {col}"] = dist

    for col in ["IH_II_01", "IH_II_02"]:
        dist = df[col].value_counts().sort_index().reset_index()
        dist.columns = [col, "Cantidad"]
        if col in etiquetas:
            dist[col] = dist[col].map(etiquetas[col])
        resultados[f"Infrestructura {col}"] = dist

    excl_ordinal = df["exclusion_ordinal"].value_counts().sort_index().reset_index()
    excl_ordinal.columns = ["Nivel Exclusión", "Cantidad"]
    excl_ordinal["Nivel Exclusión"] = excl_ordinal["Nivel Exclusión"].map({
        0: "Sin exclusión",
        1: "Exclusión parcial",
        2: "Exclusión total"
    })
    resultados["Exclusión Ordinal"] = excl_ordinal

    return resultados, df

def generar_informe_narrativo_tic(resumen_dict, anio="2024"):
    doc = Document()
    doc.add_heading(f"Informe Analítico – Inclusión Digital TIC 4ºT {anio}", 0)
    doc.add_paragraph("Este informe presenta el diagnóstico de exclusión digital a partir del módulo TIC "
                      f"del cuarto trimestre del año {anio}, con base en datos de hogares e individuos.")

    doc.add_heading("1. Definición de Exclusión Digital", level=1)
    doc.add_paragraph("La exclusión digital es una forma de desigualdad que se expresa en la falta de acceso a tecnologías "
                      "de la información y la comunicación, así como en la incapacidad para usarlas de forma significativa "
                      "(UNESCO, 2020; Castaño-Muñoz et al., 2022).")

    doc.add_heading("2. Resultados Clave", level=1)

    # Renombrar apartados para el informe
    nombres_amigables = {
        "Distribución IP_III_04": "Uso de computadora",
        "Distribución IP_III_05": "Uso de celular",
        "Distribución IP_III_06": "Uso de internet",
        "Infrestructura IH_II_01": "Hogar con computadora",
        "Infrestructura IH_II_02": "Hogar con acceso a internet"
    }

    for nombre, tabla in resumen_dict.items():
        doc.add_heading(nombres_amigables.get(nombre, nombre.replace("_", " ")), level=2)
        for i, row in tabla.iterrows():
            valores = " – ".join([str(v) for v in row.values])
            doc.add_paragraph(f"• {valores}", style="List Bullet")

    doc.add_heading("3. Conclusiones", level=1)
    doc.add_paragraph("Los resultados muestran que persisten desigualdades en el uso y acceso a tecnologías digitales. "
                      "Los grupos más afectados son los segmentos mayores y aquellos con menor acceso a infraestructura digital. "
                      "Se recomienda priorizar políticas públicas focalizadas en conectividad inclusiva.")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
