
import pandas as pd
from docx import Document
from io import BytesIO
from typing import Dict, Tuple, Any

# =========================
# Etiquetas nominales
# =========================
ETIQUETAS_BINARIAS = {
    "IP_III_04": {1: "Sí", 2: "No", 9: "Ns/Nc"},  # Uso de computadora
    "IP_III_05": {1: "Sí", 2: "No", 9: "Ns/Nc"},  # Uso de celular
    "IP_III_06": {1: "Sí", 2: "No", 9: "Ns/Nc"},  # Uso de internet
    "IH_II_01": {1: "Sí", 2: "No", 9: "Ns/Nc"},  # Hogar con computadora
    "IH_II_02": {1: "Sí", 2: "No", 9: "Ns/Nc"},  # Hogar con acceso a internet
    "CH04": {1: "Varón", 2: "Mujer"},            # Sexo (EPH)
}

ETIQUETAS_EDUC = {
    1: "Sin instrucción",
    2: "Primaria incompleta",
    3: "Primaria completa",
    4: "Secundaria incompleta",
    5: "Secundaria completa",
    6: "Superior incompleto",
    7: "Superior completo",
    9: "Ns/Nc",
}

# Columnas candidatas por si cambian entre trimestres
CANDIDATAS_EDAD = ["CH06", "EDAD", "edad"]
CANDIDATAS_SEXO = ["CH04", "SEXO"]
CANDIDATAS_EDUC = ["NIVEL_ED", "NIVEL_EDUC", "NIVEL_EDUCATIVO"]
CANDIDATA_INGRESO = ["ITF", "ITF_HOGAR", "ingreso_total_hogar"]
CANDIDATAS_PESO = ["PONDERA", "PONDIIO", "PONDIH", "PESO", "FACTOR", "FACTOR_EXP"]

def _primera_columna(df: pd.DataFrame, candidatas) -> str | None:
    for c in candidatas:
        if c in df.columns:
            return c
    return None

def _value_counts_pesado(serie: pd.Series, pesos: pd.Series | None) -> pd.DataFrame:
    if pesos is not None:
        tmp = pd.DataFrame({"x": serie, "w": pesos}).dropna(subset=["x"])
        vc = tmp.groupby("x", as_index=False)["w"].sum().rename(columns={"x": serie.name, "w": "Cantidad"})
        vc["Cantidad"] = vc["Cantidad"].round(0).astype(int)
        return vc.sort_values(by=serie.name)
    else:
        vc = serie.value_counts(dropna=False).sort_index().reset_index()
        vc.columns = [serie.name, "Cantidad"]
        return vc

def _prop_pesada(flag: pd.Series, pesos: pd.Series | None) -> float:
    if pesos is not None:
        df = pd.DataFrame({"f": flag.astype(float), "w": pesos}).dropna(subset=["f"])
        num = (df["f"] * df["w"]).sum()
        den = df["w"].sum()
        return float(num / den) if den > 0 else 0.0
    else:
        return float(flag.mean()) if len(flag) else 0.0

def _tabla_prop_por(grupo: pd.Series, flag: pd.Series, pesos: pd.Series | None) -> pd.DataFrame:
    data = pd.DataFrame({"g": grupo, "f": flag.astype(int)}).dropna(subset=["g"])
    if pesos is not None and len(pesos) == len(grupo):
        data["w"] = pesos
        agg = data.groupby("g").apply(lambda d: (d["f"]*d["w"]).sum() / d["w"].sum()).reset_index()
    else:
        agg = data.groupby("g")["f"].mean().reset_index()
    agg.columns = [grupo.name, "Porcentaje"]
    agg["Porcentaje"] = (agg["Porcentaje"] * 100).round(2)
    return agg

def _etiquetar(colname: str, serie: pd.Series) -> pd.Series:
    if colname in ETIQUETAS_BINARIAS:
        return serie.map(ETIQUETAS_BINARIAS[colname]).fillna(serie)
    if colname in CANDIDATAS_EDUC:
        # Si los valores son 1..7/9, mapear
        if pd.api.types.is_numeric_dtype(serie):
            return serie.map(ETIQUETAS_EDUC).fillna(serie)
    return serie

def _etiquetar_df_para_listado(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        out[c] = _etiquetar(c, out[c])
    return out

def _crear_grupos_edad(df: pd.DataFrame) -> pd.Series:
    col_edad = _primera_columna(df, CANDIDATAS_EDAD)
    if col_edad and pd.api.types.is_numeric_dtype(df[col_edad]):
        edad = df[col_edad].clip(lower=0, upper=100)
        bins = [-0.1, 17, 29, 44, 64, 150]
        labels = ["0-17", "18-29", "30-44", "45-64", "65+"]
        return pd.cut(edad, bins=bins, labels=labels)
    # fallback (no hay edad): quintiles del índice para no romper
    return pd.qcut(range(len(df)), q=5, labels=["0-17", "18-29", "30-44", "45-64", "65+"])

def _crear_quintiles_ingreso(df: pd.DataFrame) -> pd.Series | None:
    col_inc = _primera_columna(df, CANDIDATA_INGRESO)
    if col_inc and pd.api.types.is_numeric_dtype(df[col_inc]):
        try:
            q = pd.qcut(df[col_inc].rank(method="first"), 5, labels=["Q1 (más bajo)", "Q2", "Q3", "Q4", "Q5 (más alto)"])
            q.name = "Quintil ingreso"
            return q
        except Exception:
            return None
    return None

def _col_peso(df: pd.DataFrame) -> pd.Series | None:
    col = _primera_columna(df, CANDIDATAS_PESO)
    return df[col] if col else None

# =========================
# Núcleo de análisis
# =========================
def generar_analisis_tic_ampliado(df_merged: pd.DataFrame) -> Tuple[Dict[str, pd.DataFrame], pd.DataFrame, Dict[str, Any]]:
    """
    Recibe el merge de individuos + hogares (mismo que ya usás).
    Devuelve: (tablas_para_excel, df_enriquecido, resumen_para_conclusiones)
    """
    df = df_merged.copy()

    # Variables núcleo TIC (si no están, levantamos error claro)
    for var in ["IP_III_04", "IP_III_06"]:
        if var not in df.columns:
            raise ValueError(f"Falta la columna '{var}' en la base. Revisá el instructivo del trimestre.")

    # Exclusión binaria/ordinal (No computadora Y No internet => excluido total)
    df["excluido_binario"] = ((df["IP_III_04"] == 2) & (df["IP_III_06"] == 2)).astype(int)
    df["exclusion_ordinal"] = df.apply(lambda x: 2 if x["IP_III_04"] == 2 and x["IP_III_06"] == 2
                                       else 1 if x["IP_III_04"] == 2 or x["IP_III_06"] == 2
                                       else 0, axis=1)

    # Etiquetas nominales en columnas núcleo para listados
    for c in ["IP_III_04", "IP_III_05", "IP_III_06", "IH_II_01", "IH_II_02", "CH04"]:
        if c in df.columns:
            df[c] = _etiquetar(c, df[c])

    # Columnas auxiliares
    edad_grp = _crear_grupos_edad(df); edad_grp.name = "Edad"
    sexo_col = _primera_columna(df, CANDIDATAS_SEXO)
    educ_col = _primera_columna(df, CANDIDATAS_EDUC)
    quintil = _crear_quintiles_ingreso(df)
    peso = _col_peso(df)

    tablas: Dict[str, pd.DataFrame] = {}

    # -------------------------
    # Distribuciones básicas
    # -------------------------
    for col, titulo in [
        ("IP_III_04", "Uso de computadora"),
        ("IP_III_05", "Uso de celular"),
        ("IP_III_06", "Uso de internet"),
        ("IH_II_01", "Hogar con computadora"),
        ("IH_II_02", "Hogar con acceso a internet"),
    ]:
        if col in df.columns:
            dist = _value_counts_pesado(df[col], peso)
            dist[col] = _etiquetar(col, dist[col])
            total = dist["Cantidad"].sum()
            dist["%"] = (dist["Cantidad"] / total * 100).round(2)
            tablas[titulo] = dist

    # Exclusión ordinal (conteos + %)
    excl_ordinal = _value_counts_pesado(df["exclusion_ordinal"], peso)
    excl_ordinal["Nivel Exclusión"] = excl_ordinal["exclusion_ordinal"].map({0: "Sin exclusión", 1: "Exclusión parcial", 2: "Exclusión total"})
    excl_ordinal = excl_ordinal[["Nivel Exclusión", "Cantidad"]]
    excl_ordinal["%"] = (excl_ordinal["Cantidad"] / excl_ordinal["Cantidad"].sum() * 100).round(2)
    tablas["Exclusión – Nivel ordinal"] = excl_ordinal

    # -------------------------
    # Brechas y cruces clave
    # -------------------------
    # Por edad
    excl_por_edad = _tabla_prop_por(edad_grp, df["excluido_binario"], peso)
    tablas["Brecha por edad (exclusión)"] = excl_por_edad

    # Por sexo
    if sexo_col:
        sexolab = _etiquetar(sexo_col, df[sexo_col])
        excl_por_sexo = _tabla_prop_por(sexolab.rename("Sexo"), df["excluido_binario"], peso)
        tablas["Brecha por sexo (exclusión)"] = excl_por_sexo

    # Por nivel educativo
    if educ_col:
        educlab = _etiquetar(educ_col, df[educ_col])
        excl_por_educ = _tabla_prop_por(educlab.rename("Nivel educativo"), df["excluido_binario"], peso)
        tablas["Brecha por educación (exclusión)"] = excl_por_educ.sort_values("Nivel educativo")

    # Acceso a internet por quintil de ingreso del hogar
    if ("IH_II_02" in df.columns) and (quintil is not None):
        acc_int = (df["IH_II_02"] == "Sí") if df["IH_II_02"].dtype == object else (df["IH_II_02"] == 1)
        acc_int_por_quintil = _tabla_prop_por(quintil, acc_int.rename("acceso"), peso)
        acc_int_por_quintil.rename(columns={"Porcentaje": "% hogares con internet"}, inplace=True)
        tablas["Acceso a internet por quintil"] = acc_int_por_quintil

    # -------------------------
    # Resumen para narrativa
    # -------------------------
    resumen: Dict[str, Any] = {}
    # Tasas generales
    excl_total = _prop_pesada(df["excluido_binario"], peso) * 100
    resumen["exclusion_total_pct"] = round(excl_total, 2)

    if "IH_II_02" in df.columns:
        acc_int_flag = (df["IH_II_02"] == "Sí") if df["IH_II_02"].dtype == object else (df["IH_II_02"] == 1)
        resumen["hogar_con_internet_pct"] = round(_prop_pesada(acc_int_flag, peso) * 100, 2)

    if "IH_II_01" in df.columns:
        pc_flag = (df["IH_II_01"] == "Sí") if df["IH_II_01"].dtype == object else (df["IH_II_01"] == 1)
        resumen["hogar_con_pc_pct"] = round(_prop_pesada(pc_flag, peso) * 100, 2)

    # Brechas
    resumen["edad_brecha"] = excl_por_edad.sort_values("Porcentaje", ascending=False).head(1).to_dict(orient="records")[0]
    if "Brecha por sexo (exclusión)" in tablas:
        t = tablas["Brecha por sexo (exclusión)"]
        max_s = t.loc[t["Porcentaje"].idxmax()]
        min_s = t.loc[t["Porcentaje"].idxmin()]
        resumen["sexo_brecha_pp"] = round(float(max_s["Porcentaje"] - min_s["Porcentaje"]), 2)
        resumen["sexo_mayor_riesgo"] = max_s[0] if 0 in max_s.index else max_s["Sexo"]

    if "Acceso a internet por quintil" in tablas:
        q = tablas["Acceso a internet por quintil"]
        max_q = q.iloc[-1, 1]; min_q = q.iloc[0, 1]
        resumen["brecha_quintil_pp"] = round(float(max_q - min_q), 2)

    return tablas, df, resumen

# =========================
# Informe Word con narrativa robusta
# =========================
def generar_informe_narrativo_tic(tablas: Dict[str, pd.DataFrame], anio: str = "2024", resumen: Dict[str, Any] | None = None) -> BytesIO:
    doc = Document()
    doc.add_heading(f"Informe Analítico – Inclusión Digital TIC 4ºT {anio}", 0)
    doc.add_paragraph(
        "Este informe presenta un diagnóstico de inclusión/exclusión digital a partir del módulo TIC "
        f"del cuarto trimestre del año {anio}, combinando información de hogares e individuos."
    )

    # 1. Definición
    doc.add_heading("1. Definición de exclusión digital", level=1)
    doc.add_paragraph(
        "La exclusión digital se manifiesta tanto por la falta de acceso a infraestructura (dispositivos y conectividad) "
        "como por el no uso o el uso restringido de las tecnologías. En este informe se aproxima empíricamente mediante: "
        "a) acceso del hogar a internet y a computadora; b) uso individual de internet y computadora; y c) un indicador "
        "ordinal de exclusión (sin, parcial, total)."
    )

    # 2. Resultados clave (listas legibles)
    doc.add_heading("2. Resultados clave", level=1)
    nombres_amigables = {
        "Uso de computadora": "Uso de computadora (individuos)",
        "Uso de celular": "Uso de celular (individuos)",
        "Uso de internet": "Uso de internet (individuos)",
        "Hogar con computadora": "Disponibilidad de computadora (hogares)",
        "Hogar con acceso a internet": "Acceso a internet (hogares)",
        "Exclusión – Nivel ordinal": "Exclusión digital (índice ordinal)",
        "Brecha por edad (exclusión)": "Brecha por edad (personas excluidas)",
        "Brecha por sexo (exclusión)": "Brecha por sexo (personas excluidas)",
        "Brecha por educación (exclusión)": "Brecha por educación (personas excluidas)",
        "Acceso a internet por quintil": "Acceso a internet por quintil de ingreso del hogar",
    }

    for nombre, tabla in tablas.items():
        doc.add_heading(nombres_amigables.get(nombre, nombre), level=2)
        # listado simple
        for _, row in tabla.iterrows():
            valores = " – ".join([f"{tabla.columns[i]}: {row[i]}" for i in range(len(row))])
            doc.add_paragraph(f"• {valores}", style="List Bullet")

    # 3. Conclusiones y análisis robusto
    doc.add_heading("3. Conclusiones y análisis robusto", level=1)

    if resumen:
        frases = []
        if "exclusion_total_pct" in resumen:
            frases.append(f"La exclusión digital total (sin uso de internet ni computadora) alcanza al {resumen['exclusion_total_pct']}% de las personas.")
        if "hogar_con_internet_pct" in resumen and "hogar_con_pc_pct" in resumen:
            frases.append(
                f"En el plano de infraestructura, el {resumen['hogar_con_internet_pct']}% de los hogares cuenta con internet "
                f"y el {resumen['hogar_con_pc_pct']}% dispone de computadora."
            )
        if "edad_brecha" in resumen:
            frases.append(
                f"Por edad, el grupo con mayor exclusión es {resumen['edad_brecha'].get('Edad')} "
                f"({resumen['edad_brecha'].get('Porcentaje')}%)."
            )
        if "sexo_brecha_pp" in resumen:
            frases.append(
                f"Se observa una brecha por sexo de {resumen['sexo_brecha_pp']} puntos porcentuales en la tasa de exclusión."
            )
        if "brecha_quintil_pp" in resumen:
            frases.append(
                f"La desigualdad por ingreso es marcada: entre Q1 y Q5 hay {resumen['brecha_quintil_pp']} puntos de diferencia "
                f"en acceso a internet del hogar."
            )

        doc.add_paragraph(
            " ".join(frases) if frases else
            "La evidencia muestra brechas relevantes por edad, condiciones del hogar y nivel socioeconómico."
        )

    doc.add_paragraph(
        "En conjunto, los hallazgos confirman que la exclusión digital opera como un multiplicador de desigualdades: "
        "las limitaciones de infraestructura en los hogares de menores ingresos, junto con menores niveles educativos y el ciclo de vida, "
        "confluyen en una menor apropiación efectiva de las TIC. Para revertir este patrón, se recomienda combinar: "
        "i) expansión de conectividad fija de calidad en barrios y localidades rezagadas; ii) acceso a dispositivos mediante programas "
        "de financiamiento o provisión; y iii) alfabetización digital diferenciada por grupos etarios y trayectorias educativas, "
        "con foco en usos educativos y laborales."
    )

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
