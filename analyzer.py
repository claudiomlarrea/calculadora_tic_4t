
import pandas as pd
from docx import Document
from io import BytesIO
from typing import Dict, Tuple, Any, Optional

# =========================
# Etiquetas y utilidades
# =========================
ETIQUETAS_BINARIAS = {
    "IP_III_04": {1: "Sí", 2: "No", 9: "Ns/Nc"},  # Uso de computadora
    "IP_III_05": {1: "Sí", 2: "No", 9: "Ns/Nc"},  # Uso de celular
    "IP_III_06": {1: "Sí", 2: "No", 9: "Ns/Nc"},  # Uso de internet
    "IH_II_01": {1: "Sí", 2: "No", 9: "Ns/Nc"},  # Hogar con computadora
    "IH_II_02": {1: "Sí", 2: "No", 9: "Ns/Nc"},  # Hogar con acceso a internet
    "CH04":     {1: "Varón", 2: "Mujer"},        # Sexo
}

# Extiende mapeos aceptando claves como strings ("1","2","9") y normalizando
def _extender_mapa(m: Dict[int, str]) -> Dict[str, str]:
    out = {}
    for k, v in m.items():
        out[k] = v  # deja int por compatibilidad
        out[str(k)] = v  # agrega versión string
    # variantes comunes de texto libre
    out.update({"SI": "Sí", "Si": "Sí", "si": "Sí", "NO": "No", "no": "No"})
    return out

ETIQUETAS_BINARIAS_STR = {var: _extender_mapa(m) for var, m in ETIQUETAS_BINARIAS.items()}

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
ETIQUETAS_EDUC_STR = _extender_mapa(ETIQUETAS_EDUC)

CANDIDATAS_EDAD = ["CH06", "EDAD", "edad"]
CANDIDATAS_SEXO = ["CH04", "SEXO"]
CANDIDATAS_EDUC = ["NIVEL_ED", "NIVEL_EDUC", "NIVEL_EDUCATIVO"]
CANDIDATA_INGRESO = ["ITF", "ITF_HOGAR", "ingreso_total_hogar"]
CANDIDATAS_PESO = ["PONDERA", "PONDIIO", "PONDIH", "PESO", "FACTOR", "FACTOR_EXP"]

def _primera(df: pd.DataFrame, cols) -> Optional[str]:
    for c in cols:
        if c in df.columns:
            return c
    return None

def _peso(df: pd.DataFrame) -> Optional[pd.Series]:
    c = _primera(df, CANDIDATAS_PESO)
    return df[c] if c else None

def _edad_grupos(df: pd.DataFrame) -> pd.Series:
    c = _primera(df, CANDIDATAS_EDAD)
    if c and pd.api.types.is_numeric_dtype(df[c]):
        e = df[c].clip(lower=0, upper=100)
        bins = [-0.1, 17, 29, 44, 64, 150]
        labels = ["0-17", "18-29", "30-44", "45-64", "65+"]
        return pd.cut(e, bins=bins, labels=labels)
    # fallback si no hay edad numérica
    return pd.qcut(range(len(df)), 5, labels=["0-17", "18-29", "30-44", "45-64", "65+"])

def _quintiles_ingreso(df: pd.DataFrame) -> Optional[pd.Series]:
    c = _primera(df, CANDIDATA_INGRESO)
    if not c or not pd.api.types.is_numeric_dtype(df[c]):
        return None
    try:
        q = pd.qcut(df[c].rank(method="first"), 5, labels=["Q1 (más bajo)", "Q2", "Q3", "Q4", "Q5 (más alto)"])
        q.name = "Quintil ingreso"
        return q
    except Exception:
        return None

def _map_nominal(nombre: str, s: pd.Series) -> pd.Series:
    # convierte todo a string primero para evitar mezclas int/str
    if pd.api.types.is_categorical_dtype(s):
        s = s.astype("string")
    else:
        s = s.astype("string").str.strip()
    if nombre in ETIQUETAS_BINARIAS_STR:
        m = ETIQUETAS_BINARIAS_STR[nombre]
        return s.map(m).fillna(s)
    if nombre in CANDIDATAS_EDUC:
        return s.map(ETIQUETAS_EDUC_STR).fillna(s)
    if nombre == "CH04":
        # sexo si llega como 1/2
        return s.map(_extender_mapa({1:"Varón",2:"Mujer"})).fillna(s)
    return s

def _value_counts_w(s: pd.Series, w: Optional[pd.Series]) -> pd.DataFrame:
    # homogeneiza como texto para evitar TypeError al ordenar
    s = s.astype("string")
    if w is not None and len(w) == len(s):
        tmp = pd.DataFrame({"x": s, "w": w}).dropna(subset=["x"])
        out = tmp.groupby("x", as_index=False)["w"].sum().rename(columns={"x": s.name, "w": "Cantidad"})
        out["Cantidad"] = out["Cantidad"].round(0).astype(int)
    else:
        out = s.value_counts(dropna=False).reset_index()
        out.columns = [s.name, "Cantidad"]
    # ordenar siempre por representación de texto
    out = out.sort_values(by=s.name, key=lambda col: col.astype(str))
    return out

def _prop_w(flag: pd.Series, w: Optional[pd.Series]) -> float:
    # flag puede venir como objeto: aceptar "1","Sí","True"
    f = flag.astype("string").str.strip().str.lower().isin(["1","true","sí","si","yes"])
    if w is not None and len(w) == len(flag):
        df = pd.DataFrame({"f": f.astype(float), "w": w}).dropna(subset=["f"])
        num = (df["f"] * df["w"]).sum()
        den = df["w"].sum()
        return float(num/den) if den > 0 else 0.0
    return float(f.mean()) if len(f) else 0.0

def _tabla_prop_por(g: pd.Series, flag_bool: pd.Series, w: Optional[pd.Series]) -> pd.DataFrame:
    g = g.astype("string")
    f = flag_bool.astype(int) if flag_bool.dtype != int else flag_bool
    data = pd.DataFrame({"g": g, "f": f})
    if w is not None and len(w) == len(g):
        data["w"] = w
        agg = data.groupby("g").apply(lambda d: (d["f"]*d["w"]).sum() / d["w"].sum()).reset_index()
    else:
        agg = data.groupby("g")["f"].mean().reset_index()
    agg.columns = [g.name or "Grupo", "Porcentaje"]
    agg["Porcentaje"] = (agg["Porcentaje"] * 100).round(2)
    agg = agg.sort_values(by=agg.columns[0], key=lambda col: col.astype(str))
    return agg

# =====================================
# Núcleo de análisis
# =====================================
def generar_analisis_tic_ampliado(df_merged: pd.DataFrame) -> Tuple[Dict[str, pd.DataFrame], pd.DataFrame, Dict[str, Any]]:
    df = df_merged.copy()

    # Validación mínima
    faltantes = [v for v in ["IP_III_04", "IP_III_06"] if v not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas clave: {faltantes}. Revisá el instructivo del trimestre.")

    # Normaliza columnas núcleo a texto y mapea
    for c in ["IP_III_04","IP_III_05","IP_III_06","IH_II_01","IH_II_02","CH04"]:
        if c in df.columns:
            df[c] = _map_nominal(c, df[c])

    # Indicadores de exclusión (compara como texto)
    df["excluido_binario"] = ((df.get("IP_III_04","").astype(str) == "No") & (df.get("IP_III_06","").astype(str) == "No")).astype(int)
    df["exclusion_ordinal"] = df.apply(
        lambda x: 2 if str(x.get("IP_III_04","")) == "No" and str(x.get("IP_III_06","")) == "No"
        else 1 if (str(x.get("IP_III_04","")) == "No") or (str(x.get("IP_III_06","")) == "No")
        else 0, axis=1
    )

    # Columnas auxiliares
    edad = _edad_grupos(df); edad.name = "Edad"
    sexo_col = _primera(df, CANDIDATAS_SEXO) or "CH04" if "CH04" in df.columns else None
    educ_col = _primera(df, CANDIDATAS_EDUC)
    quintil = _quintiles_ingreso(df)
    w = _peso(df)

    tablas: Dict[str, pd.DataFrame] = {}

    # Distribuciones básicas
    for col, titulo in [
        ("IP_III_04","Uso de computadora"),
        ("IP_III_05","Uso de celular"),
        ("IP_III_06","Uso de internet"),
        ("IH_II_01","Hogar con computadora"),
        ("IH_II_02","Hogar con acceso a internet"),
    ]:
        if col in df.columns:
            dist = _value_counts_w(df[col], w)
            total = dist["Cantidad"].sum()
            dist["%"] = (dist["Cantidad"] / total * 100).round(2)
            tablas[titulo] = dist

    # Exclusión ordinal
    excl_ordinal = _value_counts_w(df["exclusion_ordinal"], w)
    excl_ordinal["Nivel"] = excl_ordinal["exclusion_ordinal"].map({0:"Sin exclusión",1:"Exclusión parcial",2:"Exclusión total"})
    excl_ordinal = excl_ordinal[["Nivel","Cantidad"]]
    excl_ordinal["%"] = (excl_ordinal["Cantidad"]/excl_ordinal["Cantidad"].sum()*100).round(2)
    tablas["Exclusión – Nivel ordinal"] = excl_ordinal

    # Brechas por edad
    tablas["Brecha por edad (exclusión)"] = _tabla_prop_por(edad, df["excluido_binario"], w)

    # Sexo (si está)
    if sexo_col and sexo_col in df.columns:
        sexolab = _map_nominal(sexo_col, df[sexo_col]).rename("Sexo")
        tablas["Brecha por sexo (exclusión)"] = _tabla_prop_por(sexolab, df["excluido_binario"], w)

    # Educación (si está)
    if educ_col and educ_col in df.columns:
        educlab = _map_nominal(educ_col, df[educ_col]).rename("Nivel educativo")
        tablas["Brecha por educación (exclusión)"] = _tabla_prop_por(educlab, df["excluido_binario"], w)

    # Acceso a internet por quintil (si hay ingreso)
    if ("IH_II_02" in df.columns) and (quintil is not None):
        acc_int = df["IH_II_02"].astype(str) == "Sí"
        acc_quintil = _tabla_prop_por(quintil, acc_int.rename("acceso"), w)
        acc_quintil.rename(columns={"Porcentaje":"% hogares con internet"}, inplace=True)
        tablas["Acceso a internet por quintil"] = acc_quintil

    # Resumen para narrativa
    resumen: Dict[str, Any] = {}
    resumen["exclusion_total_pct"] = round((df["excluido_binario"].mean() * 100), 2)
    if "IH_II_01" in df.columns:
        resumen["hogar_con_pc_pct"] = round(((df["IH_II_01"].astype(str) == "Sí").mean() * 100), 2)
    if "IH_II_02" in df.columns:
        resumen["hogar_con_internet_pct"] = round(((df["IH_II_02"].astype(str) == "Sí").mean() * 100), 2)

    # Brechas p.p.
    b_edad = tablas["Brecha por edad (exclusión)"]
    resumen["edad_peak"] = b_edad.loc[b_edad["Porcentaje"].idxmax()].to_dict()
    resumen["edad_valley"] = b_edad.loc[b_edad["Porcentaje"].idxmin()].to_dict()
    resumen["edad_brecha_pp"] = round(float(resumen["edad_peak"]["Porcentaje"] - resumen["edad_valley"]["Porcentaje"]), 2)

    if "Brecha por sexo (exclusión)" in tablas:
        t = tablas["Brecha por sexo (exclusión)"]
        pk = t.loc[t["Porcentaje"].idxmax()]; vl = t.loc[t["Porcentaje"].idxmin()]
        resumen["sexo_brecha_pp"] = round(float(pk["Porcentaje"] - vl["Porcentaje"]), 2)
        resumen["sexo_peak"] = pk.to_dict()

    if "Acceso a internet por quintil" in tablas:
        q = tablas["Acceso a internet por quintil"]
        resumen["quintil_brecha_pp"] = round(float(q.iloc[-1,1] - q.iloc[0,1]), 2)
        resumen["quintil_low"] = q.iloc[0].to_dict()
        resumen["quintil_high"] = q.iloc[-1].to_dict()

    return tablas, df, resumen

# =====================================
# Informe Word (narrativa robusta)
# =====================================
def _p(doc: Document, text: str):
    if text:
        doc.add_paragraph(text)

def generar_informe_narrativo_tic(tablas: Dict[str, pd.DataFrame], anio: str = "2024", resumen: Optional[Dict[str, Any]] = None) -> BytesIO:
    doc = Document()
    doc.add_heading(f"Informe Analítico – Inclusión Digital TIC 4ºT {anio}", 0)
    _p(doc, "Diagnóstico de inclusión/exclusión digital del cuarto trimestre, combinando información de hogares e individuos.")

    doc.add_heading("1. Definición y enfoque", level=1)
    _p(doc, "Se operacionaliza a partir de: (i) infraestructura del hogar (internet y computadora) y (ii) uso individual de internet y computadora. Se construye un índice ordinal (sin, parcial, total).")

    doc.add_heading("2. Resultados clave", level=1)
    nombres = {
        "Uso de computadora": "Uso de computadora (individuos)",
        "Uso de celular": "Uso de celular (individuos)",
        "Uso de internet": "Uso de internet (individuos)",
        "Hogar con computadora": "Disponibilidad de computadora (hogares)",
        "Hogar con acceso a internet": "Acceso a internet (hogares)",
        "Exclusión – Nivel ordinal": "Exclusión digital (índice ordinal)",
        "Brecha por edad (exclusión)": "Brecha por edad (exclusión)",
        "Brecha por sexo (exclusión)": "Brecha por sexo (exclusión)",
        "Brecha por educación (exclusión)": "Brecha por educación (exclusión)",
        "Acceso a internet por quintil": "Acceso a internet por quintil de ingreso",
    }
    for nombre, tabla in tablas.items():
        doc.add_heading(nombres.get(nombre, nombre), level=2)
        for _, row in tabla.iterrows():
            vals = [f"{col}: {row[col]}" for col in tabla.columns]
            _p(doc, "• " + " – ".join(vals))

    doc.add_heading("3. Conclusiones y análisis extendido", level=1)
    if resumen:
        frases = []
        if "exclusion_total_pct" in resumen:
            frases.append(f"La exclusión digital total (sin uso de internet ni computadora) alcanza al {resumen['exclusion_total_pct']}% de las personas.")
        if "hogar_con_internet_pct" in resumen and "hogar_con_pc_pct" in resumen:
            frases.append(f"En infraestructura, el {resumen['hogar_con_internet_pct']}% de los hogares dispone de internet y el {resumen['hogar_con_pc_pct']}% cuenta con computadora.")
        if "edad_brecha_pp" in resumen:
            frases.append(f"Por edad, la brecha entre el grupo más afectado ({resumen['edad_peak'].get('Edad','')}) y el menos afectado ({resumen['edad_valley'].get('Edad','')}) es de {resumen['edad_brecha_pp']} p.p.")
        if "sexo_brecha_pp" in resumen:
            frases.append(f"La brecha por sexo alcanza {resumen['sexo_brecha_pp']} p.p.")
        if "quintil_brecha_pp" in resumen:
            frases.append(f"Por ingreso, la diferencia entre Q1 y Q5 en acceso a internet del hogar es de {resumen['quintil_brecha_pp']} p.p.")
        _p(doc, " ".join(frases))

    _p(doc, "La evidencia confirma que la exclusión digital amplifica desigualdades socioeducativas. Es necesario integrar conectividad de calidad, acceso a dispositivos y formación orientada a usos significativos (educación, empleo y trámites).")

    doc.add_heading("4. Recomendaciones", level=1)
    _p(doc, "• Expandir conectividad fija y planes de asequibilidad en territorios rezagados.")
    _p(doc, "• Programas de acceso a computadoras para hogares de menores ingresos con estudiantes.")
    _p(doc, "• Alfabetización digital segmentada por edad y nivel educativo, con tutorías prácticas.")
    _p(doc, "• Monitoreo anual de exclusión parcial/total y brechas por subgrupos.")

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


