
import io
import re
import tempfile
from pathlib import Path
from urllib.parse import quote_plus, urlparse

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import load_workbook

st.set_page_config(
    page_title="Assessment + Plan de Carrera Técnico",
    page_icon="📈",
    layout="wide",
)

QUESTION_GROUPS = {
    "Potencial y proyección": [
        "¿Cuál es el potencial real del evaluado para asumir funciones técnicas de mayor complejidad?",
        "¿En qué áreas muy específicas o exigentes podría ser formado con mayor probabilidad de éxito?",
        "¿Qué combinación de fortalezas actuales le da más proyección a medio plazo?",
        "¿Qué evidencias indican que puede crecer por encima del estándar actual del equipo?",
        "¿Qué nivel de ambición técnica parece realista en su siguiente etapa de desarrollo?",
    ],
    "Fortalezas y especialización": [
        "Además de las áreas ya propuestas, ¿qué otras áreas tienen mayor probabilidad de éxito para especialización?",
        "¿Qué fortalezas diferenciales conviene consolidar antes de ampliar su ámbito de especialización?",
        "¿En qué tronco o conjunto de indicadores destaca de forma más consistente?",
        "¿Qué fortalezas se pueden convertir en una línea de excelencia técnica?",
        "¿Qué áreas merecen un plan de especialización avanzada y no solo de mantenimiento?",
    ],
    "Brechas y riesgos": [
        "¿Cuáles son las brechas que más limitan hoy su desarrollo profesional?",
        "¿Qué debilidades tienen mayor impacto operativo en su desempeño actual?",
        "¿Qué área débil conviene corregir primero para evitar efecto arrastre sobre otras competencias?",
        "¿Dónde aparece el mayor gap frente a referencia y frente al benchmark del equipo?",
        "¿Qué riesgos tendría forzar una especialización sin cerrar antes ciertas bases?",
    ],
    "Plan de acción y formación": [
        "¿Qué plan de desarrollo de 6 meses tendría más sentido para este perfil?",
        "¿Qué plan de desarrollo de 12 meses tendría más sentido para este perfil?",
        "¿Qué formación base debería completar antes de abordar contenidos avanzados?",
        "¿Qué combinación de formación interna y externa parece más eficiente para este caso?",
        "¿Qué indicadores deberían revisarse de nuevo en la próxima evaluación para confirmar progreso?",
    ],
}
QUESTION_TO_GROUP = {q: g for g, qs in QUESTION_GROUPS.items() for q in qs}

PROFILE_THRESHOLDS = {
    "Básico": 2.0,
    "Supera": 3.0,
    "Certificado": 4.0,
    "Excelente": 5.0,
    "Máster": 6.1,
}

AREA_KEYWORDS = {
    "Nutrición": ["nutricion", "nutrición", "producto", "aditivo", "agua", "alimentacion", "alimentación", "cebo", "racion", "ración", "pienso", "forraje", "ruminal", "ingredient", "formulacion"],
    "Patología": ["patologia", "patología", "metabolica", "metabólica", "infecciosa", "parasitaria", "antibioterapia", "diagnost", "inmunidad", "salud", "microbioma", "bioseguridad", "sanidad"],
    "Manejo": ["bioseguridad", "instalacion", "instalación", "ventilacion", "ventilación", "arranque", "produccion", "producción", "reposicion", "reposición", "manejo", "granja", "housing", "ordeño", "reproduccion", "reproducción", "bienestar"],
    "Herramientas": ["datos", "estadistica", "estadística", "crm", "power bi", "bbdd", "digitalizacion", "digitalización", "informe", "ingles", "inglés", "software", "programa", "kpi", "analytics", "robot", "automat"],
}

DEFAULT_CATALOG_PATTERNS = [
    "*Formacion*.xlsx",
    "*formacion*.xlsx",
    "*Especializacion*.xlsx",
    "*especializacion*.xlsx",
    "*Resumen*.xlsx",
    "*resumen*.xlsx",
]
SEARCH_HEADERS = {"User-Agent": "Mozilla/5.0"}

def safe_float(value):
    if value in (None, ""):
        return None
    try:
        return float(value)
    except Exception:
        return None

def pct(value):
    value = safe_float(value)
    if value is None:
        return None
    return round(value * 100.0, 1)

def format_pct(value):
    p = pct(value)
    return "No disponible" if p is None else f"{p:.1f}%"

def normalize_text(text):
    text = (text or "").strip().lower()
    replacements = {"á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ü": "u", "ñ": "n"}
    for src, dst in replacements.items():
        text = text.replace(src, dst)
    text = re.sub(r"\s+", " ", text)
    return text

def infer_area_from_text(text):
    text_n = normalize_text(text)
    scores = {area: sum(1 for kw in kws if kw in text_n) for area, kws in AREA_KEYWORDS.items()}
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "General"

def infer_species_subspecies_from_name(name):
    n = normalize_text(Path(str(name)).stem.replace("_", " ").replace("-", " "))
    species = "General"
    subspecies = "General"

    if "avicultura" in n:
        species = "Avicultura"
        if "puesta" in n:
            subspecies = "Puesta"
        elif "broiler" in n or "carne" in n:
            subspecies = "Carne"
    elif "porcino" in n:
        species = "Porcino"
        if "industrial" in n:
            subspecies = "Industrial"
    elif "vacuno" in n and "lech" in n:
        species = "Vacuno Lechero"
        if "intensivo" in n:
            subspecies = "Intensivo"
    elif "vacuno" in n and "carne" in n:
        species = "Vacuno Carne"
        if "intensivo" in n:
            subspecies = "Intensivo"
    elif "ovino" in n or "caprino" in n or "pprr" in n:
        species = "Ovino-Caprino"
    elif "rumiante" in n:
        species = "Rumiantes"
    return species, subspecies

def species_match(selected_species, row_species):
    s = normalize_text(selected_species)
    r = normalize_text(row_species)
    if not s or not r:
        return False
    return s in r or r in s

def area_match(target_area, row_area, row_text=""):
    t = normalize_text(target_area)
    r = normalize_text(row_area)
    if t == r:
        return True
    text = normalize_text(row_text)
    return any(kw in text for kw in AREA_KEYWORDS.get(target_area, []))

def detect_header_row(ws, max_scan=8):
    for r in range(1, min(ws.max_row, max_scan) + 1):
        vals = [normalize_text(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 12) + 1)]
        joined = " | ".join([v for v in vals if v])
        if "nombre del programa" in joined or "especializacion" in joined or "especialización" in joined:
            return r, vals
    return None, []

def map_headers(headers):
    mapping = {}
    for idx, header in enumerate(headers, start=1):
        h = normalize_text(header)
        if "especie" == h:
            mapping["species"] = idx
        elif "subespecie" in h:
            mapping["subspecies"] = idx
        elif "tipo de formacion" in h or "tipo de formación" in h:
            mapping["type"] = idx
        elif "nombre del programa" in h or "especializacion" in h or "especialización" in h:
            mapping["program"] = idx
        elif "duracion" in h or "duración" in h:
            mapping["duration"] = idx
        elif "modalidad" in h:
            mapping["modality"] = idx
        elif "ubicacion" in h or "ubicación" in h:
            mapping["location"] = idx
        elif "institucion" in h or "institución" in h or "centros de formacion" in h or "centros de formación" in h:
            mapping["institution"] = idx
        elif "contenido" in h:
            mapping["content"] = idx
        elif "enlace" in h or "contacto" in h or "web" in h:
            mapping["link"] = idx
    return mapping

def parse_catalog_sheet(ws, source_name):
    header_row, headers = detect_header_row(ws)
    if not header_row:
        return []
    mapping = map_headers(headers)
    if "program" not in mapping:
        return []

    default_species, default_subspecies = infer_species_subspecies_from_name(source_name)
    rows = []
    current_species = default_species

    for r in range(header_row + 1, ws.max_row + 1):
        values = {key: ws.cell(r, col).value for key, col in mapping.items()}
        if all(v in (None, "") for v in values.values()):
            continue

        species_value = values.get("species")
        if species_value not in (None, ""):
            current_species = str(species_value).strip()

        program = values.get("program")
        if program in (None, ""):
            continue

        row_species = current_species if "species" in mapping else default_species
        row_subspecies = str(values.get("subspecies")).strip() if values.get("subspecies") not in (None, "") else default_subspecies
        content = str(values.get("content")).strip() if values.get("content") not in (None, "") else ""
        institution = str(values.get("institution")).strip() if values.get("institution") not in (None, "") else "No disponible"
        row_type = str(values.get("type")).strip() if values.get("type") not in (None, "") else "Curso"
        link = str(values.get("link")).strip() if values.get("link") not in (None, "") else "No disponible"
        program_text = str(program).strip()

        area = infer_area_from_text(" ".join([program_text, content, institution, str(row_species)]))
        rows.append({
            "species": row_species,
            "subspecies": row_subspecies,
            "type": row_type,
            "program": program_text,
            "duration": str(values.get("duration")).strip() if values.get("duration") not in (None, "") else "No disponible",
            "modality": str(values.get("modality")).strip() if values.get("modality") not in (None, "") else "No disponible",
            "location": str(values.get("location")).strip() if values.get("location") not in (None, "") else "No disponible",
            "institution": institution,
            "content": content,
            "link": link,
            "area": area,
            "source_file": source_name,
        })
    return rows

@st.cache_data(show_spinner=False)
def load_catalog_from_path(path_str):
    path = Path(path_str)
    wb = load_workbook(path, data_only=True)
    rows = []
    for sheet_name in wb.sheetnames:
        rows.extend(parse_catalog_sheet(wb[sheet_name], path.name))
    return pd.DataFrame(rows)

def parse_uploaded_catalog(uploaded_file):
    suffix = Path(uploaded_file.name).suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    wb = load_workbook(tmp_path, data_only=True)
    rows = []
    for sheet_name in wb.sheetnames:
        rows.extend(parse_catalog_sheet(wb[sheet_name], uploaded_file.name))
    return pd.DataFrame(rows)

def autodiscover_catalogs(exclude_name=""):
    paths = []
    for pattern in DEFAULT_CATALOG_PATTERNS:
        paths.extend(Path(".").glob(pattern))
    unique = []
    seen = set()
    for p in paths:
        if p.name == exclude_name:
            continue
        if p.name not in seen:
            seen.add(p.name)
            unique.append(p)
    return unique

def build_indicator_frame(ref_ws, eval_ws):
    rows = []
    current_area = None
    for ref_row, eval_row in zip(range(4, 29), range(9, 34)):
        area = ref_ws[f"B{ref_row}"].value
        if area:
            current_area = str(area).strip()
        indicator = ref_ws[f"C{ref_row}"].value
        if not indicator:
            continue
        indicator = str(indicator).strip()
        weight = safe_float(ref_ws[f"D{ref_row}"].value) or 0.0
        objective_raw = safe_float(ref_ws[f"E{ref_row}"].value)
        objective_weighted = safe_float(ref_ws[f"F{ref_row}"].value)
        max_weighted = safe_float(ref_ws[f"G{ref_row}"].value)
        bbdd_raw = safe_float(ref_ws[f"H{ref_row}"].value)
        bbdd_weighted = safe_float(ref_ws[f"I{ref_row}"].value)
        raw_score = safe_float(eval_ws[f"D{eval_row}"].value)
        weighted_score = None if raw_score is None else raw_score * weight
        if objective_weighted is None and objective_raw is not None:
            objective_weighted = objective_raw * weight
        if max_weighted is None:
            max_weighted = 4 * weight
        if bbdd_weighted is None and bbdd_raw is not None:
            bbdd_weighted = bbdd_raw * weight
        rows.append({
            "tronco": current_area,
            "indicator": indicator,
            "weight": weight,
            "score_raw": raw_score,
            "score_weighted": weighted_score,
            "objective_weighted": objective_weighted,
            "max_weighted": max_weighted,
            "bbdd_weighted": bbdd_weighted,
            "vs_goal": None if not objective_weighted or weighted_score is None else weighted_score / objective_weighted,
            "vs_max": None if not max_weighted or weighted_score is None else weighted_score / max_weighted,
            "vs_bbdd": None if not bbdd_weighted or weighted_score is None else weighted_score / bbdd_weighted,
        })
    return pd.DataFrame(rows)

def summarise_trunks(indicators_df):
    rows = []
    for trunk, grp in indicators_df.groupby("tronco", dropna=False):
        score_weighted = grp["score_weighted"].sum(skipna=True)
        objective_weighted = grp["objective_weighted"].sum(skipna=True)
        max_weighted = grp["max_weighted"].sum(skipna=True)
        bbdd_weighted = grp["bbdd_weighted"].sum(skipna=True)
        benchmark_avg = None if grp["vs_bbdd"].dropna().empty else (grp["score_raw"] / grp["vs_bbdd"]).replace([pd.NA, pd.NaT, float("inf")], pd.NA).dropna().mean()
        rows.append({
            "tronco": trunk,
            "score_raw_avg": grp["score_raw"].mean(skipna=True),
            "benchmark_raw_avg": benchmark_avg,
            "vs_goal": None if not objective_weighted else score_weighted / objective_weighted,
            "vs_max": None if not max_weighted else score_weighted / max_weighted,
            "vs_bbdd": None if not bbdd_weighted else score_weighted / bbdd_weighted,
        })
    return pd.DataFrame(rows)

def summarise_global(indicators_df):
    score_weighted = indicators_df["score_weighted"].sum(skipna=True)
    objective_weighted = indicators_df["objective_weighted"].sum(skipna=True)
    max_weighted = indicators_df["max_weighted"].sum(skipna=True)
    bbdd_weighted = indicators_df["bbdd_weighted"].sum(skipna=True)
    return {
        "score_raw_avg": indicators_df["score_raw"].mean(skipna=True),
        "vs_goal": None if not objective_weighted else score_weighted / objective_weighted,
        "vs_max": None if not max_weighted else score_weighted / max_weighted,
        "vs_bbdd": None if not bbdd_weighted else score_weighted / bbdd_weighted,
    }

def classify_profile(avg_raw):
    if avg_raw is None or pd.isna(avg_raw):
        return "No disponible"
    if avg_raw < PROFILE_THRESHOLDS["Básico"]:
        return "Básico"
    if avg_raw < PROFILE_THRESHOLDS["Supera"]:
        return "Supera"
    if avg_raw < PROFILE_THRESHOLDS["Certificado"]:
        return "Certificado"
    if avg_raw < PROFILE_THRESHOLDS["Excelente"]:
        return "Excelente"
    return "Máster"

def parse_assessment(file):
    suffix = Path(file.name).suffix or ".xlsm"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(file.getvalue())
        temp_path = tmp.name
    wb = load_workbook(temp_path, data_only=True, keep_vba=True)
    if "REFERENCIAS" not in wb.sheetnames or "EVALUACION" not in wb.sheetnames:
        raise ValueError("El archivo no contiene las hojas 'REFERENCIAS' y 'EVALUACION' esperadas.")
    indicators_df = build_indicator_frame(wb["REFERENCIAS"], wb["EVALUACION"])
    if indicators_df.empty:
        raise ValueError("No se han podido leer los indicadores del assessment.")
    if indicators_df["score_raw"].notna().sum() < 20:
        raise ValueError("El assessment no tiene suficientes puntuaciones válidas. Abre y guarda el Excel en origen y vuelve a subirlo.")
    invalid = indicators_df[indicators_df["score_raw"].notna() & ((indicators_df["score_raw"] < 0) | (indicators_df["score_raw"] > 6))]
    if not invalid.empty:
        raise ValueError("Hay puntuaciones fuera del rango 0–6.")
    eval_ws = wb["EVALUACION"]
    global_summary = summarise_global(indicators_df)
    return {
        "name": str(eval_ws["C2"].value or file.name).strip(),
        "species_from_file": str(eval_ws["C4"].value or "No disponible").strip(),
        "date": str(eval_ws["C6"].value or "No disponible"),
        "indicators": indicators_df,
        "trunks": summarise_trunks(indicators_df),
        "global": global_summary,
        "profile": classify_profile(global_summary["score_raw_avg"]),
    }

def rank_strengths(indicators_df, top_n=5):
    tmp = indicators_df.copy()
    tmp["strength_score"] = tmp["vs_goal"].fillna(0) * 0.45 + tmp["vs_bbdd"].fillna(0) * 0.35 + tmp["vs_max"].fillna(0) * 0.20
    return tmp.sort_values("strength_score", ascending=False).head(top_n)

def rank_weaknesses(indicators_df, top_n=5):
    tmp = indicators_df.copy()
    tmp["gap_score"] = (
        (1 - tmp["vs_goal"].fillna(0)).clip(lower=0) * 0.55
        + (1 - tmp["vs_bbdd"].fillna(0)).clip(lower=0) * 0.35
        + ((3 - tmp["score_raw"].fillna(0)) / 3).clip(lower=0) * 0.10
    )
    return tmp.sort_values("gap_score", ascending=False).head(top_n)

def build_priority_areas(indicators_df, top_n=4):
    weak = rank_weaknesses(indicators_df, top_n=8).copy()
    weak["priority"] = weak["gap_score"]
    return weak.groupby("tronco", as_index=False)["priority"].mean().sort_values("priority", ascending=False).head(top_n)

def build_specialization_areas(indicators_df, top_n=3):
    strong = rank_strengths(indicators_df, top_n=8).copy()
    strong["priority"] = strong["strength_score"]
    return strong.groupby("tronco", as_index=False)["priority"].mean().sort_values("priority", ascending=False).head(top_n)

def recommend_internal_courses(catalog_df, species, subspecies, area, max_items=4):
    if catalog_df.empty:
        return pd.DataFrame()
    cat = catalog_df.copy()
    cat["match_score"] = 0.0
    cat.loc[cat["species"].apply(lambda x: species_match(species, x)), "match_score"] += 3
    if subspecies and normalize_text(subspecies) != "general":
        cat.loc[cat["subspecies"].apply(lambda x: species_match(subspecies, x) or normalize_text(x) == "general"), "match_score"] += 1
    cat.loc[cat.apply(lambda r: area_match(area, r["area"], " ".join([str(r["program"]), str(r["content"]), str(r["institution"])])), axis=1), "match_score"] += 3
    cat.loc[cat["type"].astype(str).str.contains("master|máster|postgrado|diplom", case=False, na=False), "match_score"] += 0.2
    filtered = cat[cat["match_score"] > 0].sort_values(["match_score", "program"], ascending=[False, True]).drop_duplicates(subset=["program", "institution"]).head(max_items)
    return filtered[["program", "institution", "duration", "modality", "location", "link", "area", "source_file"]]

@st.cache_data(show_spinner=False, ttl=21600)
def web_search_courses(species, area, subspecies, max_items=3):
    area_query = {
        "Nutrición": "nutrition animal science feed formulation course master",
        "Patología": "animal health pathology biosecurity veterinary course diploma",
        "Manejo": "livestock management welfare facilities production course",
        "Herramientas": "data analytics digitalization livestock power bi course",
        "General": "animal science veterinary production course",
    }.get(area, "animal science veterinary production course")

    queries = [
        f'"{species}" {subspecies} {area} curso master formacion veterinaria',
        f'"{species}" {area} {area_query}',
        f'"{species}" "{area}" training course university',
    ]

    results = []
    seen = set()
    for query in queries:
        try:
            url = f"https://html.duckduckgo.com/html/?q={quote_plus(query)}"
            resp = requests.get(url, headers=SEARCH_HEADERS, timeout=15)
            if resp.status_code != 200:
                continue
            soup = BeautifulSoup(resp.text, "html.parser")
            for res in soup.select(".result"):
                a = res.select_one(".result__title a") or res.select_one("a.result__a")
                if not a:
                    continue
                title = a.get_text(" ", strip=True)
                link = a.get("href") or ""
                if not title or not link:
                    continue
                norm = normalize_text(title)
                if norm in seen:
                    continue
                seen.add(norm)
                snippet_node = res.select_one(".result__snippet")
                snippet = snippet_node.get_text(" ", strip=True) if snippet_node else ""
                domain = urlparse(link).netloc.replace("www.", "") if link.startswith("http") else "No disponible"
                results.append({
                    "program": title,
                    "institution": domain or "No disponible",
                    "duration": "No disponible",
                    "modality": "No disponible",
                    "location": "No disponible",
                    "link": link,
                    "area": area,
                    "source_file": "Búsqueda web",
                    "content": snippet,
                })
                if len(results) >= max_items:
                    return pd.DataFrame(results)
        except Exception:
            continue
    return pd.DataFrame(results)

def build_question_answer(question, analysis):
    strengths = rank_strengths(analysis["indicators"], top_n=3)
    weaknesses = rank_weaknesses(analysis["indicators"], top_n=3)
    best_strengths = ", ".join(strengths["indicator"].tolist()) or "fortalezas no disponibles"
    main_weaknesses = ", ".join(weaknesses["indicator"].tolist()) or "brechas no disponibles"
    profile = analysis["profile"]
    group = QUESTION_TO_GROUP.get(question, "General")
    if group == "Potencial y proyección":
        return f"Con el perfil actual (**{profile}**) y la combinación de fortalezas ({best_strengths}), el potencial de crecimiento parece razonable si se cierran primero las brechas más limitantes ({main_weaknesses})."
    if group == "Fortalezas y especialización":
        return f"Las áreas con mayor probabilidad de consolidarse como línea de especialización son aquellas donde el técnico ya muestra mejor desempeño relativo: {best_strengths}. Conviene reforzarlas con formación avanzada sin perder equilibrio global."
    if group == "Brechas y riesgos":
        return f"Las principales brechas del caso se concentran en {main_weaknesses}. Si estas áreas no se corrigen antes de forzar una especialización avanzada, existe riesgo de crear un perfil desequilibrado."
    return f"El plan de acción debería partir del nivel actual (**{profile}**) y priorizar la mejora de {main_weaknesses}. Después tendría sentido acelerar el desarrollo sobre fortalezas como {best_strengths}."

def build_executive_report(analysis, selected_species, selected_subspecies):
    global_summary = analysis["global"]
    trunks_df = analysis["trunks"].copy()
    strengths = rank_strengths(analysis["indicators"], top_n=3)
    weaknesses = rank_weaknesses(analysis["indicators"], top_n=3)
    priority_areas = build_priority_areas(analysis["indicators"], top_n=3)
    best_trunk = trunks_df.sort_values("vs_goal", ascending=False).iloc[0]["tronco"] if not trunks_df.empty else "No disponible"
    weak_trunk = trunks_df.sort_values("vs_goal", ascending=True).iloc[0]["tronco"] if not trunks_df.empty else "No disponible"
    strength_names = ", ".join(str(x) for x in strengths["indicator"].tolist()) or "No disponible"
    weakness_names = ", ".join(str(x) for x in weaknesses["indicator"].tolist()) or "No disponible"
    area_names = ", ".join(str(x) for x in priority_areas["tronco"].tolist()) or "No disponible"
    lines = [
        f"1. El técnico evaluado es **{analysis['name']}** y el perfil global propuesto es **{analysis['profile']}**.",
        f"2. El score medio global del assessment es **{global_summary['score_raw_avg']:.2f} / 6**.",
        f"3. El resultado se sitúa en **{format_pct(global_summary['vs_goal'])}** respecto al objetivo y en **{format_pct(global_summary['vs_bbdd'])}** respecto al benchmark.",
        f"4. El tronco más sólido es **{best_trunk}** y el tronco con mayor necesidad de desarrollo es **{weak_trunk}**.",
        f"5. Las fortalezas más visibles se concentran en: **{strength_names}**.",
        f"6. Las debilidades que más condicionan el desarrollo actual son: **{weakness_names}**.",
        f"7. Las áreas prioritarias del plan de carrera deberían centrarse primero en: **{area_names}**.",
        f"8. La lectura del perfil sugiere combinar formación base para cerrar brechas y formación especializada para consolidar ventajas competitivas.",
        f"9. El plan se ha elaborado para la especie **{selected_species}** y la subespecie **{selected_subspecies}**.",
        f"10. Se recomienda revisar de nuevo los indicadores prioritarios en la próxima evaluación para confirmar evolución y ajustar el itinerario.",
    ]
    return "\n".join(lines)

def build_final_report_text(analysis, selected_species, selected_subspecies, base_internal, base_external, adv_internal, adv_external, qa_items):
    global_summary = analysis["global"]
    trunks = analysis["trunks"].copy()
    strengths = rank_strengths(analysis["indicators"], top_n=5)
    weaknesses = rank_weaknesses(analysis["indicators"], top_n=5)
    priority_areas = build_priority_areas(analysis["indicators"], top_n=4)
    specialization_areas = build_specialization_areas(analysis["indicators"], top_n=3)

    lines = []
    lines.append("# Informe de assessment y plan de carrera")
    lines.append("")
    lines.append("## 1. Resumen ejecutivo")
    lines.append(build_executive_report(analysis, selected_species, selected_subspecies))
    lines.append("")
    lines.append("## 2. Perfil del técnico")
    lines.append(f"- Nombre: {analysis['name']}")
    lines.append(f"- Especie seleccionada: {selected_species}")
    lines.append(f"- Subespecie seleccionada: {selected_subspecies}")
    lines.append(f"- Categoría propuesta: {analysis['profile']}")
    lines.append(f"- Score global medio: {global_summary['score_raw_avg']:.2f} / 6")
    lines.append(f"- Vs objetivo global: {format_pct(global_summary['vs_goal'])}")
    lines.append(f"- Vs benchmark global: {format_pct(global_summary['vs_bbdd'])}")
    lines.append("")
    lines.append("## 3. Resultados por troncos")
    for _, row in trunks.iterrows():
        lines.append(f"- {row['tronco']}: score medio {row['score_raw_avg']:.2f} | vs objetivo {format_pct(row['vs_goal'])} | vs benchmark {format_pct(row['vs_bbdd'])}")
    lines.append("")
    lines.append("## 4. Indicadores clave")
    lines.append("### Fortalezas")
    for _, row in strengths.iterrows():
        lines.append(f"- {row['indicator']} ({row['tronco']}): score {row['score_raw']:.1f} | vs objetivo {format_pct(row['vs_goal'])} | vs benchmark {format_pct(row['vs_bbdd'])}")
    lines.append("### Debilidades")
    for _, row in weaknesses.iterrows():
        lines.append(f"- {row['indicator']} ({row['tronco']}): score {row['score_raw']:.1f} | vs objetivo {format_pct(row['vs_goal'])} | vs benchmark {format_pct(row['vs_bbdd'])}")
    lines.append("")
    lines.append("## 5. Áreas prioritarias de formación")
    for _, row in priority_areas.iterrows():
        lines.append(f"- {row['tronco']}: prioridad relativa {row['priority']:.2f}")
    lines.append("")
    lines.append("## 6. Áreas de especialización recomendadas")
    for _, row in specialization_areas.iterrows():
        lines.append(f"- {row['tronco']}: potencial relativo {row['priority']:.2f}")
    lines.append("")
    lines.append("## 7. Plan de formación")
    lines.append("### 7.1 Formación base — catálogo interno")
    if base_internal.empty:
        lines.append("- No se han encontrado cursos internos en los catálogos cargados.")
    else:
        for _, row in base_internal.iterrows():
            lines.append(f"- [{row['area']}] {row['program']} | {row['institution']} | {row['duration']} | {row['modality']} | {row['source_file']}")
    lines.append("### 7.2 Formación base — búsqueda externa")
    if base_external.empty:
        lines.append("- No se han encontrado opciones externas con la búsqueda actual.")
    else:
        for _, row in base_external.iterrows():
            lines.append(f"- [{row['area']}] {row['program']} | {row['institution']} | {row['link']}")
    lines.append("### 7.3 Formación especializada — catálogo interno")
    if adv_internal.empty:
        lines.append("- No se han encontrado cursos internos de especialización en los catálogos cargados.")
    else:
        for _, row in adv_internal.iterrows():
            lines.append(f"- [{row['area']}] {row['program']} | {row['institution']} | {row['duration']} | {row['modality']} | {row['source_file']}")
    lines.append("### 7.4 Formación especializada — búsqueda externa")
    if adv_external.empty:
        lines.append("- No se han encontrado opciones externas de especialización con la búsqueda actual.")
    else:
        for _, row in adv_external.iterrows():
            lines.append(f"- [{row['area']}] {row['program']} | {row['institution']} | {row['link']}")
    included_qas = [x for x in qa_items if x.get("include", True)]
    if included_qas:
        lines.append("")
        lines.append("## 8. Preguntas y respuestas incorporadas")
        for item in included_qas:
            lines.append(f"### {item['group']} — {item['question']}")
            lines.append(item["answer"])
            lines.append("")
    return "\n".join(lines)

def make_excel_export(analysis, base_internal, base_external, adv_internal, adv_external, qa_items):
    bio = io.BytesIO()
    summary_df = pd.DataFrame([{
        "Nombre": analysis["name"],
        "Categoría": analysis["profile"],
        "Score global medio": round(analysis["global"]["score_raw_avg"], 2),
        "Vs objetivo global": pct(analysis["global"]["vs_goal"]),
        "Vs benchmark global": pct(analysis["global"]["vs_bbdd"]),
    }])
    trunks_df = analysis["trunks"].copy()
    trunks_df["vs_goal"] = trunks_df["vs_goal"].apply(pct)
    trunks_df["vs_bbdd"] = trunks_df["vs_bbdd"].apply(pct)
    indicators_df = analysis["indicators"].copy()
    indicators_df["vs_goal"] = indicators_df["vs_goal"].apply(pct)
    indicators_df["vs_bbdd"] = indicators_df["vs_bbdd"].apply(pct)
    indicators_df["vs_max"] = indicators_df["vs_max"].apply(pct)
    qa_df = pd.DataFrame(qa_items) if qa_items else pd.DataFrame(columns=["group", "question", "answer", "include"])
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, sheet_name="Resumen", index=False)
        trunks_df.to_excel(writer, sheet_name="Troncos", index=False)
        indicators_df.to_excel(writer, sheet_name="Indicadores", index=False)
        base_internal.to_excel(writer, sheet_name="Base interna", index=False)
        base_external.to_excel(writer, sheet_name="Base externa", index=False)
        adv_internal.to_excel(writer, sheet_name="Esp interna", index=False)
        adv_external.to_excel(writer, sheet_name="Esp externa", index=False)
        qa_df.to_excel(writer, sheet_name="Preguntas", index=False)
    bio.seek(0)
    return bio.getvalue()

def reset_app():
    st.session_state["reset_counter"] = st.session_state.get("reset_counter", 0) + 1
    st.session_state["qa_items"] = []
    st.session_state["final_report_text"] = ""
    st.rerun()

if "qa_items" not in st.session_state:
    st.session_state["qa_items"] = []

st.title("Assessment + Plan de Carrera Técnico")
st.caption("Carga la evaluación de un técnico, interpreta el assessment y genera un informe editable con comparativas, plan de carrera y recomendaciones formativas internas y externas.")

with st.sidebar:
    st.header("Carga y configuración")
    reset_key = st.session_state.get("reset_counter", 0)
    assessment_file = st.file_uploader("Assessment del técnico (.xlsm / .xlsx)", type=["xlsm", "xlsx"], key=f"assessment_{reset_key}")
    uploaded_catalogs = st.file_uploader("Catálogos de formación (.xlsx / .xlsm)", type=["xlsm", "xlsx"], accept_multiple_files=True, key=f"catalogs_{reset_key}")
    courses_per_area = st.number_input("Cursos internos por área", min_value=1, max_value=10, value=4, step=1)
    external_per_area = st.number_input("Cursos externos por área", min_value=1, max_value=10, value=3, step=1)
    search_web = st.checkbox("Refuerza con búsqueda web externa", value=True)
    with st.expander("Configurar umbrales de perfil", expanded=False):
        PROFILE_THRESHOLDS["Básico"] = st.slider("Límite superior Básico", 0.5, 3.0, 2.0, 0.1)
        PROFILE_THRESHOLDS["Supera"] = st.slider("Límite superior Supera", PROFILE_THRESHOLDS["Básico"] + 0.1, 4.0, 3.0, 0.1)
        PROFILE_THRESHOLDS["Certificado"] = st.slider("Límite superior Certificado", PROFILE_THRESHOLDS["Supera"] + 0.1, 5.0, 4.0, 0.1)
        PROFILE_THRESHOLDS["Excelente"] = st.slider("Límite superior Excelente", PROFILE_THRESHOLDS["Certificado"] + 0.1, 6.0, 5.0, 0.1)
    if st.button("Nuevo / borrar evaluación cargada", use_container_width=True):
        reset_app()

if not assessment_file:
    st.stop()

try:
    analysis = parse_assessment(assessment_file)
except Exception as exc:
    st.error(f"No se ha podido procesar el assessment: {exc}")
    st.stop()

catalog_frames = []
auto_catalog_paths = autodiscover_catalogs(exclude_name=assessment_file.name)
for p in auto_catalog_paths:
    try:
        df = load_catalog_from_path(str(p))
        if not df.empty:
            catalog_frames.append(df)
    except Exception:
        pass

for up in uploaded_catalogs or []:
    try:
        df = parse_uploaded_catalog(up)
        if not df.empty:
            catalog_frames.append(df)
    except Exception as exc:
        st.warning(f"No se ha podido leer el catálogo {up.name}: {exc}")

catalog_df = pd.concat(catalog_frames, ignore_index=True).drop_duplicates(subset=["program", "institution", "species", "source_file"]) if catalog_frames else pd.DataFrame()

species_options = sorted(set([analysis["species_from_file"]] + ([] if catalog_df.empty else [x for x in catalog_df["species"].dropna().tolist() if str(x).strip()])))
default_species_idx = species_options.index(analysis["species_from_file"]) if analysis["species_from_file"] in species_options else 0
selected_species = st.selectbox("Especie", species_options, index=default_species_idx)

subspecies_options = ["General"]
if not catalog_df.empty:
    subs = catalog_df[catalog_df["species"].apply(lambda x: species_match(selected_species, x))]["subspecies"].dropna().tolist()
    subspecies_options = sorted(set(["General"] + [x for x in subs if str(x).strip()]))
selected_subspecies = st.selectbox("Subespecie", subspecies_options, index=0)

global_summary = analysis["global"]
trunks_df = analysis["trunks"].copy()
indicators_df = analysis["indicators"].copy()

strengths = rank_strengths(indicators_df, top_n=5)
weaknesses = rank_weaknesses(indicators_df, top_n=5)
priority_areas = build_priority_areas(indicators_df, top_n=4)
specialization_areas = build_specialization_areas(indicators_df, top_n=3)

base_internal_frames = []
for _, row in priority_areas.iterrows():
    rec = recommend_internal_courses(catalog_df, selected_species, selected_subspecies, row["tronco"], max_items=courses_per_area)
    if not rec.empty:
        base_internal_frames.append(rec)
base_internal = pd.concat(base_internal_frames, ignore_index=True).drop_duplicates(subset=["program", "institution"]) if base_internal_frames else pd.DataFrame(columns=["program", "institution", "duration", "modality", "location", "link", "area", "source_file"])

adv_internal_frames = []
for _, row in specialization_areas.iterrows():
    rec = recommend_internal_courses(catalog_df, selected_species, selected_subspecies, row["tronco"], max_items=courses_per_area)
    if not rec.empty:
        adv_internal_frames.append(rec)
adv_internal = pd.concat(adv_internal_frames, ignore_index=True).drop_duplicates(subset=["program", "institution"]) if adv_internal_frames else pd.DataFrame(columns=["program", "institution", "duration", "modality", "location", "link", "area", "source_file"])

base_external_frames = []
adv_external_frames = []
if search_web:
    with st.spinner("Buscando formación externa en internet..."):
        for _, row in priority_areas.iterrows():
            df = web_search_courses(selected_species, row["tronco"], selected_subspecies, max_items=external_per_area)
            if not df.empty:
                base_external_frames.append(df)
        for _, row in specialization_areas.iterrows():
            df = web_search_courses(selected_species, row["tronco"], selected_subspecies, max_items=external_per_area)
            if not df.empty:
                adv_external_frames.append(df)

base_external = pd.concat(base_external_frames, ignore_index=True).drop_duplicates(subset=["program", "institution"]) if base_external_frames else pd.DataFrame(columns=["program", "institution", "duration", "modality", "location", "link", "area", "source_file", "content"])
adv_external = pd.concat(adv_external_frames, ignore_index=True).drop_duplicates(subset=["program", "institution"]) if adv_external_frames else pd.DataFrame(columns=["program", "institution", "duration", "modality", "location", "link", "area", "source_file", "content"])

st.subheader("Resumen ejecutivo del caso")
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Perfil propuesto", analysis["profile"])
m2.metric("Score global medio", f"{global_summary['score_raw_avg']:.2f} / 6")
m3.metric("Vs objetivo global", format_pct(global_summary["vs_goal"]))
m4.metric("Vs benchmark global", format_pct(global_summary["vs_bbdd"]))
m5.metric("Cursos internos detectados", len(catalog_df))

st.markdown(build_executive_report(analysis, selected_species, selected_subspecies))

with st.expander("Resumen de catálogos cargados", expanded=False):
    if catalog_df.empty:
        st.warning("No se ha cargado ningún catálogo válido.")
    else:
        st.write(f"Registros de formación detectados: **{len(catalog_df)}**")
        st.dataframe(
            catalog_df.groupby(["species", "source_file"], as_index=False)["program"].count().rename(columns={"program": "n_cursos"}),
            use_container_width=True,
            hide_index=True,
        )

st.subheader("Lectura por troncos")
bar_df = trunks_df.copy()
bar_df["Score técnico"] = bar_df["score_raw_avg"].round(2)
bar_df["Referencia"] = 3.0
bar_df["Benchmark aprox."] = bar_df["benchmark_raw_avg"].fillna(0).round(2)
bar_long = bar_df.melt(id_vars=["tronco"], value_vars=["Score técnico", "Referencia", "Benchmark aprox."], var_name="Serie", value_name="Valor")
fig_bar = px.bar(bar_long, x="Valor", y="tronco", color="Serie", barmode="group", orientation="h", text="Valor", height=430)
fig_bar.update_traces(texttemplate="%{text:.2f}", textposition="outside")
fig_bar.update_layout(xaxis_title="Puntuación media", yaxis_title="Tronco", margin=dict(l=40, r=40, t=40, b=40))
st.plotly_chart(fig_bar, use_container_width=True)

radar_df = trunks_df.copy()
radar_df["Técnico"] = radar_df["score_raw_avg"].round(2)
radar_df["Referencia"] = 3.0
radar_df["Benchmark aprox."] = radar_df["benchmark_raw_avg"].fillna(0).round(2)
categories = radar_df["tronco"].tolist()
fig_radar = go.Figure()
for series in ["Técnico", "Referencia", "Benchmark aprox."]:
    fig_radar.add_trace(go.Scatterpolar(r=radar_df[series].tolist(), theta=categories, fill="toself", name=series))
fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 6])), showlegend=True, height=520)
st.plotly_chart(fig_radar, use_container_width=True)

left, right = st.columns(2)
with left:
    st.markdown("### Top 5 fortalezas")
    strong_display = strengths[["tronco", "indicator", "score_raw", "vs_goal", "vs_bbdd"]].copy()
    strong_display["vs_goal"] = strong_display["vs_goal"].apply(format_pct)
    strong_display["vs_bbdd"] = strong_display["vs_bbdd"].apply(format_pct)
    strong_display.columns = ["Tronco", "Indicador", "Score", "Vs objetivo", "Vs benchmark"]
    st.dataframe(strong_display, use_container_width=True, hide_index=True)
with right:
    st.markdown("### Top 5 debilidades")
    weak_display = weaknesses[["tronco", "indicator", "score_raw", "vs_goal", "vs_bbdd"]].copy()
    weak_display["vs_goal"] = weak_display["vs_goal"].apply(format_pct)
    weak_display["vs_bbdd"] = weak_display["vs_bbdd"].apply(format_pct)
    weak_display.columns = ["Tronco", "Indicador", "Score", "Vs objetivo", "Vs benchmark"]
    st.dataframe(weak_display, use_container_width=True, hide_index=True)

st.subheader("Plan de carrera y formación")
c1, c2 = st.columns(2)
with c1:
    st.markdown("### Áreas prioritarias de desarrollo")
    pr_display = priority_areas.copy()
    pr_display["priority"] = pr_display["priority"].round(2)
    pr_display.columns = ["Tronco", "Prioridad relativa"]
    st.dataframe(pr_display, use_container_width=True, hide_index=True)

    st.markdown("### Formación base — catálogo interno")
    if base_internal.empty:
        st.warning("No se han encontrado cursos internos para las áreas prioritarias con los catálogos cargados.")
    else:
        st.dataframe(base_internal, use_container_width=True, hide_index=True)

    st.markdown("### Formación base — búsqueda externa")
    if base_external.empty:
        st.info("No se han encontrado opciones externas con la búsqueda actual.")
    else:
        st.dataframe(base_external[["program", "institution", "link", "area"]], use_container_width=True, hide_index=True)

with c2:
    st.markdown("### Áreas de especialización")
    sp_display = specialization_areas.copy()
    sp_display["priority"] = sp_display["priority"].round(2)
    sp_display.columns = ["Tronco", "Potencial relativo"]
    st.dataframe(sp_display, use_container_width=True, hide_index=True)

    st.markdown("### Formación especializada — catálogo interno")
    if adv_internal.empty:
        st.warning("No se han encontrado cursos internos de especialización con los catálogos cargados.")
    else:
        st.dataframe(adv_internal, use_container_width=True, hide_index=True)

    st.markdown("### Formación especializada — búsqueda externa")
    if adv_external.empty:
        st.info("No se han encontrado opciones externas de especialización con la búsqueda actual.")
    else:
        st.dataframe(adv_external[["program", "institution", "link", "area"]], use_container_width=True, hide_index=True)

st.subheader("Preguntas del usuario para enriquecer el informe")
with st.expander("Ver las 20 preguntas agrupadas", expanded=False):
    for group, questions in QUESTION_GROUPS.items():
        st.markdown(f"**{group}**")
        for q in questions:
            st.write(f"- {q}")

q_col1, q_col2 = st.columns([1, 2])
with q_col1:
    selected_group = st.selectbox("Grupo de preguntas", list(QUESTION_GROUPS.keys()))
with q_col2:
    selected_question = st.selectbox("Pregunta", QUESTION_GROUPS[selected_group])

if st.button("Añadir pregunta al informe"):
    current_questions = [x["question"] for x in st.session_state["qa_items"]]
    if selected_question not in current_questions:
        st.session_state["qa_items"].append({
            "group": selected_group,
            "question": selected_question,
            "answer": build_question_answer(selected_question, analysis),
            "include": True,
        })
        st.rerun()

if st.session_state["qa_items"]:
    for idx, item in enumerate(st.session_state["qa_items"]):
        with st.expander(f"{item['group']} — {item['question']}", expanded=False):
            include = st.checkbox("Incorporar esta respuesta al informe final", value=item.get("include", True), key=f"include_{idx}")
            answer = st.text_area("Respuesta editable", value=item.get("answer", ""), height=140, key=f"answer_{idx}")
            if st.button("Eliminar esta pregunta", key=f"remove_{idx}"):
                st.session_state["qa_items"].pop(idx)
                st.rerun()
            st.session_state["qa_items"][idx]["include"] = include
            st.session_state["qa_items"][idx]["answer"] = answer

st.subheader("Informe final editable")
auto_report = build_final_report_text(
    analysis,
    selected_species,
    selected_subspecies,
    base_internal,
    base_external,
    adv_internal,
    adv_external,
    st.session_state["qa_items"],
)
if not st.session_state.get("final_report_text"):
    st.session_state["final_report_text"] = auto_report

if st.button("Refrescar informe con el contenido actual"):
    st.session_state["final_report_text"] = build_final_report_text(
        analysis,
        selected_species,
        selected_subspecies,
        base_internal,
        base_external,
        adv_internal,
        adv_external,
        st.session_state["qa_items"],
    )
    st.rerun()

final_report_text = st.text_area(
    "Puedes editar libremente el informe antes de descargarlo",
    value=st.session_state["final_report_text"],
    height=540,
)
st.session_state["final_report_text"] = final_report_text

st.subheader("Descargas")
txt_bytes = final_report_text.encode("utf-8")
excel_bytes = make_excel_export(analysis, base_internal, base_external, adv_internal, adv_external, st.session_state["qa_items"])
d1, d2 = st.columns(2)
with d1:
    st.download_button("Descargar informe en TXT", data=txt_bytes, file_name="informe_plan_carrera.txt", mime="text/plain", use_container_width=True)
with d2:
    st.download_button("Descargar detalle en Excel", data=excel_bytes, file_name="assessment_plan_carrera.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
