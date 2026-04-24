import io
import tempfile
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
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
QUESTION_TO_GROUP = {q: group for group, questions in QUESTION_GROUPS.items() for q in questions}

PROFILE_THRESHOLDS = {
    "Básico": 2.0,
    "Supera": 3.0,
    "Certificado": 4.0,
    "Excelente": 5.0,
    "Máster": 6.1,
}

AREA_KEYWORDS = {
    "Nutrición": ["nutricion", "nutrición", "producto", "aditivo", "agua", "alimentacion", "alimentación", "cebo", "recria", "recría", "racion", "ración", "sostenibilidad"],
    "Patología": ["patologia", "patología", "metabolica", "metabólica", "infecciosa", "parasitaria", "antibioterapia", "diagnost", "inmunidad", "salud intestinal", "reproduccion", "reproducción"],
    "Manejo": ["bioseguridad", "instalacion", "instalación", "ventilacion", "ventilación", "arranque", "produccion", "producción", "reposicion", "reposición", "manejo", "granja", "housing"],
    "Herramientas": ["datos", "estadistica", "estadística", "crm", "power bi", "bbdd", "digitalizacion", "digitalización", "informe", "ingles", "inglés", "software", "programa"],
}

DEFAULT_CATALOG_FILES = [
    "Resumewn formaciones especificas.xlsx",
    "resumen_formaciones_especificas.xlsx",
    "catalogo_formacion.xlsx",
]


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
    return text


def infer_area_from_text(text):
    text_n = normalize_text(text)
    hits = {}
    for area, keywords in AREA_KEYWORDS.items():
        hits[area] = sum(1 for kw in keywords if kw in text_n)
    best = max(hits, key=hits.get)
    return best if hits[best] > 0 else "General"


def find_default_catalog_path():
    for filename in DEFAULT_CATALOG_FILES:
        p = Path(filename)
        if p.exists():
            return p
    return None


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
        rows.append({
            "tronco": trunk,
            "score_raw_avg": grp["score_raw"].mean(skipna=True),
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

    ref_ws = wb["REFERENCIAS"]
    eval_ws = wb["EVALUACION"]
    indicators_df = build_indicator_frame(ref_ws, eval_ws)

    if indicators_df.empty:
        raise ValueError("No se han podido leer los indicadores del assessment.")
    if indicators_df["score_raw"].notna().sum() < 20:
        raise ValueError("El assessment no tiene suficientes puntuaciones válidas.")
    if not indicators_df[indicators_df["score_raw"].notna() & ((indicators_df["score_raw"] < 0) | (indicators_df["score_raw"] > 6))].empty:
        raise ValueError("Hay puntuaciones fuera del rango 0–6.")

    return {
        "name": str(eval_ws["C2"].value or file.name).strip(),
        "species_from_file": str(eval_ws["C4"].value or "No disponible").strip(),
        "date": str(eval_ws["C6"].value or "No disponible"),
        "indicators": indicators_df,
        "trunks": summarise_trunks(indicators_df),
        "global": summarise_global(indicators_df),
    }


def parse_catalog(file):
    if file is None:
        return pd.DataFrame()

    if isinstance(file, Path):
        wb = load_workbook(file, data_only=True)
    else:
        suffix = Path(file.name).suffix or ".xlsx"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(file.getvalue())
            temp_path = tmp.name
        wb = load_workbook(temp_path, data_only=True)

    ws = wb[wb.sheetnames[0]]
    rows = []
    current_species = None

    for r in range(5, ws.max_row + 1):
        species = ws.cell(r, 1).value
        if species:
            current_species = str(species).strip()
        program = ws.cell(r, 2).value
        if not program:
            continue
        duration = ws.cell(r, 3).value
        modality = ws.cell(r, 4).value
        location = ws.cell(r, 5).value
        institution = ws.cell(r, 6).value
        content = ws.cell(r, 7).value
        link = ws.cell(r, 8).value

        area = infer_area_from_text(" ".join([
            str(current_species or ""), str(program or ""), str(content or ""), str(institution or "")
        ]))

        rows.append({
            "species": current_species or "GENERAL",
            "subspecies": "General",
            "program": str(program).strip(),
            "duration": str(duration).strip() if duration else "No disponible",
            "modality": str(modality).strip() if modality else "No disponible",
            "location": str(location).strip() if location else "No disponible",
            "institution": str(institution).strip() if institution else "No disponible",
            "content": str(content).strip() if content else "",
            "link": str(link).strip() if link else "No disponible",
            "area": area,
        })
    return pd.DataFrame(rows)


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


def recommend_courses(catalog_df, species, area, max_items=3):
    if catalog_df.empty:
        return pd.DataFrame()
    species_n = normalize_text(species)
    cat = catalog_df.copy()
    cat["species_n"] = cat["species"].apply(normalize_text)
    filtered = cat[cat["species_n"].str.contains(species_n, na=False)].copy() if species_n else cat.copy()
    if filtered.empty:
        filtered = cat.copy()
    area_filtered = filtered[filtered["area"].apply(normalize_text) == normalize_text(area)].copy()
    if area_filtered.empty:
        area_filtered = filtered.copy()
    return area_filtered.head(max_items)[["program", "institution", "duration", "modality", "location", "link", "area"]]


def build_executive_report(analysis, selected_species, selected_subspecies):
    global_summary = analysis["global"]
    trunks_df = analysis["trunks"].copy()
    strengths = rank_strengths(analysis["indicators"], top_n=3)
    weaknesses = rank_weaknesses(analysis["indicators"], top_n=3)
    priority_areas = build_priority_areas(analysis["indicators"], top_n=3)
    profile = classify_profile(global_summary["score_raw_avg"])

    best_trunk = trunks_df.sort_values("vs_goal", ascending=False).iloc[0]["tronco"] if not trunks_df.empty else "No disponible"
    weak_trunk = trunks_df.sort_values("vs_goal", ascending=True).iloc[0]["tronco"] if not trunks_df.empty else "No disponible"
    strength_names = ", ".join(str(x) for x in strengths["indicator"].tolist()) or "No disponible"
    weakness_names = ", ".join(str(x) for x in weaknesses["indicator"].tolist()) or "No disponible"
    area_names = ", ".join(str(x) for x in priority_areas["tronco"].tolist()) or "No disponible"

    lines = [
        f"1. El técnico evaluado es **{analysis['name']}** y el perfil global propuesto es **{profile}**.",
        f"2. El score medio global del assessment es **{global_summary['score_raw_avg']:.2f} / 6**.",
        f"3. En términos globales, el resultado se sitúa en **{format_pct(global_summary['vs_goal'])}** respecto al objetivo y en **{format_pct(global_summary['vs_bbdd'])}** respecto al benchmark.",
        f"4. El tronco más sólido es **{best_trunk}** y el tronco con mayor necesidad de desarrollo es **{weak_trunk}**.",
        f"5. Las fortalezas más visibles se concentran en: **{strength_names}**.",
        f"6. Las debilidades que más condicionan el desarrollo actual son: **{weakness_names}**.",
        f"7. Las áreas prioritarias del plan de carrera deberían centrarse primero en: **{area_names}**.",
        f"8. La lectura del perfil sugiere combinar formación base para cerrar brechas y formación especializada para consolidar ventajas competitivas.",
        f"9. El plan se ha elaborado para la especie **{selected_species}** y la subespecie **{selected_subspecies}**.",
        f"10. Se recomienda revisar nuevamente los indicadores prioritarios en la próxima evaluación para confirmar evolución y ajustar el itinerario formativo.",
    ]
    return "\n".join(lines)


def build_question_answer(question, analysis):
    strengths = rank_strengths(analysis["indicators"], top_n=3)
    weaknesses = rank_weaknesses(analysis["indicators"], top_n=3)
    best_strengths = ", ".join(strengths["indicator"].tolist()) or "fortalezas no disponibles"
    main_weaknesses = ", ".join(weaknesses["indicator"].tolist()) or "brechas no disponibles"
    profile = classify_profile(analysis["global"]["score_raw_avg"])
    group = QUESTION_TO_GROUP.get(question, "General")

    if group == "Potencial y proyección":
        return f"Con el perfil actual (**{profile}**) y la combinación de fortalezas ({best_strengths}), el potencial de crecimiento parece razonable si se cierran primero las brechas más limitantes ({main_weaknesses})."
    if group == "Fortalezas y especialización":
        return f"Las áreas con mayor probabilidad de consolidarse como línea de especialización son aquellas donde el técnico ya muestra mejor desempeño relativo: {best_strengths}. Conviene reforzarlas con formación avanzada sin perder equilibrio global."
    if group == "Brechas y riesgos":
        return f"Las principales brechas del caso se concentran en {main_weaknesses}. Si estas áreas no se corrigen antes de forzar una especialización avanzada, existe riesgo de crear un perfil desequilibrado."
    return f"El plan de acción debería partir del nivel actual (**{profile}**) y priorizar la mejora de {main_weaknesses}. Después tendría sentido acelerar el desarrollo sobre fortalezas como {best_strengths}."


def build_final_report_text(analysis, selected_species, selected_subspecies, base_courses, advanced_courses, qa_items):
    global_summary = analysis["global"]
    trunks = analysis["trunks"].copy()
    strengths = rank_strengths(analysis["indicators"], top_n=5)
    weaknesses = rank_weaknesses(analysis["indicators"], top_n=5)
    priority_areas = build_priority_areas(analysis["indicators"], top_n=4)
    specialization_areas = build_specialization_areas(analysis["indicators"], top_n=3)
    profile = classify_profile(global_summary["score_raw_avg"])

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
    lines.append(f"- Categoría propuesta: {profile}")
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
    lines.append("## 7. Formación recomendada")
    lines.append("### 7.1 Formación base")
    if base_courses.empty:
        lines.append("- No hay cursos disponibles en el catálogo cargado para las áreas prioritarias.")
    else:
        for _, row in base_courses.iterrows():
            lines.append(f"- [{row['area']}] {row['program']} | {row['institution']} | {row['duration']} | {row['modality']} | {row['location']} | {row['link']}")
    lines.append("### 7.2 Formación especializada")
    if advanced_courses.empty:
        lines.append("- No hay cursos disponibles en el catálogo cargado para las áreas de especialización.")
    else:
        for _, row in advanced_courses.iterrows():
            lines.append(f"- [{row['area']}] {row['program']} | {row['institution']} | {row['duration']} | {row['modality']} | {row['location']} | {row['link']}")
    included_qas = [x for x in qa_items if x.get("include", True)]
    if included_qas:
        lines.append("")
        lines.append("## 8. Preguntas y respuestas incorporadas")
        for item in included_qas:
            lines.append(f"### {item['group']} — {item['question']}")
            lines.append(item['answer'])
            lines.append("")
    return "\n".join(lines)


def make_excel_export(analysis, base_courses, advanced_courses, qa_items):
    bio = io.BytesIO()
    summary_df = pd.DataFrame([{
        "Nombre": analysis["name"],
        "Categoría": classify_profile(analysis["global"]["score_raw_avg"]),
        "Score global medio": round(analysis["global"]["score_raw_avg"], 2),
        "Vs objetivo global": pct(analysis["global"]["vs_goal"]),
        "Vs benchmark global": pct(analysis["global"]["vs_bbdd"]),
    }])

    trunks_df = analysis["trunks"].copy()
    trunks_df["vs_goal"] = trunks_df["vs_goal"].apply(pct)
    trunks_df["vs_bbdd"] = trunks_df["vs_bbdd"].apply(pct)
    trunks_df["vs_max"] = trunks_df["vs_max"].apply(pct)

    indicators_df = analysis["indicators"].copy()
    indicators_df["vs_goal"] = indicators_df["vs_goal"].apply(pct)
    indicators_df["vs_bbdd"] = indicators_df["vs_bbdd"].apply(pct)
    indicators_df["vs_max"] = indicators_df["vs_max"].apply(pct)

    qa_df = pd.DataFrame(qa_items) if qa_items else pd.DataFrame(columns=["group", "question", "answer", "include"])

    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, sheet_name="Resumen", index=False)
        trunks_df.to_excel(writer, sheet_name="Troncos", index=False)
        indicators_df.to_excel(writer, sheet_name="Indicadores", index=False)
        base_courses.to_excel(writer, sheet_name="Formacion base", index=False)
        advanced_courses.to_excel(writer, sheet_name="Especializacion", index=False)
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
st.caption("Carga la evaluación de un técnico, interpreta el assessment y genera un informe editable con comparativas, prioridades de desarrollo, recomendaciones formativas y preguntas-respuestas incorporables al informe.")

st.info("El error que ves (`ModuleNotFoundError: openpyxl`) suele deberse a que en GitHub/Streamlit no hay un `requirements.txt` correcto o no se ha redeployado la app después de subirlo.")

sidebar = st.sidebar
sidebar.header("Carga y configuración")
reset_key = st.session_state.get("reset_counter", 0)

assessment_file = sidebar.file_uploader("Assessment del técnico (.xlsm / .xlsx)", type=["xlsm", "xlsx"], key=f"assessment_{reset_key}")
catalog_file = sidebar.file_uploader("Catálogo de formación (.xlsx) — opcional si ya está en el repositorio", type=["xlsx", "xlsm"], key=f"catalog_{reset_key}")
courses_per_area = sidebar.number_input("Cursos por área a proponer", min_value=1, max_value=10, value=3, step=1)

with sidebar.expander("Configurar umbrales de perfil", expanded=False):
    PROFILE_THRESHOLDS["Básico"] = st.slider("Límite superior Básico", 0.5, 3.0, 2.0, 0.1)
    PROFILE_THRESHOLDS["Supera"] = st.slider("Límite superior Supera", PROFILE_THRESHOLDS["Básico"] + 0.1, 4.0, 3.0, 0.1)
    PROFILE_THRESHOLDS["Certificado"] = st.slider("Límite superior Certificado", PROFILE_THRESHOLDS["Supera"] + 0.1, 5.0, 4.0, 0.1)
    PROFILE_THRESHOLDS["Excelente"] = st.slider("Límite superior Excelente", PROFILE_THRESHOLDS["Certificado"] + 0.1, 6.0, 5.0, 0.1)

if sidebar.button("Nuevo / borrar evaluación cargada", use_container_width=True):
    reset_app()

if not assessment_file:
    st.stop()

try:
    analysis = parse_assessment(assessment_file)
except Exception as exc:
    st.error(f"No se ha podido procesar el assessment: {exc}")
    st.stop()

default_catalog = find_default_catalog_path() if catalog_file is None else None
catalog_df = pd.DataFrame()
try:
    if catalog_file is not None:
        catalog_df = parse_catalog(catalog_file)
    elif default_catalog is not None:
        catalog_df = parse_catalog(default_catalog)
except Exception as exc:
    st.warning(f"No se ha podido cargar el catálogo de formación: {exc}")

species_options = sorted(set([analysis["species_from_file"]] + ([] if catalog_df.empty else [x for x in catalog_df["species"].dropna().unique().tolist() if str(x).strip()])))
selected_species = st.selectbox("Especie", species_options if species_options else [analysis["species_from_file"]], index=0)
selected_subspecies = st.selectbox("Subespecie", ["General"], index=0)

global_summary = analysis["global"]
trunks_df = analysis["trunks"].copy()
indicators_df = analysis["indicators"].copy()
profile = classify_profile(global_summary["score_raw_avg"])

strengths = rank_strengths(indicators_df, top_n=5)
weaknesses = rank_weaknesses(indicators_df, top_n=5)
priority_areas = build_priority_areas(indicators_df, top_n=4)
specialization_areas = build_specialization_areas(indicators_df, top_n=3)

base_courses_list = []
for _, row in priority_areas.iterrows():
    rec = recommend_courses(catalog_df, selected_species, row["tronco"], max_items=courses_per_area)
    if not rec.empty:
        base_courses_list.append(rec)
base_courses = pd.concat(base_courses_list, ignore_index=True).drop_duplicates() if base_courses_list else pd.DataFrame(columns=["program", "institution", "duration", "modality", "location", "link", "area"])

advanced_courses_list = []
for _, row in specialization_areas.iterrows():
    rec = recommend_courses(catalog_df, selected_species, row["tronco"], max_items=courses_per_area)
    if not rec.empty:
        advanced_courses_list.append(rec)
advanced_courses = pd.concat(advanced_courses_list, ignore_index=True).drop_duplicates() if advanced_courses_list else pd.DataFrame(columns=["program", "institution", "duration", "modality", "location", "link", "area"])

st.subheader("Resumen ejecutivo del caso")
m1, m2, m3, m4 = st.columns(4)
m1.metric("Perfil propuesto", profile)
m2.metric("Score global medio", f"{global_summary['score_raw_avg']:.2f} / 6")
m3.metric("Vs objetivo global", format_pct(global_summary["vs_goal"]))
m4.metric("Vs benchmark global", format_pct(global_summary["vs_bbdd"]))
st.markdown(build_executive_report(analysis, selected_species, selected_subspecies))

st.subheader("Lectura por troncos")
bar_df = trunks_df.copy()
bar_df["Score técnico"] = bar_df["score_raw_avg"].round(2)
bar_df["Referencia"] = 3.0
bar_df["Benchmark aprox."] = (bar_df["score_raw_avg"] / bar_df["vs_bbdd"].replace({0: None})).fillna(0).round(2)
bar_long = bar_df.melt(id_vars=["tronco"], value_vars=["Score técnico", "Referencia", "Benchmark aprox."], var_name="Serie", value_name="Valor")
fig_bar = px.bar(bar_long, x="Valor", y="tronco", color="Serie", barmode="group", orientation="h", text="Valor", height=420)
fig_bar.update_traces(texttemplate="%{text:.2f}", textposition="outside")
fig_bar.update_layout(xaxis_title="Puntuación media", yaxis_title="Tronco", margin=dict(l=40, r=40, t=40, b=40))
st.plotly_chart(fig_bar, use_container_width=True)

radar_df = trunks_df.copy()
radar_df["Técnico"] = radar_df["score_raw_avg"].round(2)
radar_df["Referencia"] = 3.0
radar_df["Benchmark aprox."] = (radar_df["score_raw_avg"] / radar_df["vs_bbdd"].replace({0: None})).fillna(0).round(2)
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

    st.markdown("### Formación base recomendada")
    if base_courses.empty:
        st.info("No hay cursos disponibles en el catálogo cargado para las áreas prioritarias.")
    else:
        st.dataframe(base_courses, use_container_width=True, hide_index=True)

with c2:
    st.markdown("### Áreas de especialización")
    sp_display = specialization_areas.copy()
    sp_display["priority"] = sp_display["priority"].round(2)
    sp_display.columns = ["Tronco", "Potencial relativo"]
    st.dataframe(sp_display, use_container_width=True, hide_index=True)

    st.markdown("### Formación de especialización recomendada")
    if advanced_courses.empty:
        st.info("No hay cursos disponibles en el catálogo cargado para las áreas de especialización.")
    else:
        st.dataframe(advanced_courses, use_container_width=True, hide_index=True)

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
auto_report = build_final_report_text(analysis, selected_species, selected_subspecies, base_courses, advanced_courses, st.session_state["qa_items"])
if not st.session_state.get("final_report_text"):
    st.session_state["final_report_text"] = auto_report

if st.button("Refrescar informe con el contenido actual"):
    st.session_state["final_report_text"] = build_final_report_text(analysis, selected_species, selected_subspecies, base_courses, advanced_courses, st.session_state["qa_items"])
    st.rerun()

final_report_text = st.text_area("Puedes editar libremente el informe antes de descargarlo", value=st.session_state["final_report_text"], height=520)
st.session_state["final_report_text"] = final_report_text

st.subheader("Descargas")
txt_bytes = final_report_text.encode("utf-8")
excel_bytes = make_excel_export(analysis, base_courses, advanced_courses, st.session_state["qa_items"])

d1, d2 = st.columns(2)
with d1:
    st.download_button("Descargar informe en TXT", data=txt_bytes, file_name="informe_plan_carrera.txt", mime="text/plain", use_container_width=True)
with d2:
    st.download_button("Descargar detalle en Excel", data=excel_bytes, file_name="assessment_plan_carrera.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
