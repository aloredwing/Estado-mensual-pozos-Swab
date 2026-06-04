from __future__ import annotations

import re
import unicodedata
import tempfile
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
except Exception:
    Presentation = None

APP_TITLE = "Dashboard mensual de Estado de Pozos OIG"
DATA_DIR = Path(__file__).parent / "data"

MESES = {
    "ENERO": 1,
    "FEBRERO": 2,
    "MARZO": 3,
    "ABRIL": 4,
    "MAYO": 5,
    "JUNIO": 6,
    "JULIO": 7,
    "AGOSTO": 8,
    "SETIEMBRE": 9,
    "SEPTIEMBRE": 9,
    "OCTUBRE": 10,
    "NOVIEMBRE": 11,
    "DICIEMBRE": 12,
}
MESES_INV = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Setiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}
REQ = ["*PFORMACION", "*ESTADO", "*TIPO_DE_POZO", "*ULT_EST", "*BATERIA"]

COLOR_SEQ = ["#0F766E", "#2563EB", "#F97316", "#7C3AED", "#DC2626", "#059669", "#0891B2", "#CA8A04"]
COLOR_MAP = {
    "Activo": "#0F766E",
    "Inactivo": "#DC2626",
    "Observación": "#F59E0B",
    "SWAB": "#2563EB",
    "No SWAB": "#94A3B8",
}

st.set_page_config(page_title=APP_TITLE, page_icon="🛢️", layout="wide")


def norm_txt(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s)


def mes_from_text(text: str):
    t = norm_txt(text)
    for m, n in MESES.items():
        if re.search(rf"\b{m}\b", t):
            return n
    return None


def year_from_text(text: str):
    m = re.search(r"\b(20\d{2}|19\d{2})\b", str(text))
    return int(m.group(1)) if m else None


def first_date_from_cells(raw: pd.DataFrame):
    for v in raw.to_numpy().ravel():
        if isinstance(v, pd.Timestamp):
            return v
        if hasattr(v, "year") and hasattr(v, "month"):
            try:
                return pd.Timestamp(v)
            except Exception:
                pass
    return None


def infer_period(file_name: str, sheet_name: str, raw: pd.DataFrame) -> pd.Timestamp | None:
    title = " | ".join(str(x) for x in raw.iloc[:3, :12].to_numpy().ravel() if pd.notna(x) and str(x).strip())
    sheet_mes = mes_from_text(sheet_name)
    title_mes = mes_from_text(title)
    file_mes = mes_from_text(file_name)
    file_norm = norm_txt(file_name)

    mes = sheet_mes or title_mes or file_mes
    anio = year_from_text(title)

    date_cell = first_date_from_cells(raw)
    if anio is None and date_cell is not None:
        anio = int(date_cell.year)

    if anio is None:
        if "OIG" in file_norm and mes in [1, 2, 3, 4, 5]:
            anio = 2025
        elif mes:
            anio = 2024

    if not mes or not anio:
        return None
    return pd.Timestamp(anio, mes, 1)


def find_header_row(raw: pd.DataFrame) -> int | None:
    max_rows = min(len(raw), 12)
    for i in range(max_rows):
        row = [norm_txt(x) for x in raw.iloc[i].tolist()]
        if all(c in row for c in REQ):
            return i
    return None


def extract_main_table(path: Path, sheet: str, header_row: int) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet, header=header_row, engine="openpyxl")
    clean_cols = [norm_txt(c) for c in df.columns]
    positions = {}
    for req in REQ:
        positions[req] = clean_cols.index(req) if req in clean_cols else None
    if any(v is None for v in positions.values()):
        return pd.DataFrame()

    out = pd.DataFrame({
        "pozo": df.iloc[:, positions["*PFORMACION"]],
        "estado": df.iloc[:, positions["*ESTADO"]],
        "tipo_pozo": df.iloc[:, positions["*TIPO_DE_POZO"]],
        "ult_est": df.iloc[:, positions["*ULT_EST"]],
        "bateria": df.iloc[:, positions["*BATERIA"]],
    })
    out = out.dropna(subset=["pozo"])
    out = out[~out["pozo"].astype(str).str.startswith("*")].copy()
    for c in out.columns:
        out[c] = out[c].astype(str).str.strip().replace({"nan": "", "None": ""})
    out = out[out["pozo"].astype(str).str.strip() != ""]
    return out


def score_candidate(path: Path, sheet: str, periodo: pd.Timestamp, nrows: int) -> float:
    score = float(nrows) / 100000
    file_mes = mes_from_text(path.name)
    sheet_mes = mes_from_text(sheet)
    if file_mes == periodo.month:
        score += 5
    if sheet_mes == periodo.month:
        score += 4
    if "TOTAL" in norm_txt(sheet):
        score += 1
    if "OIG" in norm_txt(path.name):
        score += 0.5
    return score


def clasificar_condicion(ult_est: str, estado: str, tipo: str) -> str:
    u = norm_txt(ult_est)
    e = norm_txt(estado)
    t = norm_txt(tipo)
    inactive_words = ["INACT", "DPA", "APA", "ABAND", "SECO", "ATA"]
    active_words = ["ACT", "BM", "BCP", "BO", "GL", "PL", "PETS", "PETF", "GAS"]
    if any(w in u for w in inactive_words) or u.startswith("INACT"):
        return "Inactivo"
    if any(w in u for w in active_words) or e in ["BM", "BCP", "BO", "GL", "PL", "PETS", "PETF"] or t in ["BM", "BCP", "BO", "GL", "PL"]:
        return "Activo"
    return "Observación"


def clasificar_swab(row) -> str:
    texto = " ".join([norm_txt(row.get("bateria", "")), norm_txt(row.get("tipo_pozo", "")), norm_txt(row.get("ult_est", ""))])
    if any(x in texto for x in ["SWAB", "SUAB", "PETS", "PETF"]):
        return "SWAB"
    return "No SWAB"


@st.cache_data(show_spinner="Cargando base fija de Excel...")
def cargar_base() -> Tuple[pd.DataFrame, pd.DataFrame]:
    if not DATA_DIR.exists():
        return pd.DataFrame(), pd.DataFrame()

    files = sorted(DATA_DIR.glob("Estado de Pozos*.xlsx"))
    candidates: Dict[pd.Timestamp, Dict] = {}
    log_rows = []

    for path in files:
        try:
            xl = pd.ExcelFile(path, engine="openpyxl")
        except Exception as e:
            log_rows.append({"archivo": path.name, "hoja": "", "periodo": "", "estado_carga": f"Error al abrir: {e}"})
            continue

        for sheet in xl.sheet_names:
            try:
                raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=12, engine="openpyxl")
                header_row = find_header_row(raw)
                if header_row is None:
                    continue
                periodo = infer_period(path.name, sheet, raw)
                if periodo is None:
                    continue
                df = extract_main_table(path, sheet, header_row)
                if df.empty:
                    continue
                sc = score_candidate(path, sheet, periodo, len(df))
                log_rows.append({
                    "archivo": path.name,
                    "hoja": sheet,
                    "periodo": periodo.strftime("%Y-%m"),
                    "filas": len(df),
                    "score": round(sc, 3),
                    "estado_carga": "Candidato"
                })
                if periodo not in candidates or sc > candidates[periodo]["score"]:
                    candidates[periodo] = {"df": df, "score": sc, "archivo": path.name, "hoja": sheet}
            except Exception as e:
                log_rows.append({"archivo": path.name, "hoja": sheet, "periodo": "", "estado_carga": f"Error hoja: {e}"})

    all_rows = []
    for periodo, item in sorted(candidates.items()):
        df = item["df"].copy()
        df["fecha"] = periodo
        df["periodo"] = periodo.strftime("%Y-%m")
        df["mes"] = MESES_INV[periodo.month]
        df["anio"] = periodo.year
        df["archivo_fuente"] = item["archivo"]
        df["hoja_fuente"] = item["hoja"]
        all_rows.append(df)

    if not all_rows:
        return pd.DataFrame(), pd.DataFrame(log_rows)

    base = pd.concat(all_rows, ignore_index=True)
    base["condicion"] = base.apply(lambda r: clasificar_condicion(r["ult_est"], r["estado"], r["tipo_pozo"]), axis=1)
    base["grupo_swab"] = base.apply(clasificar_swab, axis=1)
    base["mecanismo"] = base["tipo_pozo"].map(lambda x: norm_txt(x) or "SIN DATO")
    base["ult_est_clean"] = base["ult_est"].map(lambda x: norm_txt(x) or "SIN DATO")
    base["bateria_clean"] = base["bateria"].map(lambda x: norm_txt(x) or "SIN DATO")
    base["pozo_clean"] = base["pozo"].map(lambda x: norm_txt(x))

    return base, pd.DataFrame(log_rows)


def kpi_cards(df: pd.DataFrame):
    total = df["pozo_clean"].nunique()
    activos = df.loc[df["condicion"] == "Activo", "pozo_clean"].nunique()
    inactivos = df.loc[df["condicion"] == "Inactivo", "pozo_clean"].nunique()
    swab = df.loc[df["grupo_swab"] == "SWAB", "pozo_clean"].nunique()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Pozos", f"{total:,}")
    c2.metric("Activos", f"{activos:,}")
    c3.metric("Inactivos", f"{inactivos:,}")
    c4.metric("Pozos SWAB", f"{swab:,}")


def fig_linea_condicion(df: pd.DataFrame):
    t = df.groupby(["fecha", "periodo", "condicion"], as_index=False)["pozo_clean"].nunique().rename(columns={"pozo_clean": "pozos"})
    fig = px.line(t, x="fecha", y="pozos", color="condicion", markers=True,
                  color_discrete_map=COLOR_MAP, title="Evolución mensual de pozos por condición")
    fig.update_layout(template="plotly_white", height=430, legend_title_text="Condición")
    fig.update_xaxes(title="Mes", tickformat="%b %Y")
    fig.update_yaxes(title="Cantidad de pozos")
    return fig


def fig_linea_swab(df: pd.DataFrame):
    t = df.groupby(["fecha", "periodo", "grupo_swab"], as_index=False)["pozo_clean"].nunique().rename(columns={"pozo_clean": "pozos"})
    fig = px.area(t, x="fecha", y="pozos", color="grupo_swab", line_group="grupo_swab",
                  color_discrete_map=COLOR_MAP, title="Participación mensual de pozos SWAB vs No SWAB")
    fig.update_layout(template="plotly_white", height=430, legend_title_text="Grupo")
    fig.update_xaxes(title="Mes", tickformat="%b %Y")
    fig.update_yaxes(title="Cantidad de pozos")
    return fig


def fig_estado_final(df: pd.DataFrame):
    ult_fecha = df["fecha"].max()
    d = df[df["fecha"] == ult_fecha]
    t = d.groupby("ult_est_clean", as_index=False)["pozo_clean"].nunique().rename(columns={"pozo_clean": "pozos"})
    t = t.sort_values("pozos", ascending=False).head(15)
    fig = px.bar(t, x="pozos", y="ult_est_clean", orientation="h", text="pozos",
                 color="pozos", color_continuous_scale="Tealgrn",
                 title=f"Top 15 estados operativos al cierre de {ult_fecha.strftime('%b %Y')}")
    fig.update_layout(template="plotly_white", height=520, yaxis={"categoryorder": "total ascending"}, coloraxis_showscale=False)
    fig.update_xaxes(title="Cantidad de pozos")
    fig.update_yaxes(title="Último estado")
    return fig


def fig_mecanismo_condicion(df: pd.DataFrame):
    ult_fecha = df["fecha"].max()
    d = df[df["fecha"] == ult_fecha]
    t = d.groupby(["mecanismo", "condicion"], as_index=False)["pozo_clean"].nunique().rename(columns={"pozo_clean": "pozos"})
    top_mec = d["mecanismo"].value_counts().head(10).index.tolist()
    t = t[t["mecanismo"].isin(top_mec)]
    fig = px.bar(t, x="mecanismo", y="pozos", color="condicion", barmode="group", text="pozos",
                 color_discrete_map=COLOR_MAP, title=f"Condición por tipo de pozo al cierre de {ult_fecha.strftime('%b %Y')}")
    fig.update_layout(template="plotly_white", height=450, legend_title_text="Condición")
    fig.update_xaxes(title="Tipo de pozo")
    fig.update_yaxes(title="Cantidad de pozos")
    return fig


def fig_bateria_top(df: pd.DataFrame):
    ult_fecha = df["fecha"].max()
    d = df[df["fecha"] == ult_fecha]
    t = d.groupby(["bateria_clean", "condicion"], as_index=False)["pozo_clean"].nunique().rename(columns={"pozo_clean": "pozos"})
    top = d.groupby("bateria_clean")["pozo_clean"].nunique().sort_values(ascending=False).head(15).index
    t = t[t["bateria_clean"].isin(top)]
    fig = px.bar(t, x="pozos", y="bateria_clean", color="condicion", orientation="h", text="pozos",
                 color_discrete_map=COLOR_MAP, title=f"Top baterías por condición al cierre de {ult_fecha.strftime('%b %Y')}")
    fig.update_layout(template="plotly_white", height=560, yaxis={"categoryorder": "total ascending"}, legend_title_text="Condición")
    fig.update_xaxes(title="Cantidad de pozos")
    fig.update_yaxes(title="Batería")
    return fig


def fig_matriz_heatmap(df: pd.DataFrame):
    ult_fecha = df["fecha"].max()
    d = df[df["fecha"] == ult_fecha].copy()
    top_bat = d["bateria_clean"].value_counts().head(15).index
    top_estado = d["ult_est_clean"].value_counts().head(12).index
    d = d[d["bateria_clean"].isin(top_bat) & d["ult_est_clean"].isin(top_estado)]
    pivot = d.pivot_table(index="bateria_clean", columns="ult_est_clean", values="pozo_clean", aggfunc="nunique", fill_value=0)
    fig = px.imshow(pivot, text_auto=True, aspect="auto", color_continuous_scale="YlGnBu",
                    title=f"Mapa de concentración Batería vs Último estado en {ult_fecha.strftime('%b %Y')}")
    fig.update_layout(template="plotly_white", height=620, coloraxis_colorbar_title="Pozos")
    fig.update_xaxes(title="Último estado")
    fig.update_yaxes(title="Batería")
    return fig


def cambios_entre_extremos(df: pd.DataFrame) -> pd.DataFrame:
    fechas = sorted(df["fecha"].unique())
    if len(fechas) < 2:
        return pd.DataFrame()
    ini, fin = fechas[0], fechas[-1]
    a = df[df["fecha"] == ini][["pozo_clean", "pozo", "ult_est_clean", "condicion", "grupo_swab", "bateria_clean", "mecanismo"]].copy()
    b = df[df["fecha"] == fin][["pozo_clean", "ult_est_clean", "condicion", "grupo_swab", "bateria_clean", "mecanismo"]].copy()
    a = a.rename(columns={"ult_est_clean": "estado_inicial", "condicion": "condicion_inicial", "grupo_swab": "grupo_inicial", "bateria_clean": "bateria_inicial", "mecanismo": "tipo_inicial"})
    b = b.rename(columns={"ult_est_clean": "estado_final", "condicion": "condicion_final", "grupo_swab": "grupo_final", "bateria_clean": "bateria_final", "mecanismo": "tipo_final"})
    m = a.merge(b, on="pozo_clean", how="outer")
    m["cambio_estado"] = m["estado_inicial"].fillna("NO EXISTÍA") + " → " + m["estado_final"].fillna("NO EXISTE")
    m["cambio_condicion"] = m["condicion_inicial"].fillna("NO EXISTÍA") + " → " + m["condicion_final"].fillna("NO EXISTE")
    m = m[(m["estado_inicial"] != m["estado_final"]) | (m["condicion_inicial"] != m["condicion_final"])]
    return m.sort_values(["cambio_condicion", "pozo_clean"])


def fig_transiciones(df: pd.DataFrame):
    cambios = cambios_entre_extremos(df)
    if cambios.empty:
        return None
    t = cambios.groupby("cambio_condicion", as_index=False)["pozo_clean"].nunique().rename(columns={"pozo_clean": "pozos"}).sort_values("pozos", ascending=False)
    fig = px.bar(t, x="pozos", y="cambio_condicion", orientation="h", text="pozos",
                 color="pozos", color_continuous_scale="Oranges", title="Transiciones de condición entre el primer y último mes seleccionado")
    fig.update_layout(template="plotly_white", height=430, yaxis={"categoryorder": "total ascending"}, coloraxis_showscale=False)
    fig.update_xaxes(title="Cantidad de pozos")
    fig.update_yaxes(title="Transición")
    return fig


def resumen_kpi(df: pd.DataFrame) -> Dict[str, int]:
    ult = df[df["fecha"] == df["fecha"].max()]
    return {
        "periodo_inicial": df["fecha"].min().strftime("%b %Y"),
        "periodo_final": df["fecha"].max().strftime("%b %Y"),
        "pozos_final": int(ult["pozo_clean"].nunique()),
        "activos_final": int(ult.loc[ult["condicion"] == "Activo", "pozo_clean"].nunique()),
        "inactivos_final": int(ult.loc[ult["condicion"] == "Inactivo", "pozo_clean"].nunique()),
        "swab_final": int(ult.loc[ult["grupo_swab"] == "SWAB", "pozo_clean"].nunique()),
        "meses": int(df["fecha"].nunique()),
    }


def exportar_ppt(figs: List[Tuple[str, go.Figure]], kpis: Dict[str, int], cambios: pd.DataFrame) -> bytes:
    if Presentation is None:
        raise RuntimeError("No está instalado python-pptx. Revise requirements.txt")
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    def add_title(slide, title, subtitle=None):
        tx = slide.shapes.add_textbox(Inches(0.45), Inches(0.25), Inches(12.4), Inches(0.55))
        p = tx.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = RGBColor(15, 118, 110)
        if subtitle:
            sb = slide.shapes.add_textbox(Inches(0.47), Inches(0.78), Inches(12.0), Inches(0.35))
            q = sb.text_frame.paragraphs[0]
            q.text = subtitle
            q.font.size = Pt(12)
            q.font.color.rgb = RGBColor(71, 85, 105)

    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    add_title(slide, "Estado mensual de pozos OIG", f"Rango analizado: {kpis['periodo_inicial']} a {kpis['periodo_final']}")
    caja = slide.shapes.add_textbox(Inches(0.7), Inches(1.55), Inches(12.0), Inches(3.0))
    tf = caja.text_frame
    texto = [
        f"Meses analizados: {kpis['meses']}",
        f"Pozos al cierre: {kpis['pozos_final']:,}",
        f"Activos al cierre: {kpis['activos_final']:,}",
        f"Inactivos al cierre: {kpis['inactivos_final']:,}",
        f"Pozos SWAB al cierre: {kpis['swab_final']:,}",
    ]
    for i, line in enumerate(texto):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(22)
        p.font.bold = i == 0

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        for idx, (title, fig) in enumerate(figs, start=1):
            img = tmp / f"grafico_{idx}.png"
            fig.write_image(str(img), width=1500, height=850, scale=2)
            slide = prs.slides.add_slide(blank)
            add_title(slide, title)
            slide.shapes.add_picture(str(img), Inches(0.55), Inches(1.0), width=Inches(12.2), height=Inches(5.95))

        if not cambios.empty:
            slide = prs.slides.add_slide(blank)
            add_title(slide, "Pozos con cambio de estado", "Primeros 18 registros del rango seleccionado")
            top = cambios[["pozo_clean", "estado_inicial", "estado_final", "cambio_condicion"]].head(18)
            rows, cols = len(top) + 1, 4
            table = slide.shapes.add_table(rows, cols, Inches(0.55), Inches(1.2), Inches(12.2), Inches(5.6)).table
            headers = ["Pozo", "Estado inicial", "Estado final", "Cambio condición"]
            for j, h in enumerate(headers):
                table.cell(0, j).text = h
            for i, (_, r) in enumerate(top.iterrows(), start=1):
                table.cell(i, 0).text = str(r["pozo_clean"])
                table.cell(i, 1).text = str(r["estado_inicial"])
                table.cell(i, 2).text = str(r["estado_final"])
                table.cell(i, 3).text = str(r["cambio_condicion"])
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.text_frame.paragraphs:
                        p.font.size = Pt(8)
                        p.alignment = PP_ALIGN.LEFT

        out = tmp / "estado_pozos_oig_dashboard.pptx"
        prs.save(out)
        return out.read_bytes()


def descargar_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def ui():
    st.title("🛢️ Estado mensual de pozos OIG")
    st.caption("Base fija desde archivos Excel ubicados en la carpeta data. No usa cargador de archivos.")

    base, log = cargar_base()
    if base.empty:
        st.error("No se encontró información válida. Cree una carpeta data junto a app.py y coloque allí los Excel de Estado de Pozos.")
        st.stop()

    periodos = sorted(base["fecha"].drop_duplicates().tolist())
    etiquetas = {p: f"{MESES_INV[p.month]} {p.year}" for p in periodos}

    with st.expander("Ver meses reconocidos y fuentes usadas", expanded=False):
        fuentes = base.groupby(["fecha", "archivo_fuente", "hoja_fuente"], as_index=False)["pozo_clean"].nunique()
        fuentes["periodo"] = fuentes["fecha"].dt.strftime("%Y-%m")
        st.dataframe(fuentes[["periodo", "archivo_fuente", "hoja_fuente", "pozo_clean"]].rename(columns={"pozo_clean": "pozos"}), use_container_width=True)

    with st.sidebar:
        st.header("Panel de análisis")
        with st.form("form_analisis"):
            inicio = st.selectbox("Mes inicial", periodos, index=0, format_func=lambda p: etiquetas[p])
            fin = st.selectbox("Mes final", periodos, index=len(periodos) - 1, format_func=lambda p: etiquetas[p])
            grupos = st.multiselect("Grupo", ["SWAB", "No SWAB"], default=["SWAB", "No SWAB"])
            condiciones = st.multiselect("Condición", ["Activo", "Inactivo", "Observación"], default=["Activo", "Inactivo", "Observación"])
            ejecutar = st.form_submit_button("Ejecutar análisis", type="primary")

    if not ejecutar and "ejecutado" not in st.session_state:
        st.info("Seleccione el rango de meses y presione Ejecutar análisis.")
        st.stop()

    st.session_state["ejecutado"] = True
    if inicio > fin:
        st.error("El mes inicial no puede ser mayor que el mes final.")
        st.stop()

    df = base[(base["fecha"] >= inicio) & (base["fecha"] <= fin)].copy()
    df = df[df["grupo_swab"].isin(grupos) & df["condicion"].isin(condiciones)]
    if df.empty:
        st.warning("No hay registros con los filtros seleccionados.")
        st.stop()

    st.subheader(f"Resumen del rango: {etiquetas[inicio]} a {etiquetas[fin]}")
    kpi_cards(df[df["fecha"] == df["fecha"].max()])

    figs = [
        ("Evolución mensual por condición", fig_linea_condicion(df)),
        ("SWAB vs No SWAB", fig_linea_swab(df)),
        ("Top estados operativos", fig_estado_final(df)),
        ("Condición por tipo de pozo", fig_mecanismo_condicion(df)),
        ("Top baterías por condición", fig_bateria_top(df)),
        ("Matriz Batería vs Estado", fig_matriz_heatmap(df)),
    ]
    trans_fig = fig_transiciones(df)
    if trans_fig is not None:
        figs.append(("Transiciones de condición", trans_fig))

    tabs = st.tabs(["Tendencias", "Cierre mensual", "Cambios", "Base filtrada", "Exportar"])

    with tabs[0]:
        st.plotly_chart(figs[0][1], use_container_width=True)
        st.plotly_chart(figs[1][1], use_container_width=True)

    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(figs[2][1], use_container_width=True)
            st.plotly_chart(figs[3][1], use_container_width=True)
        with c2:
            st.plotly_chart(figs[4][1], use_container_width=True)
            st.plotly_chart(figs[5][1], use_container_width=True)

    cambios = cambios_entre_extremos(df)
    with tabs[2]:
        if trans_fig is not None:
            st.plotly_chart(trans_fig, use_container_width=True)
        st.dataframe(cambios, use_container_width=True, height=480)

    with tabs[3]:
        st.dataframe(df, use_container_width=True, height=520)

    with tabs[4]:
        kpis = resumen_kpi(df)
        st.download_button("Descargar base filtrada CSV", data=descargar_csv(df), file_name="base_filtrada_estado_pozos.csv", mime="text/csv")
        st.download_button("Descargar cambios CSV", data=descargar_csv(cambios), file_name="cambios_estado_pozos.csv", mime="text/csv")
        try:
            ppt_bytes = exportar_ppt(figs, kpis, cambios)
            st.download_button(
                "Descargar PowerPoint automático",
                data=ppt_bytes,
                file_name="dashboard_estado_pozos_oig.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
            )
        except Exception as e:
            st.warning(f"No se pudo generar PPT. Revise que kaleido y python-pptx estén instalados. Detalle: {e}")

        st.write("También puede exportar cada gráfico desde el menú de la figura, opción Download plot as PNG.")


if __name__ == "__main__":
    ui()
