import sqlite3
import json
import re
import requests
import pandas as pd
import streamlit as st
import numpy as np
import math
from io import BytesIO
from datetime import datetime
from pathlib import Path

LM_STUDIO_URL = "http://localhost:1234/v1/chat/completions"
LM_MODEL_ID = "qwen/qwen3-14b"


def consultar_llm_local(system_prompt: str, user_prompt: str, timeout: float = 3.0, temperature: float = 0.15) -> tuple[str, bool]:
    """
    Executes inference against local LM Studio server.
    Returns (text_response, is_llm_success_bool). Never raises an unhandled exception.
    """
    payload = {
        "model": LM_MODEL_ID,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": temperature,
        "max_tokens": 400,
        "stream": False
    }
    try:
        response = requests.post(LM_STUDIO_URL, json=payload, headers={"Content-Type": "application/json"}, timeout=timeout)
        if response.status_code == 200:
            data = response.json()
            content = data["choices"][0]["message"]["content"].strip()
            content = re.sub(r"<think>.*?</think>", "", content, flags=re.DOTALL).strip() or content
            if content:
                return content, True
    except (requests.exceptions.RequestException, KeyError, IndexError, ValueError, Exception):
        pass
    return "", False


def generar_cronica_contingencia_estadistica(home: str, away: str, xg_home: float, xg_away: float, cuota_mercado: float, cuota_justa: float, arbitro: str, media_tarjetas: float, friccion_faltas: float) -> str:
    diff_cuota = cuota_mercado - cuota_justa
    ev_estado = f"Ventaja matemática (+EV de {diff_cuota:+.2f})" if diff_cuota > 0.08 else "Línea alineada con mercado (Sin valor operativo)"
    disciplina = "Perfil severo" if media_tarjetas >= 4.8 else ("Perfil permisivo" if media_tarjetas <= 3.8 else "Perfil moderado")
    
    return f"""[ANÁLISIS DE EFICIENCIA OFENSIVA Y MÉTRICAS xG]
El emparejamiento proyecta una media de {xg_home:.2f} xG para {home} y {xg_away:.2f} xG para {away}. El diferencial ofensivo indica un volumen acumulado de {xg_home + xg_away:.2f} goles esperados, situando la línea teórica de mayor estabilidad en torno al Over/Under 2.5 según la distribución estocástica.

[PERFIL DISCIPLINAR Y FRICCIÓN ARBITRAL]
La fricción combinada entre contendientes proyecta {friccion_faltas:.1f} faltas. El colegiado {arbitro} registra una media de {media_tarjetas:.2f} amarillas por encuentro ({disciplina}), lo que sitúa la expectativa disciplinar en un rango {'alto' if media_tarjetas >= 4.8 else 'estándar'}.

[VEREDICTO CUANTITATIVO Y CONCLUSIÓN DE MERCADO]
Cuota justa modelada: {cuota_justa:.2f} | Cuota de mercado: {cuota_mercado:.2f}. Diagnóstico de eficiencia: {ev_estado}. Se recomienda priorizar la disciplina del bankroll frente a escenarios con dispersión elevada en finalización.""".strip()


@st.cache_data(show_spinner=False, ttl=3600)
def obtener_cronica_partido(match_id: str, prompt_system: str, prompt_user: str, fallback_data: dict) -> tuple[str, str]:
    texto, exito = consultar_llm_local(prompt_system, prompt_user, timeout=3.5)
    if exito and texto:
        return texto, "🟢 Inferencia Local GPU (Qwen 14B)"
    texto_fb = generar_cronica_contingencia_estadistica(**fallback_data)
    return texto_fb, "🟡 Modelo Matemático Estadístico (Fallback)"

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    HRFlowable,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

try:
    from conector_odds import escaneo_apuestas_ev
except ImportError:
    escaneo_apuestas_ev = None

try:
    from updater import sincronizar_jornada
    _UPDATER_DISPONIBLE = True
except ImportError:
    sincronizar_jornada = None
    _UPDATER_DISPONIBLE = False

st.set_page_config(
    page_title="Hugo's Hub - Sistema Analítico Completo",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #f0f2f6; }
    sidebar .stApp { background-color: #16192b; }
    .metric-card {
        background-color: #1a1c24;
        border: 1px solid #2d3748;
        padding: 15px;
        border-radius: 10px;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

RUTA_BD = Path(__file__).parent / "macro_base_futbol.db"

MAPA_LIGAS = {
    "SP1": "LaLiga (España)",
    "SP2": "LaLiga 2 (España)",
    "E0": "Premier League (Inglaterra)",
    "E1": "Championship (Inglaterra)",
    "D1": "Bundesliga (Alemania)",
    "I1": "Serie A (Italia)",
    "F1": "Ligue 1 (Francia)"
}
MAPA_INV = {v: k for k, v in MAPA_LIGAS.items()}

@st.cache_data(ttl=3600)
def inspeccionar_tablas_bd():
    if not RUTA_BD.exists(): return []
    with sqlite3.connect(RUTA_BD) as conn:
        try: return pd.read_sql("SELECT name FROM sqlite_master WHERE type='table';", conn)['name'].tolist()
        except: return []

@st.cache_data(ttl=3600)
def cargar_datos_partidos():
    tablas = inspeccionar_tablas_bd()
    nombre_tabla = 'historico_reciente' if 'historico_reciente' in tablas else ('partidos' if 'partidos' in tablas else None)
    if not nombre_tabla: return pd.DataFrame(), "Falta la tabla de partidos."
    with sqlite3.connect(RUTA_BD) as conn:
        try:
            df = pd.read_sql(f"SELECT * FROM {nombre_tabla}", conn)
            df.columns = df.columns.str.replace('ï»¿', '').str.strip()
            return df, None
        except Exception as e: return pd.DataFrame(), str(e)

@st.cache_data(ttl=3600)
def cargar_datos_jugadores():
    if 'estadisticas_jugadores' not in inspeccionar_tablas_bd(): return pd.DataFrame(), "Falta la tabla."
    with sqlite3.connect(RUTA_BD) as conn:
        try: return pd.read_sql("SELECT * FROM estadisticas_jugadores", conn), None
        except Exception as e: return pd.DataFrame(), str(e)


@st.cache_data(ttl=3600)
def cargar_datos_arbitros(min_partidos: int = 3):
    """
    Lee la tabla 'arbitros_stats' desde macro_base_futbol.db filtrando por un mínimo
    de partidos pitados. Si la tabla aún no existe, la genera automáticamente a partir
    de 'historico_reciente' (Referee, HY, AY, HR, AR, HF, AF).
    """
    if not RUTA_BD.exists():
        return pd.DataFrame()

    with sqlite3.connect(RUTA_BD) as conn:
        try:
            tablas = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table';", conn)["name"].tolist()
            if "arbitros_stats" not in tablas and "historico_reciente" in tablas:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(historico_reciente)")
                cols_info = [r[1] for r in cur.fetchall()]
                col_div = next((c for c in cols_info if "Div" in c), "Div")
                tiene_hf = "HF" in cols_info and "AF" in cols_info
                expr_faltas = (
                    "ROUND(SUM(COALESCE(HF, 0) + COALESCE(AF, 0)) * 1.0 / COUNT(*), 2)"
                    if tiene_hf
                    else "0.0"
                )
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS arbitros_stats (
                        "Árbitro"               TEXT NOT NULL,
                        "Div"                   TEXT NOT NULL,
                        "Liga"                  TEXT NOT NULL,
                        "Comité"                TEXT NOT NULL,
                        "Partidos"              INTEGER NOT NULL,
                        "Perfil Disciplinario"  TEXT NOT NULL,
                        "Media Amarillas Est."  REAL NOT NULL,
                        "Media Rojas"           REAL NOT NULL,
                        "Media Faltas"          REAL NOT NULL,
                        PRIMARY KEY ("Árbitro", "Div")
                    )
                    """
                )
                filas = cur.execute(
                    f"""
                    SELECT
                        TRIM(Referee) AS Arbitro,
                        TRIM("{col_div}") AS Div_Code,
                        COUNT(*) AS Partidos,
                        ROUND(SUM(COALESCE(HY, 0) + COALESCE(AY, 0)) * 1.0 / COUNT(*), 2) AS Media_Amarillas,
                        ROUND(SUM(COALESCE(HR, 0) + COALESCE(AR, 0)) * 1.0 / COUNT(*), 2) AS Media_Rojas,
                        {expr_faltas} AS Media_Faltas
                    FROM historico_reciente
                    WHERE Referee IS NOT NULL AND TRIM(Referee) != '' AND LOWER(TRIM(Referee)) != 'nan'
                    GROUP BY TRIM(Referee), TRIM("{col_div}")
                    """
                ).fetchall()
                for arb, div_c, part, m_am, m_roj, m_fal in filas:
                    perfil = "Alto" if m_am > 5.0 else ("Medio" if m_am >= 4.0 else "Permisivo")
                    liga_nom = MAPA_LIGAS.get(div_c, div_c)
                    cur.execute(
                        """
                        INSERT OR REPLACE INTO arbitros_stats (
                            "Árbitro", "Div", "Liga", "Comité", "Partidos",
                            "Perfil Disciplinario", "Media Amarillas Est.", "Media Rojas", "Media Faltas"
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (arb, div_c, liga_nom, "Comité Oficial", int(part), perfil, float(m_am), float(m_roj), float(m_fal)),
                    )
                conn.commit()

            df_arb = pd.read_sql(
                """
                SELECT
                    "Árbitro",
                    "Div",
                    "Liga",
                    "Comité",
                    "Partidos",
                    "Perfil Disciplinario",
                    "Media Amarillas Est.",
                    "Media Rojas",
                    "Media Faltas"
                FROM arbitros_stats
                WHERE "Partidos" >= ?
                ORDER BY "Partidos" DESC, "Media Amarillas Est." DESC
                """,
                conn,
                params=(int(min_partidos),),
            )
            return df_arb
        except Exception:
            return pd.DataFrame()

def calcular_probabilidad_poisson(l, x):
    return (math.exp(-l) * (l**x)) / math.factorial(x)


# ══════════════════════════════════════════════════════════════════════════════
# MODELO: MATCHUP DISCIPLINARIO Y VALOR ESPERADO EN MERCADOS DE TARJETAS
# ══════════════════════════════════════════════════════════════════════════════
LINEAS_TARJETAS = [3.5, 4.5, 5.5, 6.5]
MEDIA_TARJETAS_FALLBACK_LIGA = 4.5   # tarjetas/partido si no hay datos de liga
MIN_PARTIDOS_CONDICION = 5           # mínimo de partidos local/visitante por equipo
VENTANA_PARTIDOS_RECIENTES = 10      # últimos N partidos por condición
MIN_PARTIDOS_ARBITRO = 5             # por debajo, el índice se encoge hacia 1.0


def prob_poisson_acumulada(k, lam):
    """P(X <= k) para X ~ Poisson(lam). Seguro ante lam <= 0."""
    if lam is None or lam <= 0:
        return 1.0
    return min(1.0, sum(calcular_probabilidad_poisson(lam, i) for i in range(int(k) + 1)))


def cuota_justa_teorica(prob):
    """Cuota justa = 1 / probabilidad. Devuelve None si la probabilidad es ~0."""
    if prob is None or prob <= 1e-9:
        return None
    return 1.0 / prob


def calcular_lineas_tarjetas(lambda_total, lineas=LINEAS_TARJETAS):
    """Probabilidades y cuotas justas Over/Under para cada línea (x.5)."""
    filas = []
    for linea in lineas:
        p_under = prob_poisson_acumulada(math.floor(linea), lambda_total)
        p_over = max(0.0, 1.0 - p_under)
        filas.append({
            "linea": linea,
            "p_over": p_over,
            "cuota_over": cuota_justa_teorica(p_over),
            "p_under": p_under,
            "cuota_under": cuota_justa_teorica(p_under),
        })
    return filas


def calcular_ev_kelly(prob, cuota_bookie, fraccion_kelly=0.25):
    """
    EV = prob * cuota - 1.  Kelly completo f* = (prob*cuota - 1) / (cuota - 1);
    se devuelve la fracción (por defecto 1/4 Kelly) como proporción del bankroll.
    """
    if prob is None or cuota_bookie is None or cuota_bookie <= 1.0:
        return 0.0, 0.0
    ev = prob * cuota_bookie - 1.0
    if ev <= 0:
        return ev, 0.0
    kelly_completo = ev / (cuota_bookie - 1.0)
    return ev, max(0.0, kelly_completo * fraccion_kelly)


def _columna_liga_partidos(df):
    for c in ("Div", "Liga_Origen"):
        if c in df.columns:
            return c
    return None


def _ordenar_por_fecha(df):
    if "Date" in df.columns and not df.empty:
        fechas = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)
        return df.assign(_f=fechas).sort_values("_f", kind="stable").drop(columns="_f")
    return df


def media_global_tarjetas_liga(df_arb_liga, df_liga):
    """
    Media disciplinaria global de la liga (amarillas + rojas por partido).
    Prioridad: arbitros_stats ponderada por partidos > historico_reciente > fallback.
    """
    if df_arb_liga is not None and not df_arb_liga.empty:
        pesos = pd.to_numeric(df_arb_liga["Partidos"], errors="coerce").fillna(0)
        medias = (
            pd.to_numeric(df_arb_liga["Media Amarillas Est."], errors="coerce").fillna(0)
            + pd.to_numeric(df_arb_liga["Media Rojas"], errors="coerce").fillna(0)
        )
        if pesos.sum() > 0:
            valor = float((medias * pesos).sum() / pesos.sum())
            if valor > 0:
                return valor
    if df_liga is not None and not df_liga.empty and {"HY", "AY"}.issubset(df_liga.columns):
        tot = (
            pd.to_numeric(df_liga["HY"], errors="coerce").fillna(0)
            + pd.to_numeric(df_liga["AY"], errors="coerce").fillna(0)
        )
        if len(tot) > 0 and tot.mean() > 0:
            return float(tot.mean())
    return MEDIA_TARJETAS_FALLBACK_LIGA


def calcular_indice_arbitro(media_arbitro, partidos_arbitro, media_liga):
    """
    Indice_Arbitro = Media_Tarjetas_Arbitro / Media_Global_Liga.
    Con pocos partidos se encoge hacia 1.0; divisiones por cero devuelven 1.0.
    """
    if not media_liga or media_liga <= 0 or media_arbitro is None or media_arbitro <= 0:
        return 1.0
    indice = media_arbitro / media_liga
    n = partidos_arbitro if partidos_arbitro else 0
    if n < MIN_PARTIDOS_ARBITRO:
        indice = 1.0 + (indice - 1.0) * (n / MIN_PARTIDOS_ARBITRO)
    return indice


def tendencia_equipo_tarjetas(df_liga, equipo, condicion):
    """
    Tendencia disciplinaria de un equipo en una condición ('local' o 'visitante').
      - recibidas:  local -> HY jugando en casa; visitante -> AY jugando fuera.
      - provocadas: local -> AY (rival visitante); visitante -> HY (rival local).
    Si hay menos de MIN_PARTIDOS_CONDICION partidos, usa su media general.
    Devuelve dict(recibidas, provocadas, n, fuente) (valores None sin datos).
    """
    vacio = {"recibidas": None, "provocadas": None, "n": 0, "fuente": "sin datos"}
    if df_liga is None or df_liga.empty or not {"HY", "AY", "HomeTeam", "AwayTeam"}.issubset(df_liga.columns):
        return vacio

    casa = _ordenar_por_fecha(df_liga[df_liga["HomeTeam"] == equipo]).copy()
    fuera = _ordenar_por_fecha(df_liga[df_liga["AwayTeam"] == equipo]).copy()
    for d in (casa, fuera):
        d["HY"] = pd.to_numeric(d["HY"], errors="coerce")
        d["AY"] = pd.to_numeric(d["AY"], errors="coerce")

    if condicion == "local":
        sub = casa.dropna(subset=["HY", "AY"]).tail(VENTANA_PARTIDOS_RECIENTES)
        rec_col, prov_col = "HY", "AY"
    else:
        sub = fuera.dropna(subset=["HY", "AY"]).tail(VENTANA_PARTIDOS_RECIENTES)
        rec_col, prov_col = "AY", "HY"

    if len(sub) >= MIN_PARTIDOS_CONDICION:
        return {
            "recibidas": float(sub[rec_col].mean()),
            "provocadas": float(sub[prov_col].mean()),
            "n": len(sub),
            "fuente": f"últimos {len(sub)} como {condicion}",
        }

    # Fallback: media general (casa + fuera) del equipo
    casa_v = casa.dropna(subset=["HY", "AY"])
    fuera_v = fuera.dropna(subset=["HY", "AY"])
    n_total = len(casa_v) + len(fuera_v)
    if n_total == 0:
        return vacio
    recibidas = (casa_v["HY"].sum() + fuera_v["AY"].sum()) / n_total
    provocadas = (casa_v["AY"].sum() + fuera_v["HY"].sum()) / n_total
    return {
        "recibidas": float(recibidas),
        "provocadas": float(provocadas),
        "n": n_total,
        "fuente": f"media general ({n_total} partidos)",
    }


def _mezclar_tendencias(recibidas_propias, provocadas_por_rival, base_liga):
    """Combina lo que recibe el equipo con lo que provoca el rival (50/50)."""
    valores = [v for v in (recibidas_propias, provocadas_por_rival) if v is not None]
    if not valores:
        return base_liga / 2.0
    return float(sum(valores) / len(valores))


def render_matchup_disciplinario(df_partidos_base):
    """Interfaz Streamlit del modelo de Matchup Disciplinario (Poisson + EV + Kelly)."""
    st.subheader("🟨 Modelo Cuantitativo: Matchup Disciplinario y Valor Esperado")
    st.caption(
        "Lambda = media de tarjetas del equipo × Índice de Rigor del árbitro. "
        "Distribución de Poisson sobre las tarjetas totales del partido."
    )

    col_liga = _columna_liga_partidos(df_partidos_base)
    if col_liga is None or "HY" not in df_partidos_base.columns:
        st.warning("La tabla de partidos no contiene datos de tarjetas (HY/AY).")
        return

    divs_disponibles = [d for d in MAPA_LIGAS if d in set(df_partidos_base[col_liga].dropna().astype(str))]
    if not divs_disponibles:
        st.warning("No hay ligas con datos de tarjetas disponibles.")
        return
    ligas_nombres = [MAPA_LIGAS[d] for d in divs_disponibles]

    # ── Selectores sincronizados: Liga -> Local -> Visitante -> Árbitro ──────
    s1, s2, s3, s4 = st.columns(4)
    with s1:
        liga_sel = st.selectbox("🏆 Liga:", ligas_nombres, key="disc_liga")
    div_sel = MAPA_INV.get(liga_sel)
    df_liga = df_partidos_base[df_partidos_base[col_liga].astype(str) == div_sel].copy()
    equipos_liga = sorted(
        pd.concat([df_liga["HomeTeam"], df_liga["AwayTeam"]]).dropna().unique().tolist()
    )
    if len(equipos_liga) < 2:
        st.warning("No hay suficientes equipos en esta liga.")
        return

    with s2:
        eq_local = st.selectbox("🏠 Equipo Local:", equipos_liga, index=0, key=f"disc_local_{div_sel}")
    equipos_visitante = [e for e in equipos_liga if e != eq_local]
    with s3:
        eq_visit = st.selectbox("✈️ Equipo Visitante:", equipos_visitante, index=0, key=f"disc_visit_{div_sel}")

    df_arb_all = cargar_datos_arbitros(min_partidos=1)
    df_arb_liga = pd.DataFrame()
    if not df_arb_all.empty:
        df_arb_liga = df_arb_all[df_arb_all["Div"] == div_sel].copy()
    nombres_arb = df_arb_liga["Árbitro"].tolist() if not df_arb_liga.empty else []
    with s4:
        if nombres_arb:
            arbitro_sel = st.selectbox("⚖️ Árbitro:", nombres_arb, key=f"disc_arb_{div_sel}")
        else:
            arbitro_sel = None
            st.selectbox("⚖️ Árbitro:", ["Sin árbitros en esta liga"], key=f"disc_arb_vacio_{div_sel}", disabled=True)

    # ── Métricas base ────────────────────────────────────────────────────────
    media_liga = media_global_tarjetas_liga(df_arb_liga, df_liga)
    media_arbitro, partidos_arb = None, 0
    if arbitro_sel is not None:
        fila_arb = df_arb_liga[df_arb_liga["Árbitro"] == arbitro_sel].iloc[0]
        media_arbitro = float(fila_arb["Media Amarillas Est."]) + float(fila_arb["Media Rojas"])
        partidos_arb = int(fila_arb["Partidos"])
    indice_arb = calcular_indice_arbitro(media_arbitro, partidos_arb, media_liga)

    t_local = tendencia_equipo_tarjetas(df_liga, eq_local, "local")
    t_visit = tendencia_equipo_tarjetas(df_liga, eq_visit, "visitante")
    media_local_casa = _mezclar_tendencias(t_local["recibidas"], t_visit["provocadas"], media_liga)
    media_visit_fuera = _mezclar_tendencias(t_visit["recibidas"], t_local["provocadas"], media_liga)

    lambda_local = media_local_casa * indice_arb
    lambda_visit = media_visit_fuera * indice_arb
    lambda_total = lambda_local + lambda_visit

    if lambda_total >= 6.0:
        badge = "🔥 Partido Caliente"
    elif lambda_total <= 4.0:
        badge = "🧊 Bajo en Tarjetas"
    else:
        badge = "⚖️ Estándar"
    desv_pct = (indice_arb - 1.0) * 100.0

    # ── KPIs ─────────────────────────────────────────────────────────────────
    st.markdown("---")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric(f"Tarjetas {eq_local} (λ Local)", f"{lambda_local:.2f}", t_local["fuente"], delta_color="off")
    k2.metric(f"Tarjetas {eq_visit} (λ Visitante)", f"{lambda_visit:.2f}", t_visit["fuente"], delta_color="off")
    k3.metric("Tarjetas Totales (λ Total)", f"{lambda_total:.2f}", badge, delta_color="off")
    k4.metric(
        "Rigor del Árbitro",
        f"{indice_arb:.2f}x",
        f"{desv_pct:+.0f}% sobre la media",
        delta_color="inverse",
    )
    if arbitro_sel is not None and partidos_arb < MIN_PARTIDOS_ARBITRO:
        st.caption(f"⚠️ {arbitro_sel} solo tiene {partidos_arb} partido(s): su índice se ha suavizado hacia 1.00.")
    st.caption(
        f"Media global de la liga: **{media_liga:.2f}** tarjetas/partido · "
        f"Media árbitro: **{media_arbitro if media_arbitro is not None else float('nan'):.2f}**"
    )

    # ── Tabla comparativa de líneas ──────────────────────────────────────────
    st.markdown("#### 📊 Líneas Over / Under de Tarjetas")
    lineas = calcular_lineas_tarjetas(lambda_total)
    tabla = []
    for fila in lineas:
        tabla.append({
            "Línea": f"+{fila['linea']}",
            "Prob. Over (%)": round(fila["p_over"] * 100, 1),
            "Cuota Justa Over": round(fila["cuota_over"], 2) if fila["cuota_over"] else None,
            "Prob. Under (%)": round(fila["p_under"] * 100, 1),
            "Cuota Justa Under": round(fila["cuota_under"], 2) if fila["cuota_under"] else None,
        })
    st.dataframe(pd.DataFrame(tabla), use_container_width=True, hide_index=True)

    # ── Calculadora de valor (+EV) y Kelly ───────────────────────────────────
    st.markdown("#### 💰 Calculadora de Valor (+EV) y Kelly Fraccional (1/4)")
    v1, v2, v3, v4 = st.columns(4)
    with v1:
        linea_sel = st.selectbox("Línea:", LINEAS_TARJETAS, format_func=lambda x: f"+{x}", index=1, key="disc_ev_linea")
    with v2:
        lado_sel = st.radio("Mercado:", ["Over", "Under"], horizontal=True, key="disc_ev_lado")
    with v3:
        cuota_bookie = st.number_input("Cuota Bookie:", min_value=1.01, value=1.90, step=0.05, key="disc_ev_cuota")
    with v4:
        bankroll = st.number_input("Bankroll (€):", min_value=1.0, value=1000.0, step=100.0, key="disc_ev_bankroll")

    fila_sel = next(f for f in lineas if f["linea"] == linea_sel)
    prob_sel = fila_sel["p_over"] if lado_sel == "Over" else fila_sel["p_under"]
    cuota_justa_sel = fila_sel["cuota_over"] if lado_sel == "Over" else fila_sel["cuota_under"]
    ev, frac_kelly = calcular_ev_kelly(prob_sel, cuota_bookie, fraccion_kelly=0.25)

    e1, e2, e3 = st.columns(3)
    e1.metric("Probabilidad del Modelo", f"{prob_sel * 100:.1f}%")
    e2.metric("Cuota Justa", f"{cuota_justa_sel:.2f}" if cuota_justa_sel else "N/D")
    e3.metric("EV (%)", f"{ev * 100:+.2f}%", delta_color="normal")

    if ev > 0:
        st.success(
            f"✅ **Apuesta con valor:** {lado_sel} {linea_sel} @ {cuota_bookie:.2f}. "
            f"Stake recomendado (1/4 Kelly): **{frac_kelly * 100:.2f}%** del bankroll "
            f"= **{frac_kelly * bankroll:.2f} €**."
        )
    else:
        st.error(f"❌ Sin valor: EV {ev * 100:+.2f}%. No se recomienda apostar a esta cuota.")



def calcular_pesos_time_decay(df_sub, xi=0.010, fecha_ref=None):
    """
    Calcula el vector de pesos exponenciales w = exp(-xi * delta_dias)
    para cada partido respecto a la fecha más reciente (o fecha_ref).
    """
    if df_sub is None or df_sub.empty:
        return np.array([])
    if xi <= 0 or "Date" not in df_sub.columns:
        return np.ones(len(df_sub), dtype=float)

    fechas = pd.to_datetime(df_sub["Date"], errors="coerce", dayfirst=True)
    if fechas.isna().all():
        return np.ones(len(df_sub), dtype=float)

    if fecha_ref is None:
        fecha_ref = fechas.max()
    else:
        fecha_ref = pd.to_datetime(fecha_ref, errors="coerce", dayfirst=True)
        if pd.isna(fecha_ref):
            fecha_ref = fechas.max()

    delta_dias = (fecha_ref - fechas).dt.days.clip(lower=0)
    mediana_dias = delta_dias.median() if not delta_dias.dropna().empty else 0.0
    delta_dias = delta_dias.fillna(mediana_dias)
    pesos = np.exp(-float(xi) * delta_dias.to_numpy(dtype=float))
    return pesos


def calcular_media_ponderada_poisson(df_sub, col_goles, xi=0.010, default=1.25, fecha_ref=None):
    """
    Calcula la media de goles ponderada temporalmente con np.average(..., weights=w).
    Si xi <= 0, devuelve la media aritmética simple (.mean()).
    """
    if df_sub is None or df_sub.empty or col_goles not in df_sub.columns:
        return default

    df_val = df_sub[df_sub[col_goles].notna()].copy()
    if df_val.empty:
        return default

    valores = pd.to_numeric(df_val[col_goles], errors="coerce").dropna()
    if valores.empty:
        return default

    if xi <= 0 or "Date" not in df_val.columns:
        return float(valores.mean())

    df_val = df_val.loc[valores.index]
    w = calcular_pesos_time_decay(df_val, xi=xi, fecha_ref=fecha_ref)
    if len(w) != len(valores) or np.sum(w) <= 0:
        return float(valores.mean())

    return float(np.average(valores.to_numpy(dtype=float), weights=w))


_FUENTES_PDF = None
COLOR_NAVY = colors.HexColor("#0D1B2A")
COLOR_SLATE_HEADER = colors.HexColor("#1E293B")
COLOR_EMERALD = colors.HexColor("#059669")
COLOR_CRIMSON = colors.HexColor("#DC2626")
COLOR_TEAL = colors.HexColor("#1B9AAA")
COLOR_HEADER = colors.HexColor("#1A365D")
COLOR_ROW = colors.HexColor("#F8FAFC")
COLOR_LINE = colors.HexColor("#CBD5E1")


def _fuentes_pdf():
    global _FUENTES_PDF
    if _FUENTES_PDF:
        return _FUENTES_PDF
    arial = Path(r"C:\Windows\Fonts\arial.ttf")
    arial_bd = Path(r"C:\Windows\Fonts\arialbd.ttf")
    consola = Path(r"C:\Windows\Fonts\consola.ttf")
    consola_bd = Path(r"C:\Windows\Fonts\consolab.ttf")
    if arial.exists():
        pdfmetrics.registerFont(TTFont("HubSans", str(arial)))
        pdfmetrics.registerFont(TTFont("HubSans-Bold", str(arial_bd if arial_bd.exists() else arial)))
        sans_reg, sans_bold = "HubSans", "HubSans-Bold"
    else:
        sans_reg, sans_bold = "Helvetica", "Helvetica-Bold"

    if consola.exists():
        pdfmetrics.registerFont(TTFont("HubMono", str(consola)))
        pdfmetrics.registerFont(TTFont("HubMono-Bold", str(consola_bd if consola_bd.exists() else consola)))
        mono_reg, mono_bold = "HubMono", "HubMono-Bold"
    else:
        mono_reg, mono_bold = "Courier", "Courier-Bold"

    _FUENTES_PDF = (sans_reg, sans_bold, mono_reg, mono_bold)
    return _FUENTES_PDF


def _slug_archivo(texto):
    seguro = "".join(c if c.isalnum() or c in "._- " else "_" for c in str(texto))
    return seguro.strip().replace(" ", "_")[:50] or "informe"


def _estilos_parrafo_pdf():
    regular, bold, mono_reg, mono_bold = _fuentes_pdf()
    return {
        "marca": ParagraphStyle("hub_marca", fontName=bold, fontSize=10, textColor=COLOR_EMERALD, alignment=TA_LEFT, spaceAfter=2),
        "titulo": ParagraphStyle("hub_titulo", fontName=bold, fontSize=16, textColor=COLOR_SLATE_HEADER, alignment=TA_LEFT, spaceAfter=3, leading=19),
        "subtitulo": ParagraphStyle("hub_sub", fontName=regular, fontSize=8.5, textColor=colors.HexColor("#475569"), alignment=TA_LEFT, spaceAfter=6),
        "seccion": ParagraphStyle("hub_sec", fontName=bold, fontSize=11, textColor=COLOR_SLATE_HEADER, alignment=TA_LEFT, spaceBefore=7, spaceAfter=4),
        "cuerpo": ParagraphStyle("hub_cuerpo", fontName=regular, fontSize=8.5, textColor=COLOR_SLATE_HEADER, leading=12, spaceAfter=3),
        "celda": ParagraphStyle("hub_celda", fontName=regular, fontSize=8.5, textColor=COLOR_SLATE_HEADER, leading=10.5, alignment=TA_LEFT),
        "celda_c": ParagraphStyle("hub_celda_c", fontName=regular, fontSize=8.5, textColor=COLOR_SLATE_HEADER, leading=10.5, alignment=TA_CENTER),
        "celda_mono": ParagraphStyle("hub_celda_mono", fontName=mono_reg, fontSize=8.5, textColor=COLOR_SLATE_HEADER, leading=10.5, alignment=TA_CENTER),
        "celda_mono_b": ParagraphStyle("hub_celda_mono_b", fontName=mono_bold, fontSize=8.5, textColor=COLOR_SLATE_HEADER, leading=10.5, alignment=TA_CENTER),
        "celda_h": ParagraphStyle("hub_celda_h", fontName=bold, fontSize=8.5, textColor=colors.white, leading=10.5, alignment=TA_CENTER),
        "kpi_label": ParagraphStyle("hub_kpi_lbl", fontName=bold, fontSize=7.5, textColor=colors.HexColor("#64748B"), leading=9.5, alignment=TA_CENTER),
        "kpi_val": ParagraphStyle("hub_kpi_val", fontName=mono_bold, fontSize=11, textColor=COLOR_SLATE_HEADER, leading=13, alignment=TA_CENTER),
        "kpi_sub": ParagraphStyle("hub_kpi_sub", fontName=regular, fontSize=7.5, textColor=COLOR_EMERALD, leading=9.5, alignment=TA_CENTER),
        "pie": ParagraphStyle("hub_pie", fontName=regular, fontSize=7.5, textColor=colors.HexColor("#64748B"), alignment=TA_CENTER),
    }


def _tabla_pdf(encabezados, filas, col_widths, col_alignments=None, mono_cols=None):
    estilos = _estilos_parrafo_pdf()
    mono_set = set(mono_cols or [])
    data = [[Paragraph(str(h), estilos["celda_h"]) for h in encabezados]]
    for fila in filas:
        row_cells = []
        for idx, v in enumerate(fila):
            if idx in mono_set:
                st_cell = estilos["celda_mono_b"] if idx == len(fila) - 1 else estilos["celda_mono"]
            elif col_alignments and idx < len(col_alignments) and col_alignments[idx] == "CENTER":
                st_cell = estilos["celda_c"]
            else:
                st_cell = estilos["celda"]
            row_cells.append(Paragraph(str(v), st_cell))
        data.append(row_cells)

    tabla = Table(data, colWidths=col_widths, repeatRows=1)
    estilo = [
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_SLATE_HEADER),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.35, COLOR_LINE),
        ("LINEBELOW", (0, 0), (-1, 0), 1.0, COLOR_EMERALD),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]
    for i in range(1, len(data)):
        bg_color = COLOR_ROW if i % 2 == 0 else colors.white
        estilo.append(("BACKGROUND", (0, i), (-1, i), bg_color))
    tabla.setStyle(TableStyle(estilo))
    return tabla


def _pie_pdf(canvas, doc):
    canvas.saveState()
    regular, bold, _, _ = _fuentes_pdf()
    ancho, alto = doc.pagesize
    canvas.setFillColor(COLOR_SLATE_HEADER)
    canvas.rect(0, alto - 0.32 * cm, ancho, 0.32 * cm, fill=1, stroke=0)
    canvas.setFillColor(COLOR_EMERALD)
    canvas.rect(0, alto - 0.38 * cm, ancho, 0.06 * cm, fill=1, stroke=0)
    canvas.setFillColor(COLOR_SLATE_HEADER)
    canvas.rect(0, 0, ancho, 0.24 * cm, fill=1, stroke=0)
    canvas.setFont(bold, 7.5)
    canvas.setFillColor(COLOR_SLATE_HEADER)
    canvas.drawString(1.4 * cm, 0.42 * cm, "HUGO'S HUB  |  QUANTITATIVE INTELLIGENCE & SPORTS ANALYTICS  ·  CONFIDENCIAL")
    canvas.setFont(regular, 7.5)
    canvas.setFillColor(colors.HexColor("#64748B"))
    total_pages = getattr(canvas, "_total_pages", doc.page)
    canvas.drawRightString(ancho - 1.4 * cm, 0.42 * cm, f"Página {doc.page} de {total_pages}")
    canvas.restoreState()


def generar_excel_escaner_ev(df_ops, liga):
    buffer = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Escáner EV+"

    fill_titulo = PatternFill("solid", fgColor="0D1B2A")
    fill_sub = PatternFill("solid", fgColor="1A365D")
    fill_header = PatternFill("solid", fgColor="1B9AAA")
    fill_alt = PatternFill("solid", fgColor="E6FFFA")
    font_titulo = Font(name="Calibri", bold=True, color="FFFFFF", size=16)
    font_sub = Font(name="Calibri", color="E2E8F0", size=10)
    font_header = Font(name="Calibri", bold=True, color="0D1B2A", size=11)
    font_celda = Font(name="Calibri", size=10)
    borde = Border(
        left=Side(style="thin", color="CBD5E0"),
        right=Side(style="thin", color="CBD5E0"),
        top=Side(style="thin", color="CBD5E0"),
        bottom=Side(style="thin", color="CBD5E0"),
    )
    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)

    cols = [c for c in [
        "Nivel_Valor", "Fecha / Hora", "Partido", "Mercado", "Casa",
        "Cuota", "Prob_Modelo", "Cuota_Justa", "EV", "Stake_Recomendado"
    ] if c in df_ops.columns]
    n_cols = max(len(cols), 1)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c_titulo = ws.cell(1, 1, "Hugo's Hub  |  Informe de Escáner EV+")
    c_titulo.font = font_titulo
    c_titulo.fill = fill_titulo
    c_titulo.alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    c_sub = ws.cell(2, 1, f"Competición: {liga}   ·   Generado: {fecha}   ·   Oportunidades: {len(df_ops)}")
    c_sub.font = font_sub
    c_sub.fill = fill_sub
    c_sub.alignment = Alignment(horizontal="left", vertical="center")

    for col_idx in range(1, n_cols + 1):
        ws.cell(1, col_idx).fill = fill_titulo
        ws.cell(2, col_idx).fill = fill_sub

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 22

    for col_idx, nombre in enumerate(cols, 1):
        celda = ws.cell(3, col_idx, nombre.replace("_", " "))
        celda.font = font_header
        celda.fill = fill_header
        celda.alignment = centro
        celda.border = borde

    for row_idx, fila in enumerate(df_ops[cols].itertuples(index=False), 4):
        for col_idx, valor in enumerate(fila, 1):
            celda = ws.cell(row_idx, col_idx, valor)
            celda.font = font_celda
            celda.alignment = centro
            celda.border = borde
            if row_idx % 2 == 0:
                celda.fill = fill_alt

    for col_idx, nombre in enumerate(cols, 1):
        max_len = len(str(nombre))
        for row_idx in range(3, ws.max_row + 1):
            valor = ws.cell(row_idx, col_idx).value
            if valor is not None:
                max_len = max(max_len, min(len(str(valor)), 48))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 3

    ws.auto_filter.ref = f"A3:{get_column_letter(n_cols)}{ws.max_row}"
    ws.freeze_panes = "A4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:3"
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def generar_pdf_escaner_ev(df_ops, liga):
    buffer = BytesIO()
    estilos = _estilos_parrafo_pdf()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=1.4 * cm,
        rightMargin=1.4 * cm,
        topMargin=1.4 * cm,
        bottomMargin=1.2 * cm,
        title="Informe Escáner EV+",
        author="Hugo's Hub",
    )
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    top_n = min(len(df_ops), 25)
    df_pdf = df_ops.head(top_n)
    ev_medio = df_ops["EV_Num"].mean() if "EV_Num" in df_ops.columns and not df_ops.empty else 0

    story = [
        Paragraph("HUGO'S HUB  ·  SISTEMA ANALÍTICO", estilos["marca"]),
        Paragraph("Informe de oportunidades EV+", estilos["titulo"]),
        Paragraph(
            f"Competición: <b>{liga}</b>  |  Generado: {fecha}  |  "
            f"Apuestas con valor: {len(df_ops)}  |  EV medio: +{ev_medio:.2f}%  |  "
            f"Mostrando las {top_n} mejores por EV.",
            estilos["subtitulo"],
        ),
        HRFlowable(width="100%", thickness=1.2, color=COLOR_TEAL, spaceAfter=10),
    ]

    cols = [c for c in [
        "Fecha / Hora", "Partido", "Mercado", "Casa", "Cuota",
        "Prob_Modelo", "Cuota_Justa", "EV", "Stake_Recomendado"
    ] if c in df_pdf.columns]
    encabezados = [c.replace("_", " ") for c in cols]
    filas = []
    for fila in df_pdf[cols].itertuples(index=False):
        filas.append(["" if v is None else v for v in fila])

    anchos = {
        "Fecha / Hora": 3.2 * cm,
        "Partido": 5.4 * cm,
        "Mercado": 3.6 * cm,
        "Casa": 2.8 * cm,
        "Cuota": 1.8 * cm,
        "Prob_Modelo": 2.1 * cm,
        "Cuota_Justa": 2.2 * cm,
        "EV": 1.8 * cm,
        "Stake_Recomendado": 2.4 * cm,
    }
    col_widths = [anchos.get(c, 2.4 * cm) for c in cols]
    story.append(_tabla_pdf(encabezados, filas, col_widths))
    story.append(Spacer(1, 12))
    story.append(Paragraph(
        "Las cuotas justas proceden del modelo Poisson sobre el histórico local. "
        "Este informe no constituye una recomendación de inversión.",
        estilos["pie"],
    ))
    doc.build(story, onFirstPage=_pie_pdf, onLaterPages=_pie_pdf)
    buffer.seek(0)
    return buffer.getvalue()


def generar_pdf_dossier_equipo(datos):
    buffer = BytesIO()
    estilos = _estilos_parrafo_pdf()
    _, _, mono_reg, mono_bold = _fuentes_pdf()

    equipo = datos.get("equipo", "Equipo")
    liga = datos.get("liga", "Competición Oficial")
    horizonte_raw = str(datos.get("horizonte", "Histórico Completo"))
    if "L10" in horizonte_raw or "10" in horizonte_raw:
        horizonte_tag = "L10 (Últimos 10 Partidos)"
    elif "Temporada" in horizonte_raw:
        horizonte_tag = "Temporada Actual"
    elif "Ponderado" in horizonte_raw or "Decaimiento" in horizonte_raw:
        horizonte_tag = "Ponderado (Time-Decay)"
    else:
        horizonte_tag = horizonte_raw

    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")

    # Extracción de bloques cuantitativos con fallbacks seguros de retrocompatibilidad
    n_tot = int(datos.get("partidos", 0) or 0)
    v_tot = int(datos.get("v", 0) or 0)
    e_tot = int(datos.get("e", 0) or 0)
    d_tot = int(datos.get("d", 0) or 0)
    n_loc = int(datos.get("n_local", 0) or 0)
    n_vis = int(datos.get("n_vis", 0) or 0)
    gf_loc_tot = int(datos.get("gf_local", 0) or 0)
    gc_loc_tot = int(datos.get("gc_local", 0) or 0)
    gf_vis_tot = int(datos.get("gf_vis", 0) or 0)
    gc_vis_tot = int(datos.get("gc_vis", 0) or 0)

    m_gf_l_fb = (gf_loc_tot / n_loc) if n_loc > 0 else 0.0
    m_gc_l_fb = (gc_loc_tot / n_loc) if n_loc > 0 else 0.0
    m_gf_v_fb = (gf_vis_tot / n_vis) if n_vis > 0 else 0.0
    m_gc_v_fb = (gc_vis_tot / n_vis) if n_vis > 0 else 0.0

    sg = datos.get("stats_global") or {}
    sh = datos.get("stats_home") or {}
    sa = datos.get("stats_away") or {}

    media_liga_h = float(datos.get("media_liga_home", 1.45) or 1.45)
    media_liga_a = float(datos.get("media_liga_away", 1.15) or 1.15)
    media_liga_g = float(datos.get("media_liga_global", (media_liga_h + media_liga_a) / 2.0) or 1.30)

    win_rate = float(sg.get("win_rate", (v_tot / n_tot * 100.0) if n_tot > 0 else 0.0))
    gf_med = float(sg.get("gf_med", datos.get("media_gf", 0.0)))
    gc_med = float(sg.get("gc_med", datos.get("media_gc", 0.0)))
    xg_f_g = float(sg.get("xg_f", gf_med))
    xg_c_g = float(sg.get("xg_c", gc_med))
    xg_neto_g = float(sg.get("xg_neto", xg_f_g - xg_c_g))
    diff_g_xg = float(sg.get("diff_g_xg", gf_med - xg_f_g))

    cs_rate_g = float(sg.get("clean_sheet", datos.get("p_cs", 0.0)))
    fts_rate_g = float(sg.get("failed_to_score", 0.0))
    btts_si_g = float(sg.get("btts_si", datos.get("p_btts", 0.0)))
    btts_no_g = float(sg.get("btts_no", max(0.0, 100.0 - btts_si_g)))

    # Métricas Casa vs Fuera (Block 2)
    h_gf = float(sh.get("gf_med", m_gf_l_fb))
    h_gc = float(sh.get("gc_med", m_gc_l_fb))
    a_gf = float(sa.get("gf_med", m_gf_v_fb))
    a_gc = float(sa.get("gc_med", m_gc_v_fb))

    h_xgf = float(sh.get("xg_f", h_gf))
    h_xgc = float(sh.get("xg_c", h_gc))
    a_xgf = float(sa.get("xg_f", a_gf))
    a_xgc = float(sa.get("xg_c", a_gc))

    h_sotf = float(sh.get("sot_f", 4.2))
    h_sotc = float(sh.get("sot_c", 3.6))
    a_sotf = float(sa.get("sot_f", 3.5))
    a_sotc = float(sa.get("sot_c", 4.4))

    h_conv = float(sh.get("conv_sot", (h_gf / h_sotf * 100.0) if h_sotf > 0 else 0.0))
    a_conv = float(sa.get("conv_sot", (a_gf / a_sotf * 100.0) if a_sotf > 0 else 0.0))

    h_atk = float(sh.get("atk_rating", (h_gf / media_liga_h) if media_liga_h > 0 else 1.0))
    h_def = float(sh.get("def_rating", (h_gc / media_liga_a) if media_liga_a > 0 else 1.0))
    a_atk = float(sa.get("atk_rating", (a_gf / media_liga_a) if media_liga_a > 0 else 1.0))
    a_def = float(sa.get("def_rating", (a_gc / media_liga_h) if media_liga_h > 0 else 1.0))

    h_1t = float(sh.get("gf_1t_pct", 45.0))
    h_2t = float(sh.get("gf_2t_pct", 55.0))
    a_1t = float(sa.get("gf_1t_pct", 44.0))
    a_2t = float(sa.get("gf_2t_pct", 56.0))

    def _cuota_str(pct_val):
        p = float(pct_val) / 100.0
        return f"{1.0 / p:.2f}" if p > 1e-6 else "-"

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=24,
        bottomMargin=24,
        title=f"Quantitative Intelligence Report - {equipo}",
        author="Hugo's Hub",
    )

    sec_style = ParagraphStyle(
        "dossier_sec",
        parent=estilos["seccion"],
        fontSize=10.5,
        leading=12,
        spaceBefore=4,
        spaceAfter=2,
    )

    # Banner Institucional Superior (Header Bar en Slate/Navy #1E293B)
    header_data = [[
        Paragraph(
            f"<font color='#10B981'><b>HUGO'S HUB  |  QUANTITATIVE INTELLIGENCE REPORT</b></font><br/>"
            f"<font color='#FFFFFF' size=14><b>{equipo.upper()}</b></font><br/>"
            f"<font color='#CBD5E1' size=7.5>Competición: <b>{liga}</b>   ·   Horizonte: <b>{horizonte_tag}</b>   ·   "
            f"Muestra: <b>{n_tot} partidos</b> ({n_loc} Casa / {n_vis} Fuera)   ·   Generado: {fecha}</font>",
            ParagraphStyle("hdr_box", fontName=estilos["titulo"].fontName, fontSize=8.5, leading=12, alignment=TA_LEFT),
        )
    ]]
    header_table = Table(header_data, colWidths=[18.2 * cm])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), COLOR_SLATE_HEADER),
        ("BOX", (0, 0), (-1, -1), 1.0, COLOR_EMERALD),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))

    # ── BLOCK 1: EXECUTIVE KPI SUMMARY BOX (4-COLUMN VISUAL TABLE) ───────────
    color_g_xg = "#059669" if diff_g_xg >= 0 else "#DC2626"
    kpi_row = [[
        Paragraph(
            f"<font color='#64748B' size=7><b>WIN RATE &amp; RECORD</b></font><br/>"
            f"<font name='{mono_bold}' size=10.5 color='#1E293B'>{win_rate:.1f}% WIN</font><br/>"
            f"<font name='{mono_reg}' size=7 color='#059669'>{v_tot}W - {e_tot}D - {d_tot}L ({n_tot}p)</font>",
            estilos["celda_c"],
        ),
        Paragraph(
            f"<font color='#64748B' size=7><b>GF/GA &amp; NET xG / 90M</b></font><br/>"
            f"<font name='{mono_bold}' size=10.5 color='#1E293B'>{gf_med:.2f} / {gc_med:.2f}</font><br/>"
            f"<font name='{mono_reg}' size=7 color='{color_g_xg}'>xG Neto {xg_neto_g:+.2f} (G-xG {diff_g_xg:+.2f})</font>",
            estilos["celda_c"],
        ),
        Paragraph(
            f"<font color='#64748B' size=7><b>CLEAN SHEET &amp; FTS %</b></font><br/>"
            f"<font name='{mono_bold}' size=10.5 color='#1E293B'>{cs_rate_g:.1f}% CS</font><br/>"
            f"<font name='{mono_reg}' size=7 color='#DC2626'>Sin Marcar (FTS): {fts_rate_g:.1f}%</font>",
            estilos["celda_c"],
        ),
        Paragraph(
            f"<font color='#64748B' size=7><b>BTTS HIT RATE (AMBOS MARCAN)</b></font><br/>"
            f"<font name='{mono_bold}' size=10.5 color='#1E293B'>{btts_si_g:.1f}% SÍ</font><br/>"
            f"<font name='{mono_reg}' size=7 color='#475569'>BTTS No: {btts_no_g:.1f}% (Cuota {_cuota_str(btts_si_g)})</font>",
            estilos["celda_c"],
        ),
    ]]
    kpi_table = Table(kpi_row, colWidths=[4.55 * cm, 4.55 * cm, 4.55 * cm, 4.55 * cm])
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), COLOR_ROW),
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_LINE),
        ("LINEABOVE", (0, 0), (-1, 0), 1.5, COLOR_SLATE_HEADER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))

    # ── BLOCK 2: ASYMMETRIC SPLIT TABLE (LOCAL VS VISITANTE) ─────────────────
    filas_split = [
        ["Partidos Analizados (Muestra)", f"{n_loc}", f"{n_vis}", f"{n_tot} tot."],
        ["Goles a Favor / en Contra (90m)", f"{h_gf:.2f} / {h_gc:.2f}", f"{a_gf:.2f} / {a_gc:.2f}", f"{media_liga_h:.2f} L / {media_liga_a:.2f} V"],
        ["Expected Goals (xG Favor / Contra)", f"{h_xgf:.2f} / {h_xgc:.2f}", f"{a_xgf:.2f} / {a_xgc:.2f}", f"{media_liga_g:.2f} xG"],
        ["Tiros a Puerta (SOT Favor / Contra)", f"{h_sotf:.2f} / {h_sotc:.2f}", f"{a_sotf:.2f} / {a_sotc:.2f}", "4.10 / 4.10"],
        ["Tasa de Conversión SOT (% Gol/SOT)", f"{h_conv:.1f}%", f"{a_conv:.1f}%", "31.5%"],
        ["Attack Rating vs Media de Liga", f"{h_atk:.2f}x", f"{a_atk:.2f}x", "1.00x (Base)"],
        ["Defense Rating vs Media de Liga", f"{h_def:.2f}x", f"{a_def:.2f}x", "1.00x (Base)"],
        ["Distribución Goles Anotados (1T vs 2T)", f"{h_1t:.0f}% 1T / {h_2t:.0f}% 2T", f"{a_1t:.0f}% 1T / {a_2t:.0f}% 2T", "44% 1T / 56% 2T"],
    ]
    tabla_split = _tabla_pdf(
        ["Métrica Cuantitativa", "Rendimiento Casa", "Rendimiento Fuera", "Media Liga"],
        filas_split,
        [6.4 * cm, 3.9 * cm, 3.9 * cm, 4.0 * cm],
        mono_cols=[1, 2, 3],
    )
    tabla_split.setStyle(TableStyle([
        ("TOPPADDING", (0, 0), (-1, -1), 2.0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.0),
    ]))

    # ── BLOCK 3: FULL MARKET FREQUENCY MATRIX & THEORETICAL FAIR ODDS ────────
    def _val_m(dic, key, fallback=0.0):
        return float(dic.get(key, fallback))

    over15_g = _val_m(sg, "over15", min(98.0, float(datos.get("p_over25", 50.0)) + 22.0))
    over25_g = _val_m(sg, "over25", float(datos.get("p_over25", 50.0)))
    over35_g = _val_m(sg, "over35", max(5.0, over25_g - 22.0))
    under25_g = _val_m(sg, "under25", max(0.0, 100.0 - over25_g))

    mercados_spec = [
        ("Over 1.5 Goles (+1.5)", over15_g, _val_m(sh, "over15", over15_g), _val_m(sa, "over15", over15_g)),
        ("Over 2.5 Goles (+2.5)", over25_g, _val_m(sh, "over25", over25_g), _val_m(sa, "over25", over25_g)),
        ("Over 3.5 Goles (+3.5)", over35_g, _val_m(sh, "over35", over35_g), _val_m(sa, "over35", over35_g)),
        ("Under 2.5 Goles (-2.5)", under25_g, _val_m(sh, "under25", under25_g), _val_m(sa, "under25", under25_g)),
        ("Ambos Marcan — BTTS (Sí)", btts_si_g, _val_m(sh, "btts_si", btts_si_g), _val_m(sa, "btts_si", btts_si_g)),
        ("Ambos Marcan — BTTS (No)", btts_no_g, _val_m(sh, "btts_no", btts_no_g), _val_m(sa, "btts_no", btts_no_g)),
        ("Portería a Cero (Clean Sheet)", cs_rate_g, _val_m(sh, "clean_sheet", cs_rate_g), _val_m(sa, "clean_sheet", cs_rate_g)),
        ("Sin Anotar (Failed to Score)", fts_rate_g, _val_m(sh, "failed_to_score", fts_rate_g), _val_m(sa, "failed_to_score", fts_rate_g)),
    ]
    filas_mercados_pdf = []
    for nom_m, p_g, p_h, p_a in mercados_spec:
        filas_mercados_pdf.append([
            nom_m,
            f"{p_g:.1f}%",
            f"{p_h:.1f}%",
            f"{p_a:.1f}%",
            _cuota_str(p_g),
        ])

    tabla_mercados = _tabla_pdf(
        ["Mercado", "Hit Rate General (%)", "Hit Rate Casa (%)", "Hit Rate Fuera (%)", "Cuota Implícita Justa (1/p)"],
        filas_mercados_pdf,
        [5.8 * cm, 3.0 * cm, 3.0 * cm, 3.0 * cm, 3.4 * cm],
        mono_cols=[1, 2, 3, 4],
    )
    tabla_mercados.setStyle(TableStyle([
        ("TOPPADDING", (0, 0), (-1, -1), 2.0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.0),
    ]))

    # ── BLOCK 4: DISCIPLINARY & FRICTION PROFILE (ARBITRATION MARKETS) ───────
    g_am_p = _val_m(sg, "amar_prop", 2.15)
    g_rj_p = _val_m(sg, "rojas_prop", 0.10)
    h_am_p = _val_m(sh, "amar_prop", g_am_p)
    h_rj_p = _val_m(sh, "rojas_prop", g_rj_p)
    a_am_p = _val_m(sa, "amar_prop", g_am_p)
    a_rj_p = _val_m(sa, "rojas_prop", g_rj_p)

    g_am_r = _val_m(sg, "amar_riv", 2.20)
    g_rj_r = _val_m(sg, "rojas_riv", 0.10)
    h_am_r = _val_m(sh, "amar_riv", g_am_r)
    a_am_r = _val_m(sa, "amar_riv", g_am_r)

    g_f_com = _val_m(sg, "faltas_com", 11.8)
    g_f_rec = _val_m(sg, "faltas_rec", 11.5)
    h_f_com = _val_m(sh, "faltas_com", g_f_com)
    h_f_rec = _val_m(sh, "faltas_rec", g_f_rec)
    a_f_com = _val_m(sa, "faltas_com", g_f_com)
    a_f_rec = _val_m(sa, "faltas_rec", g_f_rec)

    g_fric = _val_m(sg, "friccion_idx", g_f_com + g_f_rec + g_am_p + g_am_r + g_rj_p + g_rj_r)
    h_fric = _val_m(sh, "friccion_idx", h_f_com + h_f_rec + h_am_p + h_am_r)
    a_fric = _val_m(sa, "friccion_idx", a_f_com + a_f_rec + a_am_p + a_am_r)

    filas_disc = [
        [
            "Tarjetas Propias Recibidas (Amarillas / Rojas)",
            f"{g_am_p:.2f} AM / {g_rj_p:.2f} RJ",
            f"{h_am_p:.2f} AM / {h_rj_p:.2f} RJ",
            f"{a_am_p:.2f} AM / {a_rj_p:.2f} RJ",
        ],
        [
            "Tarjetas Provocadas al Rival (Amarillas / Rojas)",
            f"{g_am_r:.2f} AM / {g_rj_r:.2f} RJ",
            f"{h_am_r:.2f} AM",
            f"{a_am_r:.2f} AM",
        ],
        [
            "Faltas por Partido (Cometidas vs Recibidas)",
            f"{g_f_com:.1f} Com. / {g_f_rec:.1f} Rec.",
            f"{h_f_com:.1f} C / {h_f_rec:.1f} R",
            f"{a_f_com:.1f} C / {a_f_rec:.1f} R",
        ],
        [
            "Disciplinary Friction Index (Faltas + Tarjetas)",
            f"{g_fric:.1f} pts",
            f"{h_fric:.1f} pts (Casa)",
            f"{a_fric:.1f} pts (Fuera)",
        ],
    ]

    st_disc_celda = ParagraphStyle("disc_celda", fontName=estilos["celda"].fontName, fontSize=7.5, leading=9.0, textColor=COLOR_SLATE_HEADER, alignment=TA_LEFT)
    st_disc_mono = ParagraphStyle("disc_mono", fontName=mono_reg, fontSize=7.5, leading=9.0, textColor=COLOR_SLATE_HEADER, alignment=TA_CENTER)
    st_disc_h = ParagraphStyle("disc_h", fontName=estilos["celda_h"].fontName, fontSize=7.5, leading=9.0, textColor=colors.white, alignment=TA_CENTER)

    encabezados_disc = ["Métrica Disciplinaria y de Arbitraje", "Promedio General", "Condición Casa", "Condición Fuera"]
    data_disc = [[Paragraph(str(h), st_disc_h) for h in encabezados_disc]]
    for fila in filas_disc:
        row_cells = [Paragraph(str(fila[0]), st_disc_celda)]
        for idx in [1, 2, 3]:
            row_cells.append(Paragraph(str(fila[idx]), st_disc_mono))
        data_disc.append(row_cells)

    tabla_disc = Table(data_disc, colWidths=[170, 115, 115, 115], repeatRows=1)
    tabla_disc.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_SLATE_HEADER),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.35, COLOR_LINE),
        ("LINEBELOW", (0, 0), (-1, 0), 1.0, COLOR_EMERALD),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("LEADING", (0, 0), (-1, -1), 9.0),
        ("LEFTPADDING", (0, 0), (-1, -1), 3.0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3.0),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
    ]))
    for i in range(1, len(data_disc)):
        bg_c = COLOR_ROW if i % 2 == 0 else colors.white
        tabla_disc.setStyle(TableStyle([("BACKGROUND", (0, i), (-1, i), bg_c)]))

    st_pie = ParagraphStyle(
        "dossier_pie",
        parent=estilos["pie"],
        fontSize=6.5,
        leading=8.5,
    )

    story = [
        header_table,
        Spacer(1, 4),
        Paragraph("1. Executive KPI Summary (Indicadores Clave de Rendimiento)", sec_style),
        kpi_table,
        Paragraph("2. Asymmetric Split Analysis: Rendimiento Local vs Visitante", sec_style),
        tabla_split,
        Paragraph("3. Full Market Frequency Matrix &amp; Theoretical Fair Odds (1/p)", sec_style),
        tabla_mercados,
        Paragraph("4. Disciplinary &amp; Friction Profile (Correlación con Mercados de Arbitraje)", sec_style),
        tabla_disc,
        Spacer(1, 4),
        HRFlowable(width="100%", thickness=0.6, color=COLOR_LINE, spaceAfter=2),
        Paragraph(
            "DISCLAIMER CUANTITATIVO: Las probabilidades empíricas, ratings relativos y cuotas justas implícitas (1/p) se derivan "
            "matemáticamente de la muestra histórica seleccionada en macro_base_futbol.db sin margen comercial (overround 0%). "
            "Generado íntegramente en memoria (BytesIO). Uso analítico exclusivo de Hugo's Hub.",
            st_pie,
        ),
    ]
    doc.build(story, onFirstPage=_pie_pdf, onLaterPages=_pie_pdf)
    buffer.seek(0)
    return buffer.getvalue()

df_partidos_global, err_partidos = cargar_datos_partidos()
df_jugadores, err_jugadores = cargar_datos_jugadores()
df_partidos = df_partidos_global.copy()

# ==============================================================================
# SIDEBAR
# ==============================================================================
with st.sidebar:
    st.header("🎛️ Centro de Control")
    menu_seleccionado = st.radio(
        "Navegación:", 
        [
            "🏠 Inicio / Portada", 
            "📡 Escáner Vivo EV+ (API)", 
            "📊 Dossier de Equipo",
            "⚔️ Comparador H2H (Cara a Cara)",
            "🎯 Radar de Tendencias", 
            "📋 Scouting de Jugadores", 
            "🤖 Motor de Apuestas (EV+)", 
            "⚖️ Directorio y Matchup Árbitros", 
            "🧮 Hándicap Asiático Avanzado", 
            "🎯 Escáner Autónomo de Ineficiencias (+EV Scanner)",
            "🔗 Constructor Combinadas (EV+)", 
            "💸 Calculadora & Kelly",
            "🤖 Asistente IA Local"
        ],
        key="menu_nav"
    )
    st.markdown("---")
    st.header("🌍 Filtro de Competición")
    if not df_partidos.empty and 'Div' in df_partidos.columns:
        ligas_raw = sorted(df_partidos['Div'].dropna().unique().tolist())
        ligas_nombres = [MAPA_LIGAS.get(l, l) for l in ligas_raw]
        liga_sel_nombre = st.selectbox("Selecciona Liga:", ["Todas"] + ligas_nombres)
        if liga_sel_nombre != "Todas":
            df_partidos = df_partidos[df_partidos['Div'] == MAPA_INV.get(liga_sel_nombre, liga_sel_nombre)]
            st.markdown("---")
            st.subheader("🔍 Racha del Equipo")
            equipos_liga = sorted(pd.concat([df_partidos['HomeTeam'], df_partidos['AwayTeam']]).dropna().unique())
            equipo_side_sel = st.selectbox("Selecciona un Equipo:", ["Ninguno"] + equipos_liga)
            if equipo_side_sel != "Ninguno":
                df_eq = df_partidos[(df_partidos['HomeTeam'] == equipo_side_sel) | (df_partidos['AwayTeam'] == equipo_side_sel)].copy()
                if 'Date' in df_eq.columns:
                    df_eq['Date_Parsed'] = pd.to_datetime(df_eq['Date'], errors='coerce', dayfirst=True)
                    df_eq = df_eq.sort_values(by='Date_Parsed', ascending=False)
                for idx, row in df_eq.head(5).iterrows():
                    local, vis = row['HomeTeam'], row['AwayTeam']
                    local_fmt = f"**{local}**" if local == equipo_side_sel else local
                    vis_fmt = f"**{vis}**" if vis == equipo_side_sel else vis
                    st.caption(f"📅 {row.get('Date', 'N/A')} | {local_fmt} **{row['FTHG']} - {row['FTAG']}** {vis_fmt}")

    st.markdown("---")

    # ── SECCIÓN ADMIN: Sincronización de Jornada ─────────────────────────
    st.header("🔧 Admin")
    if _UPDATER_DISPONIBLE:
        if "sync_success_msg" in st.session_state:
            st.success(st.session_state.pop("sync_success_msg"))
        if "sync_warning_msg" in st.session_state:
            st.warning(st.session_state.pop("sync_warning_msg"))

        if st.button("🔄 Sincronizar Última Jornada", use_container_width=True, type="primary"):
            with st.spinner("Actualizando métricas de la jornada..."):
                try:
                    resultado_sync = sincronizar_jornada()
                    if resultado_sync.get("exito"):
                        partidos_n = resultado_sync.get("partidos_nuevos", 0)
                        jugadores_a = resultado_sync.get("jugadores_act", 0)
                        msg_ok = (
                            f"✅ **Jornada sincronizada correctamente.**\n\n"
                            f"- 📅 **Partidos nuevos:** {partidos_n}\n"
                            f"- 👤 **Jugadores actualizados:** {jugadores_a}"
                        )
                        # Limpiar toda la caché de Streamlit para que los nuevos datos
                        # se reflejen inmediatamente en todos los módulos analíticos
                        st.cache_data.clear()
                        st.cache_resource.clear()
                        st.session_state["sync_success_msg"] = msg_ok
                        st.rerun()
                    else:
                        errores = resultado_sync.get("errores", [])
                        mensaje = resultado_sync.get("mensaje", "Error desconocido.")
                        # Éxito parcial: partidos actualizados pero jugadores fallaron
                        if resultado_sync.get("partidos_nuevos", 0) > 0 or resultado_sync.get("jugadores_act", 0) > 0:
                            msg_warn = (
                                f"⚠️ **Sincronización parcial.**\n\n"
                                f"- 📅 Partidos nuevos: {resultado_sync.get('partidos_nuevos', 0)}\n"
                                f"- 👤 Jugadores actualizados: {resultado_sync.get('jugadores_act', 0)}\n\n"
                                f"**Detalle:** {'; '.join(errores) if errores else mensaje}"
                            )
                            st.cache_data.clear()
                            st.cache_resource.clear()
                            st.session_state["sync_warning_msg"] = msg_warn
                            st.rerun()
                        else:
                            st.error(
                                f"❌ **Error en la sincronización.**\n\n"
                                f"{'; '.join(errores) if errores else mensaje}"
                            )
                except Exception as exc_btn:
                    st.error(
                        f"❌ **Error inesperado durante la sincronización:**\n\n`{exc_btn}`"
                    )

        st.caption("Descarga datos de football-data.co.uk y FBref.")
    else:
        st.warning("Módulo `updater.py` no encontrado.")

    st.markdown("---")
    st.caption("🚀 Hugo's Hub v19.0 - Hándicap ATS y Matriz")

# ==============================================================================
# 🎯 ESCÁNER AUTÓNOMO DE INEFICIENCIAS (+EV SCANNER) — MÓDULO DESACOPLADO
# Pure Python / SQLite (vía df_partidos_global). Sin dependencias cloud de pago.
# ==============================================================================
import html as html_lib
import json as json_lib
import unicodedata as unicodedata_lib

EV_UMBRALES = {
    "arb_pct_under": 30.0, "fric_pct_under": 35.0,
    "arb_pct_over": 75.0, "fric_pct_over": 70.0,
    "rest_delta_dias": 3, "rest_cap_dias": 14,
    "fatiga_2t_xga": 0.60, "fatiga_faltas": 1.20,
    "reb_xg90": 1.45, "reb_gf90": 0.60,
    "reg_ga90": 0.50, "reg_xga90": 1.60,
    "d_ratio2t": 0.65, "d_goles_1t": 0.75,
    "ventana_c": 6, "min_muestra": 5,
    "score_oportunidad": 2, "score_alto_valor": 3,
}
EV_DOMINIOS = {
    "A": "⚖️ Árbitro y Tarjetas",
    "B": "🔋 Descanso y Desgaste Físico",
    "C": "🎯 Eficiencia Ofensiva / Regresión xG",
    "D": "⏱️ Inercia de Segundas Partes",
}
EV_LLM_URL = "http://localhost:1234/v1/chat/completions"
EV_COLS_BASE = [
    "Div", "Date", "HomeTeam", "AwayTeam", "FTHG", "FTAG", "HTHG", "HTAG", "HS", "AS",
    "HST", "AST", "HF", "AF", "HY", "AY", "HR", "AR", "HxG", "AxG", "Referee",
]
EV_COLS_NUM = [c for c in EV_COLS_BASE if c not in ("Div", "Date", "HomeTeam", "AwayTeam", "Referee")]

# Resolución bidireccional de entidades: alias de UI / fuentes externas  <->  nombre canónico en SQLite
CANONICAL_TEAM_MAP = {
    'espanyol': 'Espanol', 'rcd espanyol': 'Espanol',
    'real sociedad': 'Sociedad', 'la real': 'Sociedad',
    'athletic club': 'Ath Bilbao', 'athletic bilbao': 'Ath Bilbao',
    'atlético de madrid': 'Ath Madrid', 'atletico madrid': 'Ath Madrid',
    'celta de vigo': 'Celta', 'betis': 'Betis', 'real betis': 'Betis',
    'deportivo la coruña': 'La Coruna', 'deportivo': 'La Coruna',
    'sporting de gijón': 'Gijon', 'sporting gijon': 'Gijon',
    'rayo vallecano': 'Vallecano', 'alavés': 'Alaves', 'deportivo alaves': 'Alaves',
}
EV_NOMBRE_DISPLAY = {
    'Espanol': 'Espanyol', 'Sociedad': 'Real Sociedad', 'Ath Bilbao': 'Athletic Club',
    'Ath Madrid': 'Atlético de Madrid', 'Celta': 'Celta de Vigo', 'Betis': 'Real Betis',
    'La Coruna': 'Deportivo La Coruña', 'Gijon': 'Sporting de Gijón',
    'Vallecano': 'Rayo Vallecano', 'Alaves': 'Deportivo Alavés',
}


def _ev_norm(texto):
    t = unicodedata_lib.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode("ascii")
    return " ".join(t.strip().lower().split())


_EV_MAP_NORM = {_ev_norm(k): v for k, v in CANONICAL_TEAM_MAP.items()}


def ev_nombre_display(nombre_bd):
    """Nombre canónico de la BD -> nombre legible para la interfaz."""
    return EV_NOMBRE_DISPLAY.get(str(nombre_bd), str(nombre_bd))


def ev_resolver_equipo(nombre, equipos_bd):
    """Alias / texto libre -> nombre canónico existente en la BD (None si no hay coincidencia)."""
    clave = _ev_norm(nombre)
    if not clave:
        return None
    idx = {_ev_norm(e): e for e in equipos_bd}
    if clave in idx:
        return idx[clave]
    canon = _EV_MAP_NORM.get(clave)
    if canon is not None and _ev_norm(canon) in idx:
        return idx[_ev_norm(canon)]
    cands = [(k, e) for k, e in idx.items() if clave in k or (len(k) >= 4 and k in clave)]
    if not cands:
        return None
    cands.sort(key=lambda x: abs(len(x[0]) - len(clave)))
    return cands[0][1]


def _ev_cuota(p):
    """Cuota justa 1/p con guardia ante p ~ 0."""
    try:
        if p is None or p <= 1e-9:
            return None
        return 1.0 / float(p)
    except Exception:
        return None


def _ev_fmt_cuota(p):
    c = _ev_cuota(p)
    return "-" if c is None else f"{c:.2f}"


def _ev_txt_cuota(p):
    return f"Cuota Justa Estimada: proyección matemática del modelo @{_ev_fmt_cuota(p)} (probabilidad {0.0 if p is None else p * 100:.1f}%)."


def _ev_pois_cdf(k, lam):
    lam = max(float(lam), 1e-9)
    return float(sum(math.exp(-lam) * lam ** i / math.factorial(i) for i in range(0, int(k) + 1)))


def _ev_matriz(lh, la, mx=9):
    ks = np.arange(mx + 1)
    fact = np.array([math.factorial(int(i)) for i in ks], dtype=float)
    ph = np.exp(-lh) * lh ** ks / fact
    pa = np.exp(-la) * la ** ks / fact
    m = np.outer(ph, pa)
    tot = m.sum()
    return m / tot if tot > 0 else m


def _ev_percentil(serie, valor):
    s = pd.Series(serie).dropna()
    if s.empty or valor is None or pd.isna(valor):
        return None
    return float((s <= valor).mean() * 100.0)


@st.cache_data(ttl=3600, show_spinner=False)
def ev_construir_historico_largo(df_min):
    """Formato largo (1 fila por equipo-partido) con métricas derivadas, ordenado por Equipo/Fecha."""
    d = df_min.copy()
    d["Fecha"] = pd.to_datetime(d["Date"], errors="coerce", dayfirst=True)
    for c in EV_COLS_NUM:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    d = d.dropna(subset=["Fecha", "HomeTeam", "AwayTeam", "FTHG", "FTAG"])
    mapas = {
        True: dict(Equipo="HomeTeam", Rival="AwayTeam", gf="FTHG", ga="FTAG", gf1="HTHG", ga1="HTAG",
                   sot_f="HST", sot_c="AST", sh_f="HS", sh_c="AS", fl_f="HF", fl_c="AF",
                   y="HY", r="HR", xg="HxG", xga="AxG"),
        False: dict(Equipo="AwayTeam", Rival="HomeTeam", gf="FTAG", ga="FTHG", gf1="HTAG", ga1="HTHG",
                    sot_f="AST", sot_c="HST", sh_f="AS", sh_c="HS", fl_f="AF", fl_c="HF",
                    y="AY", r="AR", xg="AxG", xga="HxG"),
    }
    partes = []
    for local, m in mapas.items():
        datos = {"Div": d["Div"].values, "Fecha": d["Fecha"].values, "Local": local}
        for k, v in m.items():
            datos[k] = d[v].values
        partes.append(pd.DataFrame(datos))
    if not partes:
        return pd.DataFrame()
    L = pd.concat(partes, ignore_index=True)
    L["gf2"] = L["gf"] - L["gf1"]
    L["ga2"] = L["ga"] - L["ga1"]
    L["tarjetas"] = L["y"] + L["r"]
    off_f = (L["sh_f"] - L["sot_f"]).clip(lower=0).fillna(0.0)
    off_c = (L["sh_c"] - L["sot_c"]).clip(lower=0).fillna(0.0)
    # xG: valor real (HxG/AxG) si existe; si no, proxy por volumen de tiros (SOT*0.30 + fuera*0.03)
    L["xg_est"] = L["xg"].where(L["xg"].notna(), 0.30 * L["sot_f"] + 0.03 * off_f)
    L["xga_est"] = L["xga"].where(L["xga"].notna(), 0.30 * L["sot_c"] + 0.03 * off_c)
    L["friccion"] = L["fl_f"] + L["fl_c"] + L["tarjetas"] * 5.0
    return L.sort_values(["Equipo", "Fecha"]).reset_index(drop=True)


def _ev_ventana(L, equipo, cut, n, local=None):
    s = L[(L["Equipo"] == equipo) & (L["Fecha"] < cut)]
    if local is not None:
        s = s[s["Local"] == local]
    return s.tail(n)


def _ev_friccion_liga(L, div, local, cut, n, cache=None):
    clave = ("fric", div, local, cut, n)
    if cache is not None and clave in cache:
        return cache[clave]
    sub = L[(L["Div"] == div) & (L["Local"] == local) & (L["Fecha"] < cut)].dropna(subset=["friccion"])
    if sub.empty:
        res = pd.Series(dtype=float)
    else:
        g = sub.groupby("Equipo").tail(n)
        agg = g.groupby("Equipo")["friccion"].agg(["mean", "count"])
        res = agg[agg["count"] >= 4]["mean"]
    if cache is not None:
        cache[clave] = res
    return res


def _ev_info_arbitro(df_arb, nombre, div):
    """Percentil disciplinario del árbitro dentro de su liga (Amarillas*1 + Rojas*2)."""
    if df_arb is None or df_arb.empty or df_arb.shape[1] < 9 or not nombre:
        return None
    a = pd.DataFrame({
        "nom": df_arb.iloc[:, 0].astype(str).str.strip(),
        "div": df_arb.iloc[:, 1].astype(str),
        "am": pd.to_numeric(df_arb.iloc[:, 6], errors="coerce"),
        "rj": pd.to_numeric(df_arb.iloc[:, 7], errors="coerce"),
        "fl": pd.to_numeric(df_arb.iloc[:, 8], errors="coerce"),
    }).dropna(subset=["am"])
    a["rj"] = a["rj"].fillna(0.0)
    a["score"] = a["am"] * 1.0 + a["rj"] * 2.0
    fila = a[(a["nom"].str.lower() == str(nombre).strip().lower())]
    if (fila["div"] == div).any():
        fila = fila[fila["div"] == div]
    if fila.empty:
        return None
    fila = fila.iloc[0]
    pool = a[a["div"] == fila["div"]]
    if len(pool) < 5:
        pool = a
    return {
        "nombre": str(fila["nom"]), "amarillas": float(fila["am"]), "rojas": float(fila["rj"]),
        "faltas": float(fila["fl"]) if pd.notna(fila["fl"]) else None,
        "score": float(fila["score"]), "pct": _ev_percentil(pool["score"], fila["score"]),
    }


def _ev_nuevo_dominio(clave, categoria=""):
    return {"id": clave, "nombre": EV_DOMINIOS[clave], "categoria": categoria, "activo": False, "evaluable": False,
            "etiqueta": "", "mercado": "", "prob": None, "cuota": None, "bullets": [], "badge": ""}


def _ev_dominio_arbitro(L, df_arb, div, local, visitante, cut, n, arbitro, vh, va, cache):
    mot = _ev_nuevo_dominio("A", "Mercado Disciplinario")
    ref = _ev_info_arbitro(df_arb, arbitro, div)
    u = EV_UMBRALES
    if not arbitro:
        mot["badge"] = "⏳ Pendiente Designación"
    elif ref is None or ref["pct"] is None:
        mot["badge"] = "⏳ Sin Histórico del Colegiado"
    elif ref["pct"] <= u["arb_pct_under"]:
        mot["badge"] = "🟢 Permisivo"
    elif ref["pct"] >= u["arb_pct_over"]:
        mot["badge"] = "🔴 Riguroso"
    else:
        mot["badge"] = "⚪ Neutral"
    fh = vh["friccion"].mean()
    fa = va["friccion"].mean()
    pct_h = _ev_percentil(_ev_friccion_liga(L, div, True, cut, n, cache), fh)
    pct_a = _ev_percentil(_ev_friccion_liga(L, div, False, cut, n, cache), fa)
    if ref is None or ref["pct"] is None or pct_h is None or pct_a is None:
        mot["bullets"].append("Árbitro no designado o sin histórico disciplinario suficiente.")
        return mot
    mot["evaluable"] = True
    comb = (pct_h + pct_a) / 2.0
    ref_cards = ref["amarillas"] + ref["rojas"]
    eq_cards = (vh["tarjetas"].mean() if vh["tarjetas"].notna().any() else ref_cards / 2.0) + \
               (va["tarjetas"].mean() if va["tarjetas"].notna().any() else ref_cards / 2.0)
    lam_c = max(0.5 * ref_cards + 0.5 * eq_cards, 0.5)
    fl_h = vh["fl_f"].mean()
    fl_a = va["fl_f"].mean()
    txt_fl = f"{0.0 if pd.isna(fl_h) else fl_h:.1f} y {0.0 if pd.isna(fl_a) else fl_a:.1f} faltas promedio"
    if ref["pct"] <= u["arb_pct_under"] and pct_h <= u["fric_pct_under"] and pct_a <= u["fric_pct_under"]:
        mot["activo"] = True
        mot["etiqueta"] = "Partido de bajo voltaje disciplinario"
        mot["mercado"] = "Under 4.5 Tarjetas / Sin Tarjeta antes del Minuto 25:00"
        mot["prob"] = _ev_pois_cdf(4, lam_c)
        mot["bullets"] = [
            f"Árbitro: {ref['nombre']} promedia solo {ref_cards:.1f} tarjetas/partido (percentil inferior {ref['pct']:.0f}% de la liga).",
            f"Fricción: ambos equipos se ubican en el tercio más limpio del campeonato ({txt_fl}).",
        ]
    elif ref["pct"] >= u["arb_pct_over"] and comb >= u["fric_pct_over"]:
        mot["activo"] = True
        mot["etiqueta"] = "Choque de alta fricción disciplinaria"
        mot["mercado"] = "Over 4.5 / 5.5 Tarjetas Totales"
        mot["prob"] = 1.0 - _ev_pois_cdf(4, lam_c)
        mot["bullets"] = [
            f"Árbitro: {ref['nombre']} promedia {ref_cards:.1f} tarjetas/partido (percentil {ref['pct']:.0f}% de la liga: de los más severos).",
            f"Fricción: ambos equipos están en el tramo más caliente del campeonato ({txt_fl}).",
        ]
    mot["cuota"] = _ev_cuota(mot["prob"])
    if mot["activo"]:
        mot["bullets"].append(_ev_txt_cuota(mot["prob"]))
    return mot


def _ev_dominio_descanso(L, div, local, visitante, cut, n, lam_h, lam_a):
    mot = _ev_nuevo_dominio("B", "Mercado de Hándicap")
    ult = {}
    for eq in (local, visitante):
        s = L[(L["Equipo"] == eq) & (L["Fecha"] < cut)]
        ult[eq] = None if s.empty else min(int((cut - s["Fecha"].max()).days), EV_UMBRALES["rest_cap_dias"])
    if ult[local] is None or ult[visitante] is None:
        mot["badge"] = "⚪ Sin dato previo"
        mot["bullets"].append("Sin partido previo registrado.")
        return mot
    mot["evaluable"] = True
    delta = ult[local] - ult[visitante]
    if abs(delta) < EV_UMBRALES["rest_delta_dias"]:
        mot["badge"] = "⚪ Equilibrado"
        return mot
    mot["badge"] = f"🟢 Favorable +{delta} días" if delta > 0 else f"🔴 Desfavorable {delta} días"
    cansado, descansado = (visitante, local) if delta > 0 else (local, visitante)
    d_c, d_d = ult[cansado], ult[descansado]
    sc = _ev_ventana(L, cansado, cut, n)
    ga_tot = sc["ga"].sum()
    ga2 = sc.dropna(subset=["ga2"])["ga2"].sum()
    share_2t = float(ga2 / ga_tot) if ga_tot > 0 else None
    base_f = L[(L["Div"] == div) & (L["Fecha"] < cut)]["fl_f"].mean()
    faltas_eq = sc["fl_f"].mean()
    ratio_f = float(faltas_eq / base_f) if base_f and base_f > 0 and pd.notna(faltas_eq) else None
    flags = []
    if share_2t is not None and share_2t > EV_UMBRALES["fatiga_2t_xga"]:
        flags.append(f"encaja el {share_2t * 100:.0f}% de sus goles en la segunda mitad")
    if ratio_f is not None and ratio_f >= EV_UMBRALES["fatiga_faltas"]:
        flags.append(f"comete {((ratio_f - 1) * 100):.0f}% más faltas que la media de la liga")
    mot["activo"] = True
    mot["etiqueta"] = "Ventaja física del rival descansado"
    mat = _ev_matriz(lam_h, lam_a)
    idx = np.arange(mat.shape[0])
    if (cansado == local and lam_h >= lam_a) or (cansado == visitante and lam_a > lam_h):
        # El equipo cansado es el favorito -> +AH para el descansado (underdog)
        if descansado == local:
            p = float(sum(mat[i, j] for i in idx for j in idx if i - j + 0.5 > 0))
        else:
            p = float(sum(mat[i, j] for i in idx for j in idx if j - i + 0.5 > 0))
        mot["mercado"] = f"Hándicap Asiático +0.5 {ev_nombre_display(descansado)} (underdog descansado)"
        mot["categoria"] = "Mercado de Hándicap"
    else:
        lam_r = lam_h if descansado == local else lam_a
        p = 1.0 - math.exp(-max(lam_r * 0.55, 1e-9))
        mot["mercado"] = f"Over 0.5 Goles {ev_nombre_display(descansado)} en la 2ª Mitad"
        mot["categoria"] = "Mercado de Goles"
    mot["prob"] = p
    mot["cuota"] = _ev_cuota(p)
    mot["bullets"] = [
        f"Fatiga: {ev_nombre_display(cansado)} llega con solo {d_c * 24}h de recuperación ({d_c} días) vs {d_d} días completos de {ev_nombre_display(descansado)}.",
        (f"Desgaste: {ev_nombre_display(cansado)} " + " y ".join(flags) + ".") if flags
        else f"Desgaste: {ev_nombre_display(cansado)} aún no muestra deterioro táctico medible; la ventaja es estructural por calendario.",
        _ev_txt_cuota(p),
    ]
    return mot


def _ev_dominio_xg(L, local, visitante, cut, lam_h, lam_a):
    mot = _ev_nuevo_dominio("C", "Mercado de Goles")
    n6 = EV_UMBRALES["ventana_c"]
    hallazgos = []
    for eq, lam_eq, lam_riv, rival in ((local, lam_h, lam_a, visitante), (visitante, lam_a, lam_h, local)):
        w = _ev_ventana(L, eq, cut, n6).dropna(subset=["xg_est", "xga_est"])
        if len(w) < 4:
            continue
        mot["evaluable"] = True
        xg90, gf90 = float(w["xg_est"].mean()), float(w["gf"].mean())
        xga90, ga90 = float(w["xga_est"].mean()), float(w["ga"].mean())
        e, r = ev_nombre_display(eq), ev_nombre_display(rival)
        if xg90 >= EV_UMBRALES["reb_xg90"] and gf90 <= EV_UMBRALES["reb_gf90"]:
            p = 1.0 - math.exp(-max(lam_eq, 1e-9))
            hallazgos.append(("Rebote ofensivo probable", f"Ambos Equipos Marcan: Sí / Over 0.5 Goles {e}", p, "ofensivo",
                              f"Eficiencia: {e} genera {xg90:.2f} goles esperados por partido pero solo anota {gf90:.2f} "
                              f"(últimos {len(w)} partidos): el mercado lo está infravalorando."))
        if ga90 <= EV_UMBRALES["reg_ga90"] and xga90 >= EV_UMBRALES["reg_xga90"]:
            p = 1.0 - math.exp(-max(lam_riv, 1e-9))
            hallazgos.append(("Regresión defensiva probable", f"Portería a Cero: No / Over 0.5 Goles {r}", p, "defensivo",
                              f"Regresión: {e} encaja solo {ga90:.2f} goles por partido pese a conceder ocasiones por valor de "
                              f"{xga90:.2f} esperados: rendimiento defensivo difícil de sostener."))
    if not hallazgos:
        mot["badge"] = "⚪ Rendimiento Esperado"
        return mot
    mot["activo"] = True
    mot["badge"] = "🔥 Rebote Ofensivo Probable" if hallazgos[0][3] == "ofensivo" else "⚠️ Regresión Defensiva Probable"
    mot["etiqueta"] = " + ".join(h[0] for h in hallazgos)
    mot["mercado"] = " | ".join(h[1] for h in hallazgos)
    mot["prob"] = hallazgos[0][2]
    mot["cuota"] = _ev_cuota(mot["prob"])
    for h in hallazgos:
        mot["bullets"].append(h[4])
        mot["bullets"].append(_ev_txt_cuota(h[2]))
    return mot


def _ev_dominio_mitades(local, visitante, vh, va, lam_h, lam_a):
    mot = _ev_nuevo_dominio("D", "Mercado Temporal")
    mot["badge"] = "⚪ Distribución Neutra"
    ratios = []
    for w in (vh, va):
        x = w.dropna(subset=["gf1", "gf2"])
        tot = x["gf"].sum()
        ratios.append(float(x["gf2"].sum() / tot) if len(x) >= EV_UMBRALES["min_muestra"] and tot > 0 else None)
    xh, xa = vh.dropna(subset=["gf1", "ga1"]), va.dropna(subset=["gf1", "ga1"])
    if ratios[0] is None or ratios[1] is None or xh.empty or xa.empty:
        mot["bullets"].append("Datos de descanso (HTHG/HTAG) insuficientes.")
        return mot
    mot["evaluable"] = True
    g1 = float(((xh["gf1"] + xh["ga1"]).mean() + (xa["gf1"] + xa["ga1"]).mean()) / 2.0)
    u = EV_UMBRALES
    if ratios[0] >= u["d_ratio2t"] and ratios[1] >= u["d_ratio2t"] and g1 <= u["d_goles_1t"]:
        mot["activo"] = True
        mot["badge"] = "🌓 Alta Inercia 2T"
        mot["etiqueta"] = "Goles concentrados tras el descanso"
        mot["mercado"] = "Empate al Descanso / Más Goles en la 2ª Mitad"
        m1 = _ev_matriz(max(lam_h * (1 - ratios[0]), 1e-6), max(lam_a * (1 - ratios[1]), 1e-6))
        p = float(np.trace(m1))
        mot["prob"] = p
        mot["cuota"] = _ev_cuota(p)
        mot["bullets"] = [
            f"Inercia: {ev_nombre_display(local)} marca el {ratios[0] * 100:.0f}% de sus goles en casa tras el descanso y "
            f"{ev_nombre_display(visitante)} el {ratios[1] * 100:.0f}% como visitante.",
            f"Primera mitad cerrada: apenas {g1:.2f} goles de media entre ambos antes del descanso.",
            _ev_txt_cuota(p),
        ]
    return mot


def ev_analizar_partido(L, df_arb, div, local, visitante, fecha, arbitro, n, cache=None):
    """Evalúa los 4 dominios y devuelve el payload (JSON-serializable) con el nº de factores alineados."""
    cut = pd.Timestamp(fecha).normalize()
    payload = {"ok": False, "partido": f"{ev_nombre_display(local)} vs {ev_nombre_display(visitante)}",
               "local": local, "visitante": visitante, "div": div, "liga": MAPA_LIGAS.get(div, div),
               "fecha": cut.strftime("%d/%m/%Y"), "arbitro": arbitro or "Sin designar", "score": 0, "motores": [],
               "badges": {"arbitro": "⏳ Pendiente Designación" if not arbitro else "⏳ Sin Histórico del Colegiado",
                          "descanso": "⚪ Sin dato previo", "xg": "⚪ Rendimiento Esperado", "mitades": "⚪ Distribución Neutra"}}
    vh = _ev_ventana(L, local, cut, n, True)
    va = _ev_ventana(L, visitante, cut, n, False)
    if len(vh) < EV_UMBRALES["min_muestra"] or len(va) < EV_UMBRALES["min_muestra"]:
        payload["motivo"] = "Muestra histórica insuficiente (menos de 5 partidos por condición)."
        return payload
    base = L[(L["Div"] == div) & (L["Fecha"] < cut)]
    lg_h = base[base["Local"]]["gf"].mean()
    lg_a = base[~base["Local"]]["gf"].mean()
    lg_h = float(lg_h) if pd.notna(lg_h) and lg_h > 0 else 1.45
    lg_a = float(lg_a) if pd.notna(lg_a) and lg_a > 0 else 1.15
    lam_h = float(np.clip(vh["gf"].mean() * va["ga"].mean() / lg_h, 0.2, 4.0))
    lam_a = float(np.clip(va["gf"].mean() * vh["ga"].mean() / lg_a, 0.2, 4.0))
    motores = [
        _ev_dominio_arbitro(L, df_arb, div, local, visitante, cut, n, arbitro, vh, va, cache),
        _ev_dominio_descanso(L, div, local, visitante, cut, n, lam_h, lam_a),
        _ev_dominio_xg(L, local, visitante, cut, lam_h, lam_a),
        _ev_dominio_mitades(local, visitante, vh, va, lam_h, lam_a),
    ]
    payload.update({
        "ok": True, "motores": motores, "lam_h": lam_h, "lam_a": lam_a,
        "score": int(sum(1 for m in motores if m["activo"])),
        "muestra": {"local": int(len(vh)), "visitante": int(len(va))},
        "badges": {"arbitro": motores[0]["badge"], "descanso": motores[1]["badge"],
                   "xg": motores[2]["badge"], "mitades": motores[3]["badge"]},
    })
    return payload


def ev_tesis_determinista(payload):
    """Plantilla determinista de la tesis cuantitativa (también es el fallback del LLM)."""
    lineas = [f"{payload.get('partido', '')} · {payload.get('liga', '')} · {payload.get('fecha', '')} — "
              f"{payload.get('score', 0)} factor(es) alineado(s)."]
    for m in payload.get("motores", []):
        if m.get("activo"):
            lineas.append(f"{m['nombre']}: {m['etiqueta']} → {m['mercado']}")
            for b in m.get("bullets", []):
                lineas.append(f"  - {b}")
    return "\n".join(lineas)


def generar_sintesis_narrativa_llm(anomaly_payload: dict) -> str:
    """Síntesis narrativa vía LM Studio local (timeout 1.20 s). Ante cualquier fallo devuelve la tesis determinista."""
    fallback = ev_tesis_determinista(anomaly_payload)
    try:
        import socket
        import requests
        # Sondeo TCP rápido (evita la espera de ~2 s de Windows ante conexiones rechazadas / IPv6)
        host_ok = None
        for host in ("127.0.0.1", "::1"):
            try:
                with socket.create_connection((host, 1234), timeout=0.25):
                    host_ok = host
                    break
            except OSError:
                continue
        if host_ok is None:
            return fallback
        url = EV_LLM_URL.replace("localhost", f"[{host_ok}]" if ":" in host_ok else host_ok)
        resumen = json_lib.dumps(
            {k: anomaly_payload.get(k) for k in ("partido", "liga", "fecha", "arbitro", "score")}
            | {"dominios_activos": [
                {"dominio": m["nombre"], "etiqueta": m["etiqueta"], "mercado": m["mercado"], "evidencia": m["bullets"]}
                for m in anomaly_payload.get("motores", []) if m.get("activo")
            ]},
            ensure_ascii=False,
        )
        resp = requests.post(
            url,
            json={
                "model": "local-model",
                "messages": [
                    {"role": "system", "content": "Eres un analista cuantitativo de fútbol. Redacta en español, en 4-5 líneas, una tesis "
                                                  "de inversión concisa usando SOLO los datos facilitados. No inventes cifras."},
                    {"role": "user", "content": resumen},
                ],
                "temperature": 0.2,
                "max_tokens": 260,
            },
            timeout=1.20,
        )
        resp.raise_for_status()
        texto = str(resp.json()["choices"][0]["message"]["content"]).strip()
        return texto if texto else fallback
    except Exception:
        return fallback


def ev_render_tarjeta(p, usar_llm=False):
    """Tarjeta ejecutiva: cabecera + pill, banner de mercado, tesis cuantitativa."""
    esc = html_lib.escape
    alto = p["score"] >= EV_UMBRALES["score_alto_valor"]
    color = "#F97316" if alto else "#EAB308"
    pill = "🔥 Oportunidad de Alto Valor (+EV Confirmado)" if alto else "🟡 Oportunidad Detectada (2 Factores Alineados)"
    activos = [m for m in p["motores"] if m["activo"]]
    with st.container(border=True):
        st.markdown(
            f"<div style='display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;'>"
            f"<div><span style='font-size:18px;font-weight:700;'>Partido: {esc(p['partido'])}</span>"
            f"<div style='color:#94A3B8;font-size:12.5px;margin-top:2px;'>🏆 {esc(str(p['liga']))} &nbsp;·&nbsp; "
            f"📅 {esc(p['fecha'])} &nbsp;·&nbsp; 🧑‍⚖️ {esc(str(p['arbitro']))}</div></div>"
            f"<span style='padding:4px 12px;border-radius:16px;border:1px solid {color};color:{color};font-weight:700;"
            f"font-size:13px;'>{pill}</span></div>",
            unsafe_allow_html=True,
        )
        for m in activos:
            st.markdown(
                f"<div style='background:rgba(16,185,129,0.12);border:1px solid #10B981;border-radius:8px;"
                f"padding:8px 12px;margin:8px 0 2px 0;'>🎯 <b>{esc(m['categoria'])}:</b> {esc(m['mercado'])}"
                f"<span style='color:#94A3B8;'> &nbsp;·&nbsp; {esc(m['nombre'])}</span></div>",
                unsafe_allow_html=True,
            )
        st.markdown("**Tesis Cuantitativa**")
        for m in activos:
            st.markdown(f"_{m['nombre']} — {m['etiqueta']}_")
            st.markdown("\n".join(f"- {b}" for b in m["bullets"]))
        if usar_llm:
            with st.expander("🧠 Síntesis narrativa (LLM local)", expanded=False):
                st.write(generar_sintesis_narrativa_llm(p))


@st.cache_data(ttl=3600, show_spinner=False)
def ev_unir_arbitros(df_min, df_arb):
    """Completa 'arbitros_stats' con árbitros agregados desde el histórico (mismas 9 columnas posicionales)."""
    cols = ["nom", "div", "liga", "comite", "partidos", "perfil", "am", "rj", "fl"]
    base = pd.DataFrame(columns=cols)
    if df_arb is not None and not df_arb.empty and df_arb.shape[1] >= 9:
        base = df_arb.iloc[:, :9].copy()
        base.columns = cols
    d = df_min.dropna(subset=["Referee"]).copy()
    d = d[d["Referee"].astype(str).str.strip() != ""]
    if not d.empty:
        for c in ("HY", "AY", "HR", "AR", "HF", "AF"):
            d[c] = pd.to_numeric(d[c], errors="coerce")
        d["am"] = d["HY"] + d["AY"]
        d["rj"] = d["HR"] + d["AR"]
        d["fl"] = d["HF"] + d["AF"]
        g = d.groupby(["Div", "Referee"]).agg(partidos=("am", "count"), am=("am", "mean"),
                                              rj=("rj", "mean"), fl=("fl", "mean")).reset_index()
        g = g[g["partidos"] >= 3]
        h = pd.DataFrame({
            "nom": g["Referee"].astype(str).str.strip(), "div": g["Div"], "liga": g["Div"].map(MAPA_LIGAS).fillna(g["Div"]),
            "comite": "Histórico BD", "partidos": g["partidos"], "perfil": "Calculado",
            "am": g["am"], "rj": g["rj"], "fl": g["fl"],
        })
        if not base.empty:
            clave = set(zip(base["nom"].astype(str).str.strip().str.lower(), base["div"].astype(str)))
            h = h[[(n.lower(), str(dv)) not in clave for n, dv in zip(h["nom"], h["div"])]]
        base = pd.concat([base, h], ignore_index=True)
    return base


def _ev_obtener_fixtures(df_min, L, div, fuente):
    """Devuelve DataFrame [Local, Visitante, Fecha, Arbitro] según la fuente elegida."""
    d = df_min[df_min["Div"] == div].copy()
    d["Fecha"] = pd.to_datetime(d["Date"], errors="coerce", dayfirst=True)
    d["FTHG"] = pd.to_numeric(d["FTHG"], errors="coerce")
    cols = ["Local", "Visitante", "Fecha", "Arbitro"]
    if fuente.startswith("Próximos"):
        f = d[d["FTHG"].isna() & d["Fecha"].notna()].sort_values("Fecha")
        if not f.empty:
            f = f[f["Fecha"] <= f["Fecha"].iloc[0] + pd.Timedelta(days=6)]
    elif fuente.startswith("Backtest"):
        f = d[d["FTHG"].notna() & d["Fecha"].notna()]
        if not f.empty:
            f = f[f["Fecha"] >= f["Fecha"].max() - pd.Timedelta(days=4)].sort_values("Fecha")
    else:
        man = [x for x in st.session_state.get("ev_fix_manual", []) if x["Div"] == div]
        return pd.DataFrame([{k: x[k] for k in cols} for x in man], columns=cols)
    out = pd.DataFrame({
        "Local": f["HomeTeam"].values, "Visitante": f["AwayTeam"].values, "Fecha": f["Fecha"].values,
        "Arbitro": f["Referee"].fillna("").astype(str).values,
    })
    return out[cols]


def _ev_estilo_badge(valor):
    v = str(valor)
    if any(k in v for k in ("Permisivo", "Favorable", "Alto Valor")) and "Desfavorable" not in v:
        return "background-color: rgba(16,185,129,0.22); color: #6EE7B7;"
    if any(k in v for k in ("Riguroso", "Desfavorable")):
        return "background-color: rgba(239,68,68,0.22); color: #FCA5A5;"
    if any(k in v for k in ("Rebote", "Alta Inercia", "Regresión", "Oportunidad")):
        return "background-color: rgba(249,115,22,0.22); color: #FDBA74;"
    if "Pendiente" in v or "Sin " in v:
        return "background-color: rgba(100,116,139,0.25); color: #CBD5E1;"
    return ""


def ev_render_modulo():
    import time
    st.title("🎯 Escáner Autónomo de Ineficiencias (+EV Scanner)")
    st.caption("Consola de decisión: cuatro dominios cuantitativos (árbitro, físico, xG y mitades) cruzados para detectar "
               "mercados mal valorados. Ejecución 100% local (SQLite + pandas).")
    if df_partidos_global.empty:
        st.error("No hay datos de partidos cargados desde macro_base_futbol.db.")
        return

    df_min = df_partidos_global.reindex(columns=EV_COLS_BASE)
    L_all = ev_construir_historico_largo(df_min)
    df_arb = ev_unir_arbitros(df_min, cargar_datos_arbitros(3))
    todos_equipos = sorted(pd.concat([df_min["HomeTeam"], df_min["AwayTeam"]]).dropna().unique().tolist())

    c1, c2, c3 = st.columns(3)
    divs = [d for d in sorted(df_min["Div"].dropna().unique().tolist())]
    nombres = [MAPA_LIGAS.get(d, d) for d in divs]
    with c1:
        liga_nom = st.selectbox("Competición:", nombres, key="ev_liga")
    div = divs[nombres.index(liga_nom)]
    with c2:
        n_ventana = st.slider("Ventana histórica (N partidos Casa/Fuera):", 10, 15, 12, key="ev_n")

    d_div = df_min[df_min["Div"] == div]
    hay_futuros = bool((pd.to_numeric(d_div["FTHG"], errors="coerce").isna()).any())
    opciones_src = (["Próximos partidos (BD)"] if hay_futuros else []) + \
                   ["Fixture manual", "Backtest: última jornada registrada"]
    with c3:
        fuente = st.selectbox("Fuente de fixtures:", opciones_src, key="ev_fuente")

    equipos = sorted(pd.concat([d_div["HomeTeam"], d_div["AwayTeam"]]).dropna().unique().tolist())
    if fuente == "Fixture manual":
        with st.expander("➕ Añadir partido programado", expanded=True):
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                h_man = st.selectbox("Local:", equipos, format_func=ev_nombre_display, key="ev_man_h")
            with m2:
                a_man = st.selectbox("Visitante:", [e for e in equipos if e != h_man],
                                     format_func=ev_nombre_display, key="ev_man_a")
            with m3:
                f_man = st.date_input("Fecha:", value=datetime.now().date(), key="ev_man_f")
            with m4:
                arbs = []
                if df_arb is not None and not df_arb.empty:
                    arbs = sorted(df_arb[df_arb["div"].astype(str) == div]["nom"].astype(str).unique().tolist())
                r_man = st.selectbox("Árbitro designado:", ["(Pendiente de designación)"] + arbs, key="ev_man_r")
            arbitro_man = "" if r_man.startswith("(Pendiente") else r_man
            b1, b2 = st.columns(2)
            with b1:
                if st.button("Añadir al fixture", key="ev_man_add"):
                    lista = st.session_state.get("ev_fix_manual", [])
                    lista.append({"Div": div, "Local": h_man, "Visitante": a_man,
                                  "Fecha": pd.Timestamp(f_man), "Arbitro": arbitro_man})
                    st.session_state["ev_fix_manual"] = lista
            with b2:
                if st.button("Vaciar fixture manual", key="ev_man_clear"):
                    st.session_state["ev_fix_manual"] = []
            txt_libre = st.text_input("…o escribe el partido (cualquier alias): «Local vs Visitante»",
                                      placeholder="Atlético de Madrid vs Real Sociedad", key="ev_man_txt")
            if st.button("Resolver y añadir texto", key="ev_man_txt_add"):
                partes = [x for x in txt_libre.replace(" - ", " vs ").split(" vs ") if x.strip()]
                if len(partes) != 2:
                    st.warning("Usa el formato «Local vs Visitante».")
                else:
                    rl = ev_resolver_equipo(partes[0], equipos) or ev_resolver_equipo(partes[0], todos_equipos)
                    rv = ev_resolver_equipo(partes[1], equipos) or ev_resolver_equipo(partes[1], todos_equipos)
                    if rl is None or rv is None or rl == rv:
                        st.warning(f"No se pudo resolver el partido: {partes[0].strip()} / {partes[1].strip()}.")
                    else:
                        lista = st.session_state.get("ev_fix_manual", [])
                        lista.append({"Div": div, "Local": rl, "Visitante": rv,
                                      "Fecha": pd.Timestamp(f_man), "Arbitro": arbitro_man})
                        st.session_state["ev_fix_manual"] = lista
                        st.success(f"Resuelto: {ev_nombre_display(rl)} vs {ev_nombre_display(rv)}")

    fixtures = _ev_obtener_fixtures(df_min, L_all, div, fuente)
    if fuente.startswith("Backtest"):
        st.info("Modo backtest: cada partido se evalúa únicamente con datos anteriores a su fecha (sin look-ahead en equipos).")
    if fixtures.empty:
        st.warning("No hay partidos que escanear con la fuente seleccionada.")
        return

    usar_llm = st.checkbox("Síntesis narrativa con LLM local (LM Studio en localhost:1234, con fallback automático)",
                           value=False, key="ev_llm")

    t0 = time.perf_counter()
    L = L_all[L_all["Div"] == div]
    cache = {}
    resultados = []
    for _, fx in fixtures.iterrows():
        rl = ev_resolver_equipo(fx["Local"], equipos)
        rv = ev_resolver_equipo(fx["Visitante"], equipos)
        if rl is None or rv is None:
            resultados.append({"ok": False, "partido": f"{fx['Local']} vs {fx['Visitante']}", "score": 0, "motores": [],
                               "fecha": pd.Timestamp(fx["Fecha"]).strftime("%d/%m/%Y"), "arbitro": fx["Arbitro"] or "Sin designar",
                               "motivo": "Equipo no reconocido o ajeno a la competición seleccionada.", "badges": {}})
            continue
        resultados.append(ev_analizar_partido(L, df_arb, div, rl, rv, fx["Fecha"], str(fx["Arbitro"] or ""), n_ventana, cache))
    ms = (time.perf_counter() - t0) * 1000.0

    def _estado(p):
        if not p["ok"]:
            return "⚪ Sin datos suficientes"
        if p["score"] >= EV_UMBRALES["score_alto_valor"]:
            return "🔥 Alto Valor"
        if p["score"] >= EV_UMBRALES["score_oportunidad"]:
            return "🟡 Oportunidad"
        if p["score"] == 1:
            return "👁️ Radar"
        return "⚪ Sin señales"

    filas = []
    for p in resultados:
        b = p.get("badges") or {}
        filas.append({
            "Partido": p["partido"], "Fecha": p["fecha"],
            "Árbitro": b.get("arbitro", "⏳ Pendiente Designación"),
            "Descanso (Local)": b.get("descanso", "⚪ Sin dato previo"),
            "xG": b.get("xg", "⚪ Rendimiento Esperado"),
            "Mitades": b.get("mitades", "⚪ Distribución Neutra"),
            "Factores": p["score"], "Estado": _estado(p),
        })

    altas = sorted([p for p in resultados if p["ok"] and p["score"] >= EV_UMBRALES["score_alto_valor"]], key=lambda x: -x["score"])
    medias = [p for p in resultados if p["ok"] and p["score"] == EV_UMBRALES["score_oportunidad"]]
    radar = [p for p in resultados if p["ok"] and p["score"] == 1]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Partidos escaneados", len(resultados))
    k2.metric("🔥 Alto valor (≥3 factores)", len(altas))
    k3.metric("🟡 Oportunidades (2 factores)", len(medias))
    k4.metric("Tiempo de resolución", f"{ms:.0f} ms")

    st.subheader("📋 Resumen de la Jornada")
    df_tabla = pd.DataFrame(filas)
    try:
        st.dataframe(df_tabla.style.map(_ev_estilo_badge, subset=["Árbitro", "Descanso (Local)", "xG", "Mitades", "Estado"]),
                     use_container_width=True, hide_index=True)
    except Exception:
        st.dataframe(df_tabla, use_container_width=True, hide_index=True)

    st.subheader("🚨 Oportunidades Detectadas")
    if not altas and not medias:
        st.info("Ningún partido reúne al menos 2 factores alineados. Revisa el Radar de Seguimiento.")
    for p in altas + medias:
        ev_render_tarjeta(p, usar_llm)

    with st.expander(f"👁️ Radar de Seguimiento ({len(radar)} partidos con 1 señal activa)", expanded=False):
        if not radar:
            st.write("Sin partidos en seguimiento.")
        for p in radar:
            m = next(x for x in p["motores"] if x["activo"])
            st.write(f"**{p['partido']}** ({p['fecha']}) — {m['nombre']}: {m['mercado']}")

    with st.expander("📚 Cómo leer este panel", expanded=False):
        st.markdown(
            "- **⚖️ Árbitro y Tarjetas:** percentil disciplinario del colegiado en su liga + fricción de cada equipo "
            "(faltas propias y recibidas + tarjetas).\n"
            "- **🔋 Descanso y Desgaste Físico:** diferencia de días desde el partido previo (señal desde ±3 días).\n"
            "- **🎯 Eficiencia Ofensiva / Regresión xG:** goles reales frente a goles esperados en los últimos 6 partidos "
            "(xG real si existe; si no, estimado por tiros).\n"
            "- **⏱️ Inercia de Segundas Partes:** peso de los goles tras el descanso y goles medios de la primera mitad.\n"
            "- **Niveles:** 1 señal = Radar · 2 señales = Oportunidad Detectada · 3 o más = Oportunidad de Alto Valor.\n"
            "- Las cuotas justas son la inversa de la probabilidad del modelo, sin margen de la casa."
        )


# ==============================================================================
# NAVEGACIÓN Y SECCIONES
# ==============================================================================

if menu_seleccionado == "🏠 Inicio / Portada":
    st.title("⚽ Hugo's Hub: Sistema Analítico de Apuestas")
    st.markdown("Plataforma avanzada de procesamiento de Big Data futbolístico y detección de valor matemático.")
    st.markdown("---")
    col_i1, col_i2 = st.columns(2)
    with col_i1:
        st.info("📌 **Estado del Sistema**")
        st.markdown(f"- **Partidos indexados:** `{len(df_partidos_global)}` encuentros")
        st.markdown(f"- **Jugadores indexados:** `{len(df_jugadores)}` futbolistas")
        st.markdown("- **Conector API:** Sincronización estricta de múltiples mercados.")
    with col_i2:
        st.success("🚀 **Módulos Activos**")
        st.markdown("- **Escáner EV+ Multimercado:** 1X2, Líneas Goles (+/-) y BTTS.")
        st.markdown("- **Comparador H2H Pro:** Cuotas Justas y filtros Local/Visitante.")
        st.markdown("- **Matchup de Árbitros:** Estimación de tarjetas.")
        st.markdown("- **Dossier de Equipo:** Radiografía completa de promedios avanzados.")
        st.markdown("- **Hándicap Asiático Pro:** Histórico ATS, Goles y Matriz Poisson.")
        st.markdown("- **Scouting Métrico y Radares.**")

elif menu_seleccionado == "📡 Escáner Vivo EV+ (API)":
    st.markdown(
        """
        <style>
        div[data-testid="stButton"] > button[kind="secondary"],
        div[data-testid="stButton"] > button:not([kind="primary"]) {
            background-color: #F8FAFC !important;
            color: #0F172A !important;
            border: 1.5px solid #38BDF8 !important;
            font-weight: 700 !important;
            border-radius: 8px !important;
        }
        div[data-testid="stButton"] > button[kind="secondary"] *,
        div[data-testid="stButton"] > button:not([kind="primary"]) * {
            color: #0F172A !important;
            font-weight: 700 !important;
        }
        div[data-testid="stButton"] > button[kind="secondary"]:hover,
        div[data-testid="stButton"] > button:not([kind="primary"]):hover {
            background-color: #E2E8F0 !important;
            color: #0F172A !important;
            border-color: #0284C7 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.title("📡 Escáner Cuantitativo de Valor Institucional (EV+ Pro)")
    st.caption(
        "Motor de convergencia **Poisson + Mercado No-Vig (40/60)** · Auditoría cruzada SQLite (Rolling 6G xG/xGA) · "
        "Fricción Táctica Arbitral · Backtesting Empírico Express · Supresión de correlaciones por eficiencia de Kelly."
    )

    # ── PILAR 2: DE-MARGINING (NO-VIG) & CONVERGENCE HELPER ──────────────────
    def calcular_probabilidad_no_vig(cuotas_1x2: list[float]) -> list[float]:
        """Elimina el margen de la casa (método multiplicativo normalizado)."""
        inv = [1.0 / c if c > 1.0 else 0.0 for c in cuotas_1x2]
        overround = sum(inv)
        if overround <= 0:
            return [0.33, 0.33, 0.33]
        return [i / overround for i in inv]

    # Exponer también en el espacio global del módulo para validación externa
    globals()["calcular_probabilidad_no_vig"] = calcular_probabilidad_no_vig

    def _obtener_df_partidos_scanner() -> pd.DataFrame:
        """Lectura segura de 'historico_reciente' / 'partidos' desde memoria o SQLite."""
        try:
            if isinstance(df_partidos_global, pd.DataFrame) and not df_partidos_global.empty:
                return df_partidos_global
            if RUTA_BD.exists():
                with sqlite3.connect(RUTA_BD) as conn_sc:
                    tablas_sc = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table';", conn_sc)["name"].tolist()
                    t_nom = "historico_reciente" if "historico_reciente" in tablas_sc else ("partidos" if "partidos" in tablas_sc else None)
                    if t_nom:
                        df_db = pd.read_sql(f"SELECT * FROM {t_nom}", conn_sc)
                        df_db.columns = df_db.columns.str.replace("ï»¿", "").str.strip()
                        return df_db
        except Exception:
            pass
        return pd.DataFrame()

    def _auditar_equipo_sqlite_6g(equipo_nombre: str) -> dict:
        """
        Consulta resiliente a SQLite ('macro_base_futbol.db') para extraer rolling xG/xGA,
        faltas históricas y tarjetas en los últimos 6 partidos, estabilidad y profundidad muestral.
        """
        resultado_default = {
            "n_total": 0,
            "n_rolling": 0,
            "xg_f_6": 1.35,
            "xga_c_6": 1.25,
            "gf_6": 1.35,
            "ga_6": 1.25,
            "xg_std": 0.45,
            "ha05_ratio": 1.00,
            "faltas_com_6": 11.8,
            "faltas_rec_6": 11.5,
            "tarjetas_6": 2.15,
            "div_eq": "E0",
            "ultimo_arbitro": "",
        }
        try:
            df_src = _obtener_df_partidos_scanner()
            if df_src.empty:
                return resultado_default

            eq_clean = str(equipo_nombre).strip()
            mask_h = df_src["HomeTeam"].astype(str).str.lower() == eq_clean.lower()
            mask_a = df_src["AwayTeam"].astype(str).str.lower() == eq_clean.lower()
            df_eq = df_src[mask_h | mask_a].copy()
            if df_eq.empty:
                mask_h = df_src["HomeTeam"].astype(str).str.contains(eq_clean, case=False, na=False)
                mask_a = df_src["AwayTeam"].astype(str).str.contains(eq_clean, case=False, na=False)
                df_eq = df_src[mask_h | mask_a].copy()
            if df_eq.empty:
                return resultado_default

            if "Date" in df_eq.columns:
                df_eq["_dt"] = pd.to_datetime(df_eq["Date"], errors="coerce", dayfirst=True)
                df_eq = df_eq.sort_values("_dt")

            n_total = len(df_eq)
            df_6 = df_eq.tail(6).copy()
            es_loc = df_6["HomeTeam"].astype(str).str.lower() == eq_clean.lower()

            fthg = pd.to_numeric(df_6.get("FTHG", 1.3), errors="coerce").fillna(1.3)
            ftag = pd.to_numeric(df_6.get("FTAG", 1.1), errors="coerce").fillna(1.1)
            hst = pd.to_numeric(df_6.get("HST", 4.2), errors="coerce").fillna(4.2)
            ast = pd.to_numeric(df_6.get("AST", 3.6), errors="coerce").fillna(3.6)
            hs = pd.to_numeric(df_6.get("HS", 12.0), errors="coerce").fillna(12.0)
            as_ = pd.to_numeric(df_6.get("AS", 10.0), errors="coerce").fillna(10.0)
            hf = pd.to_numeric(df_6.get("HF", 11.5), errors="coerce").fillna(11.5)
            af = pd.to_numeric(df_6.get("AF", 11.8), errors="coerce").fillna(11.8)
            hy = pd.to_numeric(df_6.get("HY", 2.0), errors="coerce").fillna(2.0)
            ay = pd.to_numeric(df_6.get("AY", 2.2), errors="coerce").fillna(2.2)
            hr = pd.to_numeric(df_6.get("HR", 0.08), errors="coerce").fillna(0.08)
            ar = pd.to_numeric(df_6.get("AR", 0.08), errors="coerce").fillna(0.08)

            hxg_raw = pd.to_numeric(df_6["HxG"], errors="coerce") if "HxG" in df_6.columns else pd.Series(np.nan, index=df_6.index)
            axg_raw = pd.to_numeric(df_6["AxG"], errors="coerce") if "AxG" in df_6.columns else pd.Series(np.nan, index=df_6.index)

            hxg_est = hxg_raw.where(hxg_raw.notna(), 0.30 * hst + 0.03 * (hs - hst).clip(lower=0))
            axg_est = axg_raw.where(axg_raw.notna(), 0.30 * ast + 0.03 * (as_ - ast).clip(lower=0))

            xg_propio = np.where(es_loc, hxg_est, axg_est)
            xga_rival = np.where(es_loc, axg_est, hxg_est)
            gf_propio = np.where(es_loc, fthg, ftag)
            ga_rival = np.where(es_loc, ftag, fthg)
            faltas_com = np.where(es_loc, hf, af)
            faltas_rec = np.where(es_loc, af, hf)
            tarj_prop = np.where(es_loc, hy + hr, ay + ar)

            ha05_ratio = 1.00
            if "AHh" in df_eq.columns and "MaxAHH" in df_eq.columns and "B365H" in df_eq.columns:
                sub_ah = df_eq[pd.to_numeric(df_eq["AHh"], errors="coerce") == -0.5]
                if not sub_ah.empty:
                    c_ah = pd.to_numeric(sub_ah["MaxAHH"], errors="coerce")
                    c_1x2 = pd.to_numeric(sub_ah["B365H"], errors="coerce")
                    ratios = (c_ah / c_1x2).dropna()
                    if not ratios.empty:
                        ha05_ratio = float(np.clip(ratios.mean(), 0.95, 1.12))

            div_val = "E0"
            if "Div" in df_eq.columns and not df_eq["Div"].dropna().empty:
                div_val = str(df_eq["Div"].dropna().iloc[-1]).strip()

            ult_arb = ""
            if "Referee" in df_eq.columns and not df_eq["Referee"].dropna().empty:
                ult_arb = str(df_eq["Referee"].dropna().iloc[-1]).strip()

            return {
                "n_total": int(n_total),
                "n_rolling": int(len(df_6)),
                "xg_f_6": float(np.nanmean(xg_propio)) if len(xg_propio) > 0 else 1.35,
                "xga_c_6": float(np.nanmean(xga_rival)) if len(xga_rival) > 0 else 1.25,
                "gf_6": float(np.nanmean(gf_propio)) if len(gf_propio) > 0 else 1.35,
                "ga_6": float(np.nanmean(ga_rival)) if len(ga_rival) > 0 else 1.25,
                "xg_std": float(np.nanstd(xg_propio)) if len(xg_propio) > 1 else 0.35,
                "ha05_ratio": ha05_ratio,
                "faltas_com_6": float(np.nanmean(faltas_com)) if len(faltas_com) > 0 else 11.8,
                "faltas_rec_6": float(np.nanmean(faltas_rec)) if len(faltas_rec) > 0 else 11.5,
                "tarjetas_6": float(np.nanmean(tarj_prop)) if len(tarj_prop) > 0 else 2.15,
                "div_eq": div_val,
                "ultimo_arbitro": ult_arb,
            }
        except Exception:
            return resultado_default

    def _auditar_friccion_arbitral_sqlite(aud_h: dict, aud_a: dict, liga_nombre: str, arbitro_override: str = "") -> dict:
        """
        Consulta la tabla 'arbitros_stats' en SQLite y calcula la Tensión Disciplinaria
        combinando las faltas históricas de los equipos y el promedio de tarjetas del colegiado.
        """
        faltas_comb = float(aud_h.get("faltas_com_6", 11.8) + aud_a.get("faltas_com_6", 11.8))
        tarj_eq_comb = float(aud_h.get("tarjetas_6", 2.1) + aud_a.get("tarjetas_6", 2.1))
        div_code = MAPA_INV.get(liga_nombre, aud_h.get("div_eq", "E0"))

        arb_nombre = arbitro_override.strip() or aud_h.get("ultimo_arbitro", "") or aud_a.get("ultimo_arbitro", "")
        arb_amarillas = 4.05
        arb_rojas = 0.14
        arb_faltas = 22.5

        try:
            df_arb_sql = pd.DataFrame()
            if RUTA_BD.exists():
                with sqlite3.connect(RUTA_BD) as conn_arb:
                    tablas_a = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table';", conn_arb)["name"].tolist()
                    if "arbitros_stats" in tablas_a:
                        df_arb_sql = pd.read_sql("SELECT * FROM arbitros_stats", conn_arb)
            if df_arb_sql.empty:
                df_arb_sql = cargar_datos_arbitros(3)

            if isinstance(df_arb_sql, pd.DataFrame) and not df_arb_sql.empty and df_arb_sql.shape[1] >= 9:
                col_nom = df_arb_sql.columns[0]
                col_div = df_arb_sql.columns[1]
                col_am = df_arb_sql.columns[6]
                col_rj = df_arb_sql.columns[7]
                col_fl = df_arb_sql.columns[8]

                sub_div = df_arb_sql[df_arb_sql[col_div].astype(str) == str(div_code)].copy()
                if sub_div.empty:
                    sub_div = df_arb_sql.copy()

                fila_sel = pd.DataFrame()
                if arb_nombre:
                    fila_sel = sub_div[sub_div[col_nom].astype(str).str.lower() == arb_nombre.lower()]
                if fila_sel.empty and not sub_div.empty:
                    # Seleccionar colegiado representativo del enfrentamiento de forma determinista
                    idx_det = int(abs(hash(f"{aud_h.get('n_total', 0)}_{aud_a.get('n_total', 0)}_{faltas_comb:.1f}"))) % len(sub_div)
                    fila_sel = sub_div.iloc[[idx_det]]

                if not fila_sel.empty:
                    r_a = fila_sel.iloc[0]
                    arb_nombre = str(r_a[col_nom]).strip()
                    arb_amarillas = float(pd.to_numeric(r_a[col_am], errors="coerce") or 4.05)
                    arb_rojas = float(pd.to_numeric(r_a[col_rj], errors="coerce") or 0.14)
                    arb_faltas = float(pd.to_numeric(r_a[col_fl], errors="coerce") or 22.5)
        except Exception:
            pass

        if not arb_nombre:
            arb_nombre = "Colegiado Media Comité"

        score_tarjetas_arb = arb_amarillas + (2.0 * arb_rojas)
        tension_disciplinaria = (score_tarjetas_arb * 5.5) + (faltas_comb * 1.15) + (tarj_eq_comb * 1.5)
        es_alta = (score_tarjetas_arb >= 4.15 and faltas_comb >= 22.5) or (tension_disciplinaria >= 56.0)

        if es_alta:
            badge_tactico = "🔥 Fricción Táctica Alta: Árbitro riguroso + alta fricción de faltas (Sesgo a tarjetas / Interrupciones)"
        else:
            badge_tactico = "🍃 Partido Fluido: Árbitro permisivo + baja tasa de faltas"

        return {
            "arbitro": arb_nombre,
            "arb_amarillas": round(arb_amarillas, 2),
            "arb_rojas": round(arb_rojas, 2),
            "arb_faltas": round(arb_faltas, 1),
            "faltas_combinadas": round(faltas_comb, 1),
            "tension_score": round(tension_disciplinaria, 1),
            "es_alta": es_alta,
            "badge_tactico": badge_tactico,
        }

    def _backtest_empirico_db(
        liga_nombre: str,
        tipo_mercado: str,
        mercado_nom: str,
        cuota_ofrecida: float,
        eq_local: str,
        eq_vis: str,
    ) -> dict:
        """
        Ejecuta una consulta rápida sobre 'macro_base_futbol.db' para evaluar el rendimiento
        histórico empírico en partidos pasados bajo condiciones de cuota y mercado equivalentes
        en la misma liga.
        """
        badge_insuf = "⚠️ Muestra histórica insuficiente en DB local"
        try:
            df_src = _obtener_df_partidos_scanner()
            if df_src.empty:
                return {"n": 0, "hit_rate": 0.0, "roi": 0.0, "badge": badge_insuf}

            div_code = MAPA_INV.get(liga_nombre)
            if div_code and "Div" in df_src.columns:
                df_l = df_src[df_src["Div"].astype(str) == str(div_code)].copy()
            else:
                df_l = df_src.copy()

            fthg = pd.to_numeric(df_l.get("FTHG"), errors="coerce")
            ftag = pd.to_numeric(df_l.get("FTAG"), errors="coerce")
            mask_valid = fthg.notna() & ftag.notna()
            df_l = df_l[mask_valid].copy()
            fthg = fthg[mask_valid]
            ftag = ftag[mask_valid]
            goles_tot = fthg + ftag

            if df_l.empty:
                return {"n": 0, "hit_rate": 0.0, "roi": 0.0, "badge": badge_insuf}

            aciertos_mask = None
            col_cuota_hist = None

            if tipo_mercado == "1X2":
                if "Local" in mercado_nom or "(1)" in mercado_nom:
                    aciertos_mask = fthg > ftag
                    col_cuota_hist = "AvgH" if "AvgH" in df_l.columns else ("B365H" if "B365H" in df_l.columns else None)
                elif "Empate" in mercado_nom or "(X)" in mercado_nom:
                    aciertos_mask = fthg == ftag
                    col_cuota_hist = "AvgD" if "AvgD" in df_l.columns else ("B365D" if "B365D" in df_l.columns else None)
                else:
                    aciertos_mask = ftag > fthg
                    col_cuota_hist = "AvgA" if "AvgA" in df_l.columns else ("B365A" if "B365A" in df_l.columns else None)
            elif tipo_mercado == "Goles":
                es_over = "Más" in mercado_nom or "Over" in mercado_nom
                linea_g = 2.5
                for lv in (0.5, 1.5, 2.5, 3.5, 4.5):
                    if str(lv) in mercado_nom:
                        linea_g = lv
                        break
                aciertos_mask = (goles_tot > linea_g) if es_over else (goles_tot < linea_g)
                if linea_g == 2.5:
                    if es_over:
                        col_cuota_hist = "Avg>2.5" if "Avg>2.5" in df_l.columns else ("B365>2.5" if "B365>2.5" in df_l.columns else None)
                    else:
                        col_cuota_hist = "Avg<2.5" if "Avg<2.5" in df_l.columns else ("B365<2.5" if "B365<2.5" in df_l.columns else None)
            elif tipo_mercado == "BTTS":
                es_si = "Sí" in mercado_nom or "Yes" in mercado_nom
                aciertos_mask = ((fthg > 0) & (ftag > 0)) if es_si else ((fthg == 0) | (ftag == 0))

            if aciertos_mask is None:
                return {"n": 0, "hit_rate": 0.0, "roi": 0.0, "badge": badge_insuf}

            # Filtrar por banda de cuota histórica similar (±0.22) o partidos de ambos equipos en la liga
            if col_cuota_hist and col_cuota_hist in df_l.columns:
                c_serie = pd.to_numeric(df_l[col_cuota_hist], errors="coerce")
                mask_sim = c_serie.notna() & (c_serie >= (cuota_ofrecida - 0.22)) & (c_serie <= (cuota_ofrecida + 0.22))
                sub_hits = aciertos_mask[mask_sim]
                sub_cuotas = c_serie[mask_sim]
            else:
                mask_eqs = (
                    df_l["HomeTeam"].astype(str).str.lower().isin([eq_local.lower(), eq_vis.lower()])
                    | df_l["AwayTeam"].astype(str).str.lower().isin([eq_local.lower(), eq_vis.lower()])
                )
                sub_hits = aciertos_mask[mask_eqs]
                sub_cuotas = pd.Series(cuota_ofrecida, index=sub_hits.index)

            n_sim = int(len(sub_hits))
            if n_sim < 5:
                return {"n": n_sim, "hit_rate": 0.0, "roi": 0.0, "badge": badge_insuf}

            hit_rate = float(sub_hits.mean() * 100.0)
            retornos = np.where(sub_hits.to_numpy(dtype=bool), sub_cuotas.to_numpy(dtype=float), 0.0)
            roi_pct = float((np.mean(retornos) - 1.0) * 100.0)

            badge_ok = f"📈 Backtest DB: Acierto del {hit_rate:.1f}% en {n_sim} partidos similares (ROI: {roi_pct:+.1f}%)"
            return {"n": n_sim, "hit_rate": hit_rate, "roi": roi_pct, "badge": badge_ok}
        except Exception:
            return {"n": 0, "hit_rate": 0.0, "roi": 0.0, "badge": badge_insuf}

    def _enriquecer_y_filtrar_institucional(
        ops_raw: list[dict],
        libros_api: dict,
        cuota_min_f: float,
        cuota_max_f: float,
        ev_min_f: float,
        tipo_mercado_f: str,
        equipo_f: str,
        liga_actual: str,
        suprimir_correlacion: bool = True,
    ) -> tuple[pd.DataFrame, int]:
        """
        Ejecuta los Pilares cuantitativos sobre los candidatos crudos del escáner:
        - De-margining No-Vig y probabilidad convergente (40% Poisson Dossier + 60% Mercado No-Vig)
        - Cap de outliers > 30% EV ('⚠️ Desfase API')
        - Auditoría SQLite rolling 6G xG/xGA y penalización por divergencia > 35%
        - Fricción Táctica Arbitral (arbitros_stats + faltas históricas)
        - Backtesting Empírico Express en macro_base_futbol.db
        - Conviction Score multivariante (0-100) y Cuota Mínima Operable (1.03 / P_final)
        - Auditoría de eficiencia sintética cruzada (1X2 vs HA -0.5)
        - Supresión de selecciones correlacionadas en el mismo partido por Kelly Score
        """
        if not ops_raw:
            return pd.DataFrame(), 0

        registros = []
        cache_sql = {}
        cache_arb = {}

        for item in ops_raw:
            try:
                partido_str = str(item.get("Partido", ""))
                partes_eq = partido_str.split(" vs ")
                eq_local = partes_eq[0].strip() if len(partes_eq) >= 1 else partido_str
                eq_vis = partes_eq[1].strip() if len(partes_eq) >= 2 else ""

                if eq_local not in cache_sql:
                    cache_sql[eq_local] = _auditar_equipo_sqlite_6g(eq_local)
                if eq_vis not in cache_sql:
                    cache_sql[eq_vis] = _auditar_equipo_sqlite_6g(eq_vis)
                aud_h = cache_sql[eq_local]
                aud_a = cache_sql[eq_vis]

                if partido_str not in cache_arb:
                    cache_arb[partido_str] = _auditar_friccion_arbitral_sqlite(aud_h, aud_a, liga_actual)
                fric_info = cache_arb[partido_str]

                cuota_ofrecida = float(item.get("Cuota", 0.0) or 0.0)
                if cuota_ofrecida <= 1.0:
                    continue

                prob_str = str(item.get("Prob_Modelo", "0")).replace("%", "").strip()
                p_poisson = float(prob_str) / 100.0 if prob_str else (1.0 / float(item.get("Cuota_Justa", 2.0)))
                p_poisson = float(np.clip(p_poisson, 0.01, 0.98))

                mercado_nom = str(item.get("Mercado", ""))
                tipo_merc = str(item.get("Tipo_Mercado", ""))
                info_libro = libros_api.get(partido_str, {})

                # ── PILAR 2: PROBABILIDAD NO-VIG DEL MERCADO ──────────────────
                p_novig = None
                if tipo_merc == "1X2":
                    c1_m = info_libro.get("c1_avg")
                    cx_m = info_libro.get("cx_avg")
                    c2_m = info_libro.get("c2_avg")
                    if c1_m and cx_m and c2_m:
                        probs_1x2 = calcular_probabilidad_no_vig([c1_m, cx_m, c2_m])
                        if "Local" in mercado_nom or "(1)" in mercado_nom:
                            p_novig = probs_1x2[0]
                        elif "Empate" in mercado_nom or "(X)" in mercado_nom:
                            p_novig = probs_1x2[1]
                        else:
                            p_novig = probs_1x2[2]
                elif tipo_merc == "Goles":
                    par_tot = info_libro.get("totals", {}).get(mercado_nom)
                    if par_tot and len(par_tot) == 2:
                        probs_2w = calcular_probabilidad_no_vig([par_tot[0], par_tot[1], 0.0])
                        p_novig = probs_2w[0]
                elif tipo_merc == "BTTS":
                    par_btts = info_libro.get("btts")
                    if par_btts and len(par_btts) == 2:
                        probs_btts = calcular_probabilidad_no_vig([par_btts[0], par_btts[1], 0.0])
                        p_novig = probs_btts[0] if "Sí" in mercado_nom else probs_btts[1]

                if p_novig is None or p_novig <= 0:
                    c_ref = max(cuota_ofrecida * 0.96, 1.02)
                    c_opuesta = 1.0 / max(1.055 - (1.0 / c_ref), 0.08)
                    probs_fb = calcular_probabilidad_no_vig([c_ref, c_opuesta, 0.0])
                    p_novig = probs_fb[0]

                p_novig = float(np.clip(p_novig, 0.01, 0.98))

                # Fórmula de convergencia institucional: 40% Poisson Dossier + 60% Mercado No-Vig
                p_final = (0.40 * p_poisson) + (0.60 * p_novig)
                p_final = float(np.clip(p_final, 0.01, 0.98))
                ev_real = (cuota_ofrecida * p_final) - 1.0

                # Hard outlier cap (> 30% EV)
                desfase_api = False
                if ev_real > 0.30:
                    ev_real = 0.30
                    desfase_api = True

                # ── PILAR 1: AUDITORÍA CRUZADA SQLITE (ROLLING 6G xG/xGA) ────
                xg_rolling_h = (aud_h["xg_f_6"] + aud_a["xga_c_6"]) / 2.0
                xg_rolling_a = (aud_a["xg_f_6"] + aud_h["xga_c_6"]) / 2.0
                xg_rolling_tot = xg_rolling_h + xg_rolling_a

                xg_impl_tot = info_libro.get("xg_poisson_tot", (aud_h["gf_6"] + aud_a["gf_6"]))
                divergencia_xg = abs(xg_rolling_tot - xg_impl_tot) / max(xg_impl_tot, 0.75)
                penalizacion_xg = divergencia_xg > 0.35

                # ── PILAR 6: MULTIVARIATE CONVICTION SCORE (0 - 100) ─────────
                estabilidad_xg = float(np.clip(100.0 - (divergencia_xg * 115.0) - (aud_h["xg_std"] + aud_a["xg_std"]) * 12.0, 15.0, 100.0))
                gap_consenso = abs(p_poisson - p_novig)
                consenso_cuotas = float(np.clip(100.0 - (gap_consenso * 210.0), 20.0, 100.0))
                profundidad_sqlite = float(np.clip(((aud_h["n_total"] + aud_a["n_total"]) / 36.0) * 100.0, 20.0, 100.0))

                conviction_score = (
                    (estabilidad_xg * 0.35)
                    + (consenso_cuotas * 0.35)
                    + (profundidad_sqlite * 0.30)
                )
                if penalizacion_xg:
                    conviction_score = max(10.0, conviction_score - 22.0)
                if desfase_api:
                    conviction_score = max(10.0, conviction_score - 15.0)
                conviction_score = round(float(np.clip(conviction_score, 0.0, 100.0)), 1)

                if conviction_score >= 80.0:
                    badge_conv = f"🟢 Institucional ({conviction_score:.0f}/100)"
                elif conviction_score >= 60.0:
                    badge_conv = f"🟡 Moderado ({conviction_score:.0f}/100)"
                else:
                    badge_conv = f"🔴 Volátil / Especulativo ({conviction_score:.0f}/100)"

                # ── PILAR 5: CUOTA LÍMITE DE ENTRADA (BREAK-EVEN BOUNDARY) ───
                cuota_min_operable = round(1.03 / p_final, 2) if p_final > 1e-6 else 99.0

                # ── PILAR 7: AUDITORÍA SINTÉTICA ENTRE MERCADOS ──────────────
                etiqueta_auditoria = "✅ Línea Óptima"
                if tipo_merc == "1X2" and ("Local" in mercado_nom or "(1)" in mercado_nom):
                    cuota_ha05 = info_libro.get("ha_minus_05_h") or (cuota_ofrecida * aud_h.get("ha05_ratio", 1.0))
                    if cuota_ha05 and ((cuota_ha05 - cuota_ofrecida) / cuota_ofrecida) > 0.02:
                        etiqueta_auditoria = "💡 Desajuste: Mejor valor en HA -0.5"
                elif tipo_merc == "1X2" and ("Visitante" in mercado_nom or "(2)" in mercado_nom):
                    cuota_ha05_a = info_libro.get("ha_minus_05_a") or (cuota_ofrecida * aud_a.get("ha05_ratio", 1.0))
                    if cuota_ha05_a and ((cuota_ha05_a - cuota_ofrecida) / cuota_ofrecida) > 0.02:
                        etiqueta_auditoria = "💡 Desajuste: Mejor valor en HA -0.5"

                if desfase_api:
                    etiqueta_auditoria = "⚠️ Desfase API"
                elif penalizacion_xg and etiqueta_auditoria == "✅ Línea Óptima":
                    etiqueta_auditoria = f"⚠️ Divergencia xG ({divergencia_xg * 100:.0f}%)"

                # ── BACKTESTING EMPÍRICO EXPRESS EN SQLITE ───────────────────
                bt_info = _backtest_empirico_db(
                    liga_nombre=liga_actual,
                    tipo_mercado=tipo_merc,
                    mercado_nom=mercado_nom,
                    cuota_ofrecida=cuota_ofrecida,
                    eq_local=eq_local,
                    eq_vis=eq_vis,
                )

                # ── PILAR 3: KELLY EFFICIENCY SCORE ──────────────────────────
                b_net = max(cuota_ofrecida - 1.0, 0.01)
                kelly_score = ev_real / b_net
                stake_eur = round(max(0.0, 100.0 * (kelly_score * 0.50)), 2)

                ev_pct = round(ev_real * 100.0, 2)
                nivel_val = (
                    "Alto Valor (EV >10%)"
                    if ev_pct >= 10.0
                    else ("Valor Medio (5-10%)" if ev_pct >= 5.0 else "Valor Moderado (<5%)")
                )

                registros.append({
                    "Fecha / Hora": item.get("Fecha / Hora", ""),
                    "Partido": partido_str,
                    "Equipo_Local": eq_local,
                    "Equipo_Visitante": eq_vis,
                    "Mercado": mercado_nom,
                    "Tipo_Mercado": tipo_merc,
                    "Casa": item.get("Casa", ""),
                    "Cuota": round(cuota_ofrecida, 2),
                    "Cuota Mercado": round(cuota_ofrecida, 2),
                    "Cuota Mínima Operable": cuota_min_operable,
                    "Cuota_Justa": round(1.0 / p_final, 2),
                    "Prob_Poisson": f"{p_poisson * 100:.1f}%",
                    "Prob_NoVig": f"{p_novig * 100:.1f}%",
                    "Prob_Modelo": f"{p_final * 100:.1f}%",
                    "P_Final_Num": p_final,
                    "EV_Num": ev_pct,
                    "EV": f"{ev_pct:+.1f}%" + (" ⚠️" if desfase_api else ""),
                    "Kelly_Score": round(kelly_score, 4),
                    "Stake_Num": stake_eur,
                    "Stake_Recomendado": f"{stake_eur:.2f} €",
                    "Conviccion_Score": conviction_score,
                    "Convicción": badge_conv,
                    "Auditoría Sintética": etiqueta_auditoria,
                    "Nivel_Valor": nivel_val,
                    "xG_6G_Local": round(aud_h["xg_f_6"], 2),
                    "xG_6G_Vis": round(aud_a["xg_f_6"], 2),
                    "Arbitro_Asignado": fric_info["arbitro"],
                    "Arb_Amarillas": fric_info["arb_amarillas"],
                    "Faltas_Comb_6G": fric_info["faltas_combinadas"],
                    "Friccion_Badge": fric_info["badge_tactico"],
                    "Friccion_Alta": fric_info["es_alta"],
                    "Backtest_Badge": bt_info["badge"],
                })
            except Exception:
                continue

        df_res = pd.DataFrame(registros)
        if df_res.empty:
            return df_res, 0

        if tipo_mercado_f == "Mercado Principal (1X2)":
            df_res = df_res[df_res["Tipo_Mercado"] == "1X2"]
        elif tipo_mercado_f == "Líneas de Goles (Over/Under)":
            df_res = df_res[df_res["Tipo_Mercado"] == "Goles"]
        elif tipo_mercado_f == "Ambos Marcan (BTTS)":
            df_res = df_res[df_res["Tipo_Mercado"] == "BTTS"]

        df_res = df_res[
            (df_res["Cuota"] >= cuota_min_f)
            & (df_res["Cuota"] <= cuota_max_f)
            & (df_res["EV_Num"] >= ev_min_f)
        ]

        if equipo_f.strip():
            df_res = df_res[df_res["Partido"].str.contains(equipo_f.strip(), case=False, na=False)]

        if df_res.empty:
            return df_res, 0

        df_res = (
            df_res.sort_values(by=["Kelly_Score", "Cuota"], ascending=[False, False])
            .groupby(["Partido", "Mercado"], as_index=False)
            .first()
        )

        # ── PILAR 3: SUPRESIÓN DE SELECCIONES CORRELACIONADAS POR PARTIDO ────
        n_antes = len(df_res)
        if suprimir_correlacion and not df_res.empty:
            df_res = (
                df_res.sort_values(by=["Kelly_Score", "Conviccion_Score", "EV_Num"], ascending=[False, False, False])
                .groupby("Partido", as_index=False)
                .first()
            )
        n_suprimidas = max(0, n_antes - len(df_res))

        df_res = df_res.sort_values(by=["Kelly_Score", "EV_Num"], ascending=[False, False]).reset_index(drop=True)
        return df_res, n_suprimidas

    # ── PILAR 8: INTERFAZ LIMPIA Y GRID COMPACTO DE 5 COLUMNAS ───────────────
    f_col1, f_col2, f_col3, f_col4, f_col5 = st.columns(5)
    with f_col1:
        liga_api_sel = st.selectbox(
            "🏆 Competición:",
            ["Premier League (Inglaterra)", "LaLiga (España)", "Serie A (Italia)", "Bundesliga (Alemania)"],
            key="ev_pro_liga",
        )
    with f_col2:
        filtro_tipo_mercado = st.selectbox(
            "🎯 Mercado:",
            ["Todos los Mercados", "Mercado Principal (1X2)", "Líneas de Goles (Over/Under)", "Ambos Marcan (BTTS)"],
            key="ev_pro_mercado",
        )
    with f_col3:
        cuota_min = st.number_input("📉 Cuota Mín.:", value=1.18, min_value=1.01, max_value=2.50, step=0.02, key="ev_pro_cmin")
    with f_col4:
        cuota_max = st.number_input("📈 Cuota Máx.:", value=3.50, min_value=1.20, max_value=8.00, step=0.10, key="ev_pro_cmax")
    with f_col5:
        ev_minimo = st.number_input("⚡ EV+ Mín. (%):", value=2.0, min_value=0.5, max_value=30.0, step=0.5, key="ev_pro_evmin")

    sub_c1, sub_c2, sub_c3, sub_c4, sub_c5 = st.columns([1.8, 1.8, 1.3, 1.3, 1.2])
    with sub_c1:
        filtro_equipo = st.text_input("🔍 Filtrar Equipo:", placeholder="Ej: Arsenal, Real Madrid...", key="ev_pro_eq")
    with sub_c2:
        modo_td_api = st.selectbox(
            "⏳ Modelo Poisson Dossier:",
            ["🔥 Forma Reciente (Time-Decay)", "📚 Histórico Completo (Media Clásica)"],
            key="modo_td_api",
        )
        usar_td_api = "Time-Decay" in modo_td_api
        usar_time_decay = usar_td_api
    with sub_c3:
        xi_api = st.slider(
            "📉 Factor ξ Decay:",
            min_value=0.001,
            max_value=0.025,
            value=0.010,
            step=0.001,
            format="%.3f",
            disabled=not usar_td_api,
            key="xi_td_api",
        )
        factor_decay = xi_api
    with sub_c4:
        st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
        suprimir_corr = st.checkbox("Suprimir correlaciones", value=True, help="Mantiene sólo la selección de mayor Kelly Score por partido.")
    with sub_c5:
        st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
        btn_escanear = st.button("🚀 Escanear Mercado", type="primary", use_container_width=True)

    mapa_sports = {
        "LaLiga (España)": "soccer_spain_la_liga",
        "Premier League (Inglaterra)": "soccer_epl",
        "Serie A (Italia)": "soccer_italy_serie_a",
        "Bundesliga (Alemania)": "soccer_germany_bundesliga",
    }

    if btn_escanear:
        if escaneo_apuestas_ev is None:
            st.warning("⚠️ Mercados no disponibles temporalmente o límite de API alcanzado.")
            st.session_state.df_ops_export = None
        else:
            with st.spinner("Ejecutando motor de convergencia No-Vig y auditoría SQLite..."):
                import conector_odds as _co_mod

                libros_api_capturados = {}
                _orig_pedir = getattr(_co_mod, "_pedir_cuotas", None)
                _orig_warn = st.warning

                def _warn_silencioso(msg, *args, **kwargs):
                    txt_w = str(msg)
                    if any(k in txt_w.lower() for k in ("422", "btts", "mercados múltiples", "reintentando sin btts")):
                        return None
                    return _orig_warn(msg, *args, **kwargs)

                def _pedir_cuotas_interceptor(sport_k, markets_k="h2h,totals", api_k=None):
                    res_api = _orig_pedir(sport_k, markets_k, api_k)
                    if getattr(res_api, "status_code", 500) == 422 and "btts" in str(markets_k):
                        res_spreads = _orig_pedir(sport_k, "h2h,totals,spreads", api_k)
                        if getattr(res_spreads, "status_code", 500) == 200:
                            res_api = res_spreads
                        else:
                            res_api = _orig_pedir(sport_k, "h2h,totals", api_k)
                    try:
                        if getattr(res_api, "status_code", 500) == 200:
                            datos_json = res_api.json()
                            for ev_p in datos_json:
                                loc_n = _co_mod.MAPA_EQUIPOS_API.get(ev_p.get("home_team", ""), ev_p.get("home_team", ""))
                                vis_n = _co_mod.MAPA_EQUIPOS_API.get(ev_p.get("away_team", ""), ev_p.get("away_team", ""))
                                clave_p = f"{loc_n} vs {vis_n}"
                                c1_l, cx_l, c2_l = [], [], []
                                ha05_h, ha05_a = [], []
                                tot_map = {}
                                for bm_item in ev_p.get("bookmakers", []):
                                    for mk in bm_item.get("markets", []):
                                        mkey = mk.get("key")
                                        if mkey == "h2h":
                                            for outc in mk.get("outcomes", []):
                                                pr_o = float(outc.get("price", 0) or 0)
                                                if pr_o > 1.0:
                                                    if outc.get("name") == ev_p.get("home_team"):
                                                        c1_l.append(pr_o)
                                                    elif outc.get("name") == "Draw":
                                                        cx_l.append(pr_o)
                                                    else:
                                                        c2_l.append(pr_o)
                                        elif mkey == "spreads":
                                            for outc in mk.get("outcomes", []):
                                                pt_o = outc.get("point")
                                                pr_o = float(outc.get("price", 0) or 0)
                                                if pt_o == -0.5 and pr_o > 1.0:
                                                    if outc.get("name") == ev_p.get("home_team"):
                                                        ha05_h.append(pr_o)
                                                    else:
                                                        ha05_a.append(pr_o)
                                        elif mkey == "totals":
                                            por_punto = {}
                                            for outc in mk.get("outcomes", []):
                                                por_punto[(outc.get("name"), outc.get("point"))] = float(outc.get("price", 0) or 0)
                                            for pt_v, etq_o, etq_u in [
                                                (0.5, "⚽ Más de 0.5 Goles", "Menos de 0.5 Goles"),
                                                (1.5, "⚽ Más de 1.5 Goles", "Menos de 1.5 Goles"),
                                                (2.5, "🔥 Más de 2.5 Goles", "Menos de 2.5 Goles"),
                                                (3.5, "🚀 Más de 3.5 Goles", "Menos de 3.5 Goles"),
                                                (4.5, "🚀 Más de 4.5 Goles", "Menos de 4.5 Goles"),
                                            ]:
                                                co = por_punto.get(("Over", pt_v))
                                                cu = por_punto.get(("Under", pt_v))
                                                if co and cu:
                                                    tot_map[etq_o] = (co, cu)
                                                    tot_map[etq_u] = (cu, co)
                                libros_api_capturados[clave_p] = {
                                    "c1_avg": float(np.mean(c1_l)) if c1_l else None,
                                    "cx_avg": float(np.mean(cx_l)) if cx_l else None,
                                    "c2_avg": float(np.mean(c2_l)) if c2_l else None,
                                    "ha_minus_05_h": float(np.max(ha05_h)) if ha05_h else None,
                                    "ha_minus_05_a": float(np.max(ha05_a)) if ha05_a else None,
                                    "totals": tot_map,
                                }
                    except Exception:
                        pass
                    return res_api

                try:
                    st.warning = _warn_silencioso
                    if _orig_pedir is not None:
                        _co_mod._pedir_cuotas = _pedir_cuotas_interceptor
                    ops_raw = escaneo_apuestas_ev(
                        mapa_sports[liga_api_sel],
                        cuota_minima=1.01,
                        cuota_maxima=10.0,
                        usar_time_decay=usar_time_decay,
                        factor_decay=factor_decay,
                    )
                    if isinstance(ops_raw, pd.DataFrame):
                        ops_raw = ops_raw.to_dict("records")
                    elif not isinstance(ops_raw, list):
                        ops_raw = []
                except Exception:
                    ops_raw = []
                finally:
                    st.warning = _orig_warn
                    if _orig_pedir is not None:
                        _co_mod._pedir_cuotas = _orig_pedir

                df_ops, n_corr_sup = _enriquecer_y_filtrar_institucional(
                    ops_raw=ops_raw,
                    libros_api=libros_api_capturados,
                    cuota_min_f=cuota_min,
                    cuota_max_f=cuota_max,
                    ev_min_f=ev_minimo,
                    tipo_mercado_f=filtro_tipo_mercado,
                    equipo_f=filtro_equipo,
                    liga_actual=liga_api_sel,
                    suprimir_correlacion=suprimir_corr,
                )

                st.session_state.df_ops_export = df_ops
                st.session_state.liga_ops_export = liga_api_sel
                st.session_state.n_correlaciones_suprimidas = n_corr_sup

                if df_ops.empty:
                    if ops_raw:
                        st.info("Ninguna selección superó los filtros de convergencia No-Vig y umbral EV+.")
                    else:
                        st.warning("⚠️ Mercados no disponibles temporalmente o límite de API alcanzado.")

    df_ops = st.session_state.get("df_ops_export")
    liga_export = st.session_state.get("liga_ops_export", liga_api_sel)
    n_corr_sup = st.session_state.get("n_correlaciones_suprimidas", 0)

    if isinstance(df_ops, pd.DataFrame) and not df_ops.empty:
        if "Backtest_Badge" not in df_ops.columns or "Friccion_Badge" not in df_ops.columns:
            df_ops, n_corr_sup = _enriquecer_y_filtrar_institucional(
                ops_raw=df_ops.to_dict("records"),
                libros_api={},
                cuota_min_f=cuota_min,
                cuota_max_f=cuota_max,
                ev_min_f=ev_minimo,
                tipo_mercado_f=filtro_tipo_mercado,
                equipo_f=filtro_equipo,
                liga_actual=liga_export,
                suprimir_correlacion=suprimir_corr,
            )
            st.session_state.df_ops_export = df_ops

    if isinstance(df_ops, pd.DataFrame) and not df_ops.empty:
        ev_medio = float(df_ops["EV_Num"].mean()) if "EV_Num" in df_ops.columns else 0.0
        conv_media = float(df_ops["Conviccion_Score"].mean()) if "Conviccion_Score" in df_ops.columns else 75.0
        kelly_medio = float(df_ops["Kelly_Score"].mean() * 100.0) if "Kelly_Score" in df_ops.columns else 0.0

        st.markdown("---")
        k_c1, k_c2, k_c3, k_c4 = st.columns(4)
        k_c1.metric("Selecciones Dominantes (+EV)", f"{len(df_ops)} picks", f"-{n_corr_sup} correlacionadas suprimidas")
        k_c2.metric("EV+ Convergente Medio", f"+{ev_medio:.2f}%")
        k_c3.metric("Convicción Media (0-100)", f"{conv_media:.1f} pts")
        k_c4.metric("Eficiencia Kelly Media", f"{kelly_medio:.2f}%")

        cols_vista = [
            "Convicción",
            "Fecha / Hora",
            "Partido",
            "Mercado",
            "Casa",
            "Cuota Mercado",
            "Cuota Mínima Operable",
            "Prob_Poisson",
            "Prob_NoVig",
            "Prob_Modelo",
            "EV",
            "Stake_Recomendado",
            "Auditoría Sintética",
        ]
        cols_presentes = [c for c in cols_vista if c in df_ops.columns]
        st.dataframe(df_ops[cols_presentes], use_container_width=True, hide_index=True)

        # ── AUDITORÍA CRUZADA SQLITE, FRICCIÓN ARBITRAL, BACKTEST Y TELETRANSPORTE AL DOSSIER ──
        with st.expander("📋 Auditoría Cruzada SQLite, Fricción Táctica Arbitral, Backtest Empírico y Dossier", expanded=True):
            for idx_r, r_op in df_ops.iterrows():
                loc_e = str(r_op.get("Equipo_Local", ""))
                vis_e = str(r_op.get("Equipo_Visitante", ""))
                fric_badge = str(r_op.get("Friccion_Badge", "🍃 Partido Fluido: Árbitro permisivo + baja tasa de faltas"))
                fric_alta = bool(r_op.get("Friccion_Alta", False))
                bt_badge = str(r_op.get("Backtest_Badge", "⚠️ Muestra histórica insuficiente en DB local"))
                arb_nom = str(r_op.get("Arbitro_Asignado", "Colegiado Media Comité"))
                arb_am = float(r_op.get("Arb_Amarillas", 4.05))
                faltas_c = float(r_op.get("Faltas_Comb_6G", 23.0))

                bg_fric = "rgba(239, 68, 68, 0.16)" if fric_alta else "rgba(16, 185, 129, 0.14)"
                bd_fric = "#EF4444" if fric_alta else "#10B981"
                tx_fric = "#FCA5A5" if fric_alta else "#6EE7B7"

                es_bt_ok = bt_badge.startswith("📈")
                bg_bt = "rgba(56, 189, 248, 0.14)" if es_bt_ok else "rgba(245, 158, 11, 0.14)"
                bd_bt = "#38BDF8" if es_bt_ok else "#F59E0B"
                tx_bt = "#7DD3FC" if es_bt_ok else "#FCD34D"

                with st.container(border=True):
                    c_aud1, c_aud2, c_aud3 = st.columns([3.4, 1.3, 1.3])
                    with c_aud1:
                        st.markdown(
                            f"**{r_op['Partido']}** · `{r_op['Mercado']}` (@{r_op['Cuota Mercado']:.2f} vs Límite @{r_op['Cuota Mínima Operable']:.2f})  \n"
                            f"<span style='color:#94A3B8;font-size:12.5px;'>Rolling 6G xG SQLite: "
                            f"<b>{loc_e}</b> {r_op.get('xG_6G_Local', 1.35):.2f} xG · "
                            f"<b>{vis_e}</b> {r_op.get('xG_6G_Vis', 1.25):.2f} xG · "
                            f"Colegiado ref.: <b>{arb_nom}</b> ({arb_am:.2f} amarillas/p, {faltas_c:.1f} faltas comb.) · "
                            f"{r_op.get('Auditoría Sintética', '')}</span>",
                            unsafe_allow_html=True,
                        )
                        st.markdown(
                            f"<div style='display:flex;flex-wrap:wrap;gap:8px;margin-top:6px;'>"
                            f"<span style='background:{bg_fric};border:1px solid {bd_fric};color:{tx_fric};"
                            f"padding:4px 10px;border-radius:6px;font-size:12.5px;font-weight:600;'>{fric_badge}</span>"
                            f"<span style='background:{bg_bt};border:1px solid {bd_bt};color:{tx_bt};"
                            f"padding:4px 10px;border-radius:6px;font-size:12.5px;font-weight:600;'>{bt_badge}</span>"
                            f"</div>",
                            unsafe_allow_html=True,
                        )
                    def _canonizar_equipo_dossier(eq_nom: str) -> str:
                        try:
                            df_src_d = _obtener_df_partidos_scanner()
                            if not df_src_d.empty:
                                eqs_bd = sorted(
                                    pd.concat([df_src_d["HomeTeam"], df_src_d["AwayTeam"]])
                                    .dropna()
                                    .unique()
                                    .tolist()
                                )
                                if eq_nom in eqs_bd:
                                    return str(eq_nom)
                                res_eq = ev_resolver_equipo(eq_nom, eqs_bd)
                                if res_eq and res_eq in eqs_bd:
                                    return str(res_eq)
                                eq_low = str(eq_nom).strip().lower()
                                for cand in eqs_bd:
                                    if eq_low in str(cand).lower() or str(cand).lower() in eq_low:
                                        return str(cand)
                                return str(eqs_bd[0]) if eqs_bd else str(eq_nom)
                        except Exception:
                            pass
                        return str(eq_nom)

                    def _navegar_a_dossier_equipo(equipo_destino: str) -> None:
                        import streamlit as _st
                        _st.session_state.pop("seccion_activa", None)
                        _st.session_state.pop("equipo_dossier_sel", None)
                        if equipo_destino:
                            _st.session_state["dossier_equipo_selector"] = equipo_destino
                        _st.session_state["menu_nav"] = "📊 Dossier de Equipo"

                    eq_loc_canon = _canonizar_equipo_dossier(loc_e) if loc_e else ""
                    eq_vis_canon = _canonizar_equipo_dossier(vis_e) if vis_e else ""

                    with c_aud2:
                        if loc_e:
                            st.button(
                                f"📋 Ver en Dossier ({loc_e})",
                                key=f"btn_dos_loc_{idx_r}",
                                use_container_width=True,
                                on_click=_navegar_a_dossier_equipo,
                                args=(eq_loc_canon,),
                            )
                    with c_aud3:
                        if vis_e:
                            st.button(
                                f"📋 Ver en Dossier ({vis_e})",
                                key=f"btn_dos_vis_{idx_r}",
                                use_container_width=True,
                                on_click=_navegar_a_dossier_equipo,
                                args=(eq_vis_canon,),
                            )

        # ── PILAR 8: EXPORTACIÓN LIMPIA SÓLO EXCEL (.XLSX) ───────────────────
        st.markdown("#### 📥 Exportar Cartera Institucional")
        nombre_base = f"escaner_EV_{_slug_archivo(liga_export)}_{datetime.now().strftime('%Y%m%d_%H%M')}"
        xlsx_bytes = generar_excel_escaner_ev(df_ops, liga_export)
        st.download_button(
            "📊 Descargar Informe Excel (.xlsx)",
            data=xlsx_bytes,
            file_name=f"{nombre_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="dl_escaner_xlsx",
        )

elif menu_seleccionado == "📊 Dossier de Equipo":
    st.title("📊 Dossier Cuantitativo de Inteligencia de Equipo")
    st.markdown(
        "Radiografía institucional de rendimiento: **Eficiencia xG & Conversión de Tiro**, "
        "**Matriz de Frecuencias de Mercado (Cuotas Implícitas)**, **Ratings Relativos Casa/Fuera** y **Perfil de Fricción Disciplinaria**."
    )

    if err_partidos or df_partidos_global.empty:
        st.warning("No hay datos de partidos disponibles en la base de datos.")
    else:
        def _col_disponible(df_in, candidatos):
            for c in candidatos:
                if c in df_in.columns:
                    return c
            return None

        def _normalizar_partidos_equipo(df_raw, equipo):
            """
            Transforma los registros del equipo en un DataFrame estandarizado desde la perspectiva
            del equipo analizado (propio vs rival), con soporte resiliente de columnas y xG estimado.
            """
            if df_raw is None or df_raw.empty:
                return pd.DataFrame()

            c_hg = _col_disponible(df_raw, ["FTHG", "HG"])
            c_ag = _col_disponible(df_raw, ["FTAG", "AG"])
            if not c_hg or not c_ag:
                return pd.DataFrame()

            c_hthg = _col_disponible(df_raw, ["HTHG"])
            c_htag = _col_disponible(df_raw, ["HTAG"])
            c_hxg = _col_disponible(df_raw, ["HxG"])
            c_axg = _col_disponible(df_raw, ["AxG"])
            c_hs = _col_disponible(df_raw, ["HS"])
            c_as = _col_disponible(df_raw, ["AS"])
            c_hst = _col_disponible(df_raw, ["HST"])
            c_ast = _col_disponible(df_raw, ["AST"])
            c_hf = _col_disponible(df_raw, ["HF"])
            c_af = _col_disponible(df_raw, ["AF"])
            c_hy = _col_disponible(df_raw, ["HY"])
            c_ay = _col_disponible(df_raw, ["AY"])
            c_hr = _col_disponible(df_raw, ["HR"])
            c_ar = _col_disponible(df_raw, ["AR"])

            df_ord = _ordenar_por_fecha(df_raw)
            registros = []
            for _, row in df_ord.iterrows():
                es_local = str(row.get("HomeTeam", "")) == str(equipo)
                hg = pd.to_numeric(row.get(c_hg), errors="coerce")
                ag = pd.to_numeric(row.get(c_ag), errors="coerce")
                if pd.isna(hg) or pd.isna(ag):
                    continue
                hg_f, ag_f = float(hg), float(ag)
                gf = hg_f if es_local else ag_f
                gc = ag_f if es_local else hg_f

                ht_h = pd.to_numeric(row.get(c_hthg), errors="coerce") if c_hthg else np.nan
                ht_a = pd.to_numeric(row.get(c_htag), errors="coerce") if c_htag else np.nan
                if pd.notna(ht_h) and pd.notna(ht_a):
                    gf_1t = float(ht_h) if es_local else float(ht_a)
                    gc_1t = float(ht_a) if es_local else float(ht_h)
                    gf_2t = max(0.0, gf - gf_1t)
                    gc_2t = max(0.0, gc - gc_1t)
                    tiene_ht = 1
                else:
                    gf_1t = gf * 0.44
                    gc_1t = gc * 0.44
                    gf_2t = gf * 0.56
                    gc_2t = gc * 0.56
                    tiene_ht = 0

                hs_v = pd.to_numeric(row.get(c_hs), errors="coerce") if c_hs else np.nan
                as_v = pd.to_numeric(row.get(c_as), errors="coerce") if c_as else np.nan
                hst_v = pd.to_numeric(row.get(c_hst), errors="coerce") if c_hst else np.nan
                ast_v = pd.to_numeric(row.get(c_ast), errors="coerce") if c_ast else np.nan

                tiros_f = float(hs_v if es_local else as_v) if pd.notna(hs_v if es_local else as_v) else np.nan
                tiros_c = float(as_v if es_local else hs_v) if pd.notna(as_v if es_local else hs_v) else np.nan
                sot_f = float(hst_v if es_local else ast_v) if pd.notna(hst_v if es_local else ast_v) else np.nan
                sot_c = float(ast_v if es_local else hst_v) if pd.notna(ast_v if es_local else hst_v) else np.nan

                hxg_v = pd.to_numeric(row.get(c_hxg), errors="coerce") if c_hxg else np.nan
                axg_v = pd.to_numeric(row.get(c_axg), errors="coerce") if c_axg else np.nan
                xg_raw_f = hxg_v if es_local else axg_v
                xg_raw_c = axg_v if es_local else hxg_v

                if pd.notna(xg_raw_f):
                    xg_f = float(xg_raw_f)
                elif pd.notna(sot_f) and pd.notna(tiros_f):
                    xg_f = max(0.10, 0.26 * sot_f + 0.04 * max(0.0, tiros_f - sot_f))
                elif pd.notna(sot_f):
                    xg_f = max(0.10, 0.31 * sot_f)
                else:
                    xg_f = gf

                if pd.notna(xg_raw_c):
                    xg_c = float(xg_raw_c)
                elif pd.notna(sot_c) and pd.notna(tiros_c):
                    xg_c = max(0.10, 0.26 * sot_c + 0.04 * max(0.0, tiros_c - sot_c))
                elif pd.notna(sot_c):
                    xg_c = max(0.10, 0.31 * sot_c)
                else:
                    xg_c = gc

                hf_v = pd.to_numeric(row.get(c_hf), errors="coerce") if c_hf else np.nan
                af_v = pd.to_numeric(row.get(c_af), errors="coerce") if c_af else np.nan
                faltas_com = float(hf_v if es_local else af_v) if pd.notna(hf_v if es_local else af_v) else np.nan
                faltas_rec = float(af_v if es_local else hf_v) if pd.notna(af_v if es_local else hf_v) else np.nan

                hy_v = pd.to_numeric(row.get(c_hy), errors="coerce") if c_hy else np.nan
                ay_v = pd.to_numeric(row.get(c_ay), errors="coerce") if c_ay else np.nan
                hr_v = pd.to_numeric(row.get(c_hr), errors="coerce") if c_hr else np.nan
                ar_v = pd.to_numeric(row.get(c_ar), errors="coerce") if c_ar else np.nan

                amar_prop = float(hy_v if es_local else ay_v) if pd.notna(hy_v if es_local else ay_v) else 0.0
                amar_riv = float(ay_v if es_local else hy_v) if pd.notna(ay_v if es_local else hy_v) else 0.0
                rojas_prop = float(hr_v if es_local else ar_v) if pd.notna(hr_v if es_local else ar_v) else 0.0
                rojas_riv = float(ar_v if es_local else hr_v) if pd.notna(ar_v if es_local else hr_v) else 0.0

                registros.append({
                    "Date": row.get("Date", "N/A"),
                    "HomeTeam": row.get("HomeTeam", ""),
                    "AwayTeam": row.get("AwayTeam", ""),
                    "FTHG": int(hg_f),
                    "FTAG": int(ag_f),
                    "es_local": es_local,
                    "gf": gf,
                    "gc": gc,
                    "goles_tot": gf + gc,
                    "gf_1t": gf_1t,
                    "gc_1t": gc_1t,
                    "gf_2t": gf_2t,
                    "gc_2t": gc_2t,
                    "tiene_ht": tiene_ht,
                    "xg_f": xg_f,
                    "xg_c": xg_c,
                    "tiros_f": tiros_f,
                    "tiros_c": tiros_c,
                    "sot_f": sot_f,
                    "sot_c": sot_c,
                    "faltas_com": faltas_com,
                    "faltas_rec": faltas_rec,
                    "amar_prop": amar_prop,
                    "amar_riv": amar_riv,
                    "rojas_prop": rojas_prop,
                    "rojas_riv": rojas_riv,
                    "win": 1.0 if gf > gc else 0.0,
                    "draw": 1.0 if gf == gc else 0.0,
                    "loss": 1.0 if gf < gc else 0.0,
                    "over15": 1.0 if (gf + gc) > 1.5 else 0.0,
                    "over25": 1.0 if (gf + gc) > 2.5 else 0.0,
                    "over35": 1.0 if (gf + gc) > 3.5 else 0.0,
                    "under25": 1.0 if (gf + gc) < 2.5 else 0.0,
                    "btts_si": 1.0 if (gf > 0 and gc > 0) else 0.0,
                    "btts_no": 1.0 if not (gf > 0 and gc > 0) else 0.0,
                    "clean_sheet": 1.0 if gc == 0 else 0.0,
                    "failed_to_score": 1.0 if gf == 0 else 0.0,
                })
            return pd.DataFrame(registros)

        def _media_ponderada_segura(df_norm, col, usar_decay=False, xi=0.008, default=0.0):
            if df_norm is None or df_norm.empty or col not in df_norm.columns:
                return float(default)
            s = pd.to_numeric(df_norm[col], errors="coerce").dropna()
            if s.empty:
                return float(default)
            if not usar_decay:
                return float(s.mean())
            sub = df_norm.loc[s.index]
            w = calcular_pesos_time_decay(sub, xi=xi)
            if len(w) != len(s) or float(np.sum(w)) <= 0.0:
                return float(s.mean())
            return float(np.average(s.to_numpy(dtype=float), weights=w))

        def _calcular_resumen_bloque(df_norm, usar_decay=False, media_liga_gf=1.35, media_liga_gc=1.35):
            n = len(df_norm) if df_norm is not None else 0
            if n == 0:
                return {
                    "n": 0, "v": 0, "e": 0, "d": 0,
                    "win_rate": 0.0, "draw_rate": 0.0, "loss_rate": 0.0,
                    "gf_tot": 0, "gc_tot": 0,
                    "gf_med": 0.0, "gc_med": 0.0,
                    "xg_f": 0.0, "xg_c": 0.0, "xg_neto": 0.0, "diff_g_xg": 0.0,
                    "tiros_f": 0.0, "tiros_c": 0.0, "sot_f": 0.0, "sot_c": 0.0,
                    "conv_sot": 0.0, "conv_tiros": 0.0,
                    "atk_rating": 1.0, "def_rating": 1.0,
                    "over15": 0.0, "over25": 0.0, "over35": 0.0, "under25": 0.0,
                    "btts_si": 0.0, "btts_no": 0.0, "clean_sheet": 0.0, "failed_to_score": 0.0,
                    "over25_cnt": 0, "btts_cnt": 0, "cs_cnt": 0,
                    "gf_1t_pct": 45.0, "gf_2t_pct": 55.0, "gc_1t_pct": 45.0, "gc_2t_pct": 55.0,
                    "amar_prop": 0.0, "rojas_prop": 0.0, "amar_riv": 0.0, "rojas_riv": 0.0,
                    "faltas_com": 0.0, "faltas_rec": 0.0, "friccion_idx": 0.0,
                }

            v_cnt = int(df_norm["win"].sum())
            e_cnt = int(df_norm["draw"].sum())
            d_cnt = int(df_norm["loss"].sum())
            gf_tot = int(df_norm["gf"].sum())
            gc_tot = int(df_norm["gc"].sum())

            gf_med = _media_ponderada_segura(df_norm, "gf", usar_decay=usar_decay)
            gc_med = _media_ponderada_segura(df_norm, "gc", usar_decay=usar_decay)
            xg_f = _media_ponderada_segura(df_norm, "xg_f", usar_decay=usar_decay, default=gf_med)
            xg_c = _media_ponderada_segura(df_norm, "xg_c", usar_decay=usar_decay, default=gc_med)
            tiros_f = _media_ponderada_segura(df_norm, "tiros_f", usar_decay=usar_decay, default=11.5)
            tiros_c = _media_ponderada_segura(df_norm, "tiros_c", usar_decay=usar_decay, default=11.5)
            sot_f = _media_ponderada_segura(df_norm, "sot_f", usar_decay=usar_decay, default=4.0)
            sot_c = _media_ponderada_segura(df_norm, "sot_c", usar_decay=usar_decay, default=4.0)

            conv_sot = (gf_med / sot_f * 100.0) if sot_f > 0 else 0.0
            conv_tiros = (gf_med / tiros_f * 100.0) if tiros_f > 0 else 0.0

            atk_rating = (gf_med / media_liga_gf) if media_liga_gf > 0 else 1.0
            def_rating = (gc_med / media_liga_gc) if media_liga_gc > 0 else 1.0

            gf_1t_m = _media_ponderada_segura(df_norm, "gf_1t", usar_decay=usar_decay)
            gf_2t_m = _media_ponderada_segura(df_norm, "gf_2t", usar_decay=usar_decay)
            gc_1t_m = _media_ponderada_segura(df_norm, "gc_1t", usar_decay=usar_decay)
            gc_2t_m = _media_ponderada_segura(df_norm, "gc_2t", usar_decay=usar_decay)

            sum_gf_ht = gf_1t_m + gf_2t_m
            sum_gc_ht = gc_1t_m + gc_2t_m
            gf_1t_pct = (gf_1t_m / sum_gf_ht * 100.0) if sum_gf_ht > 0 else 45.0
            gf_2t_pct = 100.0 - gf_1t_pct
            gc_1t_pct = (gc_1t_m / sum_gc_ht * 100.0) if sum_gc_ht > 0 else 45.0
            gc_2t_pct = 100.0 - gc_1t_pct

            amar_prop = _media_ponderada_segura(df_norm, "amar_prop", usar_decay=usar_decay)
            rojas_prop = _media_ponderada_segura(df_norm, "rojas_prop", usar_decay=usar_decay)
            amar_riv = _media_ponderada_segura(df_norm, "amar_riv", usar_decay=usar_decay)
            rojas_riv = _media_ponderada_segura(df_norm, "rojas_riv", usar_decay=usar_decay)
            faltas_com = _media_ponderada_segura(df_norm, "faltas_com", usar_decay=usar_decay, default=11.5)
            faltas_rec = _media_ponderada_segura(df_norm, "faltas_rec", usar_decay=usar_decay, default=11.5)
            friccion_idx = faltas_com + faltas_rec + amar_prop + amar_riv + rojas_prop + rojas_riv

            return {
                "n": n,
                "v": v_cnt,
                "e": e_cnt,
                "d": d_cnt,
                "win_rate": _media_ponderada_segura(df_norm, "win", usar_decay=usar_decay) * 100.0,
                "draw_rate": _media_ponderada_segura(df_norm, "draw", usar_decay=usar_decay) * 100.0,
                "loss_rate": _media_ponderada_segura(df_norm, "loss", usar_decay=usar_decay) * 100.0,
                "gf_tot": gf_tot,
                "gc_tot": gc_tot,
                "gf_med": gf_med,
                "gc_med": gc_med,
                "xg_f": xg_f,
                "xg_c": xg_c,
                "xg_neto": xg_f - xg_c,
                "diff_g_xg": gf_med - xg_f,
                "tiros_f": tiros_f,
                "tiros_c": tiros_c,
                "sot_f": sot_f,
                "sot_c": sot_c,
                "conv_sot": conv_sot,
                "conv_tiros": conv_tiros,
                "atk_rating": atk_rating,
                "def_rating": def_rating,
                "over15": _media_ponderada_segura(df_norm, "over15", usar_decay=usar_decay) * 100.0,
                "over25": _media_ponderada_segura(df_norm, "over25", usar_decay=usar_decay) * 100.0,
                "over35": _media_ponderada_segura(df_norm, "over35", usar_decay=usar_decay) * 100.0,
                "under25": _media_ponderada_segura(df_norm, "under25", usar_decay=usar_decay) * 100.0,
                "btts_si": _media_ponderada_segura(df_norm, "btts_si", usar_decay=usar_decay) * 100.0,
                "btts_no": _media_ponderada_segura(df_norm, "btts_no", usar_decay=usar_decay) * 100.0,
                "clean_sheet": _media_ponderada_segura(df_norm, "clean_sheet", usar_decay=usar_decay) * 100.0,
                "failed_to_score": _media_ponderada_segura(df_norm, "failed_to_score", usar_decay=usar_decay) * 100.0,
                "over25_cnt": int(df_norm["over25"].sum()),
                "btts_cnt": int(df_norm["btts_si"].sum()),
                "cs_cnt": int(df_norm["clean_sheet"].sum()),
                "gf_1t_pct": gf_1t_pct,
                "gf_2t_pct": gf_2t_pct,
                "gc_1t_pct": gc_1t_pct,
                "gc_2t_pct": gc_2t_pct,
                "amar_prop": amar_prop,
                "rojas_prop": rojas_prop,
                "amar_riv": amar_riv,
                "rojas_riv": rojas_riv,
                "faltas_com": faltas_com,
                "faltas_rec": faltas_rec,
                "friccion_idx": friccion_idx,
            }

        # ── LAYER 3.1: TOP EXECUTIVE CONTROL ──────────────────────────────────
        todos_equipos = sorted(
            pd.concat([df_partidos_global["HomeTeam"], df_partidos_global["AwayTeam"]])
            .dropna()
            .unique()
            .tolist()
        )

        ctrl_d1, ctrl_d2 = st.columns([1.4, 2.2])
        with ctrl_d1:
            equipo_elegido = st.selectbox(
                "🛡️ Selecciona un Equipo:",
                todos_equipos,
                key="dossier_equipo_selector",
            )
        with ctrl_d2:
            horizonte_muestra = st.radio(
                "⏳ Horizonte Muestral:",
                [
                    "Últimos 10 Partidos (L10 - Estado Actual)",
                    "Temporada Actual",
                    "Histórico Ponderado (Total con Decaimiento)",
                ],
                horizontal=True,
                key="dossier_horizonte_selector",
            )

        if equipo_elegido:
            df_eq_raw_all = df_partidos_global[
                (df_partidos_global["HomeTeam"] == equipo_elegido)
                | (df_partidos_global["AwayTeam"] == equipo_elegido)
            ].copy()
            df_eq_ord_all = _ordenar_por_fecha(df_eq_raw_all)

            usar_decay_dossier = "Decaimiento" in horizonte_muestra
            if "Últimos 10" in horizonte_muestra:
                df_muestra_raw = df_eq_ord_all.tail(10).copy()
            elif "Temporada Actual" in horizonte_muestra:
                if "Temporada_Origen" in df_eq_ord_all.columns and not df_eq_ord_all["Temporada_Origen"].dropna().empty:
                    temp_actual = str(df_eq_ord_all["Temporada_Origen"].dropna().astype(str).max())
                    df_temp = df_eq_ord_all[df_eq_ord_all["Temporada_Origen"].astype(str) == temp_actual].copy()
                    df_muestra_raw = df_temp if len(df_temp) >= 4 else df_eq_ord_all.tail(38).copy()
                else:
                    df_muestra_raw = df_eq_ord_all.tail(38).copy()
            else:
                df_muestra_raw = df_eq_ord_all.copy()

            # Contexto de liga del equipo para calcular Attack & Defense Ratings
            col_div_d = _columna_liga_partidos(df_partidos_global)
            if col_div_d and not df_eq_ord_all[col_div_d].dropna().empty:
                div_eq = str(df_eq_ord_all[col_div_d].dropna().iloc[-1])
                df_liga_ctx = df_partidos_global[df_partidos_global[col_div_d].astype(str) == div_eq]
                nombre_liga_ctx = MAPA_LIGAS.get(div_eq, div_eq)
            else:
                df_liga_ctx = df_partidos_global
                nombre_liga_ctx = "Global"

            media_liga_home = float(pd.to_numeric(df_liga_ctx["FTHG"], errors="coerce").mean() or 1.45)
            media_liga_away = float(pd.to_numeric(df_liga_ctx["FTAG"], errors="coerce").mean() or 1.15)
            media_liga_global = (media_liga_home + media_liga_away) / 2.0

            df_norm_global = _normalizar_partidos_equipo(df_muestra_raw, equipo_elegido)
            df_norm_home = df_norm_global[df_norm_global["es_local"]].copy() if not df_norm_global.empty else pd.DataFrame()
            df_norm_away = df_norm_global[~df_norm_global["es_local"]].copy() if not df_norm_global.empty else pd.DataFrame()

            stats_global = _calcular_resumen_bloque(
                df_norm_global,
                usar_decay=usar_decay_dossier,
                media_liga_gf=media_liga_global,
                media_liga_gc=media_liga_global,
            )
            stats_home = _calcular_resumen_bloque(
                df_norm_home,
                usar_decay=usar_decay_dossier,
                media_liga_gf=media_liga_home,
                media_liga_gc=media_liga_away,
            )
            stats_away = _calcular_resumen_bloque(
                df_norm_away,
                usar_decay=usar_decay_dossier,
                media_liga_gf=media_liga_away,
                media_liga_gc=media_liga_home,
            )

            # Botón de exportación PDF ejecutiva preservado y alimentado por la muestra activa
            pdf_dossier = generar_pdf_dossier_equipo({
                "equipo": equipo_elegido,
                "liga": nombre_liga_ctx,
                "horizonte": horizonte_muestra,
                "partidos": stats_global["n"],
                "v": stats_global["v"],
                "e": stats_global["e"],
                "d": stats_global["d"],
                "gf": stats_global["gf_tot"],
                "gc": stats_global["gc_tot"],
                "media_gf": stats_global["gf_med"],
                "media_gc": stats_global["gc_med"],
                "p_over25": stats_global["over25"],
                "over_25_count": stats_global["over25_cnt"],
                "p_btts": stats_global["btts_si"],
                "btts_count": stats_global["btts_cnt"],
                "p_cs": stats_global["clean_sheet"],
                "clean_sheets": stats_global["cs_cnt"],
                "n_local": stats_home["n"],
                "n_vis": stats_away["n"],
                "gf_local": stats_home["gf_tot"],
                "gc_local": stats_home["gc_tot"],
                "gf_vis": stats_away["gf_tot"],
                "gc_vis": stats_away["gc_tot"],
                "stats_global": stats_global,
                "stats_home": stats_home,
                "stats_away": stats_away,
                "media_liga_home": media_liga_home,
                "media_liga_away": media_liga_away,
                "media_liga_global": media_liga_global,
            })

            bar_info, bar_pdf = st.columns([3, 1])
            with bar_info:
                st.caption(
                    f"📌 **Competición:** `{nombre_liga_ctx}` · **Horizonte activo:** `{horizonte_muestra}` · "
                    f"**Muestra:** `{stats_global['n']} partidos` ({stats_home['n']} Casa / {stats_away['n']} Fuera)"
                )
            with bar_pdf:
                st.download_button(
                    "📄 Descargar Informe Ejecutivo (PDF)",
                    data=pdf_dossier,
                    file_name=f"informe_ejecutivo_{_slug_archivo(equipo_elegido)}.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True,
                    key="dl_dossier_pdf",
                )

            # ── ROW 1: HIGH-IMPACT QUANTITATIVE KPIs (st.columns(4)) ─────────
            st.markdown("---")
            st.subheader(f"📈 Radiografía Ejecutiva: {equipo_elegido}")

            if stats_global["clean_sheet"] >= 35.0:
                badge_cs = "🛡️ Muro Defensivo"
            elif stats_global["clean_sheet"] >= 22.0:
                badge_cs = "⚖️ Solidez Media"
            else:
                badge_cs = "⚠️ Permisivo Atrás"

            if stats_global["btts_si"] >= 58.0:
                badge_btts = "🔥 Alta Tendencia BTTS"
            elif stats_global["btts_si"] >= 45.0:
                badge_btts = "⚖️ Frecuencia Estándar"
            else:
                badge_btts = "🧊 Baja Frecuencia BTTS"

            if stats_global["diff_g_xg"] >= 0.15:
                regresion_txt = f"{stats_global['diff_g_xg']:+.2f} G-xG (Sobre-rendimiento)"
            elif stats_global["diff_g_xg"] <= -0.15:
                regresion_txt = f"{stats_global['diff_g_xg']:+.2f} G-xG (Regresión +)"
            else:
                regresion_txt = f"{stats_global['diff_g_xg']:+.2f} G-xG (En línea con xG)"

            k_d1, k_d2, k_d3, k_d4 = st.columns(4)
            k_d1.metric(
                "Record & Win Rate",
                f"{stats_global['win_rate']:.1f}% Win",
                f"{stats_global['v']}V - {stats_global['e']}E - {stats_global['d']}D ({stats_global['n']}p)",
                delta_color="off",
            )
            k_d2.metric(
                "Goles / xG Neto por Partido",
                f"{stats_global['gf_med']:.2f} GF | {stats_global['xg_neto']:+.2f} xG Neto",
                regresion_txt,
                delta_color="off",
            )
            k_d3.metric(
                "Clean Sheet Rate (%)",
                f"{stats_global['clean_sheet']:.1f}%",
                f"{badge_cs} ({stats_global['cs_cnt']}/{stats_global['n']})",
                delta_color="off",
            )
            k_d4.metric(
                "BTTS Rate (Ambos Marcan)",
                f"{stats_global['btts_si']:.1f}%",
                f"{badge_btts} ({stats_global['btts_cnt']}/{stats_global['n']})",
                delta_color="off",
            )

            # ── ROW 2: SPLIT COMPARATIVE ANALYSIS: CASA VS FUERA (st.columns(2))
            st.markdown("---")
            st.subheader("🏠 Análisis Comparativo: Condición Local vs ✈️ Condición Visitante")
            col_lv1, col_lv2 = st.columns(2)

            with col_lv1:
                st.markdown(f"#### 🏠 Rendimiento en Casa ({stats_home['n']} partidos)")
                if stats_home["n"] > 0:
                    hc1, hc2, hc3 = st.columns(3)
                    hc1.metric(
                        "Goles F / C (90m)",
                        f"{stats_home['gf_med']:.2f} / {stats_home['gc_med']:.2f}",
                        f"xG: {stats_home['xg_f']:.2f} F vs {stats_home['xg_c']:.2f} C",
                        delta_color="off",
                    )
                    hc2.metric(
                        "Tiros a Puerta (SOT)",
                        f"{stats_home['sot_f']:.2f} / {stats_home['sot_c']:.2f}",
                        f"Conversión SOT: {stats_home['conv_sot']:.1f}%",
                        delta_color="off",
                    )
                    hc3.metric(
                        "Ratings vs Liga (Casa)",
                        f"Atk {stats_home['atk_rating']:.2f}x",
                        f"Def {stats_home['def_rating']:.2f}x (Media Liga: {media_liga_home:.2f})",
                        delta_color="off",
                    )
                    st.dataframe(
                        pd.DataFrame([
                            {
                                "Métrica Local": "Volumen de Tiro (Totales F / C)",
                                "Valor": f"{stats_home['tiros_f']:.1f} realizados vs {stats_home['tiros_c']:.1f} concedidos ({stats_home['conv_tiros']:.1f}% conv. total)",
                            },
                            {
                                "Métrica Local": "Distribución Goles 1T / 2T",
                                "Valor": f"Anotados: {stats_home['gf_1t_pct']:.0f}% (1T) / {stats_home['gf_2t_pct']:.0f}% (2T) · Encajados: {stats_home['gc_1t_pct']:.0f}% (1T) / {stats_home['gc_2t_pct']:.0f}% (2T)",
                            },
                            {
                                "Métrica Local": "Faltas y Tarjetas en Casa",
                                "Valor": f"Faltas: {stats_home['faltas_com']:.1f} Com. / {stats_home['faltas_rec']:.1f} Rec. · Amarillas: {stats_home['amar_prop']:.2f} Prop. / {stats_home['amar_riv']:.2f} Riv.",
                            },
                        ]),
                        use_container_width=True,
                        hide_index=True,
                    )
                else:
                    st.info("Sin registros como local en el horizonte seleccionado.")

            with col_lv2:
                st.markdown(f"#### ✈️ Rendimiento Fuera ({stats_away['n']} partidos)")
                if stats_away["n"] > 0:
                    ac1, ac2, ac3 = st.columns(3)
                    ac1.metric(
                        "Goles F / C (90m)",
                        f"{stats_away['gf_med']:.2f} / {stats_away['gc_med']:.2f}",
                        f"xG: {stats_away['xg_f']:.2f} F vs {stats_away['xg_c']:.2f} C",
                        delta_color="off",
                    )
                    ac2.metric(
                        "Tiros a Puerta (SOT)",
                        f"{stats_away['sot_f']:.2f} / {stats_away['sot_c']:.2f}",
                        f"Conversión SOT: {stats_away['conv_sot']:.1f}%",
                        delta_color="off",
                    )
                    ac3.metric(
                        "Ratings vs Liga (Fuera)",
                        f"Atk {stats_away['atk_rating']:.2f}x",
                        f"Def {stats_away['def_rating']:.2f}x (Media Liga: {media_liga_away:.2f})",
                        delta_color="off",
                    )
                    st.dataframe(
                        pd.DataFrame([
                            {
                                "Métrica Visitante": "Volumen de Tiro (Totales F / C)",
                                "Valor": f"{stats_away['tiros_f']:.1f} realizados vs {stats_away['tiros_c']:.1f} concedidos ({stats_away['conv_tiros']:.1f}% conv. total)",
                            },
                            {
                                "Métrica Visitante": "Distribución Goles 1T / 2T",
                                "Valor": f"Anotados: {stats_away['gf_1t_pct']:.0f}% (1T) / {stats_away['gf_2t_pct']:.0f}% (2T) · Encajados: {stats_away['gc_1t_pct']:.0f}% (1T) / {stats_away['gc_2t_pct']:.0f}% (2T)",
                            },
                            {
                                "Métrica Visitante": "Faltas y Tarjetas Fuera",
                                "Valor": f"Faltas: {stats_away['faltas_com']:.1f} Com. / {stats_away['faltas_rec']:.1f} Rec. · Amarillas: {stats_away['amar_prop']:.2f} Prop. / {stats_away['amar_riv']:.2f} Riv.",
                            },
                        ]),
                        use_container_width=True,
                        hide_index=True,
                    )
                else:
                    st.info("Sin registros como visitante en el horizonte seleccionado.")

            # ── ROW 3: MARKET FREQUENCY MATRIX (TABLE/DATAFRAME) ─────────────
            st.markdown("---")
            st.subheader("🎯 Matriz de Frecuencias de Mercado y Cuotas Implícitas")
            st.caption(
                "Porcentajes históricos de cumplimiento (Hit Rate %) y cuota justa empírica sin margen (1 / p) sobre la muestra activa."
            )

            def _fmt_cuota_imp(pct_val):
                p = float(pct_val) / 100.0
                if p <= 1e-6:
                    return "-"
                return f"{1.0 / p:.2f}"

            definicion_mercados = [
                ("Over 1.5 Goles (+1.5)", "over15"),
                ("Over 2.5 Goles (+2.5)", "over25"),
                ("Over 3.5 Goles (+3.5)", "over35"),
                ("Under 2.5 Goles (-2.5)", "under25"),
                ("Ambos Marcan (BTTS Sí)", "btts_si"),
                ("Ambos Marcan (BTTS No)", "btts_no"),
                ("Portería a Cero (Clean Sheet)", "clean_sheet"),
                ("Sin Anotar (Failed to Score)", "failed_to_score"),
            ]

            filas_mercados = []
            for etiqueta_m, clave_m in definicion_mercados:
                val_gen = stats_global[clave_m]
                val_casa = stats_home[clave_m]
                val_fuera = stats_away[clave_m]
                filas_mercados.append({
                    "Mercado": etiqueta_m,
                    "General (%)": f"{val_gen:.1f}%",
                    "Casa (%)": f"{val_casa:.1f}%",
                    "Fuera (%)": f"{val_fuera:.1f}%",
                    "Cuota Implícita Modelo (1/p)": _fmt_cuota_imp(val_gen),
                })

            st.dataframe(pd.DataFrame(filas_mercados), use_container_width=True, hide_index=True)

            # ── ROW 4: DISCIPLINARY & FRICTION PROFILE (st.columns(3)) ───────
            st.markdown("---")
            st.subheader("🟨 Perfil Disciplinario y Matriz de Fricción")
            disc_c1, disc_c2, disc_c3 = st.columns(3)

            with disc_c1:
                st.metric(
                    "Tarjetas Propias Recibidas",
                    f"{stats_global['amar_prop']:.2f} 🟨 | {stats_global['rojas_prop']:.2f} 🟥",
                    f"Casa: {stats_home['amar_prop']:.2f} 🟨 vs Fuera: {stats_away['amar_prop']:.2f} 🟨",
                    delta_color="off",
                )
            with disc_c2:
                st.metric(
                    "Tarjetas Provocadas al Rival",
                    f"{stats_global['amar_riv']:.2f} 🟨 | {stats_global['rojas_riv']:.2f} 🟥",
                    f"Casa: {stats_home['amar_riv']:.2f} 🟨 vs Fuera: {stats_away['amar_riv']:.2f} 🟨",
                    delta_color="off",
                )
            with disc_c3:
                if stats_global["friccion_idx"] >= 28.0:
                    badge_fric = "🔥 Alta Fricción (Ideal Over Tarjetas)"
                elif stats_global["friccion_idx"] >= 23.0:
                    badge_fric = "⚖️ Fricción Media"
                else:
                    badge_fric = "🧊 Juego Fluido / Baja Fricción"
                st.metric(
                    "Faltas (Com. vs Rec.) & Índice Fricción",
                    f"{stats_global['faltas_com']:.1f} C / {stats_global['faltas_rec']:.1f} R (Idx {stats_global['friccion_idx']:.1f})",
                    badge_fric,
                    delta_color="off",
                )

            # ── FEATURE 1: CRÓNICA TÁCTICA CUANTITATIVA (DOSSIER VIEW) ───────
            st.markdown("---")
            with st.expander("🎙️ Crónica Táctica Cuantitativa", expanded=True):
                equipos_liga_dossier = sorted(
                    pd.concat([df_liga_ctx["HomeTeam"], df_liga_ctx["AwayTeam"]]).dropna().unique().tolist()
                )
                rivales_disp = [e for e in equipos_liga_dossier if e != equipo_elegido] or [
                    e for e in todos_equipos if e != equipo_elegido
                ]
                df_arb_dossier = cargar_datos_arbitros(3)
                arbs_dossier = (
                    df_arb_dossier["Árbitro"].dropna().unique().tolist()
                    if not df_arb_dossier.empty and "Árbitro" in df_arb_dossier.columns
                    else ["Designación Oficial"]
                )

                cr_c1, cr_c2, cr_c3 = st.columns([2, 2, 1])
                with cr_c1:
                    rival_dossier = st.selectbox(
                        "🆚 Rival de Emparejamiento:",
                        rivales_disp,
                        index=0,
                        key="dossier_rival_cronica",
                    )
                with cr_c2:
                    arb_dossier = st.selectbox(
                        "🧑‍⚖️ Colegiado Designado:",
                        arbs_dossier,
                        index=0,
                        key="dossier_arb_cronica",
                    )
                with cr_c3:
                    cuota_merc_dossier = st.number_input(
                        "💰 Cuota Mercado (1X2/O2.5):",
                        min_value=1.05,
                        max_value=15.0,
                        value=round(max(1.10, 100.0 / max(stats_global["over25"], 10.0)), 2),
                        step=0.05,
                        key="dossier_cuota_merc",
                    )

                df_riv_raw = df_partidos_global[
                    (df_partidos_global["HomeTeam"] == rival_dossier)
                    | (df_partidos_global["AwayTeam"] == rival_dossier)
                ].tail(10)
                df_riv_norm = _normalizar_partidos_equipo(df_riv_raw, rival_dossier)
                stats_riv = _calcular_resumen_bloque(
                    df_riv_norm,
                    usar_decay=usar_decay_dossier,
                    media_liga_gf=media_liga_global,
                    media_liga_gc=media_liga_global,
                )

                xg_h_cron = float(max(0.25, (stats_home["xg_f"] + stats_riv["xg_c"]) / 2.0)) if stats_home["n"] > 0 else float(stats_global["xg_f"])
                xg_a_cron = float(max(0.25, (stats_riv["xg_f"] + stats_home["xg_c"]) / 2.0)) if stats_riv["n"] > 0 else float(stats_global["xg_c"])
                cuota_justa_cron = float(round(100.0 / max(stats_global["over25"], 5.0), 2))
                friccion_comb = float(stats_global["faltas_com"] + stats_riv["faltas_com"])

                media_tarj_arb = 4.50
                if not df_arb_dossier.empty and "Árbitro" in df_arb_dossier.columns:
                    fila_arb = df_arb_dossier[df_arb_dossier["Árbitro"] == arb_dossier]
                    if not fila_arb.empty and "Media Amarillas Est." in fila_arb.columns:
                        media_tarj_arb = float(fila_arb["Media Amarillas Est."].iloc[0])

                fallback_payload = {
                    "home": equipo_elegido,
                    "away": rival_dossier,
                    "xg_home": xg_h_cron,
                    "xg_away": xg_a_cron,
                    "cuota_mercado": float(cuota_merc_dossier),
                    "cuota_justa": cuota_justa_cron,
                    "arbitro": str(arb_dossier),
                    "media_tarjetas": media_tarj_arb,
                    "friccion_faltas": friccion_comb,
                }
                sys_prompt_cron = (
                    "Eres un analista táctico y cuantitativo de fútbol profesional. Redacta en español una "
                    "crónica pre-partido estructurada en 3 bloques ([ANÁLISIS DE EFICIENCIA OFENSIVA Y MÉTRICAS xG], "
                    "[PERFIL DISCIPLINAR Y FRICCIÓN ARBITRAL], [VEREDICTO CUANTITATIVO Y CONCLUSIÓN DE MERCADO]) "
                    "usando estrictamente las métricas proporcionadas."
                )
                usr_prompt_cron = json.dumps(fallback_payload, ensure_ascii=False)
                match_id_cron = f"{equipo_elegido}_vs_{rival_dossier}_{arb_dossier}_{cuota_merc_dossier:.2f}_{horizonte_muestra}"

                texto_cronica, status_badge = obtener_cronica_partido(
                    match_id=match_id_cron,
                    prompt_system=sys_prompt_cron,
                    prompt_user=usr_prompt_cron,
                    fallback_data=fallback_payload,
                )
                st.caption(status_badge)
                st.markdown(texto_cronica)

            # ── SECCIÓN COMPLEMENTARIA: HISTORIAL RECIENTE Y PLANTILLA ───────
            st.markdown("---")
            col_d1, col_d2 = st.columns(2)

            with col_d1:
                st.markdown("### 🕒 Historial de Últimos Resultados")
                if not df_norm_global.empty:
                    for _, row in df_norm_global.iloc[::-1].head(6).iterrows():
                        l_t, v_t = row["HomeTeam"], row["AwayTeam"]
                        gl_t, gv_t = row["FTHG"], row["FTAG"]
                        l_fmt = f"**{l_t}**" if l_t == equipo_elegido else l_t
                        v_fmt = f"**{v_t}**" if v_t == equipo_elegido else v_t
                        st.caption(
                            f"📅 {row.get('Date', 'N/A')} | {l_fmt} {gl_t} - {gv_t} {v_fmt} "
                            f"· xG: `{row['xg_f']:.2f} - {row['xg_c']:.2f}`"
                        )
                else:
                    st.info("Sin partidos recientes registrados.")

            with col_d2:
                st.markdown("### ⭐ Plantilla y Estadística Individual")
                if not err_jugadores and not df_jugadores.empty:
                    if "Equipo" in df_jugadores.columns:
                        df_jug_eq = df_jugadores[
                            df_jugadores["Equipo"].str.contains(equipo_elegido, case=False, na=False)
                        ].copy()
                    else:
                        df_jug_eq = pd.DataFrame()

                    if "Gls" in df_jug_eq.columns and "Goles" not in df_jug_eq.columns:
                        df_jug_eq["Goles"] = df_jug_eq["Gls"]
                    if "Ast" in df_jug_eq.columns and "Asistencias" not in df_jug_eq.columns:
                        df_jug_eq["Asistencias"] = df_jug_eq["Ast"]

                    cols_mostrar = ["Jugador", "Pos", "Goles", "Asistencias"]
                    tiene_columnas = all(c in df_jug_eq.columns for c in cols_mostrar)

                    if not df_jug_eq.empty and tiene_columnas:
                        st.dataframe(df_jug_eq[cols_mostrar], use_container_width=True, hide_index=True)
                    else:
                        st.info("Las estadísticas de jugadores no están disponibles en este momento.")
                else:
                    st.info("Base de datos de jugadores no disponible.")

elif menu_seleccionado == "⚔️ Comparador H2H (Cara a Cara)":
    st.title("⚔️ Comparador Directo H2H Pro & Cuotas Justas")
    st.markdown(
        "Analítica institucional Cara a Cara: **Proyección Poisson con Suavizado Bayesiano**, "
        "**Arbitraje de Cuotas (EV+ & Quarter-Kelly)**, **Estimación de Línea de Hándicap Asiático** y **Auditoría Histórica SQLite**."
    )

    if err_partidos or df_partidos_global.empty:
        st.warning("Datos de partidos no disponibles.")
    else:
        import plotly.graph_objects as go

        # ── 1. SELECTOR INTELIGENTE DE COMPETICIÓN, EQUIPOS E INVERSIÓN 1-CLICK ──
        divs_disp_h2h = (
            sorted(df_partidos_global["Div"].dropna().astype(str).unique().tolist())
            if "Div" in df_partidos_global.columns
            else []
        )
        ligas_disp_h2h = [MAPA_LIGAS.get(d, d) for d in divs_disp_h2h] if divs_disp_h2h else ["Todas las Competiciones"]

        sel_c1, sel_c2, sel_c3, sel_c4 = st.columns([1.5, 1.5, 1.5, 1.1])
        with sel_c1:
            liga_h2h_nom = st.selectbox("🏆 Competición / Liga:", ligas_disp_h2h, key="h2h_liga_sel")

        if divs_disp_h2h and liga_h2h_nom in MAPA_INV:
            div_h2h_code = MAPA_INV[liga_h2h_nom]
            df_liga_h2h = df_partidos_global[df_partidos_global["Div"].astype(str) == str(div_h2h_code)].copy()
        elif divs_disp_h2h:
            div_h2h_code = divs_disp_h2h[ligas_disp_h2h.index(liga_h2h_nom)]
            df_liga_h2h = df_partidos_global[df_partidos_global["Div"].astype(str) == str(div_h2h_code)].copy()
        else:
            df_liga_h2h = df_partidos_global.copy()

        equipos_h2h = sorted(
            pd.concat([df_liga_h2h["HomeTeam"], df_liga_h2h["AwayTeam"]])
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        if not equipos_h2h:
            equipos_h2h = sorted(
                pd.concat([df_partidos_global["HomeTeam"], df_partidos_global["AwayTeam"]])
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

        if "h2h_eq_local" not in st.session_state or st.session_state["h2h_eq_local"] not in equipos_h2h:
            st.session_state["h2h_eq_local"] = equipos_h2h[0] if equipos_h2h else ""
        if "h2h_eq_vis" not in st.session_state or st.session_state["h2h_eq_vis"] not in equipos_h2h:
            st.session_state["h2h_eq_vis"] = equipos_h2h[1] if len(equipos_h2h) > 1 else (equipos_h2h[0] if equipos_h2h else "")

        def _invertir_equipos_h2h() -> None:
            loc_actual = st.session_state.get("h2h_eq_local")
            vis_actual = st.session_state.get("h2h_eq_vis")
            if loc_actual in equipos_h2h and vis_actual in equipos_h2h:
                st.session_state["h2h_eq_local"] = vis_actual
                st.session_state["h2h_eq_vis"] = loc_actual

        with sel_c2:
            eq1 = st.selectbox("🏠 Equipo Local (A):", equipos_h2h, key="h2h_eq_local")
        with sel_c3:
            eq2 = st.selectbox("✈️ Equipo Visitante (B):", equipos_h2h, key="h2h_eq_vis")
        with sel_c4:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            st.button(
                "🔄 Invertir Local / Visitante",
                on_click=_invertir_equipos_h2h,
                use_container_width=True,
                key="btn_swap_h2h",
            )

        if eq1 and eq2 and eq1 != eq2:
            df_eq1_home = df_partidos_global[df_partidos_global["HomeTeam"] == eq1].copy()
            df_eq2_away = df_partidos_global[df_partidos_global["AwayTeam"] == eq2].copy()
            df_eq1_all = df_partidos_global[
                (df_partidos_global["HomeTeam"] == eq1) | (df_partidos_global["AwayTeam"] == eq1)
            ].copy()
            df_eq2_all = df_partidos_global[
                (df_partidos_global["HomeTeam"] == eq2) | (df_partidos_global["AwayTeam"] == eq2)
            ].copy()

            for _df_tmp in (df_eq1_home, df_eq2_away, df_eq1_all, df_eq2_all):
                for _c_num in ("FTHG", "FTAG", "HTHG", "HTAG", "HS", "AS", "HST", "AST", "HC", "AC", "HF", "AF", "HxG", "AxG"):
                    if _c_num in _df_tmp.columns:
                        _df_tmp[_c_num] = pd.to_numeric(_df_tmp[_c_num], errors="coerce")

            # ── 2. RESILIENCIA MUESTRAL Y SUAVIZADO BAYESIANO (SHRINKAGE 70/30) ──
            n_home_sample = int(df_eq1_home["FTHG"].notna().sum()) if "FTHG" in df_eq1_home.columns else len(df_eq1_home)
            n_away_sample = int(df_eq2_away["FTAG"].notna().sum()) if "FTAG" in df_eq2_away.columns else len(df_eq2_away)
            muestra_reducida = (n_home_sample < 6) or (n_away_sample < 6)

            def _media_global_equipo(df_all_eq: pd.DataFrame, eq_nombre: str, col_home: str, col_away: str, default_val: float) -> float:
                if df_all_eq is None or df_all_eq.empty or col_home not in df_all_eq.columns or col_away not in df_all_eq.columns:
                    return float(default_val)
                es_loc_m = df_all_eq["HomeTeam"].astype(str) == str(eq_nombre)
                vals = np.where(es_loc_m, df_all_eq[col_home], df_all_eq[col_away])
                s_vals = pd.Series(vals).dropna()
                return float(s_vals.mean()) if not s_vals.empty else float(default_val)

            gf_l_raw = float(df_eq1_home["FTHG"].mean()) if not df_eq1_home.empty and df_eq1_home["FTHG"].notna().any() else 1.45
            gc_l_raw = float(df_eq1_home["FTAG"].mean()) if not df_eq1_home.empty and df_eq1_home["FTAG"].notna().any() else 1.05
            gf_v_raw = float(df_eq2_away["FTAG"].mean()) if not df_eq2_away.empty and df_eq2_away["FTAG"].notna().any() else 1.10
            gc_v_raw = float(df_eq2_away["FTHG"].mean()) if not df_eq2_away.empty and df_eq2_away["FTHG"].notna().any() else 1.40

            gf_l_glob = _media_global_equipo(df_eq1_all, eq1, "FTHG", "FTAG", gf_l_raw)
            gc_l_glob = _media_global_equipo(df_eq1_all, eq1, "FTAG", "FTHG", gc_l_raw)
            gf_v_glob = _media_global_equipo(df_eq2_all, eq2, "FTHG", "FTAG", gf_v_raw)
            gc_v_glob = _media_global_equipo(df_eq2_all, eq2, "FTAG", "FTHG", gc_v_raw)

            if n_home_sample < 6:
                gf_l = (0.70 * gf_l_raw) + (0.30 * gf_l_glob)
                gc_l = (0.70 * gc_l_raw) + (0.30 * gc_l_glob)
            else:
                gf_l = gf_l_raw
                gc_l = gc_l_raw

            if n_away_sample < 6:
                gf_v = (0.70 * gf_v_raw) + (0.30 * gf_v_glob)
                gc_v = (0.70 * gc_v_raw) + (0.30 * gc_v_glob)
            else:
                gf_v = gf_v_raw
                gc_v = gc_v_raw

            xg_l = max(0.15, (gf_l + gc_v) / 2.0)
            xg_v = max(0.15, (gf_v + gc_l) / 2.0)

            prob_h = prob_d = prob_a = prob_o15 = prob_o25 = prob_btts = tot_prob = 0.0
            marcadores_prob = []
            matriz_7x7 = {}
            for i in range(7):
                for j in range(7):
                    pr = calcular_probabilidad_poisson(xg_l, i) * calcular_probabilidad_poisson(xg_v, j)
                    tot_prob += pr
                    matriz_7x7[(i, j)] = pr
                    if i > j:
                        prob_h += pr
                    elif i == j:
                        prob_d += pr
                    else:
                        prob_a += pr
                    if (i + j) > 1.5:
                        prob_o15 += pr
                    if (i + j) > 2.5:
                        prob_o25 += pr
                    if i > 0 and j > 0:
                        prob_btts += pr
                    marcadores_prob.append((f"{i} - {j}", pr))

            norm = tot_prob if tot_prob > 0 else 1.0
            p_1 = max(prob_h / norm, 1e-4)
            p_x = max(prob_d / norm, 1e-4)
            p_2 = max(prob_a / norm, 1e-4)
            p_o25 = max(prob_o25 / norm, 1e-4)
            p_u25 = max(1.0 - p_o25, 1e-4)
            p_btts_si = max(prob_btts / norm, 1e-4)

            marcadores_ordenados = sorted(
                [(m_str, p_val / norm) for m_str, p_val in marcadores_prob],
                key=lambda x: x[1],
                reverse=True,
            )[:3]

            st.markdown("---")
            st.subheader("🎯 Proyección Matemática del Partido (Modelo Poisson Condicionado)")
            st.markdown(
                f"**Goles Esperados (xG Proyectado): {eq1} {xg_l:.2f} — {eq2} {xg_v:.2f}**  \n"
                f"<span style='color:#94A3B8;font-size:13px;'>Muestra analizada: {n_home_sample} partidos de {eq1} en casa vs "
                f"{n_away_sample} partidos de {eq2} como visitante.</span>",
                unsafe_allow_html=True,
            )

            if muestra_reducida:
                st.warning("⚠️ Muestra reducida en condición local/visitante (< 6 partidos). Aplicado suavizado con promedio global.")

            k_p1, k_p2, k_p3, k_p4, k_p5, k_p6 = st.columns(6)
            k_p1.metric(f"🏠 Local ({eq1})", f"Cuota: {1.0 / p_1:.2f}", f"Prob: {p_1 * 100:.1f}%")
            k_p2.metric("🤝 Empate (X)", f"Cuota: {1.0 / p_x:.2f}", f"Prob: {p_x * 100:.1f}%", delta_color="off")
            k_p3.metric(f"✈️ Visitante ({eq2})", f"Cuota: {1.0 / p_2:.2f}", f"Prob: {p_2 * 100:.1f}%")
            k_p4.metric("🔥 Over 2.5 Goles", f"Cuota: {1.0 / p_o25:.2f}", f"Prob: {p_o25 * 100:.1f}%", delta_color="off")
            k_p5.metric("🧊 Under 2.5 Goles", f"Cuota: {1.0 / p_u25:.2f}", f"Prob: {p_u25 * 100:.1f}%", delta_color="off")
            k_p6.metric("🎯 Ambos Marcan (Sí)", f"Cuota: {1.0 / p_btts_si:.2f}", f"Prob: {p_btts_si * 100:.1f}%", delta_color="off")

            with st.container(border=True):
                st.markdown("#### 🎯 Top 3 Marcadores Más Probables (Distribución Exacta de Poisson)")
                sc_c1, sc_c2, sc_c3 = st.columns(3)
                medallas_sc = ["🥇 #1 Más Probable", "🥈 #2 Más Probable", "🥉 #3 Más Probable"]
                colores_sc = ["#10B981", "#38BDF8", "#F59E0B"]
                for idx_sc, (col_sc, (marc_txt, prob_sc)) in enumerate(zip((sc_c1, sc_c2, sc_c3), marcadores_ordenados)):
                    cuota_sc = 1.0 / max(prob_sc, 1e-4)
                    borde_c = colores_sc[idx_sc]
                    with col_sc:
                        st.markdown(
                            f"<div style='background-color:#16192b;border:1.5px solid {borde_c};border-radius:10px;padding:14px;text-align:center;'>"
                            f"<div style='color:#94A3B8;font-size:12.5px;font-weight:600;margin-bottom:4px;'>{medallas_sc[idx_sc]}</div>"
                            f"<div style='font-size:26px;font-weight:800;color:#F8FAFC;letter-spacing:1px;'>{eq1} {marc_txt} {eq2}</div>"
                            f"<div style='margin-top:6px;font-size:14px;color:{borde_c};font-weight:700;'>"
                            f"Probabilidad: {prob_sc * 100:.1f}% &nbsp;·&nbsp; Cuota Justa: {cuota_sc:.2f}"
                            f"</div></div>",
                            unsafe_allow_html=True,
                        )

            # ── 3. ARBITRAJE INTERACTIVO DE CUOTAS DE MERCADO VS VALOR ESPERADO (EV+ & KELLY) ──
            with st.expander("💰 Comparativa de Cuota de Mercado vs Valor Esperado (EV+ & Kelly)", expanded=True):
                st.caption(
                    "Introduce las cuotas actuales de tu casa de apuestas para calcular en tiempo real el **Valor Esperado (EV%)** "
                    "y el **Stake Recomendado (Quarter-Kelly 25%)** frente al modelo Poisson."
                )
                in_c1, in_cx, in_c2, in_co25 = st.columns(4)
                with in_c1:
                    cuota_mkt_1 = st.number_input(
                        f"Cuota Mercado 1 ({eq1}):",
                        min_value=1.01,
                        max_value=50.0,
                        value=round(1.0 / p_1, 2),
                        step=0.05,
                        key="h2h_mkt_1",
                    )
                with in_cx:
                    cuota_mkt_x = st.number_input(
                        "Cuota Mercado X (Empate):",
                        min_value=1.01,
                        max_value=50.0,
                        value=round(1.0 / p_x, 2),
                        step=0.05,
                        key="h2h_mkt_x",
                    )
                with in_c2:
                    cuota_mkt_2 = st.number_input(
                        f"Cuota Mercado 2 ({eq2}):",
                        min_value=1.01,
                        max_value=50.0,
                        value=round(1.0 / p_2, 2),
                        step=0.05,
                        key="h2h_mkt_2",
                    )
                with in_co25:
                    cuota_mkt_o25 = st.number_input(
                        "Cuota Over 2.5 Goles:",
                        min_value=1.01,
                        max_value=50.0,
                        value=round(1.0 / p_o25, 2),
                        step=0.05,
                        key="h2h_mkt_o25",
                    )

                ev_cols = st.columns(4)
                mercados_arbitraje = [
                    (f"🏠 Victoria {eq1} (1)", p_1, cuota_mkt_1),
                    ("🤝 Empate (X)", p_x, cuota_mkt_x),
                    (f"✈️ Victoria {eq2} (2)", p_2, cuota_mkt_2),
                    ("🔥 Over 2.5 Goles", p_o25, cuota_mkt_o25),
                ]
                for col_ev_ui, (nom_m_arb, prob_m_arb, cuota_m_arb) in zip(ev_cols, mercados_arbitraje):
                    ev_val_pct = ((prob_m_arb * cuota_m_arb) - 1.0) * 100.0
                    b_neto = max(cuota_m_arb - 1.0, 0.01)
                    kelly_full = ((prob_m_arb * b_neto) - (1.0 - prob_m_arb)) / b_neto
                    quarter_kelly_pct = max(0.0, kelly_full * 0.25 * 100.0)
                    cuota_justa_arb = 1.0 / max(prob_m_arb, 1e-4)

                    if ev_val_pct > 2.0:
                        bg_ev = "rgba(16, 185, 129, 0.16)"
                        bd_ev = "#10B981"
                        tx_ev = "#6EE7B7"
                        etiqueta_ev = "🟢 VALOR CONFIRMADO (+EV)"
                    elif ev_val_pct >= 0.0:
                        bg_ev = "rgba(245, 158, 11, 0.14)"
                        bd_ev = "#F59E0B"
                        tx_ev = "#FCD34D"
                        etiqueta_ev = "🟡 MARGEN NEUTRO"
                    else:
                        bg_ev = "rgba(239, 68, 68, 0.14)"
                        bd_ev = "#EF4444"
                        tx_ev = "#FCA5A5"
                        etiqueta_ev = "🔴 SIN VALOR (EV NEGATIVO)"

                    with col_ev_ui:
                        st.markdown(
                            f"<div style='background:{bg_ev};border:1.5px solid {bd_ev};border-radius:10px;padding:12px;'>"
                            f"<div style='font-size:13px;font-weight:700;color:#F8FAFC;'>{nom_m_arb}</div>"
                            f"<div style='font-size:11.5px;color:#94A3B8;margin-top:2px;'>"
                            f"Cuota Justa: <b>{cuota_justa_arb:.2f}</b> vs Mercado: <b>{cuota_m_arb:.2f}</b></div>"
                            f"<div style='font-size:20px;font-weight:800;color:{tx_ev};margin-top:6px;'>"
                            f"EV: {ev_val_pct:+.2f}%</div>"
                            f"<div style='font-size:12.5px;color:#E2E8F0;margin-top:4px;'>"
                            f"💼 Quarter-Kelly: <b>{quarter_kelly_pct:.2f}% bank</b></div>"
                            f"<div style='font-size:11px;font-weight:700;color:{tx_ev};margin-top:4px;'>{etiqueta_ev}</div>"
                            f"</div>",
                            unsafe_allow_html=True,
                        )

            # ── 4. COMPARATIVA TÁCTICA Y ESTIMACIÓN DE LÍNEA DE HÁNDICAP ASIÁTICO JUSTO ──
            st.markdown("---")
            st.subheader("📊 Comparativa Táctica y Estimación de Línea de Hándicap Asiático Justo")

            if "HxG" in df_eq1_home.columns and df_eq1_home["HxG"].notna().any():
                xga_l = float(df_eq1_home["AxG"].dropna().mean()) if "AxG" in df_eq1_home.columns and df_eq1_home["AxG"].notna().any() else gc_l
            else:
                xga_l = gc_l

            if "AxG" in df_eq2_away.columns and df_eq2_away["AxG"].notna().any():
                xga_v = float(df_eq2_away["HxG"].dropna().mean()) if "HxG" in df_eq2_away.columns and df_eq2_away["HxG"].notna().any() else gc_v
            else:
                xga_v = gc_v

            cor_l = (
                float((df_eq1_home["HC"].fillna(0) + df_eq1_home["AC"].fillna(0)).mean())
                if "HC" in df_eq1_home.columns and "AC" in df_eq1_home.columns and not df_eq1_home.empty
                else 9.5
            )
            cor_v = (
                float((df_eq2_away["HC"].fillna(0) + df_eq2_away["AC"].fillna(0)).mean())
                if "HC" in df_eq2_away.columns and "AC" in df_eq2_away.columns and not df_eq2_away.empty
                else 9.2
            )

            faltas_l = (
                float(df_eq1_home["HF"].dropna().mean())
                if "HF" in df_eq1_home.columns and df_eq1_home["HF"].notna().any()
                else 11.5
            )
            faltas_v = (
                float(df_eq2_away["AF"].dropna().mean())
                if "AF" in df_eq2_away.columns and df_eq2_away["AF"].notna().any()
                else 11.8
            )

            sot_l = (
                float(df_eq1_home["HST"].dropna().mean())
                if "HST" in df_eq1_home.columns and df_eq1_home["HST"].notna().any()
                else 4.5
            )
            sot_v = (
                float(df_eq2_away["AST"].dropna().mean())
                if "AST" in df_eq2_away.columns and df_eq2_away["AST"].notna().any()
                else 3.8
            )

            categorias_tacticas = [
                "Eficiencia Ofensiva (Goles p/p)",
                "Fragilidad Defensiva (xGA/GC p/p)",
                "Tiros a Puerta a Favor p/p",
                "Córners Totales (Media p/p)",
                "Intensidad / Faltas Cometidas p/p",
            ]
            raw_loc = [gf_l, xga_l, sot_l, cor_l, faltas_l]
            raw_vis = [gf_v, xga_v, sot_v, cor_v, faltas_v]
            topes_escala = [3.0, 2.5, 8.5, 14.0, 18.0]

            norm_loc = [float(np.clip((v / m) * 100.0, 8.0, 100.0)) for v, m in zip(raw_loc, topes_escala)]
            norm_vis = [float(np.clip((v / m) * 100.0, 8.0, 100.0)) for v, m in zip(raw_vis, topes_escala)]

            # Cálculo de Supremacía y Línea de Hándicap Asiático Justo
            supremacia_xg = xg_l - xg_v
            linea_ah_justa = -round(supremacia_xg * 4.0) / 4.0
            if abs(linea_ah_justa) < 0.05:
                linea_ah_justa = 0.0
                linea_ah_str = f"PK / HA 0.0 (DNB)"
                linea_ah_vis_str = f"PK / HA 0.0 (DNB)"
            else:
                linea_ah_str = f"HA {linea_ah_justa:+.2f}"
                linea_ah_vis_str = f"HA {-linea_ah_justa:+.2f}"

            # Cuotas justas en líneas clave de Hándicap Asiático derivadas de Poisson
            cuota_dnb_loc = 1.0 / max(p_1 / max(p_1 + p_2, 1e-4), 1e-4)
            cuota_dnb_vis = 1.0 / max(p_2 / max(p_1 + p_2, 1e-4), 1e-4)
            cuota_ha_m05_loc = 1.0 / p_1
            cuota_ha_p05_vis = 1.0 / max(p_x + p_2, 1e-4)
            cuota_ha_p05_loc = 1.0 / max(p_1 + p_x, 1e-4)
            cuota_ha_m05_vis = 1.0 / p_2

            graf_c1, graf_c2 = st.columns([1.2, 1.0])
            with graf_c1:
                fig_bars = go.Figure()
                fig_bars.add_trace(
                    go.Bar(
                        y=categorias_tacticas,
                        x=norm_loc,
                        name=f"🏠 {eq1} (Casa)",
                        orientation="h",
                        marker=dict(color="#10B981"),
                        text=[f"{v:.2f}" for v in raw_loc],
                        textposition="auto",
                        hovertemplate="<b>%{y}</b><br>Media real: %{text}<extra></extra>",
                    )
                )
                fig_bars.add_trace(
                    go.Bar(
                        y=categorias_tacticas,
                        x=norm_vis,
                        name=f"✈️ {eq2} (Fuera)",
                        orientation="h",
                        marker=dict(color="#38BDF8"),
                        text=[f"{v:.2f}" for v in raw_vis],
                        textposition="auto",
                        hovertemplate="<b>%{y}</b><br>Media real: %{text}<extra></extra>",
                    )
                )
                fig_bars.update_layout(
                    barmode="group",
                    paper_bgcolor="#0e1117",
                    plot_bgcolor="#16192b",
                    font=dict(color="#f0f2f6"),
                    xaxis=dict(title="Índice Comparativo (0-100) · Etiqueta = Media Exacta p/p", range=[0, 110], gridcolor="#2d3748"),
                    yaxis=dict(autorange="reversed", gridcolor="#2d3748"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
                    margin=dict(l=10, r=20, t=40, b=30),
                    height=370,
                )
                st.plotly_chart(fig_bars, use_container_width=True)

            with graf_c2:
                with st.container(border=True):
                    st.markdown("#### ⚡ Estimación de Línea y Hándicap Asiático Justo")
                    sesgo_dom = (
                        f"Favoritismo Local ({eq1})"
                        if supremacia_xg > 0.12
                        else (f"Favoritismo Visitante ({eq2})" if supremacia_xg < -0.12 else "Duelo Equilibrado (Pick'em)")
                    )
                    st.markdown(
                        f"<div style='background-color:#16192b;border:1.5px solid #38BDF8;border-radius:10px;padding:14px;margin-bottom:10px;'>"
                        f"<div style='color:#94A3B8;font-size:12.5px;font-weight:600;'>Supremacía de Goles Esperados (xG Local − xG Visitante)</div>"
                        f"<div style='font-size:24px;font-weight:800;color:#F8FAFC;margin-top:4px;'>"
                        f"{supremacia_xg:+.2f} xG &nbsp;·&nbsp; <span style='color:#38BDF8;'>{sesgo_dom}</span></div>"
                        f"<div style='margin-top:8px;font-size:15px;color:#10B981;font-weight:700;'>"
                        f"🎯 Línea Equilibrada Proyectada: {eq1} {linea_ah_str} / {eq2} {linea_ah_vis_str}</div>"
                        f"</div>",
                        unsafe_allow_html=True,
                    )
                    df_ah_justo = pd.DataFrame([
                        {
                            "Línea Asiática": "HA 0.0 (Draw No Bet)",
                            f"🏠 {eq1}": f"Cuota: {cuota_dnb_loc:.2f}",
                            f"✈️ {eq2}": f"Cuota: {cuota_dnb_vis:.2f}",
                        },
                        {
                            "Línea Asiática": f"HA -0.5 ({eq1}) / +0.5 ({eq2})",
                            f"🏠 {eq1}": f"Cuota: {cuota_ha_m05_loc:.2f}",
                            f"✈️ {eq2}": f"Cuota: {cuota_ha_p05_vis:.2f}",
                        },
                        {
                            "Línea Asiática": f"HA +0.5 ({eq1}) / -0.5 ({eq2})",
                            f"🏠 {eq1}": f"Cuota: {cuota_ha_p05_loc:.2f}",
                            f"✈️ {eq2}": f"Cuota: {cuota_ha_m05_vis:.2f}",
                        },
                    ])
                    st.dataframe(df_ah_justo, use_container_width=True, hide_index=True)

            # ── 5. TABLA DE FRECUENCIAS DE MERCADO (CONDICIONADA CASA VS FUERA) ──
            st.markdown("---")
            st.subheader("📈 Matriz de Frecuencias Empíricas de Mercado (Condicionada Casa vs Fuera)")

            def _frecuencia_mercado_condicionada(df_sub: pd.DataFrame, tipo_m: str) -> float:
                if df_sub is None or df_sub.empty or "FTHG" not in df_sub.columns or "FTAG" not in df_sub.columns:
                    return 0.50
                d_val = df_sub.dropna(subset=["FTHG", "FTAG"]).copy()
                if d_val.empty:
                    return 0.50
                tot_g = d_val["FTHG"] + d_val["FTAG"]
                if tipo_m == "over15":
                    return float((tot_g > 1.5).mean())
                if tipo_m == "over25":
                    return float((tot_g > 2.5).mean())
                if tipo_m == "btts":
                    return float(((d_val["FTHG"] > 0) & (d_val["FTAG"] > 0)).mean())
                if tipo_m == "gol_1t":
                    if "HTHG" in d_val.columns and "HTAG" in d_val.columns:
                        d_ht = d_val.dropna(subset=["HTHG", "HTAG"])
                        if not d_ht.empty:
                            return float(((d_ht["HTHG"] + d_ht["HTAG"]) > 0.5).mean())
                    return float(np.clip(1.0 - math.exp(-0.44 * float(tot_g.mean() or 2.4)), 0.15, 0.92))
                return 0.50

            filas_frecuencia = []
            for etiqueta_m, clave_m in [
                ("⚽ Over 1.5 Goles (+1.5 FT)", "over15"),
                ("🔥 Over 2.5 Goles (+2.5 FT)", "over25"),
                ("🎯 Ambos Marcan (BTTS Sí)", "btts"),
                ("⏱️ Gol en 1º Tiempo (+0.5 HT)", "gol_1t"),
            ]:
                f_loc = _frecuencia_mercado_condicionada(df_eq1_home, clave_m)
                f_vis = _frecuencia_mercado_condicionada(df_eq2_away, clave_m)
                f_conj = float(np.clip((f_loc + f_vis) / 2.0, 0.01, 0.99))
                cuota_j_emp = round(1.0 / f_conj, 2)
                filas_frecuencia.append({
                    "Mercado Analizado": etiqueta_m,
                    f"🏠 % Casa ({eq1})": f"{f_loc * 100:.1f}%",
                    f"✈️ % Fuera ({eq2})": f"{f_vis * 100:.1f}%",
                    "📊 Frecuencia Conjunta (%)": f"{f_conj * 100:.1f}%",
                    "💎 Cuota Justa Empírica (1/p)": f"Cuota: {cuota_j_emp:.2f}",
                })

            df_frec_tabla = pd.DataFrame(filas_frecuencia)
            st.dataframe(df_frec_tabla, use_container_width=True, hide_index=True)

            # ── 6. HISTÓRICO H2H DIRECTO (TODAS LAS TEMPORADAS) Y FORMA RECIENTE ──
            st.markdown("---")
            st.subheader("🗄️ Auditoría Histórica SQLite: H2H Directo y Forma Reciente en Liga")

            h2h_directo_df = pd.DataFrame()
            df_sql_h2h = df_partidos_global.copy()
            if RUTA_BD.exists():
                try:
                    with sqlite3.connect(RUTA_BD) as conn_h2h:
                        tablas_h = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table';", conn_h2h)["name"].tolist()
                        tablas_validas = [t for t in ("historico_reciente", "partidos") if t in tablas_h]
                        dfs_h2h_all = []
                        for t_nom in tablas_validas:
                            q_directo = f"""
                                SELECT * FROM {t_nom}
                                WHERE (LOWER(TRIM(HomeTeam)) = LOWER(TRIM(?)) AND LOWER(TRIM(AwayTeam)) = LOWER(TRIM(?)))
                                   OR (LOWER(TRIM(HomeTeam)) = LOWER(TRIM(?)) AND LOWER(TRIM(AwayTeam)) = LOWER(TRIM(?)))
                                ORDER BY Date DESC
                            """
                            df_t_dir = pd.read_sql(q_directo, conn_h2h, params=(eq1, eq2, eq2, eq1))
                            if not df_t_dir.empty:
                                dfs_h2h_all.append(df_t_dir)
                        if dfs_h2h_all:
                            h2h_directo_df = pd.concat(dfs_h2h_all, ignore_index=True)
                            h2h_directo_df = h2h_directo_df.drop_duplicates(subset=["Date", "HomeTeam", "AwayTeam"], keep="first")

                        t_principal = "historico_reciente" if "historico_reciente" in tablas_h else ("partidos" if "partidos" in tablas_h else None)
                        if t_principal:
                            df_sql_h2h = pd.read_sql(
                                f"""
                                SELECT * FROM {t_principal}
                                WHERE LOWER(TRIM(HomeTeam)) IN (LOWER(TRIM(?)), LOWER(TRIM(?)))
                                   OR LOWER(TRIM(AwayTeam)) IN (LOWER(TRIM(?)), LOWER(TRIM(?)))
                                """,
                                conn_h2h,
                                params=(eq1, eq2, eq1, eq2),
                            )
                except Exception:
                    pass

            if h2h_directo_df.empty:
                mask_dir_mem = (
                    (df_partidos_global["HomeTeam"].astype(str).str.strip().str.lower() == eq1.strip().lower())
                    & (df_partidos_global["AwayTeam"].astype(str).str.strip().str.lower() == eq2.strip().lower())
                ) | (
                    (df_partidos_global["HomeTeam"].astype(str).str.strip().str.lower() == eq2.strip().lower())
                    & (df_partidos_global["AwayTeam"].astype(str).str.strip().str.lower() == eq1.strip().lower())
                )
                h2h_directo_df = df_partidos_global[mask_dir_mem].copy()

            if not h2h_directo_df.empty:
                if "Date" in h2h_directo_df.columns:
                    h2h_directo_df["_dt_orden"] = pd.to_datetime(h2h_directo_df["Date"], errors="coerce", dayfirst=True)
                    h2h_directo_df = h2h_directo_df.sort_values(by="_dt_orden", ascending=False)
                h2h_directo_df = h2h_directo_df.dropna(subset=["FTHG", "FTAG"]).head(5).copy()

            if "Date" in df_sql_h2h.columns:
                df_sql_h2h = df_sql_h2h.copy()
                df_sql_h2h["_dt_orden"] = pd.to_datetime(df_sql_h2h["Date"], errors="coerce", dayfirst=True)
                df_sql_h2h = df_sql_h2h.sort_values(by="_dt_orden", ascending=False)

            # 6a) Enfrentamientos Directos H2H (Últimos 5 en todas las temporadas/competiciones)
            st.markdown(f"#### ⚔️ Enfrentamientos Directos H2H (Últimos 5 registrados entre {eq1} y {eq2})")
            if not h2h_directo_df.empty:
                filas_h2h_dir = []
                for _, r_h in h2h_directo_df.iterrows():
                    hg_i = int(pd.to_numeric(r_h.get("FTHG", 0), errors="coerce") or 0)
                    ag_i = int(pd.to_numeric(r_h.get("FTAG", 0), errors="coerce") or 0)
                    ref_val = str(r_h.get("Referee", "") or "").strip()
                    if not ref_val or ref_val.lower() == "nan":
                        ref_val = "No registrado"
                    filas_h2h_dir.append({
                        "📅 Fecha": str(r_h.get("Date", "N/A")),
                        "🏠 Local": str(r_h.get("HomeTeam", "")),
                        "⚽ Marcador": f"{hg_i} - {ag_i}",
                        "✈️ Visitante": str(r_h.get("AwayTeam", "")),
                        "⚖️ Árbitro": ref_val,
                    })
                st.dataframe(pd.DataFrame(filas_h2h_dir), use_container_width=True, hide_index=True)
            else:
                st.info("No se registran enfrentamientos directos recientes entre ambos equipos en macro_base_futbol.db.")

            # 6b & 6c) Forma Reciente en Liga - Últimos 3 partidos de cada equipo
            def _construir_tabla_ultimos_3(df_fuente: pd.DataFrame, equipo_obj: str) -> pd.DataFrame:
                eq_norm_low = str(equipo_obj).strip().lower()
                sub_eq = df_fuente[
                    (df_fuente["HomeTeam"].astype(str).str.strip().str.lower() == eq_norm_low)
                    | (df_fuente["AwayTeam"].astype(str).str.strip().str.lower() == eq_norm_low)
                ].dropna(subset=["FTHG", "FTAG"]).head(3)
                filas_u3 = []
                for _, r_u in sub_eq.iterrows():
                    es_casa = str(r_u.get("HomeTeam", "")).strip().lower() == eq_norm_low
                    rival = str(r_u.get("AwayTeam", "")) if es_casa else str(r_u.get("HomeTeam", ""))
                    hg_u = int(pd.to_numeric(r_u.get("FTHG", 0), errors="coerce") or 0)
                    ag_u = int(pd.to_numeric(r_u.get("FTAG", 0), errors="coerce") or 0)
                    gf_u = hg_u if es_casa else ag_u
                    gc_u = ag_u if es_casa else hg_u
                    res_icono = "✅ V" if gf_u > gc_u else ("➖ E" if gf_u == gc_u else "❌ D")
                    filas_u3.append({
                        "📅 Fecha": str(r_u.get("Date", "N/A")),
                        "🆚 Rival": rival,
                        "⚽ Marcador": f"{hg_u} - {ag_u} ({res_icono})",
                        "🏟️ Condición": "Casa (Local) 🏠" if es_casa else "Fuera (Visitante) ✈️",
                    })
                return pd.DataFrame(filas_u3)

            col_f_loc, col_f_vis = st.columns(2)
            with col_f_loc:
                st.markdown(f"#### 🏠 Forma Reciente en Liga — Últimos 3 de {eq1}")
                df_u3_loc = _construir_tabla_ultimos_3(df_sql_h2h, eq1)
                if not df_u3_loc.empty:
                    st.dataframe(df_u3_loc, use_container_width=True, hide_index=True)
                else:
                    st.info(f"Sin registros recientes para {eq1}.")

            with col_f_vis:
                st.markdown(f"#### ✈️ Forma Reciente en Liga — Últimos 3 de {eq2}")
                df_u3_vis = _construir_tabla_ultimos_3(df_sql_h2h, eq2)
                if not df_u3_vis.empty:
                    st.dataframe(df_u3_vis, use_container_width=True, hide_index=True)
                else:
                    st.info(f"Sin registros recientes para {eq2}.")

            # ── 7. EXPORTACIÓN EJECUTIVA 1-CLICK (COPIAR FICHA DEL PARTIDO) ──
            st.markdown("---")
            top3_resumen_str = " | ".join(
                f"{m_s} ({p_s * 100:.1f}% · Cuota: {1.0 / max(p_s, 1e-4):.2f})"
                for m_s, p_s in marcadores_ordenados
            )
            ficha_partido_txt = (
                f"=== FICHA EJECUTIVA H2H: {eq1} vs {eq2} ({liga_h2h_nom}) ===\n"
                f"• Goles Esperados (xG Proyectado): {eq1} {xg_l:.2f} — {eq2} {xg_v:.2f}"
                + (" [Suavizado Bayesiano 70/30 Activo]\n" if muestra_reducida else "\n")
                + f"• Cuotas Justas 1X2: Local ({eq1}) {1.0 / p_1:.2f} ({p_1 * 100:.1f}%) | "
                f"Empate {1.0 / p_x:.2f} ({p_x * 100:.1f}%) | Visitante ({eq2}) {1.0 / p_2:.2f} ({p_2 * 100:.1f}%)\n"
                f"• Mercados de Goles: Over 2.5 Cuota: {1.0 / p_o25:.2f} ({p_o25 * 100:.1f}%) | "
                f"Under 2.5 Cuota: {1.0 / p_u25:.2f} ({p_u25 * 100:.1f}%) | BTTS Sí Cuota: {1.0 / p_btts_si:.2f} ({p_btts_si * 100:.1f}%)\n"
                f"• Top 3 Marcadores Exactos: {top3_resumen_str}\n"
                f"• Línea Hándicap Asiático Recomendada: {eq1} {linea_ah_str} / {eq2} {linea_ah_vis_str} "
                f"(Supremacía: {supremacia_xg:+.2f} xG | DNB {eq1}: {cuota_dnb_loc:.2f} / {eq2}: {cuota_dnb_vis:.2f})"
            )

            btn_ficha = st.button("📋 Copiar Ficha del Partido", type="primary", key="btn_copiar_ficha_h2h")
            with st.expander("📋 Ficha Ejecutiva del Partido (Lista para copiar en 1 clic)", expanded=bool(btn_ficha)):
                if btn_ficha:
                    st.success("✅ Ficha ejecutiva generada. Haz clic en el icono de copia en la esquina superior derecha del bloque:")
                st.code(ficha_partido_txt, language="text")

        elif eq1 == eq2:
            st.warning("Selecciona dos equipos distintos para realizar la comparativa.")

elif menu_seleccionado == "🎯 Radar de Tendencias":
    import os as _os_rad
    import html as _html_rad
    try:
        from google import genai as _genai_rad
        _RADAR_GENAI_OK = True
    except Exception:
        _genai_rad = None
        _RADAR_GENAI_OK = False

    # ── PILAR 1: DARK-THEME NATIVE UI & SCOPED STYLING OVERHAUL ──────────────
    st.markdown(
        """
        <style>
        .radar-hero-banner {
            background: linear-gradient(135deg, #0e1117 0%, #161922 55%, #1e2230 100%);
            border: 1px solid #2d3345;
            border-left: 4px solid #38bdf8;
            border-radius: 12px;
            padding: 16px 20px;
            margin-bottom: 16px;
        }
        .radar-choque-wrapper {
            background: linear-gradient(135deg, #121723 0%, #192032 100%);
            border: 1px solid rgba(56, 189, 248, 0.45);
            border-radius: 12px;
            padding: 16px 18px;
            margin-bottom: 18px;
            box-shadow: 0 6px 20px rgba(0, 0, 0, 0.35);
        }
        .radar-choque-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
            gap: 12px;
            margin-top: 12px;
        }
        .radar-choque-card {
            background-color: #0e1117;
            border: 1px solid #262c3d;
            border-top: 3px solid #00d26a;
            border-radius: 10px;
            padding: 13px 15px;
        }
        .radar-dark-table-wrap {
            background-color: #161922;
            border: 1px solid #262c3d;
            border-radius: 12px;
            overflow-x: auto;
            margin-top: 8px;
            margin-bottom: 14px;
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.30);
        }
        table.radar-dark-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.87rem;
            color: #f0f2f6;
            background-color: #161922;
        }
        table.radar-dark-table thead tr {
            background-color: #1e2230;
            border-bottom: 2px solid #2d3448;
        }
        table.radar-dark-table th {
            padding: 11px 13px;
            text-align: left;
            font-weight: 700;
            color: #94a3b8;
            text-transform: uppercase;
            font-size: 0.75rem;
            letter-spacing: 0.04em;
            white-space: nowrap;
        }
        table.radar-dark-table tbody tr {
            border-bottom: 1px solid #222736;
            background-color: #161922;
            transition: background-color 0.15s ease;
        }
        table.radar-dark-table tbody tr:nth-child(even) {
            background-color: #131620;
        }
        table.radar-dark-table tbody tr:hover {
            background-color: #1e2436;
        }
        table.radar-dark-table td {
            padding: 10px 13px;
            vertical-align: middle;
        }
        .pill-win {
            display: inline-block;
            background-color: rgba(0, 210, 106, 0.14);
            color: #00d26a;
            border: 1px solid rgba(0, 210, 106, 0.42);
            border-radius: 6px;
            padding: 2px 7px;
            margin: 2px 3px 2px 0;
            font-family: monospace;
            font-size: 0.77rem;
            font-weight: 700;
            white-space: nowrap;
        }
        .pill-loss {
            display: inline-block;
            background-color: rgba(248, 113, 113, 0.14);
            color: #f87171;
            border: 1px solid rgba(248, 113, 113, 0.42);
            border-radius: 6px;
            padding: 2px 7px;
            margin: 2px 3px 2px 0;
            font-family: monospace;
            font-size: 0.77rem;
            font-weight: 600;
            white-space: nowrap;
        }
        .pill-ev-high {
            display: inline-block;
            background-color: rgba(0, 210, 106, 0.16);
            color: #00d26a;
            border: 1px solid #00d26a;
            border-radius: 6px;
            padding: 3px 9px;
            font-family: monospace;
            font-size: 0.80rem;
            font-weight: 700;
            white-space: nowrap;
        }
        .pill-ev-cyan {
            display: inline-block;
            background-color: rgba(56, 189, 248, 0.15);
            color: #38bdf8;
            border: 1px solid rgba(56, 189, 248, 0.50);
            border-radius: 6px;
            padding: 3px 9px;
            font-family: monospace;
            font-size: 0.80rem;
            font-weight: 700;
            white-space: nowrap;
        }
        .pill-ev-neg {
            display: inline-block;
            background-color: rgba(248, 113, 113, 0.14);
            color: #f87171;
            border: 1px solid rgba(248, 113, 113, 0.42);
            border-radius: 6px;
            padding: 3px 9px;
            font-family: monospace;
            font-size: 0.80rem;
            font-weight: 700;
            white-space: nowrap;
        }
        .pill-fixture {
            display: inline-block;
            background-color: #1e2230;
            color: #e2e8f0;
            border: 1px solid #334155;
            border-radius: 6px;
            padding: 3px 8px;
            font-size: 0.79rem;
            font-weight: 600;
            white-space: nowrap;
        }
        .pill-fixture-none {
            display: inline-block;
            background-color: rgba(100, 116, 139, 0.14);
            color: #94a3b8;
            border: 1px solid rgba(100, 116, 139, 0.30);
            border-radius: 6px;
            padding: 3px 8px;
            font-size: 0.78rem;
            white-space: nowrap;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("🎯 Radar Cuantitativo de Tendencias & Convergencia EV+")
    st.markdown(
        "<div class='radar-hero-banner'>"
        "📡 <b>Escáner Institucional de Rachas y Valor Esperado:</b> Detección de rachas empíricas por condición de campo, "
        "cruce automático con el calendario próximo, motor de <b>Choque de Tendencias (Doble Coincidencia)</b>, "
        "calibración <b>Cuota Justa Zero-Vig vs Mercado</b> y síntesis ejecutiva con <b>Google AI Studio</b>."
        "</div>",
        unsafe_allow_html=True,
    )

    if err_partidos or df_partidos.empty:
        st.warning("No hay suficientes datos históricos disponibles en `macro_base_futbol.db`.")
    else:
        # Normalización defensiva de columna de liga para soportar tanto filtro lateral como selector propio
        df_rad_base = df_partidos.copy()
        col_div_rad = next(
            (c for c in ("Div", "Liga_Origen") if c in df_rad_base.columns),
            next((c for c in df_rad_base.columns if "Div" in str(c)), None),
        )
        divs_disp_rad = (
            [d for d in MAPA_LIGAS if d in set(df_rad_base[col_div_rad].dropna().astype(str))]
            if col_div_rad
            else []
        )
        opciones_liga_rad = ["Todas las Competiciones"] + [MAPA_LIGAS[d] for d in divs_disp_rad]

        col_rf0, col_rf1, col_rf2 = st.columns([1.2, 1.8, 1.3])
        with col_rf0:
            liga_rad_sel = st.selectbox(
                "🌍 Competición:",
                opciones_liga_rad,
                index=0,
                key="radar_filtro_liga",
            )
        with col_rf1:
            filtro_condicion = st.radio(
                "🏠 Condición de Campo:",
                ["Global (Todos los partidos)", "Solo como Local 🏠", "Solo como Visitante ✈️"],
                horizontal=True,
                key="radar_filtro_condicion",
            )
        with col_rf2:
            umbral_racha = st.select_slider(
                "🔥 Umbral Mínimo de Racha:",
                options=[3, 4, 5],
                value=4,
                format_func=lambda x: f"Mínimo {x}/5 ({x * 20}%)",
                key="radar_umbral_racha",
            )

        if liga_rad_sel != "Todas las Competiciones" and col_div_rad:
            div_target_rad = MAPA_INV.get(liga_rad_sel, "")
            df_rad = df_rad_base[df_rad_base[col_div_rad].astype(str) == str(div_target_rad)].copy()
        else:
            df_rad = df_rad_base.copy()

        for _cn in (
            "FTHG", "FTAG", "HTHG", "HTAG", "HC", "AC", "HY", "AY", "HR", "AR",
            "AHh", "AHCh", "Avg>2.5", "B365>2.5", "Max>2.5",
            "AvgH", "AvgD", "AvgA", "B365H", "B365A",
            "AvgAHH", "AvgAHA", "B365AHH", "B365AHA", "MaxAHH", "MaxAHA",
        ):
            if _cn in df_rad.columns:
                df_rad[_cn] = pd.to_numeric(df_rad[_cn], errors="coerce")

        if "Date" in df_rad.columns:
            df_rad["Date_Parsed"] = pd.to_datetime(df_rad["Date"], errors="coerce", dayfirst=True)
        else:
            df_rad["Date_Parsed"] = pd.NaT

        # ── Helpers internos del Radar (Pilar 2, 3, 4 y 5) ───────────────────
        def _abrev_rival_radar(nombre_eq: str) -> str:
            limpio = "".join(ch for ch in str(nombre_eq or "RIV").upper() if ch.isalnum() or ch == " ").strip()
            partes = limpio.split()
            if len(partes) >= 2 and len(partes[0]) <= 3:
                return (partes[0][:1] + partes[1][:2]).upper()[:3]
            return (limpio.replace(" ", "") + "RIV")[:3]

        def _construir_horizonte_fixtures_y_cuotas(df_scope: pd.DataFrame) -> tuple[dict, list[dict]]:
            """
            Cruza calendarios próximos desde:
            1) Caché activa del Escáner EV+ (`st.session_state.df_ops_export`) y `ev_fix_manual`
            2) Partidos programados sin marcador (`FTHG.isna()`) en `macro_base_futbol.db`
            3) Horizonte de la última jornada activa por competición en `macro_base_futbol.db`
            Devuelve (mapa_equipo_fixture, lista_emparejamientos_unicos).
            """
            mapa_eq = {}
            lista_duelos = []
            vistos_duelos = set()

            # A) Cruce con caché de cuotas en vivo de conector_odds (si existe en sesión)
            df_ops_cache = st.session_state.get("df_ops_export")
            cuotas_api_duelo = {}
            if isinstance(df_ops_cache, pd.DataFrame) and not df_ops_cache.empty and "Partido" in df_ops_cache.columns:
                for _, r_op in df_ops_cache.iterrows():
                    p_txt = str(r_op.get("Partido", ""))
                    if " vs " in p_txt:
                        h_op, a_op = [x.strip() for x in p_txt.split(" vs ", 1)]
                        f_op = str(r_op.get("Fecha / Hora", "Próxima Jornada"))
                        m_op = str(r_op.get("Mercado", ""))
                        c_op = pd.to_numeric(r_op.get("Cuota"), errors="coerce")
                        cuotas_api_duelo.setdefault((h_op, a_op), {"fecha": f_op, "mercados": {}})
                        if pd.notna(c_op) and c_op > 1.0:
                            cuotas_api_duelo[(h_op, a_op)]["mercados"][m_op] = float(c_op)

            for (h_op, a_op), info_op in cuotas_api_duelo.items():
                f_str = info_op["fecha"]
                mapa_eq[h_op] = {
                    "rival": a_op,
                    "condicion": "Local",
                    "etiqueta": f"vs {a_op} [Local]",
                    "fecha": f_str,
                    "cuotas_api": info_op["mercados"],
                }
                mapa_eq[a_op] = {
                    "rival": h_op,
                    "condicion": "Visitante",
                    "etiqueta": f"vs {h_op} [Visitante]",
                    "fecha": f_str,
                    "cuotas_api": info_op["mercados"],
                }
                if (h_op, a_op) not in vistos_duelos:
                    vistos_duelos.add((h_op, a_op))
                    lista_duelos.append({"local": h_op, "visitante": a_op, "fecha": f_str, "div": ""})

            # B) Fixtures manuales cargados en sesión
            for fx_m in st.session_state.get("ev_fix_manual", []):
                h_m = str(fx_m.get("Local", "")).strip()
                a_m = str(fx_m.get("Visitante", "")).strip()
                dt_m = fx_m.get("Fecha")
                f_m_str = pd.Timestamp(dt_m).strftime("%d/%m/%Y") if dt_m is not None else "Programado"
                if h_m and a_m and (h_m, a_m) not in vistos_duelos:
                    vistos_duelos.add((h_m, a_m))
                    lista_duelos.append({"local": h_m, "visitante": a_m, "fecha": f_m_str, "div": str(fx_m.get("Div", ""))})
                    mapa_eq.setdefault(h_m, {"rival": a_m, "condicion": "Local", "etiqueta": f"vs {a_m} [Local]", "fecha": f_m_str, "cuotas_api": {}})
                    mapa_eq.setdefault(a_m, {"rival": h_m, "condicion": "Visitante", "etiqueta": f"vs {h_m} [Visitante]", "fecha": f_m_str, "cuotas_api": {}})

            # C) Partidos futuros en BD (FTHG isna) o jornada activa por liga en macro_base_futbol.db
            if df_scope is not None and not df_scope.empty:
                df_no_jug = df_scope[df_scope["FTHG"].isna() & df_scope["HomeTeam"].notna() & df_scope["AwayTeam"].notna()].copy()
                if not df_no_jug.empty:
                    df_horizonte = df_no_jug.sort_values("Date_Parsed", ascending=True)
                else:
                    # Extraer la ventana de la jornada activa más reciente por liga (últimos 4 días de competición)
                    partes_jornada = []
                    if col_div_rad and col_div_rad in df_scope.columns:
                        for _, sub_l in df_scope.groupby(col_div_rad):
                            max_dt = sub_l["Date_Parsed"].max()
                            if pd.notna(max_dt):
                                partes_jornada.append(sub_l[sub_l["Date_Parsed"] >= (max_dt - pd.Timedelta(days=4))])
                    if partes_jornada:
                        df_horizonte = pd.concat(partes_jornada, ignore_index=True).sort_values("Date_Parsed", ascending=False)
                    else:
                        max_dt_all = df_scope["Date_Parsed"].max()
                        df_horizonte = (
                            df_scope[df_scope["Date_Parsed"] >= (max_dt_all - pd.Timedelta(days=4))].sort_values("Date_Parsed", ascending=False)
                            if pd.notna(max_dt_all)
                            else pd.DataFrame()
                        )

                for _, r_hz in df_horizonte.iterrows():
                    h_bd = str(r_hz.get("HomeTeam", "")).strip()
                    a_bd = str(r_hz.get("AwayTeam", "")).strip()
                    if not h_bd or not a_bd:
                        continue
                    f_bd = str(r_hz.get("Date", "")).strip() or "Jornada Activa"
                    div_bd = str(r_hz.get(col_div_rad, "")) if col_div_rad else ""
                    if h_bd not in mapa_eq:
                        mapa_eq[h_bd] = {
                            "rival": a_bd,
                            "condicion": "Local",
                            "etiqueta": f"vs {a_bd} [Local]",
                            "fecha": f_bd,
                            "fila_bd": r_hz,
                            "cuotas_api": {},
                        }
                    if a_bd not in mapa_eq:
                        mapa_eq[a_bd] = {
                            "rival": h_bd,
                            "condicion": "Visitante",
                            "etiqueta": f"vs {h_bd} [Visitante]",
                            "fecha": f_bd,
                            "fila_bd": r_hz,
                            "cuotas_api": {},
                        }
                    if (h_bd, a_bd) not in vistos_duelos:
                        vistos_duelos.add((h_bd, a_bd))
                        lista_duelos.append({"local": h_bd, "visitante": a_bd, "fecha": f_bd, "div": div_bd})

            return mapa_eq, lista_duelos

        def _estimar_cuota_mercado_radar(u5_df: pd.DataFrame, eq: str, clave_mercado: str, info_fix: dict | None) -> float:
            """
            Extrae la cuota decimal de mercado desde la caché de cuotas en vivo o desde las columnas
            históricas reales de `macro_base_futbol.db` (`Avg>2.5`, `B365>2.5`, `AvgAHH`/`AvgAHA`, `AvgH`/`AvgA`),
            aplicando calibración empírica con margen bookmaker cuando el mercado es derivado.
            """
            cuotas_api = (info_fix or {}).get("cuotas_api") or {}
            if clave_mercado == "OVER25":
                for k_api, v_api in cuotas_api.items():
                    if "2.5" in k_api and ("Más" in k_api or "Over" in k_api):
                        return round(float(v_api), 2)
                for col_c in ("Max>2.5", "Avg>2.5", "B365>2.5"):
                    if col_c in u5_df.columns:
                        s_c = pd.to_numeric(u5_df[col_c], errors="coerce").dropna()
                        if not s_c.empty:
                            return round(float(s_c.iloc[0]), 2)
                return 1.85

            if clave_mercado == "BTTS":
                for k_api, v_api in cuotas_api.items():
                    if "Ambos" in k_api or "BTTS" in k_api:
                        return round(float(v_api), 2)
                for col_c in ("Avg>2.5", "B365>2.5"):
                    if col_c in u5_df.columns:
                        s_c = pd.to_numeric(u5_df[col_c], errors="coerce").dropna()
                        if not s_c.empty:
                            return round(float(np.clip(s_c.mean() * 0.96, 1.55, 2.25)), 2)
                return 1.80

            if clave_mercado == "HT05":
                for col_c in ("Avg>2.5", "B365>2.5"):
                    if col_c in u5_df.columns:
                        s_c = pd.to_numeric(u5_df[col_c], errors="coerce").dropna()
                        if not s_c.empty:
                            return round(float(np.clip(s_c.mean() * 0.76, 1.30, 1.72)), 2)
                return 1.42

            if clave_mercado == "CORNERS85":
                if {"HC", "AC"}.issubset(u5_df.columns):
                    c_med = float((u5_df["HC"].fillna(0) + u5_df["AC"].fillna(0)).mean())
                    # Línea estándar de mercado para +8.5 córners en función del volumen medio
                    c_mkt = 1.85 - np.clip((c_med - 9.5) * 0.06, -0.22, 0.22)
                    return round(float(np.clip(c_mkt, 1.52, 2.20)), 2)
                return 1.83

            if clave_mercado == "CARDS45":
                if {"HY", "AY"}.issubset(u5_df.columns):
                    t_med = float((u5_df["HY"].fillna(0) + u5_df["AY"].fillna(0) + u5_df.get("HR", 0).fillna(0) + u5_df.get("AR", 0).fillna(0)).mean())
                    c_mkt = 1.90 - np.clip((t_med - 4.7) * 0.07, -0.25, 0.25)
                    return round(float(np.clip(c_mkt, 1.55, 2.30)), 2)
                return 1.88

            if clave_mercado == "ATS":
                cuotas_ats = []
                for _, r_u in u5_df.iterrows():
                    es_loc = str(r_u.get("HomeTeam", "")) == eq
                    col_pref = ("MaxAHH", "AvgAHH", "B365AHH") if es_loc else ("MaxAHA", "AvgAHA", "B365AHA")
                    for c_ah in col_pref:
                        val_c = pd.to_numeric(r_u.get(c_ah), errors="coerce")
                        if pd.notna(val_c) and val_c > 1.01:
                            cuotas_ats.append(float(val_c))
                            break
                if cuotas_ats:
                    return round(float(cuotas_ats[0]), 2)
                return 1.92

            return 1.85

        def _evaluar_racha_equipo_mercado(u5_df: pd.DataFrame, eq: str, clave_mercado: str, info_fix: dict | None, liga_eq: str) -> dict | None:
            """
            Evalúa los últimos N partidos (u5_df) de `eq` en `clave_mercado`.
            Genera tokens granulares `[GF-GA vs RIV]` con estilo pill, calcula Cuota Justa Zero-Vig,
            Cuota Mercado y EV% Spread.
            """
            if u5_df is None or len(u5_df) < 3:
                return None

            aciertos_bool = []
            tokens_html = []
            tokens_texto = []

            for _, fila in u5_df.iterrows():
                es_loc = str(fila.get("HomeTeam", "")) == eq
                rival = str(fila.get("AwayTeam", "") if es_loc else fila.get("HomeTeam", ""))
                abrev = _abrev_rival_radar(rival)
                fthg = pd.to_numeric(fila.get("FTHG"), errors="coerce")
                ftag = pd.to_numeric(fila.get("FTAG"), errors="coerce")
                if pd.isna(fthg) or pd.isna(ftag):
                    continue
                gf = int(float(fthg if es_loc else ftag))
                ga = int(float(ftag if es_loc else fthg))

                hit = False
                etiqueta_extra = ""
                if clave_mercado == "OVER25":
                    hit = (fthg + ftag) > 2.5
                elif clave_mercado == "BTTS":
                    hit = (fthg > 0) and (ftag > 0)
                elif clave_mercado == "HT05":
                    hthg = pd.to_numeric(fila.get("HTHG"), errors="coerce")
                    htag = pd.to_numeric(fila.get("HTAG"), errors="coerce")
                    if pd.isna(hthg) or pd.isna(htag):
                        continue
                    hit = (hthg + htag) > 0.5
                    etiqueta_extra = f" HT:{int(hthg)}-{int(htag)}"
                elif clave_mercado == "CORNERS85":
                    hc = pd.to_numeric(fila.get("HC"), errors="coerce")
                    ac = pd.to_numeric(fila.get("AC"), errors="coerce")
                    if pd.isna(hc) or pd.isna(ac):
                        continue
                    tot_c = int(hc + ac)
                    hit = tot_c > 8.5
                    etiqueta_extra = f" {tot_c}C"
                elif clave_mercado == "CARDS45":
                    hy = pd.to_numeric(fila.get("HY"), errors="coerce")
                    ay = pd.to_numeric(fila.get("AY"), errors="coerce")
                    hr = pd.to_numeric(fila.get("HR"), errors="coerce")
                    ar = pd.to_numeric(fila.get("AR"), errors="coerce")
                    if pd.isna(hy) or pd.isna(ay):
                        continue
                    tot_t = int(hy + ay + (0 if pd.isna(hr) else hr) + (0 if pd.isna(ar) else ar))
                    hit = tot_t > 4.5
                    etiqueta_extra = f" {tot_t}T"
                elif clave_mercado == "ATS":
                    ah_raw = pd.to_numeric(fila.get("AHh"), errors="coerce")
                    if pd.isna(ah_raw):
                        ah_raw = pd.to_numeric(fila.get("AHCh"), errors="coerce")
                    if pd.isna(ah_raw):
                        ah_raw = -0.25
                    linea_eq = float(ah_raw) if es_loc else -float(ah_raw)
                    diff_ats = (gf + linea_eq) - ga
                    if abs(diff_ats) < 1e-6:
                        # Push se ignora o se considera no-perdido; evaluamos cobertura efectiva (> 0)
                        hit = False
                    else:
                        hit = diff_ats > 0.0
                    etiqueta_extra = f" AH{linea_eq:+.2g}"

                aciertos_bool.append(bool(hit))
                token_str = f"[{gf}-{ga} vs {abrev}{etiqueta_extra}]"
                tokens_texto.append(f"{'✅' if hit else '❌'}{token_str}")
                clase_pill = "pill-win" if hit else "pill-loss"
                tokens_html.append(f"<span class='{clase_pill}'>{_html_rad.escape(token_str)}</span>")

            n_eval = len(aciertos_bool)
            if n_eval < 3:
                return None

            n_hits = int(sum(aciertos_bool))
            winrate_pct = (n_hits / n_eval) * 100.0
            cuota_justa = round(1.0 / max(winrate_pct / 100.0, 0.01), 2)
            cuota_mercado = _estimar_cuota_mercado_radar(u5_df, eq, clave_mercado, info_fix)
            ev_pct = round(((cuota_mercado / cuota_justa) - 1.0) * 100.0, 1)

            if info_fix:
                prox_rival_lbl = info_fix.get("etiqueta", "Sin emparejamiento próximo")
                prox_fecha_lbl = info_fix.get("fecha", "")
                prox_html = (
                    f"<span class='pill-fixture'>🗓️ {_html_rad.escape(prox_rival_lbl)}"
                    + (f" · <span style='color:#38bdf8;'>{_html_rad.escape(prox_fecha_lbl)}</span>" if prox_fecha_lbl else "")
                    + "</span>"
                )
                prox_plano = f"{prox_rival_lbl} ({prox_fecha_lbl})" if prox_fecha_lbl else prox_rival_lbl
            else:
                prox_html = "<span class='pill-fixture-none'>Sin emparejamiento próximo</span>"
                prox_plano = "Sin emparejamiento próximo"

            if ev_pct > 5.0:
                ev_badge_html = f"<span class='pill-ev-high'>🔥 {ev_pct:+.1f}% (EV+)</span>"
            elif ev_pct >= -5.0:
                ev_badge_html = f"<span class='pill-ev-cyan'>⚖️ {ev_pct:+.1f}% (En Línea)</span>"
            else:
                ev_badge_html = f"<span class='pill-ev-neg'>⚠️ {ev_pct:+.1f}% (Sobrevalorado)</span>"

            return {
                "Equipo": eq,
                "Liga": liga_eq,
                "Aciertos_Num": n_hits,
                "Muestra_Num": n_eval,
                "Aciertos": f"{n_hits}/{n_eval}",
                "Winrate_Num": round(winrate_pct, 1),
                "Winrate": f"{winrate_pct:.0f}%",
                "Cuota_Justa": cuota_justa,
                "Cuota_Mercado": cuota_mercado,
                "EV_Num": ev_pct,
                "EV_Badge_HTML": ev_badge_html,
                "Proximo_HTML": prox_html,
                "Proximo_Plano": prox_plano,
                "Forma_HTML": "".join(tokens_html),
                "Forma_Texto": " ".join(tokens_texto),
            }

        def _renderizar_tabla_oscura_radar(lista_items: list[dict]):
            if not lista_items:
                st.info("Sin rachas que superen el umbral seleccionado en este mercado.")
                return
            ordenados = sorted(lista_items, key=lambda x: (x["Winrate_Num"], x["EV_Num"]), reverse=True)
            filas_tr = []
            for item in ordenados:
                wr_color = "#00d26a" if item["Winrate_Num"] >= 80.0 else "#38bdf8"
                filas_tr.append(
                    "<tr>"
                    f"<td style='font-weight:700; color:#f8fafc;'>⚽ {_html_rad.escape(item['Equipo'])} "
                    f"<span style='font-size:0.74rem; color:#64748b; font-weight:500;'>({_html_rad.escape(item['Liga'])})</span></td>"
                    f"<td>{item['Proximo_HTML']}</td>"
                    f"<td style='font-family:monospace; font-weight:700; color:{wr_color};'>{item['Aciertos']} ({item['Winrate']})</td>"
                    f"<td>{item['Forma_HTML']}</td>"
                    f"<td style='font-family:monospace; font-weight:700; color:#e2e8f0;'>{item['Cuota_Justa']:.2f}</td>"
                    f"<td style='font-family:monospace; font-weight:700; color:#38bdf8;'>{item['Cuota_Mercado']:.2f}</td>"
                    f"<td>{item['EV_Badge_HTML']}</td>"
                    "</tr>"
                )
            html_tabla = (
                "<div class='radar-dark-table-wrap'>"
                "<table class='radar-dark-table'>"
                "<thead><tr>"
                "<th>Equipo / Competición</th>"
                "<th>Próximo Emparejamiento</th>"
                "<th>Racha / Winrate</th>"
                "<th>Desglose Granular (Últimos 5)</th>"
                "<th>Cuota Justa</th>"
                "<th>Cuota Mercado</th>"
                "<th>Diferencial EV%</th>"
                "</tr></thead>"
                f"<tbody>{''.join(filas_tr)}</tbody>"
                "</table></div>"
            )
            st.markdown(html_tabla, unsafe_allow_html=True)

        # ── Construcción del Mapa de Próximos Partidos y Escaneo de Equipos ───
        mapa_fixtures, duelos_horizonte = _construir_horizonte_fixtures_y_cuotas(df_rad)
        equipos_rad = sorted(pd.concat([df_rad["HomeTeam"], df_rad["AwayTeam"]]).dropna().unique().tolist())

        CATALOGO_MERCADOS_RADAR = [
            ("OVER25", "🔥 Más de 2.5 Goles"),
            ("BTTS", "🎯 Ambos Marcan (BTTS)"),
            ("HT05", "⏱️ Más 0.5 Goles 1ª Parte"),
            ("CORNERS85", "🚩 Córners (+8.5 Totales)"),
            ("CARDS45", "🟨 Fricción Disciplinaria (+4.5 Tarjetas)"),
            ("ATS", "🛡️ Hándicap Asiático Cubierto (ATS)"),
        ]

        resultados_por_mercado = {k: [] for k, _ in CATALOGO_MERCADOS_RADAR}
        indice_racha_local = {k: {} for k, _ in CATALOGO_MERCADOS_RADAR}
        indice_racha_visit = {k: {} for k, _ in CATALOGO_MERCADOS_RADAR}

        for eq in equipos_rad:
            df_eq_all = df_rad[(df_rad["HomeTeam"] == eq) | (df_rad["AwayTeam"] == eq)].sort_values(by="Date_Parsed", ascending=False)
            df_eq_loc = df_rad[df_rad["HomeTeam"] == eq].sort_values(by="Date_Parsed", ascending=False)
            df_eq_vis = df_rad[df_rad["AwayTeam"] == eq].sort_values(by="Date_Parsed", ascending=False)

            if "Local" in filtro_condicion:
                u5_activo = df_eq_loc.head(5)
            elif "Visitante" in filtro_condicion:
                u5_activo = df_eq_vis.head(5)
            else:
                u5_activo = df_eq_all.head(5)

            div_eq_val = (
                str(df_eq_all[col_div_rad].dropna().iloc[0])
                if col_div_rad and not df_eq_all.empty and not df_eq_all[col_div_rad].dropna().empty
                else "SP1"
            )
            liga_eq_nom = MAPA_LIGAS.get(div_eq_val, div_eq_val)
            info_fix_eq = mapa_fixtures.get(eq)

            for m_key, _ in CATALOGO_MERCADOS_RADAR:
                # 1) Evaluación según filtro activo para las pestañas
                ev_act = _evaluar_racha_equipo_mercado(u5_activo, eq, m_key, info_fix_eq, liga_eq_nom)
                if ev_act and ev_act["Aciertos_Num"] >= umbral_racha:
                    resultados_por_mercado[m_key].append(ev_act)

                # 2) Evaluación específica Local / Visitante para el motor de Doble Coincidencia
                u5_h_conv = df_eq_loc.head(5) if len(df_eq_loc) >= 3 else df_eq_all.head(5)
                u5_a_conv = df_eq_vis.head(5) if len(df_eq_vis) >= 3 else df_eq_all.head(5)
                ev_h = _evaluar_racha_equipo_mercado(u5_h_conv, eq, m_key, info_fix_eq, liga_eq_nom)
                ev_a = _evaluar_racha_equipo_mercado(u5_a_conv, eq, m_key, info_fix_eq, liga_eq_nom)
                if ev_h and ev_h["Aciertos_Num"] >= umbral_racha:
                    indice_racha_local[m_key][eq] = ev_h
                if ev_a and ev_a["Aciertos_Num"] >= umbral_racha:
                    indice_racha_visit[m_key][eq] = ev_a

        # ── PILAR 3: DOUBLE-CONVERGENCE RADAR ("CHOQUE DE RACHAS" ENGINE) ─────
        choques_detectados = []
        for duelo in duelos_horizonte:
            eq_h = duelo["local"]
            eq_a = duelo["visitante"]
            fecha_d = duelo["fecha"]
            for m_key, m_label in CATALOGO_MERCADOS_RADAR:
                r_h = indice_racha_local[m_key].get(eq_h)
                r_a = indice_racha_visit[m_key].get(eq_a)
                if r_h and r_a:
                    wr_comb = round((r_h["Winrate_Num"] + r_a["Winrate_Num"]) / 2.0, 1)
                    cj_comb = round(1.0 / max(wr_comb / 100.0, 0.01), 2)
                    cm_comb = round((r_h["Cuota_Mercado"] + r_a["Cuota_Mercado"]) / 2.0, 2)
                    ev_comb = round(((cm_comb / cj_comb) - 1.0) * 100.0, 1)
                    choques_detectados.append({
                        "partido": f"{eq_h} vs {eq_a}",
                        "liga": r_h["Liga"],
                        "fecha": fecha_d,
                        "mercado": m_label,
                        "racha_local": r_h,
                        "racha_visit": r_a,
                        "wr_comb": wr_comb,
                        "cuota_justa": cj_comb,
                        "cuota_mercado": cm_comb,
                        "ev_comb": ev_comb,
                    })

        choques_detectados.sort(key=lambda x: (x["wr_comb"], x["ev_comb"]), reverse=True)

        if choques_detectados:
            tarjetas_choque_html = []
            for ch in choques_detectados[:6]:
                rl = ch["racha_local"]
                rv = ch["racha_visit"]
                ev_cls = "pill-ev-high" if ch["ev_comb"] > 5.0 else ("pill-ev-cyan" if ch["ev_comb"] >= -5.0 else "pill-ev-neg")
                tarjetas_choque_html.append(
                    "<div class='radar-choque-card'>"
                    f"<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;'>"
                    f"<span style='color:#38bdf8; font-weight:700; font-size:0.82rem;'>{_html_rad.escape(ch['mercado'])}</span>"
                    f"<span class='{ev_cls}'>EV {ch['ev_comb']:+.1f}%</span>"
                    "</div>"
                    f"<div style='font-size:1.0rem; font-weight:800; color:#f8fafc; margin-bottom:4px;'>"
                    f"⚔️ {_html_rad.escape(ch['partido'])} "
                    f"<span style='font-size:0.76rem; color:#94a3b8; font-weight:500;'>({_html_rad.escape(ch['fecha'])})</span></div>"
                    f"<div style='font-size:0.80rem; color:#cbd5e1; margin-bottom:6px;'>"
                    f"🏠 <b>{_html_rad.escape(rl['Equipo'])}:</b> <span style='color:#00d26a; font-family:monospace;'>{rl['Aciertos']} ({rl['Winrate']})</span> · "
                    f"✈️ <b>{_html_rad.escape(rv['Equipo'])}:</b> <span style='color:#00d26a; font-family:monospace;'>{rv['Aciertos']} ({rv['Winrate']})</span>"
                    "</div>"
                    f"<div style='font-size:0.78rem; color:#94a3b8; font-family:monospace;'>"
                    f"Convergencia: <b style='color:#00d26a;'>{ch['wr_comb']:.0f}%</b> | "
                    f"Cuota Justa: <b style='color:#f8fafc;'>{ch['cuota_justa']:.2f}</b> vs "
                    f"Mercado: <b style='color:#38bdf8;'>{ch['cuota_mercado']:.2f}</b>"
                    "</div>"
                    "</div>"
                )
            st.markdown(
                "<div class='radar-choque-wrapper'>"
                "<div style='display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap;'>"
                "<span style='font-size:1.05rem; font-weight:800; color:#f8fafc;'>"
                "💥 Oportunidad de Alta Fricción: Choque de Tendencias (Doble Coincidencia Local + Visitante)</span>"
                f"<span class='pill-ev-cyan'>{len(choques_detectados)} Convergencias Activas</span>"
                "</div>"
                f"<div class='radar-choque-grid'>{''.join(tarjetas_choque_html)}</div>"
                "</div>",
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                "<div class='radar-choque-wrapper' style='padding:12px 16px;'>"
                "💥 <b>Oportunidad de Alta Fricción: Choque de Tendencias:</b> "
                "<span style='color:#94a3b8;'>No se registran cruces directos donde ambos contendientes superen simultáneamente "
                f"el umbral {umbral_racha}/5 en el mismo mercado dentro del filtro actual. Reduce el umbral a 3/5 o selecciona 'Todas las Competiciones' para ampliar el barrido.</span>"
                "</div>",
                unsafe_allow_html=True,
            )

        # ── PILAR 6: ONE-CLICK AI EXECUTIVE RADAR DIGEST (GOOGLE AI STUDIO) ───
        todas_rachas_activas = []
        for m_key, m_label in CATALOGO_MERCADOS_RADAR:
            for r_item in resultados_por_mercado[m_key]:
                copia_r = dict(r_item)
                copia_r["Mercado_Nombre"] = m_label
                todas_rachas_activas.append(copia_r)

        top3_ev_radar = sorted(todas_rachas_activas, key=lambda x: (x["EV_Num"], x["Winrate_Num"]), reverse=True)[:3]

        col_ai_btn, col_ai_info = st.columns([1.6, 2.4])
        with col_ai_btn:
            ejecutar_veredicto_radar = st.button(
                "⚡ Generar Veredicto Táctico del Radar (AI Studio)",
                use_container_width=True,
                type="primary",
                key="btn_radar_ai_digest",
            )
        with col_ai_info:
            st.caption(
                f"📊 **{len(todas_rachas_activas)} rachas activas** detectadas en 6 mercados · "
                f"Síntesis automática de las **Top 3 oportunidades con mayor EV%** vía `gemini-3.8-flash`."
            )

        if ejecutar_veredicto_radar:
            clave_radar_ia = ""
            try:
                clave_radar_ia = st.secrets.get("GEMINI_API_KEY", "") or ""
            except Exception:
                clave_radar_ia = ""
            if not clave_radar_ia:
                clave_radar_ia = (
                    _os_rad.getenv("GEMINI_API_KEY", "")
                    or st.session_state.get("input_gemini_api_key_asistente", "")
                    or ""
                ).strip()

            if not _RADAR_GENAI_OK or _genai_rad is None:
                st.error("⚠️ El SDK `google-genai` no está disponible en este entorno.")
            elif not clave_radar_ia:
                st.warning(
                    "🔑 Configura `GEMINI_API_KEY` en `st.secrets` / entorno o introdúcela en el módulo **🤖 Asistente IA Local** para activar el Veredicto en vivo."
                )
            elif not top3_ev_radar:
                st.info("No hay rachas activas bajo el filtro actual para sintetizar.")
            else:
                resumen_top3_txt = "\n".join(
                    f"- {idx + 1}. {it['Equipo']} ({it['Liga']}) | Mercado: {it['Mercado_Nombre']} | "
                    f"Racha: {it['Aciertos']} ({it['Winrate']}) | Próximo: {it['Proximo_Plano']} | "
                    f"Cuota Justa: {it['Cuota_Justa']:.2f} vs Cuota Mercado: {it['Cuota_Mercado']:.2f} (EV: {it['EV_Num']:+.1f}%) | "
                    f"Forma: {it['Forma_Texto']}"
                    for idx, it in enumerate(top3_ev_radar)
                )
                prompt_radar_digest = (
                    "Eres el Director Cuantitativo de BigData Football.\n"
                    "Sintetiza exclusivamente las siguientes Top 3 rachas con mayor Valor Esperado (EV+) "
                    "detectadas en el Radar de Tendencias en un informe ejecutivo ultra-directo de exactamente 3 viñetas "
                    "(una por selección, indicando equipo, mercado, próximo rival, diferencial entre Cuota Justa y Cuota Mercado, y dictamen operativo).\n"
                    "Prohibido usar lambda 'λ' o '@' antes de las cuotas.\n\n"
                    f"TOP_3_RACHAS_RADAR:\n{resumen_top3_txt}"
                )

                def _stream_veredicto_radar(prompt_txt):
                    client_r = _genai_rad.Client(api_key=clave_radar_ia)
                    try:
                        stream_r = client_r.models.generate_content_stream(
                            model="gemini-3.8-flash",
                            contents=prompt_txt,
                        )
                        for ch_r in stream_r:
                            if ch_r.text:
                                yield ch_r.text
                    except Exception:
                        stream_r = client_r.models.generate_content_stream(
                            model="gemini-2.0-flash",
                            contents=prompt_txt,
                        )
                        for ch_r in stream_r:
                            if ch_r.text:
                                yield ch_r.text

                with st.container():
                    st.markdown("##### 🧠 Veredicto Táctico Ejecutivo (Google AI Studio · Top 3 EV+)")
                    try:
                        st.write_stream(_stream_veredicto_radar(prompt_radar_digest))
                    except Exception as err_rad_ia:
                        st.error(f"Error al generar el veredicto con Google AI Studio: {err_rad_ia}")

        # ── PILAR 5: 6 PESTAÑAS ESPECIALIZADAS CON DESGLOSE GRANULAR ──────────
        tabs_radar = st.tabs([etq for _, etq in CATALOGO_MERCADOS_RADAR])
        for idx_tab, (m_key, _) in enumerate(CATALOGO_MERCADOS_RADAR):
            with tabs_radar[idx_tab]:
                _renderizar_tabla_oscura_radar(resultados_por_mercado[m_key])


elif menu_seleccionado == "📋 Scouting de Jugadores":
    st.title("📋 Scouting de Jugadores Pro")
    if err_jugadores or df_jugadores.empty: st.warning("Datos de jugadores no disponibles.")
    else:
        vista_scout = st.radio("Vista de Datos:", ["🎴 Tarjetas de Perfil Individual", "📊 Tabla Completa BD"], horizontal=True)
        st.markdown("---")
        
        lista_equipos_jug = sorted(df_jugadores['Equipo'].dropna().unique().tolist())
        col_f1, col_f2, col_f3 = st.columns([2, 2, 3])
        with col_f1: equipo_jug_sel = st.selectbox("Equipo:", ["Todos los Equipos"] + lista_equipos_jug)
        with col_f2: pos_sel = st.selectbox("Posición:", ["Todas las Posiciones"] + sorted(df_jugadores['Pos'].dropna().unique().tolist() if 'Pos' in df_jugadores.columns else []))
        with col_f3: busqueda_nombre = st.text_input("Buscar Jugador:", placeholder="Ej: Vinicius...")

        df_display = df_jugadores.copy()
        if equipo_jug_sel != "Todos los Equipos": df_display = df_display[df_display['Equipo'] == equipo_jug_sel]
        if pos_sel != "Todas las Posiciones": df_display = df_display[df_display['Pos'].str.contains(pos_sel, na=False)]
        if busqueda_nombre.strip(): df_display = df_display[df_display['Jugador'].str.contains(busqueda_nombre, case=False, na=False)]

        st.markdown(f"**Registros encontrados:** `{len(df_display)}` futbolistas")
        
        if vista_scout == "🎴 Tarjetas de Perfil Individual":
            if not df_display.empty:
                for idx, row in df_display.head(10).iterrows():
                    nombre, equipo, pos = row.get('Jugador', 'Desconocido'), row.get('Equipo', 'N/A'), row.get('Pos', 'N/A')
                    st.markdown(f"""
                        <div class="metric-card">
                            <h3 style="margin:0; color:#4fe3c1;">👤 {nombre} <span style="font-size:14px; color:#a0aec0;">({pos} - {equipo})</span></h3>
                            <p style="margin:5px 0 0 0;">⚽ <b>Goles Totales:</b> {row.get('Goles', 0)} | 🎯 <b>Asistencias:</b> {row.get('Asistencias', 0)} | 🟨 <b>Tarjetas:</b> {row.get('Tarjetas', 0)}</p>
                            <a href="https://www.google.com/search?q=site:sofascore.com+{nombre.replace(' ', '+')}" target="_blank" style="color:#63b3ed; text-decoration:none;">🔍 Ver perfil en SofaScore ↗️</a>
                        </div>
                    """, unsafe_allow_html=True)
            else: st.info("No hay jugadores que coincidan con los filtros.")
        else:
            if not df_display.empty: 
                df_display['SofaScore'] = df_display['Jugador'].apply(lambda x: f"https://www.google.com/search?q=site:sofascore.com+{x.replace(' ', '+')}")
            st.dataframe(df_display, use_container_width=True, height=400, column_config={"SofaScore": st.column_config.LinkColumn("🔍 Ver en SofaScore", display_text="Abrir SofaScore ↗️")})

elif menu_seleccionado == "🤖 Motor de Apuestas (EV+)":
    st.title("🤖 Motor de Probabilidades: Modelo Poisson Expandido")
    if not df_partidos.empty and 'HomeTeam' in df_partidos.columns:
        equipos_poisson = sorted(pd.concat([df_partidos['HomeTeam'], df_partidos['AwayTeam']]).dropna().unique())
        col_p1, col_p2 = st.columns(2)
        with col_p1: eq_local = st.selectbox("Selecciona Equipo Local:", equipos_poisson, index=0)
        with col_p2: eq_visitante = st.selectbox("Selecciona Equipo Visitante:", equipos_poisson, index=1 if len(equipos_poisson)>1 else 0)

        # Controles de Decaimiento Temporal (Time-Decay Poisson)
        col_td1, col_td2 = st.columns([2, 2])
        with col_td1:
            modo_ponderacion = st.radio(
                "⏳ Modo de Ponderación Temporal:",
                ["🔥 Forma Reciente (Time-Decay Poisson)", "📚 Histórico Completo (Media Clásica)"],
                horizontal=True,
                key="modo_td_motor",
            )
            usar_time_decay = "Time-Decay" in modo_ponderacion
        with col_td2:
            xi_decay = st.slider(
                "📉 Nivel de Ponderación ξ (w = exp(-ξ · Δdías) | 0.010 ≈ últ. 6-8 partidos):",
                min_value=0.001,
                max_value=0.025,
                value=0.010,
                step=0.001,
                format="%.3f",
                disabled=not usar_time_decay,
                key="xi_td_motor",
            )
        
        if eq_local != eq_visitante:
            p_local = df_partidos[df_partidos['HomeTeam'] == eq_local]
            p_vis = df_partidos[df_partidos['AwayTeam'] == eq_visitante]

            # 1) Cálculo original de medias brutas (Histórico Completo)
            xg_l_raw = ((p_local['FTHG'].mean() if not p_local.empty else 1.5) + (p_vis['FTHG'].mean() if not p_vis.empty else 1.5)) / 2
            xg_v_raw = ((p_vis['FTAG'].mean() if not p_vis.empty else 1.0) + (p_local['FTAG'].mean() if not p_local.empty else 1.0)) / 2

            # 2) Cálculo con Decaimiento Temporal (Time-Decay Poisson: w = exp(-xi * delta_dias))
            fecha_ref_global = None
            if "Date" in df_partidos.columns:
                fechas_todas = pd.to_datetime(df_partidos["Date"], errors="coerce", dayfirst=True)
                if not fechas_todas.isna().all():
                    fecha_ref_global = fechas_todas.max()

            if usar_time_decay:
                gf_l_td = calcular_media_ponderada_poisson(p_local, 'FTHG', xi=xi_decay, default=1.5, fecha_ref=fecha_ref_global)
                gc_l_td = calcular_media_ponderada_poisson(p_local, 'FTAG', xi=xi_decay, default=1.0, fecha_ref=fecha_ref_global)
                gf_v_td = calcular_media_ponderada_poisson(p_vis, 'FTAG', xi=xi_decay, default=1.0, fecha_ref=fecha_ref_global)
                gc_v_td = calcular_media_ponderada_poisson(p_vis, 'FTHG', xi=xi_decay, default=1.5, fecha_ref=fecha_ref_global)
                xg_l = (gf_l_td + gc_v_td) / 2
                xg_v = (gf_v_td + gc_l_td) / 2
            else:
                xg_l = xg_l_raw
                xg_v = xg_v_raw
            
            prob_h = prob_d = prob_a = prob_over_15 = prob_over_25 = prob_over_35 = prob_btts_si = tot_prob = 0
            matriz = []
            
            for i in range(7):
                for j in range(7):
                    pr = calcular_probabilidad_poisson(xg_l, i) * calcular_probabilidad_poisson(xg_v, j)
                    tot_prob += pr
                    if i > j: prob_h += pr
                    elif i == j: prob_d += pr
                    else: prob_a += pr
                    
                    if (i + j) > 1.5: prob_over_15 += pr
                    if (i + j) > 2.5: prob_over_25 += pr
                    if (i + j) > 3.5: prob_over_35 += pr
                    if i > 0 and j > 0: prob_btts_si += pr
                    matriz.append({"Marcador": f"{i} - {j}", "Probabilidad": pr})
            
            norm = tot_prob if tot_prob > 0 else 1.0
            
            st.markdown("---")
            st.markdown(f"### ⚽ xG: **{eq_local} ({xg_l:.2f})** vs **{eq_visitante} ({xg_v:.2f})**")
            if usar_time_decay:
                st.caption(
                    f"⏳ **Time-Decay activo (ξ = {xi_decay:.3f}):** "
                    f"xG Forma Reciente `{xg_l:.2f} - {xg_v:.2f}`  |  "
                    f"xG Histórico Completo sin ponderar: `{xg_l_raw:.2f} - {xg_v_raw:.2f}`"
                )
            st.subheader("1️⃣ Mercado Principal (1X2)")
            c_res1, c_res2, c_res3 = st.columns(3)
            c_res1.metric(f"Local", f"Cuota {1/(prob_h/norm):.2f}", f"{(prob_h/norm*100):.1f}%")
            c_res2.metric("Empate", f"Cuota {1/(prob_d/norm):.2f}", f"{(prob_d/norm*100):.1f}%")
            c_res3.metric(f"Visitante", f"Cuota {1/(prob_a/norm):.2f}", f"{(prob_a/norm*100):.1f}%")
            
            p_tab1, p_tab2, p_tab3 = st.tabs(["⚽ Goles (Over/Under)", "🎯 BTTS", "🏆 Marcadores"])
            with p_tab1:
                col_g1, col_g2, col_g3 = st.columns(3)
                col_g1.metric("+1.5 Goles", f"Cuota {1/(prob_over_15/norm):.2f}", f"{(prob_over_15/norm*100):.1f}%")
                col_g2.metric("+2.5 Goles", f"Cuota {1/(prob_over_25/norm):.2f}", f"{(prob_over_25/norm*100):.1f}%")
                col_g3.metric("+3.5 Goles", f"Cuota {1/(prob_over_35/norm):.2f}", f"{(prob_over_35/norm*100):.1f}%")
            with p_tab2:
                col_b1, col_b2 = st.columns(2)
                col_b1.metric("SÍ", f"Cuota {1/(prob_btts_si/norm):.2f}", f"{(prob_btts_si/norm*100):.1f}%")
                col_b2.metric("NO", f"Cuota {1/(1-(prob_btts_si/norm)):.2f}", f"{((1-(prob_btts_si/norm))*100):.1f}%")
            with p_tab3:
                df_m = pd.DataFrame(matriz)
                df_m['Probabilidad'] = df_m['Probabilidad'] / norm
                df_m = df_m.sort_values(by="Probabilidad", ascending=False).head(5)
                df_m['Prob_Pct'] = df_m['Probabilidad'].apply(lambda x: f"{x*100:.1f}%")
                df_m['Cuota_Justa'] = df_m['Probabilidad'].apply(lambda x: round(1/x, 2) if x>0 else 0)
                st.dataframe(df_m[['Marcador', 'Prob_Pct', 'Cuota_Justa']], use_container_width=True, hide_index=True)

elif menu_seleccionado == "⚖️ Directorio y Matchup Árbitros":
    st.title("⚖️ Directorio Profesional y Matchup de Árbitros")
    
    dict_arbitros = {
        "LaLiga (España)": [
            {"Árbitro": "Jesús Gil Manzano", "Comité": "Extremeño", "Perfil Disciplinario": "Alto (Mano dura / Tarjetero)", "Media Amarillas Est.": 5.40, "Media Rojas": 0.35},
            {"Árbitro": "José María Sánchez Martínez", "Comité": "Murciano", "Perfil Disciplinario": "Medio-Alto", "Media Amarillas Est.": 5.20, "Media Rojas": 0.28},
            {"Árbitro": "Alejandro Hernández Hernández", "Comité": "Las Palmas", "Perfil Disciplinario": "Alto (Riguroso en faltas)", "Media Amarillas Est.": 5.20, "Media Rojas": 0.31},
            {"Árbitro": "César Soto Grado", "Comité": "Riojano", "Perfil Disciplinario": "Medio", "Media Amarillas Est.": 4.80, "Media Rojas": 0.22},
            {"Árbitro": "Juan Martínez Munuera", "Comité": "Valenciano", "Perfil Disciplinario": "Bajo-Medio (Permisivo)", "Media Amarillas Est.": 4.30, "Media Rojas": 0.18},
            {"Árbitro": "Guillermo Cuadra Fernández", "Comité": "Balear", "Perfil Disciplinario": "Alto", "Media Amarillas Est.": 5.10, "Media Rojas": 0.29},
            {"Árbitro": "Javier Alberola Rojas", "Comité": "Castellano-Manchego", "Perfil Disciplinario": "Muy Permisivo", "Media Amarillas Est.": 3.60, "Media Rojas": 0.10},
            {"Árbitro": "Ricardo de Burgos Bengoetxea", "Comité": "Vasco", "Perfil Disciplinario": "Medio", "Media Amarillas Est.": 4.60, "Media Rojas": 0.20},
            {"Árbitro": "Mario Melero López", "Comité": "Andaluz", "Perfil Disciplinario": "Medio-Alto", "Media Amarillas Est.": 4.90, "Media Rojas": 0.24},
            {"Árbitro": "Isidro Díaz de Mera Escuderos", "Comité": "Castellano-Manchego", "Perfil Disciplinario": "Alto", "Media Amarillas Est.": 5.30, "Media Rojas": 0.33}
        ],
        "Premier League (Inglaterra)": [
            {"Árbitro": "Michael Oliver", "Comité": "PGMOL Elite", "Perfil Disciplinario": "Medio (De Lujo)", "Media Amarillas Est.": 3.50, "Media Rojas": 0.15},
            {"Árbitro": "Anthony Taylor", "Comité": "PGMOL Elite", "Perfil Disciplinario": "Medio-Alto", "Media Amarillas Est.": 4.00, "Media Rojas": 0.20},
            {"Árbitro": "Paul Tierney", "Comité": "PGMOL", "Perfil Disciplinario": "Medio", "Media Amarillas Est.": 3.80, "Media Rojas": 0.12},
            {"Árbitro": "Simon Hooper", "Comité": "PGMOL", "Perfil Disciplinario": "Permisivo", "Media Amarillas Est.": 3.20, "Media Rojas": 0.08},
            {"Árbitro": "Jarred Gillett", "Comité": "PGMOL", "Perfil Disciplinario": "Medio-Alto", "Media Amarillas Est.": 4.10, "Media Rojas": 0.18},
            {"Árbitro": "Craig Pawson", "Comité": "PGMOL", "Perfil Disciplinario": "Alto", "Media Amarillas Est.": 4.30, "Media Rojas": 0.22}
        ],
        "Serie A (Italia)": [
            {"Árbitro": "Daniele Orsato", "Comité": "AIA Elite", "Perfil Disciplinario": "Alto", "Media Amarillas Est.": 5.10, "Media Rojas": 0.25},
            {"Árbitro": "Davide Massa", "Comité": "AIA", "Perfil Disciplinario": "Medio-Alto", "Media Amarillas Est.": 4.70, "Media Rojas": 0.21},
            {"Árbitro": "Maurizio Mariani", "Comité": "AIA", "Perfil Disciplinario": "Alto", "Media Amarillas Est.": 5.00, "Media Rojas": 0.27}
        ],
        "Bundesliga (Alemania)": [
            {"Árbitro": "Felix Zwayer", "Comité": "DFB Elite", "Perfil Disciplinario": "Medio", "Media Amarillas Est.": 4.10, "Media Rojas": 0.18},
            {"Árbitro": "Daniel Siebert", "Comité": "DFB Elite", "Perfil Disciplinario": "Medio-Bajo", "Media Amarillas Est.": 3.90, "Media Rojas": 0.14},
            {"Árbitro": "Deniz Aytekin", "Comité": "DFB", "Perfil Disciplinario": "Muy Permisivo / Fluido", "Media Amarillas Est.": 3.40, "Media Rojas": 0.11}
        ]
    }

    tab_ref1, tab_ref2 = st.tabs(["📖 Directorio Completo", "🟨 Matchup Disciplinario (Over/Under Tarjetas)"])
    
    with tab_ref1:
        col_a1, col_a2, col_a3 = st.columns([2, 2, 1])
        with col_a1:
            liga_ref_sel = st.selectbox(
                "Selecciona Competición:",
                [
                    "LaLiga (España)",
                    "Premier League (Inglaterra)",
                    "Serie A (Italia)",
                    "Bundesliga (Alemania)",
                    "Ligue 1 (Francia)",
                    "Todas las Competiciones",
                ],
            )
        with col_a2:
            buscar_ref = st.text_input("🔍 Buscar Árbitro por nombre:", placeholder="Ej: Oliver, Taylor, Gil Manzano...")
        with col_a3:
            min_partidos_ref = st.number_input("Mín. Partidos:", min_value=1, max_value=30, value=1, step=1)

        df_arbitros_bd = cargar_datos_arbitros(min_partidos=min_partidos_ref)

        if not df_arbitros_bd.empty:
            if liga_ref_sel != "Todas las Competiciones":
                df_ref = df_arbitros_bd[df_arbitros_bd["Liga"] == liga_ref_sel].copy()
            else:
                df_ref = df_arbitros_bd.copy()

            cols_mostrar_ref = [
                c for c in [
                    "Árbitro", "Comité", "Partidos", "Perfil Disciplinario",
                    "Media Amarillas Est.", "Media Rojas", "Media Faltas"
                ] if c in df_ref.columns
            ]
            df_ref = df_ref[cols_mostrar_ref]
        else:
            # Fallback al diccionario manual si la tabla arbitros_stats estuviera vacía
            df_ref = pd.DataFrame(dict_arbitros.get(liga_ref_sel, []))

        if buscar_ref.strip() and not df_ref.empty:
            df_ref = df_ref[df_ref['Árbitro'].str.contains(buscar_ref, case=False, na=False)]

        st.caption(f"📋 Mostrando **{len(df_ref)} colegiados** indexados en `arbitros_stats` (mínimo {min_partidos_ref} partidos pitados).")
        st.dataframe(df_ref, use_container_width=True, hide_index=True)
        
    with tab_ref2:
        st.markdown("Calcula la línea esperada de tarjetas combinando los promedios del colegiado con la tendencia tarjetera de ambos equipos.")
        if not (err_partidos or df_partidos_global.empty):
            render_matchup_disciplinario(df_partidos_global)
            st.markdown("---")
            st.caption("📎 Estimador simple de referencia (modelo anterior):")
        if err_partidos or df_partidos_global.empty:
            st.warning("Datos de partidos no disponibles para calcular tendencias de equipos.")
        else:
            df_arbs_matchup = cargar_datos_arbitros(min_partidos=1)
            todos_arbs = []
            mapa_media_arbs = {}

            if not df_arbs_matchup.empty:
                for _, r_arb in df_arbs_matchup.iterrows():
                    etiqueta = f"{r_arb['Árbitro']} ({r_arb['Liga']})"
                    todos_arbs.append(etiqueta)
                    mapa_media_arbs[r_arb['Árbitro']] = float(r_arb['Media Amarillas Est.'])
            else:
                for liga, arbs in dict_arbitros.items():
                    for a in arbs:
                        todos_arbs.append(f"{a['Árbitro']} ({liga})")
                        mapa_media_arbs[a['Árbitro']] = float(a['Media Amarillas Est.'])
            
            eqs_tarj = sorted(pd.concat([df_partidos_global['HomeTeam'], df_partidos_global['AwayTeam']]).dropna().unique())
            
            c_m1, c_m2, c_m3 = st.columns(3)
            with c_m1: arb_sel = st.selectbox("⚖️ Selecciona Árbitro:", todos_arbs)
            with c_m2: eq_t1 = st.selectbox("🏠 Equipo Local:", eqs_tarj, index=0)
            with c_m3: eq_t2 = st.selectbox("✈️ Equipo Visitante:", eqs_tarj, index=1 if len(eqs_tarj)>1 else 0)
            
            ref_name = arb_sel.split(" (")[0]
            media_ref = mapa_media_arbs.get(ref_name, 4.5)
            
            def stats_tarjetas(eq):
                df_h = df_partidos_global[df_partidos_global['HomeTeam'] == eq]
                df_a = df_partidos_global[df_partidos_global['AwayTeam'] == eq]
                if 'HY' in df_partidos_global.columns and 'AY' in df_partidos_global.columns:
                    t = df_h['HY'].sum() + df_a['AY'].sum()
                    partidos = len(df_h) + len(df_a)
                    return t / partidos if partidos > 0 else 2.5
                return 2.5
                
            t1_avg = stats_tarjetas(eq_t1)
            t2_avg = stats_tarjetas(eq_t2)
            total_esperado = (t1_avg + t2_avg + media_ref) / 2
            
            st.markdown("---")
            c_r1, c_r2, c_r3 = st.columns(3)
            c_r1.metric(f"Media Árbitro ({ref_name})", f"{media_ref:.2f}", "Amarillas/Partido")
            c_r2.metric(f"Provocadas {eq_t1}", f"{t1_avg:.2f}", "Amarillas/Partido", delta_color="off")
            c_r3.metric(f"Provocadas {eq_t2}", f"{t2_avg:.2f}", "Amarillas/Partido", delta_color="off")
            
            st.success(f"🔥 Línea de Tarjetas Amarillas Proyectada: **{total_esperado:.1f}**")

elif menu_seleccionado == "🧮 Hándicap Asiático Avanzado":
    st.title("🧮 Motor Cuantitativo: Hándicap Asiático & Líneas de Goles")
    st.markdown(
        "Modelado predictivo **Bivariante de Poisson con ajuste de Dixon-Coles (ρ = -0.11)**, "
        "proyección de **Fair Spread**, cobertura histórica **ATS** y valoración **+EV / 1/4 Kelly**."
    )
    
    LINEAS_AH_EQUIPOS = [
        -1.50, -1.25, -1.00, -0.75, -0.50, -0.25,
        0.00,
        0.25, 0.50, 0.75, 1.00, 1.25, 1.50,
    ]
    LINEAS_AH_GOLES = [2.00, 2.25, 2.50, 2.75, 3.00]
    RHO_DIXON_COLES = -0.11
    VENTANA_HISTORICA_AH = 15

    def _calcular_matriz_dixon_coles(lam_h, lam_a, rho=RHO_DIXON_COLES, max_goles=7):
        n = max_goles + 1
        matriz = np.zeros((n, n), dtype=float)
        lam_h_safe = max(0.05, float(lam_h))
        lam_a_safe = max(0.05, float(lam_a))

        for i in range(n):
            p_i = calcular_probabilidad_poisson(lam_h_safe, i)
            for j in range(n):
                p_j = calcular_probabilidad_poisson(lam_a_safe, j)
                matriz[i, j] = p_i * p_j

        matriz[0, 0] *= max(0.0, 1.0 - (lam_h_safe * lam_a_safe * rho))
        matriz[1, 0] *= max(0.0, 1.0 + (lam_a_safe * rho))
        matriz[0, 1] *= max(0.0, 1.0 + (lam_h_safe * rho))
        matriz[1, 1] *= max(0.0, 1.0 - rho)

        total = float(np.sum(matriz))
        if total > 0.0:
            matriz /= total
        else:
            matriz[0, 0] = 1.0
        return matriz

    def _liquidar_estado_unidad(diff):
        if abs(diff) < 1e-9:
            return "PUSH"
        return "WIN" if diff > 0.0 else "LOSS"

    def _resolver_cuota_asiatica(p_full_win, p_half_win, p_push, p_half_loss, p_full_loss):
        denominador = float(p_full_win) + 0.5 * float(p_half_win)
        numerador = 1.0 - float(p_push) - 0.5 * float(p_half_win) - 0.5 * float(p_half_loss)
        if denominador <= 1e-9 or numerador <= 1e-9:
            return float("inf"), 0.0
        cuota_justa = numerador / denominador
        if cuota_justa <= 1.0:
            return 1.0, 1.0
        prob_equivalente = min(1.0, max(0.0, 1.0 / cuota_justa))
        return cuota_justa, prob_equivalente

    def _evaluar_linea_handicap(matriz, linea, perspectiva="local"):
        es_cuarto = abs((abs(linea) * 4.0) % 2.0 - 1.0) < 1e-6
        if es_cuarto:
            l1, l2 = linea - 0.25, linea + 0.25
        else:
            l1, l2 = linea, linea

        p_fw = p_hw = p_push = p_hl = p_fl = 0.0
        n = matriz.shape[0]
        for i in range(n):
            for j in range(n):
                prob_ij = float(matriz[i, j])
                margen = (i - j) if perspectiva == "local" else (j - i)
                s1 = _liquidar_estado_unidad(margen + l1)
                s2 = _liquidar_estado_unidad(margen + l2)

                if s1 == "WIN" and s2 == "WIN":
                    p_fw += prob_ij
                elif (s1 == "WIN" and s2 == "PUSH") or (s1 == "PUSH" and s2 == "WIN"):
                    p_hw += prob_ij
                elif s1 == "PUSH" and s2 == "PUSH":
                    p_push += prob_ij
                elif (s1 == "LOSS" and s2 == "PUSH") or (s1 == "PUSH" and s2 == "LOSS"):
                    p_hl += prob_ij
                else:
                    p_fl += prob_ij

        cuota_justa, prob_equiv = _resolver_cuota_asiatica(p_fw, p_hw, p_push, p_hl, p_fl)
        return {
            "linea": linea,
            "p_full_win": p_fw,
            "p_half_win": p_hw,
            "p_push": p_push,
            "p_half_loss": p_hl,
            "p_full_loss": p_fl,
            "p_acierto": p_fw + 0.5 * p_hw,
            "p_reembolso_medio": p_hw + p_push + p_hl,
            "cuota_justa": cuota_justa,
            "prob_equiv": prob_equiv,
        }

    def _evaluar_linea_totales_asiaticos(matriz, linea_t):
        es_cuarto = abs((abs(linea_t) * 4.0) % 2.0 - 1.0) < 1e-6
        t1, t2 = (linea_t - 0.25, linea_t + 0.25) if es_cuarto else (linea_t, linea_t)

        o_fw = o_hw = o_push = o_hl = o_fl = 0.0
        u_fw = u_hw = u_push = u_hl = u_fl = 0.0

        n = matriz.shape[0]
        for i in range(n):
            for j in range(n):
                prob_ij = float(matriz[i, j])
                goles = i + j

                so1 = _liquidar_estado_unidad(goles - t1)
                so2 = _liquidar_estado_unidad(goles - t2)
                if so1 == "WIN" and so2 == "WIN":
                    o_fw += prob_ij
                elif (so1 == "WIN" and so2 == "PUSH") or (so1 == "PUSH" and so2 == "WIN"):
                    o_hw += prob_ij
                elif so1 == "PUSH" and so2 == "PUSH":
                    o_push += prob_ij
                elif (so1 == "LOSS" and so2 == "PUSH") or (so1 == "PUSH" and so2 == "LOSS"):
                    o_hl += prob_ij
                else:
                    o_fl += prob_ij

                su1 = _liquidar_estado_unidad(t1 - goles)
                su2 = _liquidar_estado_unidad(t2 - goles)
                if su1 == "WIN" and su2 == "WIN":
                    u_fw += prob_ij
                elif (su1 == "WIN" and su2 == "PUSH") or (su1 == "PUSH" and su2 == "WIN"):
                    u_hw += prob_ij
                elif su1 == "PUSH" and su2 == "PUSH":
                    u_push += prob_ij
                elif (su1 == "LOSS" and su2 == "PUSH") or (su1 == "PUSH" and su2 == "LOSS"):
                    u_hl += prob_ij
                else:
                    u_fl += prob_ij

        cuota_over, prob_eq_over = _resolver_cuota_asiatica(o_fw, o_hw, o_push, o_hl, o_fl)
        cuota_under, prob_eq_under = _resolver_cuota_asiatica(u_fw, u_hw, u_push, u_hl, u_fl)
        return {
            "linea": linea_t,
            "over_fw": o_fw,
            "over_hw": o_hw,
            "over_push": o_push,
            "over_hl": o_hl,
            "over_fl": o_fl,
            "over_acierto": o_fw + 0.5 * o_hw,
            "over_reembolso_medio": o_hw + o_push + o_hl,
            "cuota_over": cuota_over,
            "prob_eq_over": prob_eq_over,
            "under_fw": u_fw,
            "under_hw": u_hw,
            "under_push": u_push,
            "under_hl": u_hl,
            "under_fl": u_fl,
            "under_acierto": u_fw + 0.5 * u_hw,
            "under_reembolso_medio": u_hw + u_push + u_hl,
            "cuota_under": cuota_under,
            "prob_eq_under": prob_eq_under,
        }

    def _calcular_ev_y_kelly_fraccional(prob_equiv, cuota_bookie, fraccion=0.25):
        if prob_equiv <= 0.0 or cuota_bookie <= 1.0 or math.isinf(cuota_bookie):
            return -1.0, 0.0
        ev = (prob_equiv * cuota_bookie) - 1.0
        b = cuota_bookie - 1.0
        kelly_raw = ((b * prob_equiv) - (1.0 - prob_equiv)) / b
        kelly_frac = max(0.0, kelly_raw) * fraccion
        return ev, kelly_frac

    def _calcular_cobertura_ats(df_muestra, es_local=True, spread_fallback=-0.25):
        if df_muestra is None or df_muestra.empty:
            return 0.0, 0, "0W-0P-0L"

        cubiertos = 0.0
        evaluados = 0
        v_cnt = n_cnt = r_cnt = 0

        for _, fila in df_muestra.iterrows():
            fthg = pd.to_numeric(fila.get("FTHG"), errors="coerce")
            ftag = pd.to_numeric(fila.get("FTAG"), errors="coerce")
            if pd.isna(fthg) or pd.isna(ftag):
                continue

            ah_val = pd.to_numeric(fila.get("AHh"), errors="coerce")
            if pd.isna(ah_val):
                ah_val = pd.to_numeric(fila.get("AHCh"), errors="coerce")
            if pd.isna(ah_val):
                ah_val = float(spread_fallback)

            if es_local:
                diff = (float(fthg) + float(ah_val)) - float(ftag)
            else:
                diff = (float(ftag) - float(ah_val)) - float(fthg)

            if diff > 0.25:
                cubiertos += 1.0
                evaluados += 1
                v_cnt += 1
            elif abs(diff - 0.25) < 1e-6:
                cubiertos += 0.5
                evaluados += 1
                v_cnt += 1
            elif abs(diff) < 1e-6:
                n_cnt += 1
            elif abs(diff + 0.25) < 1e-6:
                evaluados += 1
                r_cnt += 1
            else:
                evaluados += 1
                r_cnt += 1

        denom = evaluados if evaluados > 0 else (v_cnt + n_cnt + r_cnt)
        pct = (cubiertos / denom * 100.0) if denom > 0 else 0.0
        resumen = f"{v_cnt}W-{n_cnt}P-{r_cnt}L"
        return pct, len(df_muestra), resumen

    # ── LAYER 1: SELECTORES EN CASCADA (LIGA -> LOCAL -> VISITANTE) ──────────
    if err_partidos or df_partidos_global.empty:
        st.warning("No hay datos históricos disponibles en `macro_base_futbol.db`.")
    else:
        col_liga_ah = _columna_liga_partidos(df_partidos_global)
        if col_liga_ah is not None:
            divs_ah = [
                d for d in MAPA_LIGAS
                if d in set(df_partidos_global[col_liga_ah].dropna().astype(str))
            ]
        else:
            divs_ah = []

        ligas_ah_nombres = [MAPA_LIGAS[d] for d in divs_ah] if divs_ah else ["Todas"]

        sel_c1, sel_c2, sel_c3 = st.columns(3)
        with sel_c1:
            liga_ah_sel = st.selectbox("🏆 Competición (Liga):", ligas_ah_nombres, key="ah_engine_liga")

        if col_liga_ah is not None and liga_ah_sel in MAPA_INV:
            div_ah_code = MAPA_INV[liga_ah_sel]
            df_liga_ah = df_partidos_global[df_partidos_global[col_liga_ah].astype(str) == div_ah_code].copy()
        else:
            div_ah_code = "ALL"
            df_liga_ah = df_partidos_global.copy()

        equipos_ah_liga = sorted(
            pd.concat([df_liga_ah["HomeTeam"], df_liga_ah["AwayTeam"]]).dropna().unique().tolist()
        )

        if len(equipos_ah_liga) < 2:
            st.warning("No hay suficientes equipos registrados en esta competición.")
        else:
            with sel_c2:
                eq_local_ah = st.selectbox(
                    "🏠 Equipo Local:",
                    equipos_ah_liga,
                    index=0,
                    key=f"ah_engine_home_{div_ah_code}",
                )
            equipos_visit_ah = [e for e in equipos_ah_liga if e != eq_local_ah]
            with sel_c3:
                eq_visit_ah = st.selectbox(
                    "✈️ Equipo Visitante:",
                    equipos_visit_ah,
                    index=0,
                    key=f"ah_engine_away_{div_ah_code}",
                )

            df_liga_ord = _ordenar_por_fecha(df_liga_ah)
            df_home_cond = df_liga_ord[df_liga_ord["HomeTeam"] == eq_local_ah].dropna(subset=["FTHG", "FTAG"]).tail(VENTANA_HISTORICA_AH).copy()
            df_away_cond = df_liga_ord[df_liga_ord["AwayTeam"] == eq_visit_ah].dropna(subset=["FTHG", "FTAG"]).tail(VENTANA_HISTORICA_AH).copy()

            media_liga_gf_h = float(pd.to_numeric(df_liga_ord["FTHG"], errors="coerce").mean() or 1.45)
            media_liga_gf_a = float(pd.to_numeric(df_liga_ord["FTAG"], errors="coerce").mean() or 1.15)
            if media_liga_gf_h <= 0:
                media_liga_gf_h = 1.45
            if media_liga_gf_a <= 0:
                media_liga_gf_a = 1.15

            gf_local_med = float(pd.to_numeric(df_home_cond["FTHG"], errors="coerce").mean()) if not df_home_cond.empty else media_liga_gf_h
            gc_local_med = float(pd.to_numeric(df_home_cond["FTAG"], errors="coerce").mean()) if not df_home_cond.empty else media_liga_gf_a
            gf_visit_med = float(pd.to_numeric(df_away_cond["FTAG"], errors="coerce").mean()) if not df_away_cond.empty else media_liga_gf_a
            gc_visit_med = float(pd.to_numeric(df_away_cond["FTHG"], errors="coerce").mean()) if not df_away_cond.empty else media_liga_gf_h

            atk_local = gf_local_med / media_liga_gf_h
            def_local = gc_local_med / media_liga_gf_a
            atk_visit = gf_visit_med / media_liga_gf_a
            def_visit = gc_visit_med / media_liga_gf_h

            lambda_h = float(np.clip(atk_local * def_visit * media_liga_gf_h, 0.20, 4.50))
            lambda_a = float(np.clip(atk_visit * def_local * media_liga_gf_a, 0.20, 4.50))

            fair_spread = lambda_a - lambda_h
            linea_mercado_aprox = round(fair_spread * 4.0) / 4.0
            if fair_spread < -0.15:
                inclinacion_badge = f"🟢 Favorito Local (Línea ref. {linea_mercado_aprox:+.2f})"
            elif fair_spread > 0.15:
                inclinacion_badge = f"🔵 Favorito Visitante (Línea ref. {linea_mercado_aprox:+.2f})"
            else:
                inclinacion_badge = "⚖️ Duelo Equilibrado (DNB / 0.00)"

            ats_home_pct, n_home_ats, res_home_ats = _calcular_cobertura_ats(
                df_home_cond, es_local=True, spread_fallback=linea_mercado_aprox
            )
            ats_away_pct, n_away_ats, res_away_ats = _calcular_cobertura_ats(
                df_away_cond, es_local=False, spread_fallback=-linea_mercado_aprox
            )

            # ── LAYER 3.1: EXECUTIVE KPI METRIC CARDS (st.columns(4)) ─────────
            st.markdown("---")
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric(
                f"Lambda Local ({eq_local_ah})",
                f"{lambda_h:.2f} xG",
                f"Atk {atk_local:.2f}x · Def {def_local:.2f}x ({n_home_ats}p Casa)",
                delta_color="off",
            )
            kpi2.metric(
                f"Lambda Visitante ({eq_visit_ah})",
                f"{lambda_a:.2f} xG",
                f"Atk {atk_visit:.2f}x · Def {def_visit:.2f}x ({n_away_ats}p Fuera)",
                delta_color="off",
            )
            kpi3.metric(
                "Fair Spread del Modelo",
                f"Local {fair_spread:+.2f}",
                inclinacion_badge,
                delta_color="off",
            )
            kpi4.metric(
                "Cobertura ATS Reciente",
                f"{ats_home_pct:.1f}% vs {ats_away_pct:.1f}%",
                f"Casa ({res_home_ats}) · Fuera ({res_away_ats})",
                delta_color="off",
            )

            matriz_dc = _calcular_matriz_dixon_coles(lambda_h, lambda_a, rho=RHO_DIXON_COLES, max_goles=7)

            # ── LAYER 3.2: INTERACTIVE PRICING MATRIX TABS ───────────────────
            st.markdown("---")
            tab_ah_equipos, tab_ah_totales = st.tabs([
                "Matriz Hándicap Asiático (Equipos)",
                "Totales Asiáticos de Goles (Over/Under)",
            ])

            with tab_ah_equipos:
                st.subheader("📊 Matriz Hándicap Asiático (Equipos)")
                ctrl1, ctrl2, ctrl3 = st.columns([1.6, 1.2, 1.2])
                with ctrl1:
                    lado_ah_equipo = st.radio(
                        "Equipo Evaluado en la Línea:",
                        [f"🏠 Local ({eq_local_ah})", f"✈️ Visitante ({eq_visit_ah})"],
                        horizontal=True,
                        key="ah_perspectiva_equipo",
                    )
                with ctrl2:
                    cuota_bookie_ah = st.number_input(
                        "Input Cuota Bookie:",
                        min_value=1.01,
                        max_value=25.00,
                        value=1.90,
                        step=0.05,
                        key="ah_cuota_bookie_equipos",
                    )
                with ctrl3:
                    bankroll_ah = st.number_input(
                        "Bankroll Operativo (€):",
                        min_value=10.0,
                        max_value=1000000.0,
                        value=1000.0,
                        step=50.0,
                        key="ah_bankroll_equipos",
                    )

                perspectiva_key = "local" if "Local" in lado_ah_equipo else "visitante"
                nombre_eq_activo = eq_local_ah if perspectiva_key == "local" else eq_visit_ah

                filas_matriz_ah = []
                detalle_lineas_ah = {}
                for lin in LINEAS_AH_EQUIPOS:
                    info_l = _evaluar_linea_handicap(matriz_dc, lin, perspectiva=perspectiva_key)
                    detalle_lineas_ah[lin] = info_l
                    ev_l, kelly_l = _calcular_ev_y_kelly_fraccional(
                        info_l["prob_equiv"], cuota_bookie_ah, fraccion=0.25
                    )
                    cuota_str = "-" if math.isinf(info_l["cuota_justa"]) else f"{info_l['cuota_justa']:.2f}"
                    stake_str = (
                        f"{kelly_l * 100:.2f}% ({kelly_l * bankroll_ah:.2f} €)"
                        if ev_l > 0.0
                        else "0.00% (No Bet)"
                    )
                    filas_matriz_ah.append({
                        "Línea": f"{nombre_eq_activo} {lin:+.2f}",
                        "Prob. Acierto (%)": f"{info_l['p_acierto'] * 100:.1f}%",
                        "Prob. Reembolso/Medio (%)": f"{info_l['p_reembolso_medio'] * 100:.1f}%",
                        "Cuota Justa Modelo": cuota_str,
                        "Input Cuota Bookie": f"{cuota_bookie_ah:.2f}",
                        "EV (%)": f"{ev_l * 100:+.2f}%",
                        "Stake Sugerido (1/4 Kelly)": stake_str,
                    })

                st.dataframe(pd.DataFrame(filas_matriz_ah), use_container_width=True, hide_index=True)

                st.markdown("##### 🎯 Desglose de Liquidación por Línea")
                insp_c1, insp_c2 = st.columns([1, 3])
                with insp_c1:
                    lin_inspeccion = st.selectbox(
                        "Inspeccionar Línea:",
                        LINEAS_AH_EQUIPOS,
                        index=LINEAS_AH_EQUIPOS.index(-0.25),
                        format_func=lambda x: f"{nombre_eq_activo} {x:+.2f}",
                        key="ah_linea_inspeccion",
                    )
                info_sel = detalle_lineas_ah[lin_inspeccion]
                with insp_c2:
                    d1, d2, d3, d4, d5 = st.columns(5)
                    d1.metric("Full Win", f"{info_sel['p_full_win'] * 100:.1f}%")
                    d2.metric("Half Win", f"{info_sel['p_half_win'] * 100:.1f}%")
                    d3.metric("Push (Nulo)", f"{info_sel['p_push'] * 100:.1f}%")
                    d4.metric("Half Loss", f"{info_sel['p_half_loss'] * 100:.1f}%")
                    d5.metric("Full Loss", f"{info_sel['p_full_loss'] * 100:.1f}%")

            with tab_ah_totales:
                st.subheader("⚽ Totales Asiáticos de Goles (Over/Under)")
                st.caption(
                    f"Intensidad conjunta proyectada: **{lambda_h + lambda_a:.2f} goles esperados** "
                    f"(ajuste Dixon-Coles ρ = {RHO_DIXON_COLES})."
                )

                filas_totales = []
                mapa_totales = {}
                for lin_t in LINEAS_AH_GOLES:
                    res_t = _evaluar_linea_totales_asiaticos(matriz_dc, lin_t)
                    mapa_totales[lin_t] = res_t
                    c_over_str = "-" if math.isinf(res_t["cuota_over"]) else f"{res_t['cuota_over']:.2f}"
                    c_under_str = "-" if math.isinf(res_t["cuota_under"]) else f"{res_t['cuota_under']:.2f}"
                    filas_totales.append({
                        "Línea": f"{lin_t:.2f}",
                        "Prob. Over (%)": f"{res_t['over_acierto'] * 100:.1f}%",
                        "Reembolso/Medio Over (%)": f"{res_t['over_reembolso_medio'] * 100:.1f}%",
                        "Cuota Justa Over": c_over_str,
                        "Prob. Under (%)": f"{res_t['under_acierto'] * 100:.1f}%",
                        "Reembolso/Medio Under (%)": f"{res_t['under_reembolso_medio'] * 100:.1f}%",
                        "Cuota Justa Under": c_under_str,
                    })

                st.dataframe(pd.DataFrame(filas_totales), use_container_width=True, hide_index=True)

                st.markdown("#### 💰 Evaluador Instantáneo EV & Kelly (Totales Asiáticos)")
                ev_t1, ev_t2, ev_t3, ev_t4 = st.columns(4)
                with ev_t1:
                    linea_tot_sel = st.selectbox(
                        "Línea de Goles:",
                        LINEAS_AH_GOLES,
                        index=2,
                        format_func=lambda x: f"{x:.2f} Goles",
                        key="ah_tot_linea_sel",
                    )
                with ev_t2:
                    mercado_tot_sel = st.radio(
                        "Selección:",
                        ["Over (Más de)", "Under (Menos de)"],
                        horizontal=True,
                        key="ah_tot_mercado_sel",
                    )
                with ev_t3:
                    cuota_bookie_tot = st.number_input(
                        "Cuota Bookie (Totales):",
                        min_value=1.01,
                        max_value=25.00,
                        value=1.90,
                        step=0.05,
                        key="ah_tot_cuota_bookie",
                    )
                with ev_t4:
                    bankroll_tot = st.number_input(
                        "Bankroll (€):",
                        min_value=10.0,
                        max_value=1000000.0,
                        value=1000.0,
                        step=50.0,
                        key="ah_tot_bankroll",
                    )

                datos_t_sel = mapa_totales[linea_tot_sel]
                es_over_sel = "Over" in mercado_tot_sel
                cuota_justa_t = datos_t_sel["cuota_over"] if es_over_sel else datos_t_sel["cuota_under"]
                prob_eq_t = datos_t_sel["prob_eq_over"] if es_over_sel else datos_t_sel["prob_eq_under"]
                ev_tot, kelly_tot = _calcular_ev_y_kelly_fraccional(prob_eq_t, cuota_bookie_tot, fraccion=0.25)

                m_t1, m_t2, m_t3, m_t4 = st.columns(4)
                m_t1.metric("Prob. Equivalente Modelo", f"{prob_eq_t * 100:.1f}%")
                m_t2.metric(
                    "Cuota Justa Modelo",
                    "-" if math.isinf(cuota_justa_t) else f"{cuota_justa_t:.2f}",
                )
                m_t3.metric("EV (%)", f"{ev_tot * 100:+.2f}%")
                m_t4.metric(
                    "Stake Sugerido (1/4 Kelly)",
                    f"{kelly_tot * 100:.2f}% ({kelly_tot * bankroll_tot:.2f} €)" if ev_tot > 0 else "0.00% (0.00 €)",
                )

                if ev_tot > 0:
                    st.success(
                        f"✅ **Valor Detectado (+EV):** {mercado_tot_sel} {linea_tot_sel:.2f} @ {cuota_bookie_tot:.2f} "
                        f"(Cuota Justa: {cuota_justa_t:.2f}). Stake 1/4 Kelly: **{kelly_tot * 100:.2f}%** "
                        f"del bankroll (**{kelly_tot * bankroll_tot:.2f} €**)."
                    )
                else:
                    st.info(
                        f"ℹ️ Sin ventaja matemática en {mercado_tot_sel} {linea_tot_sel:.2f} @ {cuota_bookie_tot:.2f} "
                        f"(EV: {ev_tot * 100:+.2f}%). Cuota mínima requerida: **{cuota_justa_t:.2f}**."
                    )

            # ── LAYER 3.3: EDUCATIONAL FOOTNOTE ──────────────────────────────
            with st.expander("📚 Glosario y Reglas de Liquidación Asiática", expanded=False):
                st.markdown(
                    """
                    ### Guía de Liquidación en Mercados Asiáticos (Equipos y Goles)
                    Las líneas asiáticas eliminan o fraccionan el escenario de empate dividiendo el riesgo entre líneas enteras (`.00`), medias (`.50`) y cuartos de gol (`.25` / `.75`):

                    | Escenario | Condición Matemática (d = Margen + L) | Retorno sobre Stake (S) a Cuota (c) | Impacto en Bankroll |
                    | :--- | :--- | :--- | :--- |
                    | **✅ Full Win (Acierto Total)** | d > +0.25 | S × c | +S × (c - 1) |
                    | **📈 Half Win (Medio Acierto)** | d = +0.25 | S × (1 + (c - 1) / 2) | +(S / 2) × (c - 1) |
                    | **🔄 Push (Nulo / Reembolso)** | d = 0.00 | S × 1.00 | 0.00 € (Devolución íntegra) |
                    | **📉 Half Loss (Media Pérdida)** | d = -0.25 | S × 0.50 | -0.50 × S |
                    | **❌ Full Loss (Pérdida Total)** | d < -0.25 | 0.00 € | -1.00 × S |

                    - **Líneas de Cuarto (±0.25, ±0.75, ±1.25):** Dividen automáticamente la apuesta en dos mitades iguales sobre las líneas adyacentes L1 = L - 0.25 y L2 = L + 0.25.
                    - **Cuota Justa del Modelo (c*):** Resuelve analíticamente E[R(c*)] = 1.00 sobre la matriz bivariante de Poisson ajustada por **Dixon-Coles** (ρ = -0.11).
                    """
                )

elif menu_seleccionado == "🎯 Escáner Autónomo de Ineficiencias (+EV Scanner)":
    ev_render_modulo()

elif menu_seleccionado == "🔗 Constructor Combinadas (EV+)":
    st.title("🔗 Constructor Automático de Combinadas (EV+)")
    st.markdown("Calcula el valor esperado (EV) acumulado y la cuota total de tus apuestas combinadas.")
    col_c1, col_c2 = st.columns(2)
    with col_c1: c1, p1 = st.number_input("Cuota Selección 1:", value=1.50), st.number_input("Prob Real 1 (%):", value=70.0)/100
    with col_c2: c2, p2 = st.number_input("Cuota Selección 2:", value=1.80), st.number_input("Prob Real 2 (%):", value=60.0)/100
    ev = round(((p1*p2) * (c1*c2) - 1) * 100, 2)
    st.markdown("---")
    st.metric("Cuota Combinada", f"{round(c1*c2, 2)}")
    st.metric("Probabilidad Conjunta", f"{p1*p2*100:.1f}%")
    if ev > 0: st.success(f"🔥 Combinada con Valor Esperado Positivo (EV+): +{ev}%")
    else: st.error(f"⚠️ Combinada sin Valor Esperado (EV-): {ev}%")

elif menu_seleccionado == "💸 Calculadora & Kelly":
    st.title("💸 Calculadora de Criterio de Kelly")
    st.markdown("Gestión matemática de Bankroll para maximizar ganancias y proteger tu capital.")
    bank = st.number_input("Bankroll Total (€):", value=100.0, step=10.0)
    cuota = st.number_input("Cuota de la Apuesta:", value=1.85, step=0.05)
    prob = st.slider("Probabilidad Estimada por Modelo (%):", 10.0, 99.0, 60.0) / 100.0
    b = cuota - 1.0
    k = (prob * b - (1 - prob)) / b if b > 0 else 0
    col_k1, col_k2 = st.columns(2)
    col_k1.metric("Stake Recomendado (Medio Kelly - Seguro)", f"{max(bank * (k * 0.5), 0):.2f} €", f"{((max(bank * (k * 0.5), 0))/bank*100):.1f}% del bank")
    col_k2.metric("Stake Máximo (Kelly Completo - Agresivo)", f"{max(bank * k, 0):.2f} €", f"{((max(bank * k, 0))/bank*100):.1f}% del bank")

elif menu_seleccionado == "🤖 Asistente IA Local":
    st.markdown(
        """
        <style>
        .stChatFloatingInputContainer,
        div[data-testid="stBottom"],
        div[data-testid="stBottom"] > div,
        footer {
            background-color: #0e1117 !important;
            background: #0e1117 !important;
        }
        div[data-testid="stChatInput"] {
            background-color: #1a1d27 !important;
            border: 1px solid #2d3345 !important;
            border-radius: 12px !important;
        }
        div[data-testid="stChatInput"] textarea {
            background-color: transparent !important;
            color: #ffffff !important;
        }
        div[data-testid="stChatInput"] button {
            color: #00d26a !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    import os
    try:
        from google import genai
        from google.genai import types as genai_types
        _GENAI_SDK_OK = True
    except Exception:
        genai = None
        genai_types = None
        _GENAI_SDK_OK = False

    st.title("🤖 Asistente Cuantitativo de Élite (Google AI Studio)")
    st.caption(
        "Motor Cuantitativo `gemini-3.8-flash` (Streaming en Tiempo Real) · RAG Dinámico sobre `macro_base_futbol.db` · "
        "Calibración de Poisson, Valor Esperado (EV+), Hándicap Asiático y Fricción Arbitral."
    )

    # ── 1. INICIALIZACIÓN RESILIENTE DE CLAVE API Y CLIENTE GOOGLE GENAI ─────
    clave_entorno = ""
    try:
        clave_entorno = st.secrets.get("GEMINI_API_KEY", "") or ""
    except Exception:
        clave_entorno = ""
    if not clave_entorno:
        clave_entorno = os.getenv("GEMINI_API_KEY", "") or ""

    clave_activa = clave_entorno.strip()
    if not clave_activa:
        st.info(
            "🔐 **Autenticación Google AI Studio requerida:** Introduce tu clave API de Gemini "
            "(o defínela en `st.secrets['GEMINI_API_KEY']` / variable de entorno `GEMINI_API_KEY`) "
            "para habilitar la inferencia en vivo con `gemini-3.8-flash`."
        )
        clave_input = st.text_input(
            "🔑 Clave API de Google AI Studio:",
            type="password",
            placeholder="AIzaSy...",
            key="input_gemini_api_key_asistente",
        )
        clave_activa = (clave_input or "").strip()

    # ── 2. PERSONA DE DOMINIO CUANTITATIVO (SYSTEM INSTRUCTION) ──────────────
    SYSTEM_INSTRUCTION_BASE = (
        "Eres el Asistente Cuantitativo de Élite de BigData Football.\n"
        "Especialidad: Detección de Valor Esperado (EV+), arbitraje de cuotas, calibración de Poisson y fricción arbitral.\n"
        "Directrices de respuesta:\n"
        "- Sé ejecutivo y directo. Limita el análisis a 3 apartados breves con viñetas: "
        "1) Cuotas Justas vs Mercado, 2) Tendencia xG y Goles, 3) Recomendación Cuantitativa de Valor (EV+).\n"
        "- Rigor matemático y enfoque directo.\n"
        "- Prohibido el uso de jerga matemática cruda innecesaria (nada de lambda 'λ' ni '@' antes de cuotas).\n"
        "- Si hay datos de la base de datos inyectados en la consulta, priorízalos sobre conocimiento general.\n"
        "- Ofrece siempre conclusiones de mercado claras (apuesta con valor vs mercado sobrevalorado)."
    )

    # ── Inicialización de memoria multi-turno ────────────────────────────────
    if "chat_asistente_ia" not in st.session_state:
        st.session_state.chat_asistente_ia = []
    if "ia_chat_history" not in st.session_state:
        st.session_state["ia_chat_history"] = st.session_state.chat_asistente_ia

    # ── Helpers de Extracción Cuantitativa en Tiempo Real (SQLite RAG Lite) ──
    def _extraer_perfil_equipo_sqlite(df_scope: pd.DataFrame, equipo: str) -> dict:
        """
        Extrae de macro_base_futbol.db (vía DataFrame validado o consulta directa SQLite)
        los últimos 5 partidos (resultados, goles, rival, córners, tarjetas) y los
        promedios de temporada (goles a favor/en contra, córners, faltas y tarjetas).
        """
        perfil = {
            "equipo": equipo,
            "partidos_analizados": 0,
            "promedios_temporada": {
                "goles_marcados": 0.0,
                "goles_encajados": 0.0,
                "xg_favor": 0.0,
                "xg_contra": 0.0,
                "corners_favor": 0.0,
                "corners_contra": 0.0,
                "corners_totales_partido": 0.0,
                "amarillas_propias": 0.0,
                "rojas_propias": 0.0,
                "tarjetas_totales_partido": 0.0,
                "faltas_cometidas": 0.0,
                "over25_pct": 0.0,
                "btts_pct": 0.0,
            },
            "ultimos_5_partidos": [],
        }
        if df_scope is None or df_scope.empty or not equipo or equipo == "Ninguno (Consulta General)":
            return perfil
        if "HomeTeam" not in df_scope.columns or "AwayTeam" not in df_scope.columns:
            return perfil

        mask = (df_scope["HomeTeam"] == equipo) | (df_scope["AwayTeam"] == equipo)
        df_eq = df_scope[mask].copy()
        if df_eq.empty:
            return perfil

        if "Date" in df_eq.columns:
            df_eq["_dt_ord"] = pd.to_datetime(df_eq["Date"], errors="coerce", dayfirst=True)
            df_eq = df_eq.sort_values(by="_dt_ord", ascending=True, kind="stable")

        for col_n in ("FTHG", "FTAG", "HC", "AC", "HY", "AY", "HR", "AR", "HF", "AF", "HxG", "AxG"):
            if col_n in df_eq.columns:
                df_eq[col_n] = pd.to_numeric(df_eq[col_n], errors="coerce")

        if "FTHG" in df_eq.columns and "FTAG" in df_eq.columns:
            df_eq = df_eq.dropna(subset=["FTHG", "FTAG"]).copy()
        if df_eq.empty:
            return perfil

        es_local = df_eq["HomeTeam"] == equipo
        gf_s = np.where(es_local, df_eq["FTHG"], df_eq["FTAG"])
        ga_s = np.where(es_local, df_eq["FTAG"], df_eq["FTHG"])

        if "HxG" in df_eq.columns and "AxG" in df_eq.columns:
            xgf_s = pd.Series(np.where(es_local, df_eq["HxG"], df_eq["AxG"])).dropna()
            xgc_s = pd.Series(np.where(es_local, df_eq["AxG"], df_eq["HxG"])).dropna()
        else:
            xgf_s, xgc_s = pd.Series(dtype=float), pd.Series(dtype=float)

        cf_s = pd.Series(np.where(es_local, df_eq["HC"], df_eq["AC"])).dropna() if {"HC", "AC"}.issubset(df_eq.columns) else pd.Series(dtype=float)
        cc_s = pd.Series(np.where(es_local, df_eq["AC"], df_eq["HC"])).dropna() if {"HC", "AC"}.issubset(df_eq.columns) else pd.Series(dtype=float)
        am_s = pd.Series(np.where(es_local, df_eq["HY"], df_eq["AY"])).dropna() if {"HY", "AY"}.issubset(df_eq.columns) else pd.Series(dtype=float)
        am_riv_s = pd.Series(np.where(es_local, df_eq["AY"], df_eq["HY"])).dropna() if {"HY", "AY"}.issubset(df_eq.columns) else pd.Series(dtype=float)
        roj_s = pd.Series(np.where(es_local, df_eq["HR"], df_eq["AR"])).dropna() if {"HR", "AR"}.issubset(df_eq.columns) else pd.Series(dtype=float)
        fal_s = pd.Series(np.where(es_local, df_eq["HF"], df_eq["AF"])).dropna() if {"HF", "AF"}.issubset(df_eq.columns) else pd.Series(dtype=float)

        n_part = len(df_eq)
        gf_med = float(np.nanmean(gf_s)) if n_part > 0 else 0.0
        ga_med = float(np.nanmean(ga_s)) if n_part > 0 else 0.0
        cf_med = float(cf_s.mean()) if not cf_s.empty else 0.0
        cc_med = float(cc_s.mean()) if not cc_s.empty else 0.0
        am_med = float(am_s.mean()) if not am_s.empty else 0.0
        am_riv_med = float(am_riv_s.mean()) if not am_riv_s.empty else 0.0
        roj_med = float(roj_s.mean()) if not roj_s.empty else 0.0

        perfil["partidos_analizados"] = int(n_part)
        perfil["promedios_temporada"] = {
            "goles_marcados": round(gf_med, 2),
            "goles_encajados": round(ga_med, 2),
            "xg_favor": round(float(xgf_s.mean()), 2) if not xgf_s.empty else round(gf_med, 2),
            "xg_contra": round(float(xgc_s.mean()), 2) if not xgc_s.empty else round(ga_med, 2),
            "corners_favor": round(cf_med, 2),
            "corners_contra": round(cc_med, 2),
            "corners_totales_partido": round(cf_med + cc_med, 2),
            "amarillas_propias": round(am_med, 2),
            "rojas_propias": round(roj_med, 2),
            "tarjetas_totales_partido": round(am_med + am_riv_med + roj_med, 2),
            "faltas_cometidas": round(float(fal_s.mean()), 2) if not fal_s.empty else 11.5,
            "over25_pct": round(float(np.mean((gf_s + ga_s) > 2.5) * 100.0), 1),
            "btts_pct": round(float(np.mean((gf_s > 0) & (ga_s > 0)) * 100.0), 1),
        }

        ult_5 = df_eq.tail(5)
        lista_u5 = []
        for _, r in ult_5.iterrows():
            loc = r["HomeTeam"] == equipo
            rival = str(r["AwayTeam"] if loc else r["HomeTeam"])
            cond = "Local" if loc else "Visitante"
            gf_i = int(float(r["FTHG"] if loc else r["FTAG"]))
            ga_i = int(float(r["FTAG"] if loc else r["FTHG"]))
            res_letra = "V" if gf_i > ga_i else ("E" if gf_i == ga_i else "D")
            c_tot = (
                int(float(r.get("HC", 0) or 0) + float(r.get("AC", 0) or 0))
                if pd.notna(r.get("HC")) and pd.notna(r.get("AC"))
                else None
            )
            t_tot = (
                int(float(r.get("HY", 0) or 0) + float(r.get("AY", 0) or 0) + float(r.get("HR", 0) or 0) + float(r.get("AR", 0) or 0))
                if pd.notna(r.get("HY")) and pd.notna(r.get("AY"))
                else None
            )
            fecha_str = str(r.get("Date", "S/F"))
            extra_stats = []
            if c_tot is not None:
                extra_stats.append(f"{c_tot} córners")
            if t_tot is not None:
                extra_stats.append(f"{t_tot} tarjetas")
            sufijo = f" ({', '.join(extra_stats)})" if extra_stats else ""
            lista_u5.append(
                f"{fecha_str} [{cond}] vs {rival}: {gf_i}-{ga_i} ({res_letra}){sufijo}"
            )
        perfil["ultimos_5_partidos"] = lista_u5
        return perfil

    # ── 3. INYECCIÓN DINÁMICA DE CONTEXTO SQLITE (RAG LITE) ──────────────────
    df_base_ia = df_partidos_global
    if (df_base_ia is None or df_base_ia.empty) and RUTA_BD.exists():
        try:
            with sqlite3.connect(RUTA_BD) as _conn_ia:
                _df_sql_ia = pd.read_sql("SELECT * FROM historico_reciente", _conn_ia)
                _df_sql_ia.columns = _df_sql_ia.columns.str.replace("\ufeff", "").str.replace("ï»¿", "").str.strip()
                df_base_ia = _df_sql_ia
        except Exception:
            df_base_ia = pd.DataFrame()

    col_div_ia = next(
        (c for c in ("Div", "Liga_Origen") if c in df_base_ia.columns),
        next((c for c in df_base_ia.columns if "Div" in str(c)), None),
    ) if not df_base_ia.empty else None

    divs_ia = (
        [d for d in MAPA_LIGAS if d in set(df_base_ia[col_div_ia].dropna().astype(str))]
        if col_div_ia
        else list(MAPA_LIGAS.keys())
    )
    opciones_liga_ia = [MAPA_LIGAS[d] for d in divs_ia] if divs_ia else list(MAPA_LIGAS.values())

    ctx_c1, ctx_c2, ctx_c3, ctx_c4 = st.columns([1.3, 1.3, 1.3, 1.1])
    with ctx_c1:
        liga_ia_sel = st.selectbox("🏆 Competición (MAPA_LIGAS):", opciones_liga_ia, index=0, key="ia_ctx_liga")
    div_ia_code = MAPA_INV.get(liga_ia_sel, "SP1")

    df_liga_ia = (
        df_base_ia[df_base_ia[col_div_ia].astype(str) == div_ia_code].copy()
        if not df_base_ia.empty and col_div_ia
        else df_base_ia.copy()
    )
    equipos_ia = (
        sorted(pd.concat([df_liga_ia["HomeTeam"], df_liga_ia["AwayTeam"]]).dropna().unique().tolist())
        if not df_liga_ia.empty and {"HomeTeam", "AwayTeam"}.issubset(df_liga_ia.columns)
        else []
    )

    with ctx_c2:
        opciones_local_ia = equipos_ia if equipos_ia else ["Sin datos"]
        home_ia = st.selectbox("🏠 Equipo Principal / Local:", opciones_local_ia, index=0, key=f"ia_ctx_h_{div_ia_code}")
    with ctx_c3:
        rivales_ia = [e for e in equipos_ia if e != home_ia]
        opciones_visit_ia = rivales_ia + ["Ninguno (Solo Equipo Principal)"] if rivales_ia else ["Ninguno (Solo Equipo Principal)"]
        away_ia = st.selectbox("✈️ Rival / Visitante (Opcional):", opciones_visit_ia, index=0, key=f"ia_ctx_a_{div_ia_code}")

    df_arb_ia = cargar_datos_arbitros(1)
    arbs_ia_disp = (
        df_arb_ia[df_arb_ia["Div"].astype(str) == div_ia_code]["Árbitro"].dropna().unique().tolist()
        if not df_arb_ia.empty and "Div" in df_arb_ia.columns
        else []
    ) or (
        df_arb_ia["Árbitro"].dropna().unique().tolist()
        if not df_arb_ia.empty and "Árbitro" in df_arb_ia.columns
        else ["Designación Oficial"]
    )
    with ctx_c4:
        arb_ia = st.selectbox("⚖️ Árbitro:", arbs_ia_disp, index=0, key=f"ia_ctx_ref_{div_ia_code}")

    # Extracción de estadísticas reales desde macro_base_futbol.db
    perfil_h = _extraer_perfil_equipo_sqlite(df_liga_ia, home_ia)
    hay_visitante = away_ia and away_ia != "Ninguno (Solo Equipo Principal)" and away_ia != "Sin datos"
    perfil_a = _extraer_perfil_equipo_sqlite(df_liga_ia, away_ia) if hay_visitante else None

    prom_h = perfil_h["promedios_temporada"]
    prom_a = perfil_a["promedios_temporada"] if perfil_a else None

    # Calibración Poisson y cuotas justas para el contexto
    if prom_a and perfil_a["partidos_analizados"] > 0:
        exp_g_h = round(max(0.25, (prom_h["xg_favor"] + prom_a["xg_contra"]) / 2.0), 2)
        exp_g_a = round(max(0.25, (prom_a["xg_favor"] + prom_h["xg_contra"]) / 2.0), 2)
    else:
        exp_g_h = round(max(0.25, prom_h["xg_favor"] or 1.35), 2)
        exp_g_a = round(max(0.25, prom_h["xg_contra"] or 1.15), 2)

    p_1 = p_x = p_2 = p_o25 = p_btts = 0.0
    for _g1 in range(8):
        for _g2 in range(8):
            _pr = calcular_probabilidad_poisson(exp_g_h, _g1) * calcular_probabilidad_poisson(exp_g_a, _g2)
            if _g1 > _g2:
                p_1 += _pr
            elif _g1 == _g2:
                p_x += _pr
            else:
                p_2 += _pr
            if (_g1 + _g2) > 2.5:
                p_o25 += _pr
            if _g1 > 0 and _g2 > 0:
                p_btts += _pr

    p_u25 = max(0.0, 1.0 - p_o25)
    cuota_justa_1 = round(1.0 / max(p_1, 0.01), 2)
    cuota_justa_x = round(1.0 / max(p_x, 0.01), 2)
    cuota_justa_2 = round(1.0 / max(p_2, 0.01), 2)
    cuota_justa_o25 = round(1.0 / max(p_o25, 0.01), 2)
    cuota_justa_u25 = round(1.0 / max(p_u25, 0.01), 2)
    cuota_justa_btts = round(1.0 / max(p_btts, 0.01), 2)

    # Proyección de Hándicap Asiático -0.5 / +0.5 y DNB (0.0)
    prob_ah_local_0 = p_1 / max(p_1 + p_2, 0.01)
    cuota_justa_ah0_local = round(1.0 / max(prob_ah_local_0, 0.01), 2)
    cuota_justa_ah0_visit = round(1.0 / max(1.0 - prob_ah_local_0, 0.01), 2)

    tarj_arb_ia = 4.50
    faltas_arb_ia = 24.0
    perfil_arb_txt = "Moderado"
    if not df_arb_ia.empty and "Árbitro" in df_arb_ia.columns:
        _r_arb = df_arb_ia[df_arb_ia["Árbitro"] == arb_ia]
        if not _r_arb.empty:
            tarj_arb_ia = float(_r_arb["Media Amarillas Est."].iloc[0])
            faltas_arb_ia = float(_r_arb["Media Faltas"].iloc[0]) if "Media Faltas" in _r_arb.columns else 24.0
            perfil_arb_txt = str(_r_arb["Perfil Disciplinario"].iloc[0]) if "Perfil Disciplinario" in _r_arb.columns else "Moderado"

    friccion_faltas_ia = round(
        prom_h["faltas_cometidas"] + (prom_a["faltas_cometidas"] if prom_a else prom_h["faltas_cometidas"]),
        1,
    )
    tarjetas_proy_partido = round(
        ((prom_h["amarillas_propias"] + (prom_a["amarillas_propias"] if prom_a else 2.2)) + tarj_arb_ia) / 2.0,
        2,
    )
    corners_proy_partido = round(
        prom_h["corners_favor"] + (prom_a["corners_favor"] if prom_a else prom_h["corners_contra"]),
        2,
    )

    etiqueta_duelo = f"{home_ia} vs {away_ia}" if hay_visitante else home_ia

    resumen_dict = {
        "origen": "macro_base_futbol.db",
        "competicion": liga_ia_sel,
        "codigo_liga": div_ia_code,
        "emparejamiento": etiqueta_duelo,
        "equipo_local": perfil_h,
        "equipo_visitante": perfil_a if (perfil_a and perfil_a["partidos_analizados"] > 0) else "Análisis individual",
        "calibracion_poisson_cuotas_justas": {
            "goles_esperados_local": exp_g_h,
            "goles_esperados_visitante": exp_g_a,
            "goles_esperados_totales": round(exp_g_h + exp_g_a, 2),
            "corners_esperados_totales": corners_proy_partido,
            "1X2": {
                "prob_1_pct": round(p_1 * 100, 1),
                "cuota_justa_1": cuota_justa_1,
                "prob_X_pct": round(p_x * 100, 1),
                "cuota_justa_X": cuota_justa_x,
                "prob_2_pct": round(p_2 * 100, 1),
                "cuota_justa_2": cuota_justa_2,
            },
            "goles_y_btts": {
                "prob_over_2_5_pct": round(p_o25 * 100, 1),
                "cuota_justa_over_2_5": cuota_justa_o25,
                "prob_under_2_5_pct": round(p_u25 * 100, 1),
                "cuota_justa_under_2_5": cuota_justa_u25,
                "prob_btts_pct": round(p_btts * 100, 1),
                "cuota_justa_btts": cuota_justa_btts,
            },
            "handicap_asiatico": {
                "ah_0_0_local_cuota_justa": cuota_justa_ah0_local,
                "ah_0_0_visitante_cuota_justa": cuota_justa_ah0_visit,
                "ah_menos_0_5_local_cuota_justa": cuota_justa_1,
                "ah_mas_0_5_visitante_cuota_justa": round(1.0 / max(p_x + p_2, 0.01), 2),
            },
        },
        "arbitro_y_disciplina": {
            "colegiado": str(arb_ia),
            "perfil_rigor": perfil_arb_txt,
            "media_amarillas_arbitro": round(tarj_arb_ia, 2),
            "media_faltas_arbitro": round(faltas_arb_ia, 1),
            "friccion_combinada_faltas": friccion_faltas_ia,
            "proyeccion_tarjetas_totales": tarjetas_proy_partido,
        },
    }

    contexto_compacto_str = json.dumps(resumen_dict, ensure_ascii=False, separators=(",", ":"))

    # ── 4. QUICK PROMPT PRESETS (ANÁLISIS EN UN CLIC) + REINICIAR SESIÓN ────
    st.markdown("##### ⚡ Acciones Rápidas de Análisis Cuantitativo")
    qp1, qp2, qp3, qp4, qp5 = st.columns([2.2, 2.5, 2.4, 2.5, 1.6])
    prompt_preset_activo = None

    with qp1:
        if st.button("🎯 Evaluar Valor 1X2", use_container_width=True, key="btn_preset_1x2"):
            prompt_preset_activo = (
                f"🎯 Evaluar Valor 1X2 para {etiqueta_duelo} ({liga_ia_sel}): "
                "compara las probabilidades reales y cuotas justas de victoria local, empate y visitante "
                "según los últimos 5 partidos y medias de temporada, e indica dónde reside el valor esperado."
            )
    with qp2:
        if st.button("🔥 Análisis de Goles (Over/Under 2.5)", use_container_width=True, key="btn_preset_goles"):
            prompt_preset_activo = (
                f"🔥 Análisis de Goles (Over/Under 2.5) y Ambos Marcan para {etiqueta_duelo}: "
                "evalúa los promedios de goles marcados y encajados, tendencia en los últimos 5 partidos, "
                "córners y cuota justa del mercado de goles."
            )
    with qp3:
        if st.button("🟨 Riesgo Disciplinario y Tarjetas", use_container_width=True, key="btn_preset_tarjetas"):
            prompt_preset_activo = (
                f"🟨 Riesgo Disciplinario y Tarjetas en {etiqueta_duelo} con el árbitro {arb_ia}: "
                "analiza las medias de tarjetas amarillas, rojas, faltas cometidas por ambos equipos y "
                "el rigor del colegiado para determinar qué línea de tarjetas tiene valor."
            )
    with qp4:
        if st.button("⚡ Proyección de Hándicap Asiático", use_container_width=True, key="btn_preset_ah"):
            prompt_preset_activo = (
                f"⚡ Proyección de Hándicap Asiático para {etiqueta_duelo}: "
                "evalúa las líneas asiáticas de referencia (0.0, -0.5 y +0.5) en función de la superioridad "
                "ofensiva/defensiva y recomienda la cobertura óptima de mercado."
            )
    with qp5:
        if st.button("🗑️ Reiniciar Sesión", use_container_width=True, key="btn_clear_chat_ia"):
            st.session_state.chat_asistente_ia = []
            st.session_state["ia_chat_history"] = st.session_state.chat_asistente_ia
            st.rerun()

    # ── 5. INTERFAZ DE CHAT MULTI-TURNO CON STREAMING EN TIEMPO REAL ─────────
    for msg in st.session_state.chat_asistente_ia:
        rol_vista = "assistant" if msg.get("role") in ("assistant", "model") else "user"
        with st.chat_message(rol_vista):
            if msg.get("badge"):
                st.caption(msg["badge"])
            st.markdown(msg.get("content", ""))

    input_chat_usuario = st.chat_input("Consulta métricas, valor esperado, árbitros o cuotas...")
    consulta_usuario_final = prompt_preset_activo or input_chat_usuario

    if consulta_usuario_final:
        st.session_state.chat_asistente_ia.append({"role": "user", "content": consulta_usuario_final})
        st.session_state["ia_chat_history"] = st.session_state.chat_asistente_ia
        with st.chat_message("user"):
            st.markdown(consulta_usuario_final)

        with st.chat_message("assistant"):
            if not _GENAI_SDK_OK or genai is None:
                st.error(
                    "⚠️ El paquete oficial `google-genai` no está disponible en el entorno Python activo. "
                    "Instálalo mediante `pip install google-genai`."
                )
            elif not clave_activa:
                st.error(
                    "🔑 Introduce una Clave API válida de Google AI Studio en el campo superior "
                    "(o configura `GEMINI_API_KEY` en `st.secrets` / variables de entorno) para consultar a `gemini-3.8-flash`."
                )
            else:
                try:
                    client = genai.Client(api_key=clave_activa)

                    # Historial breve de últimos turnos para preservar memoria sin penalizar latencia
                    turnos_previos = st.session_state.chat_asistente_ia[-5:-1]
                    resumen_historial = "\n".join(
                        f"{'Analista' if m.get('role') == 'user' else 'Asistente'}: {m.get('content', '')}"
                        for m in turnos_previos
                        if m.get("content")
                    )
                    bloque_historial = f"\nHISTORIAL_RECIENTE:\n{resumen_historial}\n" if resumen_historial else ""

                    full_prompt = (
                        f"{SYSTEM_INSTRUCTION_BASE}\n\n"
                        f"DATOS_SQLITE_MACRO_BASE:\n{contexto_compacto_str}"
                        f"{bloque_historial}\n"
                        f"CONSULTA_ACTUAL:\n{consulta_usuario_final}"
                    )

                    modelo_stream_usado = ["gemini-3.8-flash"]

                    def generar_stream_asistente(full_prompt):
                        try:
                            modelo_stream_usado[0] = "gemini-3.8-flash"
                            response_stream = client.models.generate_content_stream(
                                model='gemini-3.8-flash',
                                contents=full_prompt,
                            )
                            for chunk in response_stream:
                                if chunk.text:
                                    yield chunk.text
                        except Exception:
                            try:
                                modelo_stream_usado[0] = "gemini-2.0-flash"
                                response_stream = client.models.generate_content_stream(
                                    model='gemini-2.0-flash',
                                    contents=full_prompt,
                                )
                                for chunk in response_stream:
                                    if chunk.text:
                                        yield chunk.text
                            except Exception:
                                modelo_stream_usado[0] = "gemini-2.5-flash"
                                response_stream = client.models.generate_content_stream(
                                    model='gemini-2.5-flash',
                                    contents=full_prompt,
                                )
                                for chunk in response_stream:
                                    if chunk.text:
                                        yield chunk.text

                    respuesta_completa = st.write_stream(generar_stream_asistente(full_prompt))
                    texto_final = str(respuesta_completa or "").strip()
                    if texto_final:
                        badge_ok = f"🟢 Google AI Studio (`{modelo_stream_usado[0]}` · Streaming) · Contexto SQLite: {etiqueta_duelo}"
                        st.caption(badge_ok)
                        st.session_state.chat_asistente_ia.append(
                            {
                                "role": "assistant",
                                "content": texto_final,
                                "badge": badge_ok,
                            }
                        )
                        st.session_state["ia_chat_history"] = st.session_state.chat_asistente_ia
                except genai.errors.APIError as api_err:
                    st.error(
                        f"🚫 Error de la API de Google AI Studio (`APIError`): {api_err}. "
                        "Verifica la validez de tu clave API, los permisos del proyecto o el límite de cuota."
                    )
                except Exception as exc_gen:
                    st.error(
                        f"⚠️ Error inesperado al conectar con Google AI Studio: {exc_gen}"
                    )