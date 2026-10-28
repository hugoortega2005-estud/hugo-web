import sqlite3
import pandas as pd
import streamlit as st
import numpy as np
import math
from io import BytesIO
from datetime import datetime
from pathlib import Path

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
COLOR_TEAL = colors.HexColor("#1B9AAA")
COLOR_HEADER = colors.HexColor("#1A365D")
COLOR_ROW = colors.HexColor("#F7FAFC")
COLOR_LINE = colors.HexColor("#CBD5E0")


def _fuentes_pdf():
    global _FUENTES_PDF
    if _FUENTES_PDF:
        return _FUENTES_PDF
    arial = Path(r"C:\Windows\Fonts\arial.ttf")
    arial_bd = Path(r"C:\Windows\Fonts\arialbd.ttf")
    if arial.exists():
        pdfmetrics.registerFont(TTFont("HubSans", str(arial)))
        pdfmetrics.registerFont(TTFont("HubSans-Bold", str(arial_bd if arial_bd.exists() else arial)))
        _FUENTES_PDF = ("HubSans", "HubSans-Bold")
    else:
        _FUENTES_PDF = ("Helvetica", "Helvetica-Bold")
    return _FUENTES_PDF


def _slug_archivo(texto):
    seguro = "".join(c if c.isalnum() or c in "._- " else "_" for c in str(texto))
    return seguro.strip().replace(" ", "_")[:50] or "informe"


def _estilos_parrafo_pdf():
    regular, bold = _fuentes_pdf()
    return {
        "marca": ParagraphStyle("hub_marca", fontName=bold, fontSize=11, textColor=COLOR_TEAL, alignment=TA_LEFT, spaceAfter=2),
        "titulo": ParagraphStyle("hub_titulo", fontName=bold, fontSize=16, textColor=COLOR_NAVY, alignment=TA_LEFT, spaceAfter=4, leading=20),
        "subtitulo": ParagraphStyle("hub_sub", fontName=regular, fontSize=9, textColor=colors.HexColor("#4A5568"), alignment=TA_LEFT, spaceAfter=8),
        "seccion": ParagraphStyle("hub_sec", fontName=bold, fontSize=12, textColor=COLOR_HEADER, alignment=TA_LEFT, spaceBefore=10, spaceAfter=6),
        "cuerpo": ParagraphStyle("hub_cuerpo", fontName=regular, fontSize=9, textColor=COLOR_NAVY, leading=13, spaceAfter=3),
        "celda": ParagraphStyle("hub_celda", fontName=regular, fontSize=7, textColor=COLOR_NAVY, leading=9, alignment=TA_LEFT),
        "celda_h": ParagraphStyle("hub_celda_h", fontName=bold, fontSize=7.5, textColor=colors.white, leading=10, alignment=TA_CENTER),
        "pie": ParagraphStyle("hub_pie", fontName=regular, fontSize=8, textColor=colors.HexColor("#718096"), alignment=TA_CENTER),
    }


def _tabla_pdf(encabezados, filas, col_widths):
    estilos = _estilos_parrafo_pdf()
    data = [[Paragraph(str(h), estilos["celda_h"]) for h in encabezados]]
    for fila in filas:
        data.append([Paragraph(str(v), estilos["celda"]) for v in fila])
    tabla = Table(data, colWidths=col_widths, repeatRows=1)
    estilo = [
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_HEADER),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.3, COLOR_LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    for i in range(1, len(data)):
        if i % 2 == 0:
            estilo.append(("BACKGROUND", (0, i), (-1, i), COLOR_ROW))
    tabla.setStyle(TableStyle(estilo))
    return tabla


def _pie_pdf(canvas, doc):
    canvas.saveState()
    regular, _ = _fuentes_pdf()
    ancho, alto = doc.pagesize
    canvas.setFillColor(COLOR_NAVY)
    canvas.rect(0, alto - 0.35 * cm, ancho, 0.35 * cm, fill=1, stroke=0)
    canvas.setFillColor(COLOR_TEAL)
    canvas.rect(0, 0, ancho, 0.28 * cm, fill=1, stroke=0)
    canvas.setFont(regular, 8)
    canvas.setFillColor(colors.HexColor("#4A5568"))
    canvas.drawString(1.5 * cm, 0.45 * cm, "Hugo's Hub  |  Informe confidencial de analisis")
    canvas.drawRightString(ancho - 1.5 * cm, 0.45 * cm, f"Pagina {doc.page}")
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
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=1.6 * cm,
        rightMargin=1.6 * cm,
        topMargin=1.6 * cm,
        bottomMargin=1.4 * cm,
        title=f"Informe Ejecutivo {datos['equipo']}",
        author="Hugo's Hub",
    )
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    equipo = datos["equipo"]
    n_loc = datos["n_local"]
    n_vis = datos["n_vis"]
    m_gf_l = (datos["gf_local"] / n_loc) if n_loc else 0
    m_gc_l = (datos["gc_local"] / n_loc) if n_loc else 0
    m_gf_v = (datos["gf_vis"] / n_vis) if n_vis else 0
    m_gc_v = (datos["gc_vis"] / n_vis) if n_vis else 0

    story = [
        Paragraph("HUGO'S HUB  ·  INFORME EJECUTIVO", estilos["marca"]),
        Paragraph(f"Dossier de equipo: {equipo}", estilos["titulo"]),
        Paragraph(f"Generado el {fecha}  |  Muestra analizada: {datos['partidos']} partidos", estilos["subtitulo"]),
        HRFlowable(width="100%", thickness=1.2, color=COLOR_TEAL, spaceAfter=6),
        Paragraph("1. Resumen de estadísticas globales", estilos["seccion"]),
        _tabla_pdf(
            ["Métrica", "Valor"],
            [
                ["Partidos analizados", datos["partidos"]],
                ["Balance (V-E-D)", f"{datos['v']} - {datos['e']} - {datos['d']}"],
                ["Goles a favor", f"{datos['gf']}  (media {datos['media_gf']:.2f})"],
                ["Goles en contra", f"{datos['gc']}  (media {datos['media_gc']:.2f})"],
                ["Diferencia de goles", f"{datos['gf'] - datos['gc']:+d}"],
            ],
            [9 * cm, 8 * cm],
        ),
        Paragraph("2. Porcentajes clave de mercado", estilos["seccion"]),
        _tabla_pdf(
            ["Mercado", "Porcentaje", "Frecuencia"],
            [
                ["Más de 2.5 goles", f"{datos['p_over25']:.1f}%", f"{datos['over_25_count']}/{datos['partidos']}"],
                ["Ambos marcan (BTTS)", f"{datos['p_btts']:.1f}%", f"{datos['btts_count']}/{datos['partidos']}"],
                ["Porterías a cero", f"{datos['p_cs']:.1f}%", f"{datos['clean_sheets']}/{datos['partidos']}"],
            ],
            [6.5 * cm, 5 * cm, 5.5 * cm],
        ),
        Paragraph("3. Desglose Local vs Visitante", estilos["seccion"]),
        _tabla_pdf(
            ["Concepto", "Local", "Visitante"],
            [
                ["Partidos", n_loc, n_vis],
                ["Goles anotados (total)", datos["gf_local"], datos["gf_vis"]],
                ["Goles anotados (media)", f"{m_gf_l:.2f}", f"{m_gf_v:.2f}"],
                ["Goles recibidos (total)", datos["gc_local"], datos["gc_vis"]],
                ["Goles recibidos (media)", f"{m_gc_l:.2f}", f"{m_gc_v:.2f}"],
            ],
            [7 * cm, 5 * cm, 5 * cm],
        ),
        Spacer(1, 14),
        Paragraph(
            "Fuente: base histórica local de Hugo's Hub. Las medias se calculan sobre los partidos indexados "
            "del equipo seleccionado. Documento generado en memoria (BytesIO) sin archivos temporales.",
            estilos["pie"],
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
            "🔗 Constructor Combinadas (EV+)", 
            "💸 Calculadora & Kelly"
        ]
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
    st.title("📡 Escáner de Cuotas de Valor Reales (EV+ Pro)")
    st.markdown("Encuentra oportunidades de valor verificado en **Mercado 1X2, Líneas de Goles Completas y Ambos Marcan**.")
    
    col_e1, col_e2, col_e3 = st.columns([2, 2, 1])
    with col_e1:
        liga_api_sel = st.selectbox("Selecciona Competición:", ["Premier League (Inglaterra)", "LaLiga (España)", "Serie A (Italia)", "Bundesliga (Alemania)"])
    with col_e2:
        filtro_equipo = st.text_input("🔍 Buscar por Equipo:", placeholder="Ej: Arsenal, Liverpool...")
    with col_e3:
        st.markdown("<br>", unsafe_allow_html=True)
        btn_escanear = st.button("🚀 Escanear Cuotas Reales", type="primary")

    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    with col_f1:
        filtro_tipo_mercado = st.selectbox("🎯 Mercado a Mostrar:", [
            "Todos los Mercados", "Mercado Principal (1X2)", "Líneas de Goles (Over/Under)", "Ambos Marcan (BTTS)"
        ])
    with col_f2:
        cuota_min = st.number_input("🎯 Cuota Mínima:", value=1.12, min_value=1.01, max_value=2.00, step=0.01)
        cuota_max = st.slider("🎯 Cuota Máxima Realista:", 1.20, 5.00, 3.50, step=0.10)
    with col_f3:
        ev_minimo = st.slider("📈 EV+ Mínimo (%):", 1.0, 30.0, 3.0, step=1.0)
    with col_f4:
        st.markdown("<br>", unsafe_allow_html=True)
        agrupar_casas = st.checkbox("Mejor casa por apuesta", value=True)

    col_td_api1, col_td_api2 = st.columns([2, 2])
    with col_td_api1:
        modo_td_api = st.radio(
            "⏳ Ponderación del Modelo Poisson (Escáner):",
            ["🔥 Forma Reciente (Time-Decay)", "📚 Histórico Completo (Media Clásica)"],
            horizontal=True,
            key="modo_td_api",
        )
        usar_td_api = "Time-Decay" in modo_td_api
    with col_td_api2:
        xi_api = st.slider(
            "📉 Factor Decaimiento Temporal ξ (prioridad últ. 6-8 partidos ≈ 0.010):",
            min_value=0.001,
            max_value=0.025,
            value=0.010,
            step=0.001,
            format="%.3f",
            disabled=not usar_td_api,
            key="xi_td_api",
        )

    mapa_sports = {
        "LaLiga (España)": "soccer_spain_la_liga",
        "Premier League (Inglaterra)": "soccer_epl",
        "Serie A (Italia)": "soccer_italy_serie_a",
        "Bundesliga (Alemania)": "soccer_germany_bundesliga"
    }

    if btn_escanear:
        if escaneo_apuestas_ev is None:
            st.error("No se encuentra el conector API.")
            st.session_state.df_ops_export = None
        elif "ODDS_API_KEY" not in st.secrets or not st.secrets["ODDS_API_KEY"]:
            st.error("Falta `ODDS_API_KEY` en `.streamlit/secrets.toml`.")
            st.session_state.df_ops_export = None
        else:
            with st.spinner("Escaneando casas de apuestas en vivo..."):
                ops = escaneo_apuestas_ev(
                    mapa_sports[liga_api_sel],
                    cuota_minima=cuota_min,
                    cuota_maxima=cuota_max,
                    usar_time_decay=usar_td_api,
                    xi=xi_api,
                )
                df_ops = pd.DataFrame(ops) if ops else pd.DataFrame()

                if not df_ops.empty:
                    if filtro_tipo_mercado == "Mercado Principal (1X2)":
                        df_ops = df_ops[df_ops['Tipo_Mercado'] == '1X2']
                    elif filtro_tipo_mercado == "Líneas de Goles (Over/Under)":
                        df_ops = df_ops[df_ops['Tipo_Mercado'] == 'Goles']
                    elif filtro_tipo_mercado == "Ambos Marcan (BTTS)":
                        df_ops = df_ops[df_ops['Tipo_Mercado'] == 'BTTS']

                    df_ops = df_ops[df_ops['EV_Num'] >= ev_minimo]

                    if filtro_equipo.strip():
                        df_ops = df_ops[df_ops['Partido'].str.contains(filtro_equipo, case=False, na=False)]

                    if agrupar_casas and not df_ops.empty:
                        df_ops = df_ops.sort_values(by="Cuota", ascending=False).groupby(["Partido", "Mercado"]).first().reset_index()

                    df_ops = df_ops.sort_values(by="EV_Num", ascending=False)
                    if not df_ops.empty:
                        df_ops['Nivel_Valor'] = df_ops['EV_Num'].apply(
                            lambda x: "Alto Valor (EV >10%)" if x >= 10.0 else ("Valor Medio (5-10%)" if x >= 5.0 else "Valor Moderado (<5%)")
                        )

                st.session_state.df_ops_export = df_ops
                st.session_state.liga_ops_export = liga_api_sel

                if df_ops.empty:
                    if ops:
                        st.warning("No hay selecciones que cumplan los filtros.")
                    else:
                        st.info("No se encontraron partidos/cuotas.")

    df_ops = st.session_state.get("df_ops_export")
    liga_export = st.session_state.get("liga_ops_export", liga_api_sel)
    if isinstance(df_ops, pd.DataFrame) and not df_ops.empty:
        ev_medio = df_ops['EV_Num'].mean() if 'EV_Num' in df_ops.columns else 0
        st.markdown("---")
        m_c1, m_c2, m_c3 = st.columns(3)
        m_c1.metric("Oportunidades Reales EV+", f"{len(df_ops)} apuestas")
        m_c2.metric("EV+ Promedio del Escaneo", f"+{ev_medio:.2f}%")
        m_c3.metric("Yield Esperado Estimado", f"+{ev_medio:.2f}%")
        st.markdown("---")

        cols = ['Nivel_Valor', 'Fecha / Hora', 'Partido', 'Mercado', 'Casa', 'Cuota', 'Prob_Modelo', 'Cuota_Justa', 'EV', 'Stake_Recomendado']
        st.success(f"🔥 ¡Se han detectado **{len(df_ops)} oportunidades reales multimercado**!")
        st.dataframe(df_ops[[c for c in cols if c in df_ops.columns]], use_container_width=True, hide_index=True)

        st.markdown("#### Exportar informe")
        col_dl1, col_dl2 = st.columns(2)
        nombre_base = f"escaner_EV_{_slug_archivo(liga_export)}_{datetime.now().strftime('%Y%m%d_%H%M')}"
        xlsx_bytes = generar_excel_escaner_ev(df_ops, liga_export)
        pdf_bytes = generar_pdf_escaner_ev(df_ops, liga_export)
        with col_dl1:
            st.download_button(
                "Descargar Excel (.xlsx)",
                data=xlsx_bytes,
                file_name=f"{nombre_base}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="dl_escaner_xlsx",
            )
        with col_dl2:
            st.download_button(
                "Descargar PDF (.pdf)",
                data=pdf_bytes,
                file_name=f"{nombre_base}.pdf",
                mime="application/pdf",
                key="dl_escaner_pdf",
            )

elif menu_seleccionado == "📊 Dossier de Equipo":
    st.title("📊 Dossier Analítico Avanzado de Equipo")
    st.markdown("Radiografía estadística exhaustiva: rendimiento global, desglose local/visitante, porcentajes de mercado y disciplina.")
    
    if err_partidos or df_partidos_global.empty:
        st.warning("No hay datos de partidos disponibles en la base de datos.")
    else:
        todos_equipos = sorted(pd.concat([df_partidos_global['HomeTeam'], df_partidos_global['AwayTeam']]).dropna().unique().tolist())
        
        col_sel1, col_sel2 = st.columns([2, 1])
        with col_sel1:
            equipo_elegido = st.selectbox("Selecciona un equipo para generar el dossier:", todos_equipos)
        
        if equipo_elegido:
            df_eq_partidos = df_partidos_global[
                (df_partidos_global['HomeTeam'] == equipo_elegido) | 
                (df_partidos_global['AwayTeam'] == equipo_elegido)
            ].copy()
            
            if 'Date' in df_eq_partidos.columns:
                df_eq_partidos['Date_Parsed'] = pd.to_datetime(df_eq_partidos['Date'], errors='coerce', dayfirst=True)
                df_sort = df_eq_partidos.sort_values(by='Date_Parsed', ascending=False)
            else:
                df_sort = df_eq_partidos.copy()
            
            total_partidos = len(df_eq_partidos)
            local_partidos = df_eq_partidos[df_eq_partidos['HomeTeam'] == equipo_elegido]
            vis_partidos = df_eq_partidos[df_eq_partidos['AwayTeam'] == equipo_elegido]
            
            # Goles
            gf_local = int(local_partidos['FTHG'].sum()) if not local_partidos.empty else 0
            gc_local = int(local_partidos['FTAG'].sum()) if not local_partidos.empty else 0
            gf_vis = int(vis_partidos['FTAG'].sum()) if not vis_partidos.empty else 0
            gc_vis = int(vis_partidos['FTHG'].sum()) if not vis_partidos.empty else 0
            
            goles_favor = gf_local + gf_vis
            goles_contra = gc_local + gc_vis
            media_gf = goles_favor / total_partidos if total_partidos > 0 else 0
            media_gc = goles_contra / total_partidos if total_partidos > 0 else 0
            
            # Victorias, Empates, Derrotas
            v = e = d = 0
            for _, r in df_eq_partidos.iterrows():
                is_home = (r['HomeTeam'] == equipo_elegido)
                hg, ag = r['FTHG'], r['FTAG']
                if hg == ag: e += 1
                elif (is_home and hg > ag) or (not is_home and ag > hg): v += 1
                else: d += 1

            # Porcentajes de mercado (Over 2.5, BTTS, Portería a cero)
            over_25_count = 0
            btts_count = 0
            clean_sheets = 0
            for _, r in df_eq_partidos.iterrows():
                is_home = (r['HomeTeam'] == equipo_elegido)
                hg, ag = r['FTHG'], r['FTAG']
                if (hg + ag) > 2.5: over_25_count += 1
                if hg > 0 and ag > 0: btts_count += 1
                if (is_home and ag == 0) or (not is_home and hg == 0): clean_sheets += 1
            
            p_over25 = (over_25_count / total_partidos * 100) if total_partidos > 0 else 0
            p_btts = (btts_count / total_partidos * 100) if total_partidos > 0 else 0
            p_cs = (clean_sheets / total_partidos * 100) if total_partidos > 0 else 0

            st.markdown("---")
            st.markdown(f"### 📈 Radiografía Global: **{equipo_elegido}**")
            
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Partidos", f"{total_partidos}")
            c2.metric("Balance (V-E-D)", f"{v} - {e} - {d}")
            c3.metric("Goles Favor (Media)", f"{goles_favor} ({media_gf:.2f})")
            c4.metric("Goles Contra (Media)", f"{goles_contra} ({media_gc:.2f})")
            c5.metric("Diferencia", f"{goles_favor - goles_contra:+d}")

            pdf_dossier = generar_pdf_dossier_equipo({
                "equipo": equipo_elegido,
                "partidos": total_partidos,
                "v": v, "e": e, "d": d,
                "gf": goles_favor, "gc": goles_contra,
                "media_gf": media_gf, "media_gc": media_gc,
                "p_over25": p_over25, "over_25_count": over_25_count,
                "p_btts": p_btts, "btts_count": btts_count,
                "p_cs": p_cs, "clean_sheets": clean_sheets,
                "n_local": len(local_partidos), "n_vis": len(vis_partidos),
                "gf_local": gf_local, "gc_local": gc_local,
                "gf_vis": gf_vis, "gc_vis": gc_vis,
            })
            st.download_button(
                "Descargar Informe Ejecutivo (PDF)",
                data=pdf_dossier,
                file_name=f"informe_ejecutivo_{_slug_archivo(equipo_elegido)}.pdf",
                mime="application/pdf",
                type="primary",
                key="dl_dossier_pdf",
            )
            
            # Desglose Local / Visitante
            st.markdown("---")
            st.subheader("🏠 Desglose: Local vs ✈️ Visitante")
            col_lv1, col_lv2 = st.columns(2)
            
            with col_lv1:
                st.markdown(f"#### 🏠 Rendimiento en Casa ({len(local_partidos)} partidos)")
                if not local_partidos.empty:
                    m_gf_l = gf_local / len(local_partidos)
                    m_gc_l = gc_local / len(local_partidos)
                    st.markdown(f"- **Goles Anotados (Total / Media):** `{gf_local}` (`{m_gf_l:.2f}` por partido)")
                    st.markdown(f"- **Goles Recibidos (Total / Media):** `{gc_local}` (`{m_gc_l:.2f}` por partido)")
                else:
                    st.info("Sin registros como local.")
                    
            with col_lv2:
                st.markdown(f"#### ✈️ Rendimiento Fuera ({len(vis_partidos)} partidos)")
                if not vis_partidos.empty:
                    m_gf_v = gf_vis / len(vis_partidos)
                    m_gc_v = gc_vis / len(vis_partidos)
                    st.markdown(f"- **Goles Anotados (Total / Media):** `{gf_vis}` (`{m_gf_v:.2f}` por partido)")
                    st.markdown(f"- **Goles Recibidos (Total / Media):** `{gc_vis}` (`{m_gc_v:.2f}` por partido)")
                else:
                    st.info("Sin registros como visitante.")

            # Métricas de Mercado y Tendencias
            st.markdown("---")
            st.subheader("🎯 Probabilidades y Tendencias de Mercado")
            m1, m2, m3 = st.columns(3)
            m1.metric("Porcentaje Más de 2.5 Goles", f"{p_over25:.1f}%", f"{over_25_count}/{total_partidos} partidos")
            m2.metric("Porcentaje Ambos Marcan (BTTS)", f"{p_btts:.1f}%", f"{btts_count}/{total_partidos} partidos")
            m3.metric("Porcentaje Porterías a Cero", f"{p_cs:.1f}%", f"{clean_sheets}/{total_partidos} partidos")

            st.markdown("---")
            col_d1, col_d2 = st.columns(2)
            
            with col_d1:
                st.markdown("### 🕒 Historial de Últimos Resultados")
                if not df_sort.empty:
                    for idx, row in df_sort.head(6).iterrows():
                        l, v_team = row['HomeTeam'], row['AwayTeam']
                        gl, gv = row['FTHG'], row['FTAG']
                        l_fmt = f"**{l}**" if l == equipo_elegido else l
                        v_fmt = f"**{v_team}**" if v_team == equipo_elegido else v_team
                        st.caption(f"📅 {row.get('Date', 'N/A')} | {l_fmt} {gl} - {gv} {v_fmt}")
                else:
                    st.info("Sin partidos recientes registrados.")
            
            with col_d2:
                st.markdown("### ⭐ Plantilla y Estadística Individual")
                if not err_jugadores and not df_jugadores.empty:
                    if 'Equipo' in df_jugadores.columns:
                        df_jug_eq = df_jugadores[df_jugadores['Equipo'].str.contains(equipo_elegido, case=False, na=False)].copy()
                    else:
                        df_jug_eq = pd.DataFrame()

                    # Mapear Gls/Ast de FBref a Goles/Asistencias si es necesario
                    if 'Gls' in df_jug_eq.columns and 'Goles' not in df_jug_eq.columns:
                        df_jug_eq['Goles'] = df_jug_eq['Gls']
                    if 'Ast' in df_jug_eq.columns and 'Asistencias' not in df_jug_eq.columns:
                        df_jug_eq['Asistencias'] = df_jug_eq['Ast']

                    cols_mostrar = ['Jugador', 'Pos', 'Goles', 'Asistencias']
                    tiene_columnas = all(c in df_jug_eq.columns for c in cols_mostrar)

                    if not df_jug_eq.empty and tiene_columnas:
                        st.dataframe(df_jug_eq[cols_mostrar], use_container_width=True, hide_index=True)
                    else:
                        st.info("Las estadísticas de jugadores no están disponibles en este momento.")
                else:
                    st.info("Base de datos de jugadores no disponible.")

elif menu_seleccionado == "⚔️ Comparador H2H (Cara a Cara)":
    st.title("⚔️ Comparador Directo H2H Pro & Cuotas Justas")
    st.markdown("Analiza el enfrentamiento aislando el rendimiento del Local en su estadio vs el Visitante a domicilio, proyectando cuotas justas por Poisson.")
    
    if err_partidos or df_partidos_global.empty:
        st.warning("Datos de partidos no disponibles.")
    else:
        equipos_h2h = sorted(pd.concat([df_partidos_global['HomeTeam'], df_partidos_global['AwayTeam']]).dropna().unique().tolist())
        col1, col2 = st.columns(2)
        with col1: eq1 = st.selectbox("🏠 Equipo Local (A):", equipos_h2h, index=0)
        with col2: eq2 = st.selectbox("✈️ Equipo Visitante (B):", equipos_h2h, index=1 if len(equipos_h2h)>1 else 0)

        if eq1 and eq2 and eq1 != eq2:
            df_eq1_home = df_partidos_global[df_partidos_global['HomeTeam'] == eq1].copy()
            df_eq2_away = df_partidos_global[df_partidos_global['AwayTeam'] == eq2].copy()
            
            def get_form(df, team, is_home):
                if df.empty: return []
                if 'Date' in df.columns:
                    df['Date_Parsed'] = pd.to_datetime(df['Date'], errors='coerce', dayfirst=True)
                    df = df.sort_values(by='Date_Parsed', ascending=False)
                form = []
                for _, r in df.head(5).iterrows():
                    hg, ag = r['FTHG'], r['FTAG']
                    if hg == ag: form.append('E')
                    elif (is_home and hg > ag) or (not is_home and ag > hg): form.append('V')
                    else: form.append('D')
                return form

            forma_eq1 = get_form(df_eq1_home, eq1, True)
            forma_eq2 = get_form(df_eq2_away, eq2, False)
            
            st.markdown("---")
            st.subheader("🎯 Proyección Matemática del Partido (Modelo Poisson)")
            st.markdown(f"Cálculo estadístico cruzando **la fuerza ofensiva/defensiva de {eq1} (solo en casa)** contra **{eq2} (solo fuera)**.")
            
            xg_l = ((df_eq1_home['FTHG'].mean() if not df_eq1_home.empty else 1.5) + (df_eq2_away['FTHG'].mean() if not df_eq2_away.empty else 1.5)) / 2
            xg_v = ((df_eq2_away['FTAG'].mean() if not df_eq2_away.empty else 1.0) + (df_eq1_home['FTAG'].mean() if not df_eq1_home.empty else 1.0)) / 2
            
            prob_h = prob_d = prob_a = prob_o25 = prob_btts = tot_prob = 0
            for i in range(7):
                for j in range(7):
                    pr = calcular_probabilidad_poisson(xg_l, i) * calcular_probabilidad_poisson(xg_v, j)
                    tot_prob += pr
                    if i > j: prob_h += pr
                    elif i == j: prob_d += pr
                    else: prob_a += pr
                    if (i + j) > 2.5: prob_o25 += pr
                    if i > 0 and j > 0: prob_btts += pr
                    
            norm = tot_prob if tot_prob > 0 else 1.0
            
            p_c1, p_c2, p_c3 = st.columns(3)
            p_c1.metric("Cuota Justa Local (1)", f"{1/(prob_h/norm):.2f}", f"Prob: {(prob_h/norm*100):.1f}%")
            p_c2.metric("Cuota Justa +2.5 Goles", f"{1/(prob_o25/norm):.2f}", f"Prob: {(prob_o25/norm*100):.1f}%", delta_color="off")
            p_c3.metric("Cuota Justa BTTS (Sí)", f"{1/(prob_btts/norm):.2f}", f"Prob: {(prob_btts/norm*100):.1f}%", delta_color="off")

            st.markdown("---")
            st.subheader("📊 Rendimiento Estricto (Local vs Visitante)")
            
            gf_l = df_eq1_home['FTHG'].mean() if not df_eq1_home.empty else 0
            gc_l = df_eq1_home['FTAG'].mean() if not df_eq1_home.empty else 0
            cor_l = (df_eq1_home['HC'] + df_eq1_home['AC']).mean() if 'HC' in df_eq1_home.columns else 0
            
            gf_v = df_eq2_away['FTAG'].mean() if not df_eq2_away.empty else 0
            gc_v = df_eq2_away['FTHG'].mean() if not df_eq2_away.empty else 0
            cor_v = (df_eq2_away['HC'] + df_eq2_away['AC']).mean() if 'AC' in df_eq2_away.columns else 0

            c_st1, c_st2 = st.columns(2)
            with c_st1:
                st.markdown(f"#### 🏠 {eq1} (Como Local)")
                st.write(f"**Forma (Últimos 5):** {' - '.join(forma_eq1) if forma_eq1 else 'N/D'}")
                st.write(f"**Goles Anotados / Recibidos:** `{gf_l:.2f}` / `{gc_l:.2f}`")
                st.write(f"**Media Córners Totales:** `{cor_l:.1f}`")
            with c_st2:
                st.markdown(f"#### ✈️ {eq2} (Como Visitante)")
                st.write(f"**Forma (Últimos 5):** {' - '.join(forma_eq2) if forma_eq2 else 'N/D'}")
                st.write(f"**Goles Anotados / Recibidos:** `{gf_v:.2f}` / `{gc_v:.2f}`")
                st.write(f"**Media Córners Totales:** `{cor_v:.1f}`")
            
            h2h_df = df_partidos_global[((df_partidos_global['HomeTeam'] == eq1) & (df_partidos_global['AwayTeam'] == eq2)) | ((df_partidos_global['HomeTeam'] == eq2) & (df_partidos_global['AwayTeam'] == eq1))].copy()
            
            st.markdown("---")
            st.subheader(f"⚔️ Últimos Enfrentamientos Directos")
            if not h2h_df.empty:
                if 'Date' in h2h_df.columns:
                    h2h_df['Date_Parsed'] = pd.to_datetime(h2h_df['Date'], errors='coerce', dayfirst=True)
                    h2h_df = h2h_df.sort_values(by='Date_Parsed', ascending=False)
                for _, row in h2h_df.head(5).iterrows():
                    l_h2h, v_h2h = row['HomeTeam'], row['AwayTeam']
                    gl_h2h, gv_h2h = row['FTHG'], row['FTAG']
                    res_str = f"📅 {row.get('Date', 'N/A')} | **{l_h2h}** {gl_h2h} - {gv_h2h} **{v_h2h}**"
                    st.info(res_str)
            else:
                st.warning("No hay enfrentamientos directos recientes en la base de datos.")
        elif eq1 == eq2:
            st.warning("Selecciona dos equipos distintos para realizar la comparativa.")

elif menu_seleccionado == "🎯 Radar de Tendencias":
    st.title("🎯 Radar de Tendencias Avanzado")
    st.markdown("Escaneo inteligente de rachas con filtros por **condición de campo (Local/Visitante/Global)** y **umbral de aciertos**.")
    
    col_rf1, col_rf2 = st.columns(2)
    with col_rf1:
        filtro_condicion = st.radio("🏠 Condición de Campo:", ["Global (Todos los partidos)", "Solo como Local 🏠", "Solo como Visitante ✈️"], horizontal=True)
    with col_rf2:
        umbral_racha = st.select_slider("🔥 Umbral Mínimo de Racha:", options=[3, 4, 5], value=4, format_func=lambda x: f"Mínimo {x}/5 aciertos")

    if err_partidos or df_partidos.empty: st.warning("No hay suficientes datos.")
    else:
        equipos = pd.concat([df_partidos['HomeTeam'], df_partidos['AwayTeam']]).dropna().unique()
        tend_over, tend_btts, tend_ht, tend_corners = [], [], [], []
        
        for eq in equipos:
            df_eq = df_partidos[df_partidos['HomeTeam'] == eq].copy() if "Local" in filtro_condicion else (df_partidos[df_partidos['AwayTeam'] == eq].copy() if "Visitante" in filtro_condicion else df_partidos[(df_partidos['HomeTeam'] == eq) | (df_partidos['AwayTeam'] == eq)].copy())
            if 'Date' in df_eq.columns:
                df_eq['Date_Parsed'] = pd.to_datetime(df_eq['Date'], errors='coerce', dayfirst=True)
                df_eq = df_eq.sort_values(by='Date_Parsed', ascending=False)
            
            u5 = df_eq.head(5)
            if len(u5) >= 3:
                if 'FTHG' in u5.columns and 'FTAG' in u5.columns:
                    res_o = ((u5['FTHG'] + u5['FTAG']) > 2.5).tolist()
                    if sum(res_o) >= umbral_racha: tend_over.append({"Equipo": eq, "Aciertos": f"{sum(res_o)}/{len(u5)}", "Winrate": f"{(sum(res_o)/len(u5))*100:.0f}%", "Forma": "".join(["✅" if x else "❌" for x in res_o])})
                    
                    res_b = ((u5['FTHG'] > 0) & (u5['FTAG'] > 0)).tolist()
                    if sum(res_b) >= umbral_racha: tend_btts.append({"Equipo": eq, "Aciertos": f"{sum(res_b)}/{len(u5)}", "Winrate": f"{(sum(res_b)/len(u5))*100:.0f}%", "Forma": "".join(["✅" if x else "❌" for x in res_b])})
                
                if 'HTHG' in u5.columns and 'HTAG' in u5.columns:
                    res_ht = ((u5['HTHG'] + u5['HTAG']) > 0).tolist()
                    if sum(res_ht) >= umbral_racha: tend_ht.append({"Equipo": eq, "Aciertos": f"{sum(res_ht)}/{len(u5)}", "Winrate": f"{(sum(res_ht)/len(u5))*100:.0f}%", "Forma": "".join(["✅" if x else "❌" for x in res_ht])})
                
                if 'HC' in u5.columns and 'AC' in u5.columns:
                    res_c = ((u5['HC'] + u5['AC']) > 8.5).tolist()
                    if sum(res_c) >= umbral_racha: tend_corners.append({"Equipo": eq, "Aciertos": f"{sum(res_c)}/{len(u5)}", "Winrate": f"{(sum(res_c)/len(u5))*100:.0f}%", "Forma": "".join(["✅" if x else "❌" for x in res_c])})

        tab1, tab2, tab3, tab4 = st.tabs(["🔥 Más de 2.5 Goles", "🎯 Ambos Marcan (BTTS)", "⏱️ Más 0.5 Goles 1ª Parte", "🚩 Córners (+8.5 Totales)"])
        with tab1:
            if tend_over:
                st.dataframe(pd.DataFrame(tend_over).sort_values(by="Winrate", ascending=False), use_container_width=True, hide_index=True)
            else:
                st.info("Sin rachas destacadas.")
        with tab2:
            if tend_btts:
                st.dataframe(pd.DataFrame(tend_btts).sort_values(by="Winrate", ascending=False), use_container_width=True, hide_index=True)
            else:
                st.info("Sin rachas destacadas.")
        with tab3:
            if tend_ht:
                st.dataframe(pd.DataFrame(tend_ht).sort_values(by="Winrate", ascending=False), use_container_width=True, hide_index=True)
            else:
                st.info("Sin rachas destacadas.")
        with tab4:
            if tend_corners:
                st.dataframe(pd.DataFrame(tend_corners).sort_values(by="Winrate", ascending=False), use_container_width=True, hide_index=True)
            else:
                st.info("Sin rachas destacadas.")

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
    st.title("🧮 Módulo Pro: Hándicap Asiático & Value Betting")
    st.markdown("Calculadora de totales asiáticos, escáner histórico ATS (Against The Spread) y Matriz Completa de Cuotas Justas Poisson.")
    
    sub_tab1, sub_tab2, sub_tab3 = st.tabs(["📊 Calculadora Asiática Total", "📈 Histórico ATS (Against The Spread)", "🤖 Matriz de Cobertura Poisson"])
    
    with sub_tab1:
        st.subheader("📊 Calculadora Universal (Equipos y Goles)")
        tipo_ah = st.radio("Selecciona el Tipo de Mercado Asiático:", ["Hándicap de Equipo (ej. -0.75, +0.25)", "Línea de Goles Asiática (Over/Under ej. 2.25, 2.75)"], horizontal=True)
        
        c_c1, c_c2 = st.columns(2)
        with c_c1: stake_ha = st.number_input("Cantidad Apostada (€):", value=10.0, step=5.0)
        with c_c2: cuota_ha = st.number_input("Cuota de la Casa:", value=1.95, step=0.05)
        
        st.markdown("---")
        if "Equipo" in tipo_ah:
            linea_ha = st.selectbox("Línea de Hándicap Seleccionada:", [
                "Hándicap 0.0 (Apuesta Sin Empate)", "Hándicap -0.25 (Asiático -0.25)", "Hándicap -0.50 (Victoria Simple)", "Hándicap -0.75 (Asiático -0.75)", "Hándicap -1.00 (Victoria por 2+)",
                "Hándicap +0.25 (Asiático +0.25)", "Hándicap +0.50 (Doble Oportunidad X2)", "Hándicap +0.75 (Asiático +0.75)", "Hándicap +1.00 (Puede perder de 1)"
            ])
            ganancia_total = round(stake_ha * cuota_ha, 2)
            beneficio_neto = round(ganancia_total - stake_ha, 2)

            if "0.0" in linea_ha:
                st.success(f"✅ **Si tu equipo gana:** Acierto Total. Cobras **{ganancia_total} €** (Beneficio: +{beneficio_neto} €).")
                st.info(f"🔄 **Si hay empate:** Apuesta Nula (Void). Te devuelven tus **{stake_ha} €**.")
                st.error(f"❌ **Si tu equipo pierde:** Apuesta Perdida (0 €).")
            elif "-0.25" in linea_ha:
                st.success(f"✅ **Si tu equipo gana:** Acierto Total. Cobras **{ganancia_total} €**.")
                st.warning(f"📉 **Si hay empate:** Medio Acierto Nulo / Media Pérdida. Recuperas **{stake_ha / 2} €**.")
                st.error(f"❌ **Si tu equipo pierde:** Apuesta Perdida (0 €).")
            elif "-0.50" in linea_ha:
                st.success(f"✅ **Si tu equipo gana:** Acierto Total. Cobras **{ganancia_total} €**.")
                st.error(f"❌ **Si empatan o pierden:** Apuesta Perdida (0 €).")
            elif "-0.75" in linea_ha:
                st.success(f"✅ **Si gana por 2 o más goles:** Acierto Total. Cobras **{ganancia_total} €**.")
                st.warning(f"📈 **Si gana por SOLO 1 gol:** Medio Acierto. Ganas la mitad del beneficio: **{round(stake_ha + (beneficio_neto/2), 2)} €**.")
                st.error(f"❌ **Si empatan o pierden:** Apuesta Perdida (0 €).")
            elif "-1.00" in linea_ha:
                st.success(f"✅ **Si gana por 2 o más goles:** Acierto Total. Cobras **{ganancia_total} €**.")
                st.info(f"🔄 **Si gana por SOLO 1 gol:** Apuesta Nula (Void). Te devuelven tus **{stake_ha} €**.")
                st.error(f"❌ **Si empatan o pierden:** Apuesta Perdida (0 €).")
            elif "+0.25" in linea_ha:
                st.success(f"✅ **Si tu equipo gana:** Acierto Total. Cobras **{ganancia_total} €**.")
                st.warning(f"📈 **Si hay empate:** Medio Acierto. Ganas la mitad del beneficio: **{round(stake_ha + (beneficio_neto/2), 2)} €**.")
                st.error(f"❌ **Si tu equipo pierde:** Apuesta Perdida (0 €).")
            elif "+0.50" in linea_ha:
                st.success(f"✅ **Si gana o empata:** Acierto Total. Cobras **{ganancia_total} €**.")
                st.error(f"❌ **Si tu equipo pierde:** Apuesta Perdida (0 €).")
            elif "+0.75" in linea_ha:
                st.success(f"✅ **Si gana o empata:** Acierto Total. Cobras **{ganancia_total} €**.")
                st.warning(f"📉 **Si pierde por SOLO 1 gol:** Media Pérdida. Recuperas la mitad: **{stake_ha / 2} €**.")
                st.error(f"❌ **Si pierde por 2 o más goles:** Apuesta Perdida (0 €).")
            elif "+1.00" in linea_ha:
                st.success(f"✅ **Si gana o empata:** Acierto Total. Cobras **{ganancia_total} €**.")
                st.info(f"🔄 **Si pierde por SOLO 1 gol:** Apuesta Nula (Void). Te devuelven tus **{stake_ha} €**.")
                st.error(f"❌ **Si pierde por 2 o más goles:** Apuesta Perdida (0 €).")
        else:
            linea_goles = st.selectbox("Línea de Goles Asiática:", ["Over 2.0 / Under 2.0", "Over 2.25 / Under 2.25", "Over 2.5 / Under 2.5", "Over 2.75 / Under 2.75", "Over 3.0 / Under 3.0"])
            ou_sel = st.radio("Tu Elección:", ["Apostar al OVER (Más de)", "Apostar al UNDER (Menos de)"], horizontal=True)
            ganancia_total = round(stake_ha * cuota_ha, 2)
            beneficio_neto = round(ganancia_total - stake_ha, 2)

            if "2.0" in linea_goles and "2.25" not in linea_goles:
                if "OVER" in ou_sel:
                    st.success(f"✅ **3 o más goles:** Acierto Total. Cobras **{ganancia_total} €**.")
                    st.info(f"🔄 **Exactamente 2 goles:** Nulo. Devuelven **{stake_ha} €**.")
                    st.error(f"❌ **0 o 1 gol:** Perdida (0 €).")
                else:
                    st.success(f"✅ **0 o 1 gol:** Acierto Total. Cobras **{ganancia_total} €**.")
                    st.info(f"🔄 **Exactamente 2 goles:** Nulo. Devuelven **{stake_ha} €**.")
                    st.error(f"❌ **3 o más goles:** Perdida (0 €).")
            elif "2.25" in linea_goles:
                if "OVER" in ou_sel:
                    st.success(f"✅ **3 o más goles:** Acierto Total. Cobras **{ganancia_total} €**.")
                    st.warning(f"📉 **Exactamente 2 goles:** Media Pérdida. Recuperas **{stake_ha / 2} €**.")
                    st.error(f"❌ **0 o 1 gol:** Perdida (0 €).")
                else:
                    st.success(f"✅ **0 o 1 gol:** Acierto Total. Cobras **{ganancia_total} €**.")
                    st.warning(f"📈 **Exactamente 2 goles:** Medio Acierto. Ganas la mitad del beneficio: **{round(stake_ha + (beneficio_neto/2), 2)} €**.")
                    st.error(f"❌ **3 o más goles:** Perdida (0 €).")
            elif "2.5" in linea_goles and "2.75" not in linea_goles:
                if "OVER" in ou_sel:
                    st.success(f"✅ **3 o más goles:** Acierto Total. Cobras **{ganancia_total} €**.")
                    st.error(f"❌ **0, 1 o 2 goles:** Perdida (0 €).")
                else:
                    st.success(f"✅ **0, 1 o 2 goles:** Acierto Total. Cobras **{ganancia_total} €**.")
                    st.error(f"❌ **3 o más goles:** Perdida (0 €).")
            elif "2.75" in linea_goles:
                if "OVER" in ou_sel:
                    st.success(f"✅ **4 o más goles:** Acierto Total. Cobras **{ganancia_total} €**.")
                    st.warning(f"📈 **Exactamente 3 goles:** Medio Acierto. Ganas la mitad del beneficio: **{round(stake_ha + (beneficio_neto/2), 2)} €**.")
                    st.error(f"❌ **0, 1 o 2 goles:** Perdida (0 €).")
                else:
                    st.success(f"✅ **0, 1 o 2 goles:** Acierto Total. Cobras **{ganancia_total} €**.")
                    st.warning(f"📉 **Exactamente 3 goles:** Media Pérdida. Recuperas **{stake_ha / 2} €**.")
                    st.error(f"❌ **4 o más goles:** Perdida (0 €).")
            elif "3.0" in linea_goles:
                if "OVER" in ou_sel:
                    st.success(f"✅ **4 o más goles:** Acierto Total. Cobras **{ganancia_total} €**.")
                    st.info(f"🔄 **Exactamente 3 goles:** Nulo. Devuelven **{stake_ha} €**.")
                    st.error(f"❌ **0, 1 o 2 goles:** Perdida (0 €).")
                else:
                    st.success(f"✅ **0, 1 o 2 goles:** Acierto Total. Cobras **{ganancia_total} €**.")
                    st.info(f"🔄 **Exactamente 3 goles:** Nulo. Devuelven **{stake_ha} €**.")
                    st.error(f"❌ **4 o más goles:** Perdida (0 €).")

    with sub_tab2:
        st.subheader("📈 Escáner de Rendimiento ATS (Against The Spread)")
        st.markdown("Evalúa estadísticamente qué porcentaje de veces un equipo ha cubierto una línea asiática específica en sus últimos partidos.")
        if not df_partidos.empty:
            equipos_ats = sorted(pd.concat([df_partidos['HomeTeam'], df_partidos['AwayTeam']]).dropna().unique())
            c_ats1, c_ats2, c_ats3 = st.columns(3)
            with c_ats1: eq_ats = st.selectbox("Equipo a Evaluar:", equipos_ats)
            with c_ats2: cond_ats = st.selectbox("Condición:", ["Global (Local y Visitante)", "Solo Local", "Solo Visitante"])
            with c_ats3: val_linea_ats = st.number_input("Línea Hándicap Asiático (ej. -0.75, +1.0):", value=-0.50, step=0.25)
            
            df_ats = df_partidos.copy()
            if cond_ats == "Solo Local": df_ats = df_ats[df_ats['HomeTeam'] == eq_ats]
            elif cond_ats == "Solo Visitante": df_ats = df_ats[df_ats['AwayTeam'] == eq_ats]
            else: df_ats = df_ats[(df_ats['HomeTeam'] == eq_ats) | (df_ats['AwayTeam'] == eq_ats)]
            
            if 'Date' in df_ats.columns:
                df_ats['Date_Parsed'] = pd.to_datetime(df_ats['Date'], errors='coerce', dayfirst=True)
                df_ats = df_ats.sort_values(by='Date_Parsed', ascending=False).head(20) # Ultimos 20
                
            resultados_ats = {"Verde (Win)": 0, "Medio Verde": 0, "Nulo (Push)": 0, "Medio Rojo": 0, "Rojo (Loss)": 0}
            
            for _, row in df_ats.iterrows():
                if row['HomeTeam'] == eq_ats: g_fav, g_con = row['FTHG'], row['FTAG']
                else: g_fav, g_con = row['FTAG'], row['FTHG']
                
                diff = (g_fav + val_linea_ats) - g_con
                if diff > 0.25: resultados_ats["Verde (Win)"] += 1
                elif diff == 0.25: resultados_ats["Medio Verde"] += 1
                elif diff == 0.0: resultados_ats["Nulo (Push)"] += 1
                elif diff == -0.25: resultados_ats["Medio Rojo"] += 1
                else: resultados_ats["Rojo (Loss)"] += 1
                
            total_eval = sum(resultados_ats.values())
            if total_eval > 0:
                win_rate = ((resultados_ats["Verde (Win)"] + (resultados_ats["Medio Verde"] * 0.5)) / total_eval) * 100
                st.markdown("---")
                st.metric(f"Tasa de Cobertura ATS ({val_linea_ats:+.2f}) - Últ. {total_eval} Partidos", f"{win_rate:.1f}%", f"{resultados_ats['Verde (Win)']} V | {resultados_ats['Medio Verde']} MV | {resultados_ats['Nulo (Push)']} N | {resultados_ats['Medio Rojo']} MR | {resultados_ats['Rojo (Loss)']} R")
                
                df_res_ats = pd.DataFrame([resultados_ats])
                st.dataframe(df_res_ats, use_container_width=True, hide_index=True)
            else:
                st.info("No hay suficientes datos para evaluar este equipo.")

    with sub_tab3:
        st.subheader("🤖 Matriz Poisson de Cobertura de Líneas")
        st.markdown("Cálculo de Cuotas Justas para todo el abanico del Hándicap Asiático basado en el xG cruzado.")
        if not df_partidos.empty and 'HomeTeam' in df_partidos.columns:
            eqs_ha = sorted(pd.concat([df_partidos['HomeTeam'], df_partidos['AwayTeam']]).dropna().unique())
            c_ha1, c_ha2 = st.columns(2)
            with c_ha1: eq_l_ha = st.selectbox("Equipo Local (Para Hándicap):", eqs_ha, index=0, key="ha_loc")
            with c_ha2: eq_v_ha = st.selectbox("Rival Visitante:", eqs_ha, index=1 if len(eqs_ha)>1 else 0, key="ha_vis")
            
            if eq_l_ha != eq_v_ha:
                p_l = df_partidos[df_partidos['HomeTeam'] == eq_l_ha]
                p_v = df_partidos[df_partidos['AwayTeam'] == eq_v_ha]
                
                gf_l = p_l['FTHG'].mean() if not p_l.empty else 1.5
                gc_v = p_v['FTHG'].mean() if not p_v.empty else 1.5
                gf_v = p_v['FTAG'].mean() if not p_v.empty else 1.0
                gc_l = p_l['FTAG'].mean() if not p_l.empty else 1.0
                
                xg_local = (gf_l + gc_v) / 2
                xg_vis = (gf_v + gc_l) / 2
                st.info(f"💡 **Fuerza Calculada (xG):** {eq_l_ha} (`{xg_local:.2f}`) vs {eq_v_ha} (`{xg_vis:.2f}`)")
                
                lineas_a_calcular = [-1.5, -1.0, -0.75, -0.5, -0.25, 0.0, 0.25, 0.5, 0.75, 1.0, 1.5]
                resultados_matriz = []
                
                for linea in lineas_a_calcular:
                    ev_win = ev_half_win = ev_push = ev_half_loss = ev_loss = 0
                    for i in range(7):
                        for j in range(7):
                            pr = calcular_probabilidad_poisson(xg_local, i) * calcular_probabilidad_poisson(xg_vis, j)
                            diff = (i + linea) - j
                            if diff > 0.25: ev_win += pr
                            elif diff == 0.25: ev_half_win += pr
                            elif diff == 0.0: ev_push += pr
                            elif diff == -0.25: ev_half_loss += pr
                            else: ev_loss += pr
                    
                    # Probabilidad efectiva de cobro (Win + 0.5 * Half Win)
                    prob_efectiva = ev_win + (ev_half_win * 0.5)
                    cuota_justa = 1 / prob_efectiva if prob_efectiva > 0 else 0
                    
                    resultados_matriz.append({
                        "Línea Hándicap": f"{linea:+.2f}",
                        "Prob. Efectiva": f"{prob_efectiva*100:.1f}%",
                        "Cuota Justa": f"{cuota_justa:.2f}",
                        "Win Puro": f"{ev_win*100:.1f}%",
                        "Nulo/Medios": f"{(ev_half_win+ev_push+ev_half_loss)*100:.1f}%",
                        "Loss Puro": f"{ev_loss*100:.1f}%"
                    })
                    
                st.dataframe(pd.DataFrame(resultados_matriz), use_container_width=True, hide_index=True)

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