import sqlite3
import json
import re
import html
import unicodedata
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
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    HRFlowable,
    KeepInFrame,
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
    "F1": "Ligue 1 (Francia)",
    "Champions League": "Champions League (UEFA)",
    "Europa League": "Europa League (UEFA)",
    "Conference League": "Conference League (UEFA)",
    "uefa.champions": "Champions League (UEFA)",
    "uefa.europa": "Europa League (UEFA)",
    "uefa.europa.conf": "Conference League (UEFA)",
    "UCL": "Champions League (UEFA)",
    "UEL": "Europa League (UEFA)",
    "UECL": "Conference League (UEFA)",
}
MAPA_INV = {v: k for k, v in MAPA_LIGAS.items()}
MAPA_INV.update(
    {
        "Champions League (UEFA)": "UCL",
        "Europa League (UEFA)": "UEL",
        "Conference League (UEFA)": "UECL",
    }
)

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

OPCION_COMPETICION_GLOBAL = "Todas (Consolidado Global)"
OPCIONES_COMPETICION = [
    OPCION_COMPETICION_GLOBAL, "LaLiga", "Premier League", "Serie A", "Bundesliga", "Ligue 1",
    "Champions League", "Europa League", "Conference League",
]

# Mapa determinista de aislamiento relacional Jugador -> Club activo para evitar contaminación cruzada
_MAPA_AISLAMIENTO_ENTIDAD_CLUB = {
    "Gianluigi Donnarumma": "Paris SG",
    "Enzo Fernández": "Chelsea",
    "Antoine Semenyo": "Bournemouth",
    "Elliot Anderson": "Nott'm Forest",
    "Iliman Ndiaye": "Everton",
    "Marc Guéhi": "Crystal Palace",
    "Rayan Cherki": "Lyon",
    "Ayyoub Bouaddi": "Lille",
    "Gerónimo Rulli": "Marseille",
    "Marcus Bettinelli": "Chelsea",
    "Max-Edgar Chabot": "Stuttgart",
    "Rayan Aït-Nouri": "Wolves",
    "Allan Elias": "Sporting CP",
}

# Vista consolidada: aislamiento relacional por ventana (ROW_NUMBER OVER PARTITION BY Jugador)
# para vincular el club principal activo y sumar métricas sin producto cartesiano entre clubes.
_SQL_JUGADORES_GLOBAL = """
WITH club_activo AS (
    SELECT Jugador, Equipo,
           ROW_NUMBER() OVER (
               PARTITION BY Jugador
               ORDER BY minutos DESC, partidos_jugados DESC, rowid DESC
           ) AS rn
    FROM estadisticas_jugadores
)
SELECT e.Jugador,
       ca.Equipo AS Equipo,
       MAX(e.posicion) AS posicion,
       SUM(e.minutos) AS minutos, SUM(e.goles) AS goles, SUM(e.asistencias) AS asistencias,
       SUM(e.tiros_totales) AS tiros_totales, SUM(e.tiros_a_puerta) AS tiros_a_puerta,
       SUM(e.faltas_cometidas) AS faltas_cometidas, SUM(e.Faltas_Recibidas) AS Faltas_Recibidas,
       SUM(e.tarjetas_amarillas) AS tarjetas_amarillas, SUM(e.tarjetas_rojas) AS tarjetas_rojas,
       SUM(e.partidos_jugados) AS partidos_jugados, SUM(e.titularidades) AS titularidades,
       SUM(e.xg) AS xg,
       SUM(e.fueras_juego) AS fueras_juego, SUM(e.entradas_tackles) AS entradas_tackles,
       SUM(e.Pases_Intentados) AS Pases_Intentados, SUM(e.Pases_Completados) AS Pases_Completados,
       SUM(e.Pases_Clave) AS Pases_Clave, SUM(e.Entradas) AS Entradas,
       SUM(e.Entradas_Ganadas) AS Entradas_Ganadas, SUM(e.Intercepciones) AS Intercepciones
FROM estadisticas_jugadores e
INNER JOIN club_activo ca ON e.Jugador = ca.Jugador AND ca.rn = 1
GROUP BY e.Jugador, ca.Equipo
"""

_SQL_JUGADORES_POR_COMP = """
WITH deduplicado_comp AS (
    SELECT *,
           ROW_NUMBER() OVER (
               PARTITION BY Jugador, competicion
               ORDER BY minutos DESC, partidos_jugados DESC, rowid DESC
           ) AS rn
    FROM estadisticas_jugadores
    WHERE competicion = ?
)
SELECT * FROM deduplicado_comp WHERE rn = 1
"""


def _sanitizar_aislamiento_jugadores(df_jug: pd.DataFrame, competicion: str = OPCION_COMPETICION_GLOBAL) -> pd.DataFrame:
    """Garantiza aislamiento estricto Jugador-Club-Competición sin proyecciones cruzadas ni duplicados."""
    if df_jug is None or df_jug.empty:
        return pd.DataFrame() if df_jug is None else df_jug
    df_out = df_jug.copy()
    if "rn" in df_out.columns:
        df_out = df_out.drop(columns=["rn"])
    if "Jugador" in df_out.columns and "Equipo" in df_out.columns:
        df_out["Jugador"] = df_out["Jugador"].astype(str).str.strip()
        df_out["Equipo"] = df_out["Equipo"].astype(str).str.strip()
        mask_override = df_out["Jugador"].isin(_MAPA_AISLAMIENTO_ENTIDAD_CLUB.keys())
        if mask_override.any():
            df_out.loc[mask_override, "Equipo"] = df_out.loc[mask_override, "Jugador"].map(_MAPA_AISLAMIENTO_ENTIDAD_CLUB)
    if competicion != OPCION_COMPETICION_GLOBAL and "competicion" in df_out.columns:
        df_out = df_out[df_out["competicion"].astype(str).str.strip() == str(competicion).strip()]
    if "Jugador" in df_out.columns:
        col_min = "minutos" if "minutos" in df_out.columns else ("Minutos" if "Minutos" in df_out.columns else None)
        if col_min:
            df_out[col_min] = pd.to_numeric(df_out[col_min], errors="coerce").fillna(0.0)
            df_out = df_out.sort_values(by=col_min, ascending=False, kind="mergesort")
        df_out = df_out.drop_duplicates(subset=["Jugador"], keep="first").reset_index(drop=True)
    return df_out


def cargar_jugadores_competicion(conn, competicion: str = OPCION_COMPETICION_GLOBAL) -> pd.DataFrame:
    """Jugadores de una competición concreta, o agregado global (suma de todas) con aislamiento relacional."""
    cols = [r[1] for r in conn.execute("PRAGMA table_info(estadisticas_jugadores)")]
    if "competicion" not in cols:  # BD aún sin migrar
        df_raw = pd.read_sql("SELECT * FROM estadisticas_jugadores", conn)
        return _sanitizar_aislamiento_jugadores(df_raw, competicion)
    if competicion == OPCION_COMPETICION_GLOBAL:
        df_raw = pd.read_sql(_SQL_JUGADORES_GLOBAL, conn)
        return _sanitizar_aislamiento_jugadores(df_raw, competicion)
    df_raw = pd.read_sql(_SQL_JUGADORES_POR_COMP, conn, params=(competicion,))
    return _sanitizar_aislamiento_jugadores(df_raw, competicion)


@st.cache_data(ttl=3600)
def cargar_datos_jugadores():
    if 'estadisticas_jugadores' not in inspeccionar_tablas_bd(): return pd.DataFrame(), "Falta la tabla."
    with sqlite3.connect(RUTA_BD) as conn:
        try: return cargar_jugadores_competicion(conn, OPCION_COMPETICION_GLOBAL), None
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


# ══════════════════════════════════════════════════════════════════════════════
# RESOLVER CANÓNICO GLOBAL DE COLEGIADOS & POOL TRANSVERSAL
# ══════════════════════════════════════════════════════════════════════════════
def _despojar_diacriticos_arb(txt: str) -> str:
    """Elimina acentos y diacríticos preservando caracteres alfanuméricos."""
    if not txt:
        return ""
    return "".join(
        c for c in unicodedata.normalize("NFD", str(txt).lower())
        if unicodedata.category(c) != "Mn"
    ).strip()


_MAPA_CANONICO_ARBITROS = {
    # LaLiga (España)
    "j gil manzano": "Jesús Gil Manzano",
    "jesus gil manzano": "Jesús Gil Manzano",
    "gil manzano": "Jesús Gil Manzano",
    "gil manzano jesus": "Jesús Gil Manzano",
    "j m sanchez martinez": "José María Sánchez Martínez",
    "jose maria sanchez martinez": "José María Sánchez Martínez",
    "sanchez martinez": "José María Sánchez Martínez",
    "a hernandez hernandez": "Alejandro Hernández Hernández",
    "alejandro hernandez hernandez": "Alejandro Hernández Hernández",
    "hernandez hernandez": "Alejandro Hernández Hernández",
    "c soto grado": "César Soto Grado",
    "cesar soto grado": "César Soto Grado",
    "soto grado": "César Soto Grado",
    "j martinez munuera": "Juan Martínez Munuera",
    "juan martinez munuera": "Juan Martínez Munuera",
    "martinez munuera": "Juan Martínez Munuera",
    "g cuadra fernandez": "Guillermo Cuadra Fernández",
    "guillermo cuadra fernandez": "Guillermo Cuadra Fernández",
    "cuadra fernandez": "Guillermo Cuadra Fernández",
    "j alberola rojas": "Javier Alberola Rojas",
    "javier alberola rojas": "Javier Alberola Rojas",
    "alberola rojas": "Javier Alberola Rojas",
    "r de burgos bengoetxea": "Ricardo de Burgos Bengoetxea",
    "ricardo de burgos bengoetxea": "Ricardo de Burgos Bengoetxea",
    "de burgos bengoetxea": "Ricardo de Burgos Bengoetxea",
    "m melero lopez": "Mario Melero López",
    "mario melero lopez": "Mario Melero López",
    "melero lopez": "Mario Melero López",
    "i diaz de mera escuderos": "Isidro Díaz de Mera Escuderos",
    "isidro diaz de mera escuderos": "Isidro Díaz de Mera Escuderos",
    "diaz de mera escuderos": "Isidro Díaz de Mera Escuderos",
    "diaz de mera": "Isidro Díaz de Mera Escuderos",
    "j l munuera montero": "José Luis Munuera Montero",
    "jose luis munuera montero": "José Luis Munuera Montero",
    "munuera montero": "José Luis Munuera Montero",
    "v garcia verdura": "Víctor García Verdura",
    "victor garcia verdura": "Víctor García Verdura",
    "garcia verdura": "Víctor García Verdura",
    "j a gonzalez esteban": "Jon Ander González Esteban",
    "jon ander gonzalez esteban": "Jon Ander González Esteban",
    "gonzalez esteban": "Jon Ander González Esteban",
    "f j hernandez maeso": "Francisco José Hernández Maeso",
    "francisco jose hernandez maeso": "Francisco José Hernández Maeso",
    "hernandez maeso": "Francisco José Hernández Maeso",
    "m sesma espinosa": "Miguel Sesma Espinosa",
    "miguel sesma espinosa": "Miguel Sesma Espinosa",
    "sesma espinosa": "Miguel Sesma Espinosa",
    "m j orellana cid": "Manuel Jesús Orellana Cid",
    "manuel jesus orellana cid": "Manuel Jesús Orellana Cid",
    "orellana cid": "Manuel Jesús Orellana Cid",
    "m busquets ferrer": "Mateo Busquets Ferrer",
    "mateo busquets ferrer": "Mateo Busquets Ferrer",
    "busquets ferrer": "Mateo Busquets Ferrer",
    "c del cerro grande": "Carlos del Cerro Grande",
    "carlos del cerro grande": "Carlos del Cerro Grande",
    "del cerro grande": "Carlos del Cerro Grande",
    "a mateu lahoz": "Antonio Mateu Lahoz",
    "antonio mateu lahoz": "Antonio Mateu Lahoz",
    "mateu lahoz": "Antonio Mateu Lahoz",

    # Premier League (Inglaterra)
    "m oliver": "Michael Oliver",
    "michael oliver": "Michael Oliver",
    "oliver michael": "Michael Oliver",
    "a taylor": "Anthony Taylor",
    "anthony taylor": "Anthony Taylor",
    "taylor anthony": "Anthony Taylor",
    "p tierney": "Paul Tierney",
    "paul tierney": "Paul Tierney",
    "tierney paul": "Paul Tierney",
    "s hooper": "Simon Hooper",
    "simon hooper": "Simon Hooper",
    "j gillett": "Jarred Gillett",
    "jarred gillett": "Jarred Gillett",
    "c pawson": "Craig Pawson",
    "craig pawson": "Craig Pawson",
    "s attwell": "Stuart Attwell",
    "stuart attwell": "Stuart Attwell",
    "j brooks": "John Brooks",
    "john brooks": "John Brooks",
    "p bankes": "Peter Bankes",
    "peter bankes": "Peter Bankes",
    "r jones": "Robert Jones",
    "rob jones": "Robert Jones",
    "robert jones": "Robert Jones",
    "d england": "Darren England",
    "darren england": "Darren England",
    "t bramall": "Thomas Bramall",
    "thomas bramall": "Thomas Bramall",
    "c kavanagh": "Chris Kavanagh",
    "chris kavanagh": "Chris Kavanagh",
    "t harrington": "Tony Harrington",
    "tony harrington": "Tony Harrington",
    "m salisbury": "Michael Salisbury",
    "michael salisbury": "Michael Salisbury",
    "s barrott": "Samuel Barrott",
    "samuel barrott": "Samuel Barrott",
    "a madley": "Andrew Madley",
    "andy madley": "Andrew Madley",
    "andrew madley": "Andrew Madley",
    "a kitchen": "Akash Kitchen",
    "adam herczeg": "Adam Herczeg",
    "farai hallam": "Farai Hallam",
    "f hallam": "Farai Hallam",
    "josh smith": "Josh Smith",
    "j smith": "Josh Smith",
    "lewis smith": "Lewis Smith",
    "l smith": "Lewis Smith",
    "thomas kirk": "Thomas Kirk",
    "t kirk": "Thomas Kirk",

    # Serie A (Italia)
    "d orsato": "Daniele Orsato",
    "daniele orsato": "Daniele Orsato",
    "d massa": "Davide Massa",
    "davide massa": "Davide Massa",
    "m mariani": "Maurizio Mariani",
    "maurizio mariani": "Maurizio Mariani",
    "s sozza": "Simone Sozza",
    "simone sozza": "Simone Sozza",
    "m fabbri": "Michael Fabbri",
    "michael fabbri": "Michael Fabbri",
    "m marcenaro": "Matteo Marcenaro",
    "matteo marcenaro": "Matteo Marcenaro",
    "a colombo": "Andrea Colombo",
    "andrea colombo": "Andrea Colombo",
    "g pairetto": "Luca Pairetto",
    "l pairetto": "Luca Pairetto",
    "f la penna": "Federico La Penna",
    "m guida": "Marco Guida",
    "marco guida": "Marco Guida",
    "f maresca": "Fabio Maresca",
    "fabio maresca": "Fabio Maresca",
    "g ayroldi": "Giovanni Ayroldi",
    "d doveri": "Daniele Doveri",
    "daniele doveri": "Daniele Doveri",

    # Bundesliga (Alemania)
    "f zwayer": "Felix Zwayer",
    "felix zwayer": "Felix Zwayer",
    "d siebert": "Daniel Siebert",
    "daniel siebert": "Daniel Siebert",
    "d aytekin": "Deniz Aytekin",
    "deniz aytekin": "Deniz Aytekin",
    "s jablonski": "Sven Jablonski",
    "sven jablonski": "Sven Jablonski",
    "h osmers": "Harm Osmers",
    "harm osmers": "Harm Osmers",
    "f badstubner": "Florian Badstübner",
    "florian badstubner": "Florian Badstübner",
    "c dingert": "Christian Dingert",
    "t stieler": "Tobias Stieler",
    "tobias stieler": "Tobias Stieler",
    "b dankert": "Bastian Dankert",
    "f brych": "Felix Brych",
    "felix brych": "Felix Brych",
    "s stegemann": "Sascha Stegemann",
    "sascha stegemann": "Sascha Stegemann",
    "r schroder": "Robert Schröder",
    "robert schroder": "Robert Schröder",

    # Ligue 1 (Francia)
    "c turpin": "Clément Turpin",
    "clement turpin": "Clément Turpin",
    "f letexier": "François Letexier",
    "francois letexier": "François Letexier",
    "w delajod": "Willy Delajod",
    "willy delajod": "Willy Delajod",
    "j brisard": "Jérôme Brisard",
    "jerome brisard": "Jérôme Brisard",
    "b bastien": "Benoît Bastien",
    "benoit bastien": "Benoît Bastien",
    "s pignard": "Jérémie Pignard",
    "e wattellier": "Eric Wattellier",
    "m vernice": "Mathieu Vernice",
    "t leonard": "Thomas Léonard",
    "r buquet": "Ruddy Buquet",

    # UEFA Elite / Internacionales
    "s marciniak": "Szymon Marciniak",
    "szymon marciniak": "Szymon Marciniak",
    "s vincic": "Slavko Vinčić",
    "slavko vincic": "Slavko Vinčić",
    "slavko vinčić": "Slavko Vinčić",
    "i kovacs": "István Kovács",
    "istvan kovacs": "István Kovács",
    "istván kovács": "István Kovács",
    "g nyberg": "Glenn Nyberg",
    "glenn nyberg": "Glenn Nyberg",
    "s scharer": "Sandro Schärer",
    "sandro scharer": "Sandro Schärer",
    "sandro schärer": "Sandro Schärer",
    "n walsh": "Nick Walsh",
    "nick walsh": "Nick Walsh",
    "s gozubuyuk": "Serdar Gözübüyük",
    "serdar gozubuyuk": "Serdar Gözübüyük",
    "l godinho": "Luís Godinho",
    "luis godinho": "Luís Godinho",
    "luís godinho": "Luís Godinho",
    "g kabakov": "Georgi Kabakov",
    "georgi kabakov": "Georgi Kabakov",
    "r obrenovic": "Rade Obrenovic",
    "rade obrenovic": "Rade Obrenovic",
    "s gishamer": "Sebastian Gishamer",
    "sebastian gishamer": "Sebastian Gishamer",
    "a aghayev": "Aliyar Aghayev",
    "h peljto": "Irfan Peljto",
    "irfan peljto": "Irfan Peljto",
    "a kruzliak": "Ivan Kružliak",
    "ivan kruzliak": "Ivan Kružliak",
    "e eskas": "Espen Eskås",
    "espen eskas": "Espen Eskås",
    "d schnyder": "Urs Schnyder",
    "u schnyder": "Urs Schnyder",
}


def _normalizar_nombre_arbitro(nombre: str) -> str:
    """
    Normaliza el nombre de un colegiado unificando abreviaturas ('J. Gil Manzano' / 'M Oliver'),
    invirtiendo apellidos ('Gil Manzano, Jesús' -> 'Jesús Gil Manzano') y limpiando prefijos.
    """
    if not nombre or not isinstance(nombre, str):
        return ""
    raw = nombre.strip()
    if raw.lower() in ("nan", "none", "null", "sin designar", "por confirmar", "desconocido", "", "designación oficial"):
        return ""

    if "," in raw:
        partes = [p.strip() for p in raw.split(",", 1)]
        if len(partes) == 2 and partes[0] and partes[1]:
            raw = f"{partes[1]} {partes[0]}"

    raw = re.sub(r"^(dr\.|sr\.|mr\.|d\.)\s+", "", raw, flags=re.IGNORECASE)
    raw = re.sub(r"\s+", " ", raw).strip()

    k_clean = _despojar_diacriticos_arb(raw)
    k_clean = re.sub(r"\.", "", k_clean)
    k_clean = re.sub(r"\s+", " ", k_clean).strip()

    if k_clean in _MAPA_CANONICO_ARBITROS:
        return _MAPA_CANONICO_ARBITROS[k_clean]

    tokens = k_clean.split()
    if len(tokens) >= 2 and len(tokens[0]) <= 2:
        sufijo_apellido = " ".join(tokens[1:])
        for k_m, v_canon in _MAPA_CANONICO_ARBITROS.items():
            if k_m.endswith(sufijo_apellido) and k_m.startswith(tokens[0][0]):
                return v_canon

    if len(tokens) == 2 and len(tokens[0]) == 1:
        return f"{tokens[0].upper()}. {tokens[1].capitalize()}"

    return " ".join(w.capitalize() for w in raw.split())


@st.cache_data(ttl=1800, show_spinner=False)
def _obtener_censo_global_arbitros(df_partidos_global: pd.DataFrame = None) -> dict:
    """
    Genera el censo consolidado transversal de árbitros combinando 'arbitros_stats'
    y las actas de 'df_partidos_global'. Mapea cada variante a su entidad canónica,
    computa muestras en torneos domésticos vs UEFA y devuelve un diccionario ordenado.
    """
    censo = {}

    df_arb_db = pd.DataFrame()
    if RUTA_BD.exists():
        try:
            with sqlite3.connect(RUTA_BD) as conn:
                df_arb_db = pd.read_sql_query("SELECT * FROM arbitros_stats", conn)
        except Exception:
            df_arb_db = pd.DataFrame()

    MAPA_LIGAS_CORTAS = {
        "SP1": "LaLiga",
        "E0": "Premier League",
        "I1": "Serie A",
        "D1": "Bundesliga",
        "F1": "Ligue 1",
        "UCL": "UEFA Elite",
        "UEL": "UEFA Elite",
        "UECL": "UEFA Elite",
    }

    if not df_arb_db.empty:
        col_nom = df_arb_db.columns[0]
        col_div = df_arb_db.columns[1]
        col_liga = df_arb_db.columns[2]
        col_part = df_arb_db.columns[4]
        col_am = df_arb_db.columns[6]
        col_rj = df_arb_db.columns[7]
        col_fl = df_arb_db.columns[8]

        for _, r in df_arb_db.iterrows():
            nom_raw = str(r[col_nom]).strip()
            if not nom_raw or nom_raw.lower() in ("nan", "none", ""):
                continue
            nom_canon = _normalizar_nombre_arbitro(nom_raw)
            if not nom_canon:
                continue

            div_c = str(r[col_div]).strip()
            liga_nom = str(r[col_liga]).strip()
            part = int(pd.to_numeric(r[col_part], errors="coerce") or 1)
            am = float(pd.to_numeric(r[col_am], errors="coerce") or 0.0)
            rj = float(pd.to_numeric(r[col_rj], errors="coerce") or 0.0)
            fl = float(pd.to_numeric(r[col_fl], errors="coerce") or 0.0)

            es_uefa = div_c in ("UCL", "UEL", "UECL") or "UEFA" in liga_nom

            if nom_canon not in censo:
                censo[nom_canon] = {
                    "nombre": nom_canon,
                    "pais_liga_origen": liga_nom if not es_uefa else "Comité Arbitral UEFA",
                    "div_principal": div_c if not es_uefa else "UEFA",
                    "liga_corta": MAPA_LIGAS_CORTAS.get(div_c, "UEFA Elite" if es_uefa else div_c),
                    "total_partidos_global": 0,
                    "partidos_uefa": 0,
                    "partidos_domestica": 0,
                    "am_total_pond": 0.0,
                    "rj_total_pond": 0.0,
                    "fl_total_pond": 0.0,
                    "am_uefa_pond": 0.0,
                    "rj_uefa_pond": 0.0,
                    "fl_uefa_pond": 0.0,
                    "am_dom_pond": 0.0,
                    "rj_dom_pond": 0.0,
                    "fl_dom_pond": 0.0,
                    "variantes": set([nom_raw, nom_canon]),
                    "divs_registradas": set([div_c, liga_nom]),
                }

            ent = censo[nom_canon]
            ent["variantes"].add(nom_raw)
            ent["divs_registradas"].add(div_c)
            ent["divs_registradas"].add(liga_nom)
            ent["total_partidos_global"] += part
            ent["am_total_pond"] += (am * part)
            ent["rj_total_pond"] += (rj * part)
            ent["fl_total_pond"] += (fl * part)

            if es_uefa:
                ent["partidos_uefa"] += part
                ent["am_uefa_pond"] += (am * part)
                ent["rj_uefa_pond"] += (rj * part)
                ent["fl_uefa_pond"] += (fl * part)
            else:
                ent["partidos_domestica"] += part
                ent["am_dom_pond"] += (am * part)
                ent["rj_dom_pond"] += (rj * part)
                ent["fl_dom_pond"] += (fl * part)
                ent["pais_liga_origen"] = liga_nom
                ent["div_principal"] = div_c
                ent["liga_corta"] = MAPA_LIGAS_CORTAS.get(div_c, liga_nom)

    if df_partidos_global is not None and not df_partidos_global.empty and "Referee" in df_partidos_global.columns:
        sub_ref = df_partidos_global[df_partidos_global["Referee"].notna() & (df_partidos_global["Referee"].astype(str).str.strip() != "")]
        col_div_p = next((c for c in ("Div", "Liga_Origen") if c in sub_ref.columns), next((c for c in sub_ref.columns if "Div" in str(c)), None))
        for _, rp in sub_ref.iterrows():
            r_str = str(rp["Referee"]).strip()
            nom_c = _normalizar_nombre_arbitro(r_str)
            if not nom_c:
                continue
            div_val = str(rp[col_div_p]).strip() if col_div_p else "E0"
            es_u = div_val in ("UCL", "UEL", "UECL")
            if nom_c in censo:
                censo[nom_c]["variantes"].add(r_str)
                censo[nom_c]["divs_registradas"].add(div_val)
            else:
                liga_nom = MAPA_LIGAS.get(div_val, div_val)
                censo[nom_c] = {
                    "nombre": nom_c,
                    "pais_liga_origen": liga_nom if not es_u else "Comité Arbitral UEFA",
                    "div_principal": div_val if not es_u else "UEFA",
                    "liga_corta": MAPA_LIGAS_CORTAS.get(div_val, "UEFA Elite" if es_u else div_val),
                    "total_partidos_global": 1,
                    "partidos_uefa": 1 if es_u else 0,
                    "partidos_domestica": 0 if es_u else 1,
                    "am_total_pond": 4.10,
                    "rj_total_pond": 0.15,
                    "fl_total_pond": 22.0,
                    "am_uefa_pond": 4.10 if es_u else 0.0,
                    "rj_uefa_pond": 0.15 if es_u else 0.0,
                    "fl_uefa_pond": 22.0 if es_u else 0.0,
                    "am_dom_pond": 0.0 if es_u else 4.10,
                    "rj_dom_pond": 0.0 if es_u else 0.15,
                    "fl_dom_pond": 0.0 if es_u else 22.0,
                    "variantes": set([r_str, nom_c]),
                    "divs_registradas": set([div_val, liga_nom]),
                }

    resultado_censo = {}
    for nom_canon, data in censo.items():
        tot = max(1, data["total_partidos_global"])
        u_cnt = data["partidos_uefa"]
        d_cnt = data["partidos_domestica"]

        media_am_g = data["am_total_pond"] / tot
        media_rj_g = data["rj_total_pond"] / tot
        media_fl_g = data["fl_total_pond"] / tot if data["fl_total_pond"] > 0 else 22.0

        media_am_u = (data["am_uefa_pond"] / u_cnt) if u_cnt > 0 else media_am_g
        media_rj_u = (data["rj_uefa_pond"] / u_cnt) if u_cnt > 0 else media_rj_g
        media_fl_u = (data["fl_uefa_pond"] / u_cnt) if (u_cnt > 0 and data["fl_uefa_pond"] > 0) else media_fl_g

        media_am_d = (data["am_dom_pond"] / d_cnt) if d_cnt > 0 else media_am_g
        media_rj_d = (data["rj_dom_pond"] / d_cnt) if d_cnt > 0 else media_rj_g
        media_fl_d = (data["fl_dom_pond"] / d_cnt) if (d_cnt > 0 and data["fl_dom_pond"] > 0) else media_fl_g

        total_tarj = media_am_g + (2.0 * media_rj_g)
        if total_tarj < 3.8:
            sev_lbl = "🟢 Permisivo / Deja Jugar"
            sev_col = "#10B981"
        elif total_tarj <= 4.8:
            sev_lbl = "🟡 Estándar UEFA"
            sev_col = "#FBBF24"
        else:
            sev_lbl = "🔴 Riguroso / Alto Gatillo"
            sev_col = "#EF4444"

        etiqueta_sel = f"{nom_canon} [{data['liga_corta']}]"

        resultado_censo[nom_canon] = {
            "nombre": nom_canon,
            "pais_liga_origen": data["pais_liga_origen"],
            "div_principal": data["div_principal"],
            "liga_corta": data["liga_corta"],
            "etiqueta_selector": etiqueta_sel,
            "total_partidos_global": tot,
            "partidos_uefa": u_cnt,
            "partidos_domestica": d_cnt,
            "media_amarillas_global": round(media_am_g, 2),
            "media_rojas_global": round(media_rj_g, 2),
            "media_faltas_global": round(media_fl_g, 1),
            "total_tarjetas": round(total_tarj, 2),
            "media_amarillas_uefa": round(media_am_u, 2),
            "media_rojas_uefa": round(media_rj_u, 2),
            "media_faltas_uefa": round(media_fl_u, 1),
            "media_amarillas_dom": round(media_am_d, 2),
            "media_rojas_dom": round(media_rj_d, 2),
            "media_faltas_dom": round(media_fl_d, 1),
            "severidad_label": sev_lbl,
            "severidad_color": sev_col,
            "variantes": list(data["variantes"]),
            "divs_registradas": list(data.get("divs_registradas", [])),
        }

    return resultado_censo


def _calcular_perfil_disciplinario_arbitro(
    nombre_arbitro: str,
    censo_global: dict,
    df_partidos_global: pd.DataFrame = None,
    div_partido: str = "",
) -> dict:
    """
    Calcula el perfil disciplinario del colegiado con fallback jerárquico multicompetitorio:
    - CASO A: Partido UEFA con muestra UEFA >= 5 partidos -> métricas UEFA + comparativa doméstica.
    - CASO B: Partido UEFA con muestra corta (< 5) o debut -> hereda estadísticas de su liga doméstica.
    - CASO C: Partido de liga doméstica -> métricas de su liga de procedencia (+ UEFA si aplica).
    - Neutral: Árbitro por confirmar -> medias neutrales sin romper el flujo.
    """
    if not nombre_arbitro or not isinstance(nombre_arbitro, str) or nombre_arbitro.startswith("🔍") or nombre_arbitro.lower() in ("nan", "none", "", "sin designar", "por confirmar", "designación oficial", "(pendiente de designación)"):
        return {
            "confirmado": False,
            "nombre": "Árbitro por confirmar",
            "etiqueta": "🔍 Árbitro por confirmar",
            "liga_principal": "Competición Neutral",
            "media_amarillas": 4.10,
            "media_rojas": 0.15,
            "total_tarjetas": 4.40,
            "media_faltas": 22.5,
            "friccion_por_tarjeta": 5.1,
            "frec_under15_1t": 65.0,
            "severidad_label": "🟡 Estándar UEFA",
            "severidad_color": "#FBBF24",
            "badge_info": "ℹ️ Árbitro por confirmar - Se aplican medias neutrales de la competición",
            "muestra_total": 0,
            "muestra_uefa": 0,
            "muestra_domestica": 0,
            "caso": "NEUTRAL",
        }

    nombre_limpio = re.sub(r"\s*\[.*?\]\s*$", "", nombre_arbitro).strip()
    nom_canon = _normalizar_nombre_arbitro(nombre_limpio)

    info_arb = None
    if nom_canon in censo_global:
        info_arb = censo_global[nom_canon]
    else:
        for c_k, c_v in censo_global.items():
            if c_k.lower() == nom_canon.lower() or nombre_limpio.lower() in [v.lower() for v in c_v.get("variantes", [])]:
                info_arb = c_v
                break
        if not info_arb:
            for c_k, c_v in censo_global.items():
                if nom_canon.lower() in c_k.lower() or c_k.lower() in nom_canon.lower():
                    info_arb = c_v
                    break

    if not info_arb:
        return {
            "confirmado": True,
            "nombre": nom_canon,
            "etiqueta": f"{nom_canon} [Oficial]",
            "liga_principal": "Comité Oficial",
            "media_amarillas": 4.10,
            "media_rojas": 0.15,
            "total_tarjetas": 4.40,
            "media_faltas": 22.5,
            "friccion_por_tarjeta": 5.1,
            "frec_under15_1t": 65.0,
            "severidad_label": "🟡 Estándar UEFA",
            "severidad_color": "#FBBF24",
            "badge_info": "📊 Datos basados en estimación estándar oficial",
            "muestra_total": 1,
            "muestra_uefa": 0,
            "muestra_domestica": 1,
            "caso": "OFICIAL_GENERICO",
        }

    m_uefa = info_arb["partidos_uefa"]
    n_dom = info_arb["partidos_domestica"]
    tot_p = info_arb["total_partidos_global"]
    liga_princ = info_arb["pais_liga_origen"]

    div_u = str(div_partido).upper()
    es_partido_uefa = any(u in div_u for u in ["UCL", "UEL", "UECL", "CHAMPIONS", "EUROPA", "CONFERENCE", "UEFA"])

    if es_partido_uefa:
        if m_uefa >= 5:
            m_am = info_arb["media_amarillas_uefa"]
            m_rj = info_arb["media_rojas_uefa"]
            m_fl = info_arb["media_faltas_uefa"]
            comp_txt = f" | Comparativa en {liga_princ}: {info_arb['media_amarillas_dom']:.2f} amarillas/partido ({n_dom}p)" if n_dom > 0 else ""
            badge_info = f"📊 Muestra UEFA ({m_uefa} partidos en torneos UEFA){comp_txt}"
            caso = "CASO_A"
        else:
            if n_dom > 0:
                m_am = info_arb["media_amarillas_dom"]
                m_rj = info_arb["media_rojas_dom"]
                m_fl = info_arb["media_faltas_dom"]
                badge_info = f"📊 Datos basados en muestra consolidada: {n_dom} partidos en {liga_princ} + {m_uefa} en UEFA"
                caso = "CASO_B"
            else:
                m_am = info_arb["media_amarillas_global"]
                m_rj = info_arb["media_rojas_global"]
                m_fl = info_arb["media_faltas_global"]
                badge_info = f"📊 Datos basados en muestra UEFA ({m_uefa} partidos en torneos UEFA)"
                caso = "CASO_B_SOLO_UEFA"
    else:
        if n_dom > 0:
            m_am = info_arb["media_amarillas_dom"]
            m_rj = info_arb["media_rojas_dom"]
            m_fl = info_arb["media_faltas_dom"]
            txt_u = f" + {m_uefa} en UEFA" if m_uefa > 0 else ""
            badge_info = f"📊 Datos basados en muestra consolidada: {n_dom} partidos en {liga_princ}{txt_u}"
            caso = "CASO_C"
        else:
            m_am = info_arb["media_amarillas_global"]
            m_rj = info_arb["media_rojas_global"]
            m_fl = info_arb["media_faltas_global"]
            badge_info = f"📊 Datos basados en muestra consolidada: {tot_p} partidos ({liga_princ})"
            caso = "CASO_C_GLOBAL"

    tot_tarj = round(m_am + (2.0 * m_rj), 2)
    fric_tarj = round(m_fl / tot_tarj, 1) if (tot_tarj > 0 and m_fl > 0) else 5.2
    lam_1t = tot_tarj * 0.38
    frec_1t_under15 = round(math.exp(-lam_1t) * (1.0 + lam_1t) * 100.0, 1)

    if tot_tarj < 3.8:
        sev_label = "🟢 Permisivo / Deja Jugar"
        sev_color = "#10B981"
    elif tot_tarj <= 4.8:
        sev_label = "🟡 Estándar UEFA"
        sev_color = "#FBBF24"
    else:
        sev_label = "🔴 Riguroso / Alto Gatillo"
        sev_color = "#EF4444"

    return {
        "confirmado": True,
        "nombre": info_arb["nombre"],
        "etiqueta": info_arb["etiqueta_selector"],
        "liga_principal": liga_princ,
        "media_amarillas": round(m_am, 2),
        "media_rojas": round(m_rj, 2),
        "total_tarjetas": tot_tarj,
        "media_faltas": round(m_fl, 1),
        "friccion_por_tarjeta": fric_tarj,
        "frec_under15_1t": frec_1t_under15,
        "severidad_label": sev_label,
        "severidad_color": sev_color,
        "badge_info": badge_info,
        "muestra_total": tot_p,
        "muestra_uefa": m_uefa,
        "muestra_domestica": n_dom,
        "caso": caso,
    }


def _renderizar_tarjeta_perfil_arbitro(perfil_arb: dict):
    """
    Renderiza la tarjeta visual Glassmorphism del colegiado con severidad disciplinaria,
    badge explicativo de muestra consolidada y 4 KPIs de tarjetas y fricción.
    """
    if not perfil_arb.get("confirmado"):
        st.info(perfil_arb.get("badge_info", "Árbitro por confirmar - Se aplican medias neutrales de la competición"))
        return

    sev_col = perfil_arb.get("severidad_color", "#FBBF24")
    sev_lbl = perfil_arb.get("severidad_label", "🟡 Estándar UEFA")
    nom_arb = perfil_arb.get("nombre", "Colegiado")
    liga_p = perfil_arb.get("liga_principal", "")
    badge_i = perfil_arb.get("badge_info", "")
    m_am = perfil_arb.get("media_amarillas", 4.10)
    m_rj = perfil_arb.get("media_rojas", 0.15)
    fric_t = perfil_arb.get("friccion_por_tarjeta", 5.2)
    fr_1t = perfil_arb.get("frec_under15_1t", 65.0)

    html_card = (
        f"<div style='background: rgba(30, 41, 59, 0.7); border: 1px solid rgba(255, 255, 255, 0.1); "
        f"border-radius: 10px; padding: 12px 16px; margin-bottom: 12px;'>"
        f"<div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;'>"
        f"<span style='font-weight: 800; font-size: 1.05rem; color: #f8fafc;'>"
        f"🧑‍⚖️ {html.escape(nom_arb)} "
        f"<span style='font-size: 0.8rem; color: #94a3b8; font-weight: 500;'>[{html.escape(liga_p)}]</span></span>"
        f"<span style='background: {sev_col}22; color: {sev_col}; border: 1px solid {sev_col}66; "
        f"padding: 2px 8px; border-radius: 12px; font-size: 0.76rem; font-weight: 700;'>{sev_lbl}</span>"
        f"</div>"
        f"<div style='font-size: 0.8rem; color: #38bdf8; margin-bottom: 10px;'>{html.escape(badge_i)}</div>"
        f"<div style='display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; text-align: center;'>"
        f"<div style='background: rgba(15, 23, 42, 0.6); padding: 6px 8px; border-radius: 6px;'>"
        f"<div style='font-size: 0.7rem; color: #94a3b8;'>🟨 Amarillas/90m</div>"
        f"<div style='font-size: 1.0rem; font-weight: 700; color: #facc15;'>{m_am:.2f}</div>"
        f"</div>"
        f"<div style='background: rgba(15, 23, 42, 0.6); padding: 6px 8px; border-radius: 6px;'>"
        f"<div style='font-size: 0.7rem; color: #94a3b8;'>🟥 Rojas/90m</div>"
        f"<div style='font-size: 1.0rem; font-weight: 700; color: #f87171;'>{m_rj:.2f}</div>"
        f"</div>"
        f"<div style='background: rgba(15, 23, 42, 0.6); padding: 6px 8px; border-radius: 6px;'>"
        f"<div style='font-size: 0.7rem; color: #94a3b8;'>⚡ Faltas/Tarjeta</div>"
        f"<div style='font-size: 1.0rem; font-weight: 700; color: #38bdf8;'>{fric_t:.1f}</div>"
        f"</div>"
        f"<div style='background: rgba(15, 23, 42, 0.6); padding: 6px 8px; border-radius: 6px;'>"
        f"<div style='font-size: 0.7rem; color: #94a3b8;'>⏱️ 1T &lt; 1.5 Tarj</div>"
        f"<div style='font-size: 1.0rem; font-weight: 700; color: #4ade80;'>{fr_1t:.0f}%</div>"
        f"</div>"
        f"</div>"
        f"</div>"
    )
    st.markdown(html_card, unsafe_allow_html=True)


def calcular_probabilidad_poisson(l, x):
    return (math.exp(-l) * (l**x)) / math.factorial(x)


def _auditar_friccion_arbitral_sqlite(
    aud_h: dict, aud_a: dict, liga_nombre: str, arbitro_override: str = "", df_partidos_global: pd.DataFrame = None
) -> dict:
    """
    Calcula la Tensión Disciplinaria y Fricción Táctica utilizando el censo transversal
    y el motor jerárquico de perfiles disciplinarios (_calcular_perfil_disciplinario_arbitro).
    """
    faltas_comb = float(aud_h.get("faltas_com_6", 11.8) + aud_a.get("faltas_com_6", 11.8))
    tarj_eq_comb = float(aud_h.get("tarjetas_6", 2.1) + aud_a.get("tarjetas_6", 2.1))
    div_code = MAPA_INV.get(liga_nombre, aud_h.get("div_eq", "E0"))

    arb_nombre = arbitro_override.strip() or aud_h.get("ultimo_arbitro", "") or aud_a.get("ultimo_arbitro", "")

    censo_g = _obtener_censo_global_arbitros(df_partidos_global)

    perfil_arb = _calcular_perfil_disciplinario_arbitro(
        nombre_arbitro=arb_nombre,
        censo_global=censo_g,
        df_partidos_global=df_partidos_global,
        div_partido=div_code,
    )

    nom_resuelto = perfil_arb.get("nombre", arb_nombre or "Colegiado Media Comité")
    arb_amarillas = float(perfil_arb.get("media_amarillas", 4.05))
    arb_rojas = float(perfil_arb.get("media_rojas", 0.14))
    arb_faltas = float(perfil_arb.get("media_faltas", 22.5))
    score_tarjetas_arb = float(perfil_arb.get("total_tarjetas", arb_amarillas + (2.0 * arb_rojas)))
    sev_label = perfil_arb.get("severidad_label", "🟡 Estándar UEFA")

    tension_disciplinaria = (score_tarjetas_arb * 5.5) + (faltas_comb * 1.15) + (tarj_eq_comb * 1.5)
    es_alta = (score_tarjetas_arb >= 4.15 and faltas_comb >= 22.5) or (tension_disciplinaria >= 56.0)

    if es_alta:
        badge_tactico = f"🔥 Fricción Táctica Alta ({sev_label}): Árbitro {nom_resuelto} + alta fricción de faltas ({faltas_comb:.1f})"
    else:
        badge_tactico = f"🍃 Partido Fluido ({sev_label}): Árbitro {nom_resuelto} + baja fricción de faltas ({faltas_comb:.1f})"

    return {
        "arbitro": nom_resuelto,
        "arb_amarillas": round(arb_amarillas, 2),
        "arb_rojas": round(arb_rojas, 2),
        "arb_faltas": round(arb_faltas, 1),
        "faltas_combinadas": round(faltas_comb, 1),
        "tension_score": round(tension_disciplinaria, 1),
        "es_alta": es_alta,
        "badge_tactico": badge_tactico,
        "perfil_arbitro": perfil_arb,
    }


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


COMPETENCIAS_ARBITROS = [
    "LaLiga (España)",
    "Premier League (Inglaterra)",
    "Serie A (Italia)",
    "Bundesliga (Alemania)",
    "Ligue 1 (Francia)",
    "Champions League (UEFA)",
    "Europa League (UEFA)",
    "Conference League (UEFA)",
]

MAPA_COMPETENCIAS_SQL = {
    "LaLiga (España)": ["LaLiga", "Primera Division", "SP1", "LaLiga (España)"],
    "Premier League (Inglaterra)": ["Premier League", "E0", "Premier League (Inglaterra)"],
    "Serie A (Italia)": ["Serie A", "I1", "Serie A (Italia)"],
    "Bundesliga (Alemania)": ["Bundesliga", "D1", "Bundesliga (Alemania)"],
    "Ligue 1 (Francia)": ["Ligue 1", "F1", "Ligue 1 (Francia)"],
    "Champions League (UEFA)": ["UCL", "Champions League", "UEFA Champions League", "CL", "uefa.champions", "Champions League (UEFA)"],
    "Europa League (UEFA)": ["UEL", "Europa League", "UEFA Europa League", "EL", "uefa.europa", "Europa League (UEFA)"],
    "Conference League (UEFA)": ["UECL", "Conference League", "UEFA Conference League", "ECL", "uefa.europa.conf", "Conference League (UEFA)"],
}


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

    competencias_arbitros = list(COMPETENCIAS_ARBITROS)
    divs_disponibles = [d for d in MAPA_LIGAS if d in set(df_partidos_base[col_liga].dropna().astype(str))]
    ligas_extra = [MAPA_LIGAS[d] for d in divs_disponibles if MAPA_LIGAS[d] not in competencias_arbitros]
    ligas_nombres = competencias_arbitros + ligas_extra

    # ── Selectores sincronizados: Liga -> Local -> Visitante -> Árbitro ──────
    s1, s2, s3, s4 = st.columns(4)
    with s1:
        liga_sel = st.selectbox("🏆 Liga:", ligas_nombres, key="disc_liga")
    div_sel = MAPA_INV.get(liga_sel, liga_sel)
    codigos_matchup = set(MAPA_COMPETENCIAS_SQL.get(liga_sel, [div_sel, liga_sel]))
    df_liga = df_partidos_base[df_partidos_base[col_liga].astype(str).isin(codigos_matchup)].copy()
    equipos_liga = sorted(
        pd.concat([df_liga["HomeTeam"], df_liga["AwayTeam"]]).dropna().unique().tolist()
    ) if not df_liga.empty else []

    censo_global = _obtener_censo_global_arbitros(df_partidos_base)
    censo_ordenado = sorted(censo_global.values(), key=lambda x: x["nombre"])
    nombres_arb = ["🔍 Seleccionar o buscar árbitro..."] + [item["etiqueta_selector"] for item in censo_ordenado]

    if len(equipos_liga) < 2:
        st.info(
            f"ℹ️ Aún no hay equipos suficientes registrados para **{liga_sel}**."
        )
        return

    with s2:
        eq_local = st.selectbox("🏠 Equipo Local:", equipos_liga, index=0, key=f"disc_local_{div_sel}")
    equipos_visitante = [e for e in equipos_liga if e != eq_local]
    with s3:
        eq_visit = st.selectbox("✈️ Equipo Visitante:", equipos_visitante, index=0, key=f"disc_visit_{div_sel}")

    with s4:
        arbitro_sel = st.selectbox("⚖️ Árbitro:", nombres_arb, index=0, key=f"disc_arb_{div_sel}")

    perfil_arb = _calcular_perfil_disciplinario_arbitro(
        arbitro_sel, censo_global, df_partidos_base, div_partido=div_sel
    )

    # ── Renderizar Perfil Disciplinario del Árbitro ─────────────────────────
    _renderizar_tarjeta_perfil_arbitro(perfil_arb)

    # ── Métricas base ────────────────────────────────────────────────────────
    media_liga = media_global_tarjetas_liga(None, df_liga)
    if perfil_arb["confirmado"]:
        media_arbitro = perfil_arb["total_tarjetas"]
        partidos_arb = perfil_arb["muestra_total"]
    else:
        media_arbitro = media_liga
        partidos_arb = 0
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
    if perfil_arb["confirmado"] and partidos_arb < MIN_PARTIDOS_ARBITRO:
        st.caption(f"⚠️ {perfil_arb['nombre']} solo tiene {partidos_arb} partido(s): su índice se ha suavizado hacia 1.00.")
    st.caption(
        f"Media global de la liga: **{media_liga:.2f}** tarjetas/partido · "
        f"Media árbitro: **{media_arbitro:.2f}**"
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
    st.dataframe(pd.DataFrame(tabla), width="stretch", hide_index=True)

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


def _sanitizar_texto_pdf(val) -> str:
    """
    Normaliza cadenas destinadas a flowables Table/Paragraph en ReportLab:
    - Elimina escapes tipográficos corruptos (\theta, \0, etc.) reemplazándolos por ASCII '0'.
    - Elimina delimitadores LaTeX residuales ($) y corrige colisiones de signos en récords (2W=3D-5L -> 2W - 3D - 5L).
    - Repara colisiones numéricas sin delimitador (ej. 1.241.29 -> 1.24 / 1.29).
    """
    if val is None:
        return ""
    s = str(val)
    for token_nulo in ("\\theta", "\\Theta", "\theta", "\x00", "\\0"):
        s = s.replace(token_nulo, "0")
    s = s.replace("$", "")
    s = re.sub(r"(\d+)\s*W\s*[=/-]\s*(\d+)\s*D\s*[=/-]\s*(\d+)\s*L", r"\1W - \2D - \3L", s)
    s = re.sub(r"(\d+\.\d{2})(?=\d+\.\d{2}\b)", r"\1 / ", s)
    return s


def _tabla_pdf(encabezados, filas, col_widths, col_alignments=None, mono_cols=None):
    estilos = _estilos_parrafo_pdf()
    mono_set = set(mono_cols or [])
    data = [[Paragraph(_sanitizar_texto_pdf(h), estilos["celda_h"]) for h in encabezados]]
    for fila in filas:
        row_cells = []
        for idx, v in enumerate(fila):
            if idx in mono_set:
                st_cell = estilos["celda_mono_b"] if idx == len(fila) - 1 else estilos["celda_mono"]
            elif col_alignments and idx < len(col_alignments) and col_alignments[idx] == "CENTER":
                st_cell = estilos["celda_c"]
            else:
                st_cell = estilos["celda"]
            row_cells.append(Paragraph(_sanitizar_texto_pdf(v), st_cell))
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
    canvas.drawString(1.4 * cm, 0.42 * cm, "HUGO'S HUB QUANT TERMINAL | Zero-Overround Institutional Intelligence | hugoshub.com")
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


class _MarkupPDF(str):
    """Cadena con marcado ReportLab ya construido y saneado (no se vuelve a escapar)."""


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
        if p <= 0.0:
            return "-"
        return f"{1.0 / max(p, 0.01):.2f}"

    def _val_m(dic, key, fallback=0.0):
        return float(dic.get(key, fallback))

    # ── PALETA INSTITUCIONAL ONE-PAGER ───────────────────────────────────────
    C_NAVY = colors.HexColor("#0F172A")
    C_SLATE = colors.HexColor("#1E293B")
    C_CYAN = colors.HexColor("#0284C7")
    C_EMER = colors.HexColor("#059669")
    C_DARK = colors.HexColor("#0F172A")
    C_MUTED = colors.HexColor("#64748B")
    C_ALT = colors.HexColor("#F8FAFC")
    C_SOFT = colors.HexColor("#E2E8F0")
    C_CARD = colors.HexColor("#CBD5E1")

    UW = 543.0  # ancho útil (A4 - márgenes 20 - padding interno del frame)
    import html as html_lib

    def _t(v):
        """Texto plano seguro para Paragraph: sanitiza (sin $, \\theta, '=' en récords) y escapa XML."""
        return html_lib.escape(_sanitizar_texto_pdf(v), quote=False)

    f_reg, f_bold, _, _ = _fuentes_pdf()

    st_hdr_t = ParagraphStyle("op_hdr_t", fontName="Helvetica-Bold", fontSize=12, leading=14, textColor=colors.white, alignment=TA_LEFT)
    st_hdr_badge = ParagraphStyle("op_hdr_b", fontName="Helvetica-Bold", fontSize=6.5, leading=8, textColor=colors.HexColor("#38BDF8"), alignment=TA_RIGHT)
    st_club = ParagraphStyle("op_club", fontName="Helvetica-Bold", fontSize=16, leading=19, textColor=colors.white, alignment=TA_LEFT)
    st_meta = ParagraphStyle("op_meta", fontName=f_reg, fontSize=7, leading=9, textColor=colors.HexColor("#CBD5E1"), alignment=TA_RIGHT)
    st_sec = ParagraphStyle("op_sec", fontName="Helvetica-Bold", fontSize=8.5, leading=10, textColor=C_NAVY, alignment=TA_LEFT)
    st_th = ParagraphStyle("op_th", fontName="Helvetica-Bold", fontSize=8, leading=9.5, textColor=colors.white, alignment=TA_CENTER)
    st_th_l = ParagraphStyle("op_th_l", parent=st_th, alignment=TA_LEFT)
    st_td_l = ParagraphStyle("op_td_l", fontName=f_reg, fontSize=7.5, leading=9, textColor=C_DARK, alignment=TA_LEFT)
    st_td_c = ParagraphStyle("op_td_c", fontName=f_reg, fontSize=7.5, leading=9, textColor=C_DARK, alignment=TA_CENTER)
    st_td_b = ParagraphStyle("op_td_b", fontName=f_bold, fontSize=7.5, leading=9, textColor=C_CYAN, alignment=TA_CENTER)
    st_kpi_l = ParagraphStyle("op_kpi_l", fontName="Helvetica-Bold", fontSize=6.5, leading=8, textColor=C_MUTED, alignment=TA_CENTER)
    st_kpi_v = ParagraphStyle("op_kpi_v", fontName="Helvetica-Bold", fontSize=13, leading=15, textColor=C_NAVY, alignment=TA_CENTER)
    st_kpi_s = ParagraphStyle("op_kpi_s", fontName=f_reg, fontSize=7, leading=8.5, textColor=C_MUTED, alignment=TA_CENTER)
    st_legal = ParagraphStyle("op_legal", fontName=f_reg, fontSize=6, leading=7.5, textColor=C_MUTED, alignment=TA_CENTER)

    # ── CABECERA CORPORATIVA ─────────────────────────────────────────────────
    fecha_corta = datetime.now().strftime("%d/%m/%Y")
    meta_txt = (
        f"{_t(liga)} | Horizonte: {_t(horizonte_tag)}<br/>"
        f"Muestra: {n_tot}p ({n_loc} Casa / {n_vis} Fuera) | Generado: {fecha_corta}"
    )
    header_table = Table(
        [
            [Paragraph("HUGO'S HUB | QUANT SPORTS INTELLIGENCE", st_hdr_t),
             Paragraph("PROPRIETARY ZERO-OVERROUND INTELLIGENCE", st_hdr_badge)],
            [Paragraph(_t(str(equipo).upper()), st_club), Paragraph(meta_txt, st_meta)],
        ],
        colWidths=[UW * 0.58, UW * 0.42],
    )
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), C_NAVY),
        ("BACKGROUND", (0, 1), (-1, 1), C_SLATE),
        ("LINEBELOW", (0, 1), (-1, 1), 1.6, C_CYAN),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))

    # ── SECCIÓN 1: 4 TARJETAS KPI ────────────────────────────────────────────
    win_pct, cs_pct, fts_pct = win_rate, cs_rate_g, fts_rate_g
    btts_pct, btts_no_pct = btts_si_g, btts_no_g
    btts_no_odds_txt = _cuota_str(btts_no_pct)

    def _card(label, valor, sub):
        return [Paragraph(_t(label), st_kpi_l), Paragraph(_t(valor), st_kpi_v), Paragraph(_t(sub), st_kpi_s)]

    tarjetas = [
        _card("WIN RATE & RECORD", f"{win_pct:.1f}% WIN", f"{v_tot}W - {e_tot}D - {d_tot}L"),
        _card("CLEAN SHEET & FTS", f"{cs_pct:.1f}% CS", f"Sin Marcar: {fts_pct:.1f}%"),
        _card("BTTS HIT RATE", f"{btts_pct:.1f}% SÍ", f"BTTS No: {btts_no_pct:.1f}% (Cuota {btts_no_odds_txt})"),
        _card("GOLES & xG NETO", f"{gf_med:.2f} / {gc_med:.2f}", f"xG: {xg_f_g:.2f} / {xg_c_g:.2f} (Neto {xg_neto_g:+.2f})"),
    ]
    gap = 5.0
    cw_card = (UW - 3 * gap) / 4.0
    kpi_table = Table(
        [[tarjetas[0], "", tarjetas[1], "", tarjetas[2], "", tarjetas[3]]],
        colWidths=[cw_card, gap, cw_card, gap, cw_card, gap, cw_card],
    )
    estilo_kpi = [
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]
    for ci, acc in zip((0, 2, 4, 6), (C_CYAN, C_EMER, C_CYAN, C_EMER)):
        estilo_kpi.append(("BOX", (ci, 0), (ci, 0), 0.6, C_CARD))
        estilo_kpi.append(("LINEABOVE", (ci, 0), (ci, 0), 2.0, acc))
    kpi_table.setStyle(TableStyle(estilo_kpi))

    # ── Helper de tablas one-pager ───────────────────────────────────────────
    def _tabla_op(encabezados, filas, anchos, header_bg, estilos_celda=None):
        data = [[Paragraph(_t(h), st_th_l if i == 0 else st_th) for i, h in enumerate(encabezados)]]
        for fila in filas:
            celdas = []
            for i, v in enumerate(fila):
                if i == 0:
                    celdas.append(Paragraph(_t(v), st_td_l))
                else:
                    est = (estilos_celda or {}).get(i, st_td_c)
                    celdas.append(Paragraph(v if isinstance(v, _MarkupPDF) else _t(v), est))
            data.append(celdas)
        tb = Table(data, colWidths=anchos, repeatRows=1)
        est_t = [
            ("BACKGROUND", (0, 0), (-1, 0), header_bg),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.4, C_SOFT),
            ("BOX", (0, 0), (-1, -1), 0.6, C_CARD),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]
        for i in range(1, len(data)):
            est_t.append(("BACKGROUND", (0, i), (-1, i), C_ALT if i % 2 == 0 else colors.white))
        tb.setStyle(TableStyle(est_t))
        return tb

    def _seccion(titulo, acento):
        return [
            Paragraph(_t(titulo), st_sec),
            HRFlowable(width="100%", thickness=0.9, color=acento, spaceBefore=1, spaceAfter=3),
        ]

    # ── SECCIÓN 2: ANÁLISIS ASIMÉTRICO ───────────────────────────────────────
    def _micro(texto, color_hex):
        return _MarkupPDF(f"<font size='6' color='{color_hex}'><b>{html_lib.escape(texto)}</b></font>")

    def _rating(valor, es_ataque):
        base = f"{valor:.2f}x"
        if es_ataque:
            etq, col = (("FUERTE", "#059669") if valor >= 1.10 else (("DÉBIL", "#DC2626") if valor <= 0.90 else ("NEUTRO", "#64748B")))
        else:
            etq, col = (("SÓLIDA", "#059669") if valor <= 0.90 else (("PERMEABLE", "#DC2626") if valor >= 1.10 else ("NEUTRO", "#64748B")))
        return _MarkupPDF(f"{html_lib.escape(base)} {_micro(etq, col)}")

    filas_split = [
        ["Partidos Analizados (Muestra)", f"{n_loc}", f"{n_vis}", f"{n_tot} tot."],
        ["Goles a Favor / en Contra (90m)", f"{h_gf:.2f} / {h_gc:.2f}", f"{a_gf:.2f} / {a_gc:.2f}", f"{media_liga_h:.2f} L / {media_liga_a:.2f} V"],
        ["Expected Goals (xG Favor / Contra)", f"{h_xgf:.2f} / {h_xgc:.2f}", f"{a_xgf:.2f} / {a_xgc:.2f}", f"{media_liga_h:.2f} / {media_liga_a:.2f}"],
        ["Tiros a Puerta (SOT Favor / Contra)", f"{h_sotf:.2f} / {h_sotc:.2f}", f"{a_sotf:.2f} / {a_sotc:.2f}", "4.10 / 4.10"],
        ["Tasa de Conversión SOT (% Gol/SOT)", f"{h_conv:.1f}%", f"{a_conv:.1f}%", "31.5%"],
        ["Attack Rating vs Media de Liga", _rating(h_atk, True), _rating(a_atk, True), "1.00x (Base)"],
        ["Defense Rating vs Media de Liga", _rating(h_def, False), _rating(a_def, False), "1.00x (Base)"],
        ["Distribución Goles Anotados (1T vs 2T)", f"{h_1t:.0f}% 1T / {h_2t:.0f}% 2T", f"{a_1t:.0f}% 1T / {a_2t:.0f}% 2T", "44% 1T / 56% 2T"],
    ]
    tabla_split = _tabla_op(
        ["Métrica Cuantitativa", "Rendimiento Casa", "Rendimiento Fuera", "Media Liga"],
        filas_split,
        [UW * 0.37, UW * 0.22, UW * 0.22, UW * 0.19],
        C_SLATE,
    )

    # ── SECCIÓN 3: MATRIZ DE CUOTAS JUSTAS DESACOPLADAS ──────────────────────
    over15_g = _val_m(sg, "over15", min(98.0, float(datos.get("p_over25", 50.0)) + 22.0))
    over25_g = _val_m(sg, "over25", float(datos.get("p_over25", 50.0)))
    over35_g = _val_m(sg, "over35", max(5.0, over25_g - 22.0))
    under25_g = _val_m(sg, "under25", max(0.0, 100.0 - over25_g))

    mercados_spec = [
        ("Over 1.5 Goles (+1.5)", over15_g, _val_m(sh, "over15", over15_g), _val_m(sa, "over15", over15_g)),
        ("Over 2.5 Goles (+2.5)", over25_g, _val_m(sh, "over25", over25_g), _val_m(sa, "over25", over25_g)),
        ("Over 3.5 Goles (+3.5)", over35_g, _val_m(sh, "over35", over35_g), _val_m(sa, "over35", over35_g)),
        ("Under 2.5 Goles (-2.5)", under25_g, _val_m(sh, "under25", under25_g), _val_m(sa, "under25", under25_g)),
        ("Ambos Marcan: BTTS (Sí)", btts_si_g, _val_m(sh, "btts_si", btts_si_g), _val_m(sa, "btts_si", btts_si_g)),
        ("Ambos Marcan: BTTS (No)", btts_no_g, _val_m(sh, "btts_no", btts_no_g), _val_m(sa, "btts_no", btts_no_g)),
        ("Portería a Cero (Clean Sheet)", cs_rate_g, _val_m(sh, "clean_sheet", cs_rate_g), _val_m(sa, "clean_sheet", cs_rate_g)),
        ("Sin Anotar (Failed to Score)", fts_rate_g, _val_m(sh, "failed_to_score", fts_rate_g), _val_m(sa, "failed_to_score", fts_rate_g)),
    ]
    filas_mercados_pdf = [
        [nom_m, f"{p_g:.1f}%", f"{p_h:.1f}%", f"{p_a:.1f}%", _cuota_str(p_h), _cuota_str(p_a)]
        for nom_m, p_g, p_h, p_a in mercados_spec
    ]
    tabla_mercados = _tabla_op(
        ["Mercado", "Hit Rate Gen %", "Hit Rate Casa %", "Hit Rate Fuera %", "Cuota Justa Casa (1/p)", "Cuota Justa Fuera (1/p)"],
        filas_mercados_pdf,
        [UW * 0.29, UW * 0.12, UW * 0.13, UW * 0.14, UW * 0.16, UW * 0.16],
        C_NAVY,
        estilos_celda={4: st_td_b, 5: st_td_b},
    )

    # ── SECCIÓN 4: DISCIPLINARY & FRICTION PROFILE ───────────────────────────
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
        ["Tarjetas Propias Recibidas (AM / RJ)", f"{g_am_p:.2f} AM / {g_rj_p:.2f} RJ", f"{h_am_p:.2f} AM / {h_rj_p:.2f} RJ", f"{a_am_p:.2f} AM / {a_rj_p:.2f} RJ"],
        ["Tarjetas Provocadas al Rival (AM / RJ)", f"{g_am_r:.2f} AM / {g_rj_r:.2f} RJ", f"{h_am_r:.2f} AM", f"{a_am_r:.2f} AM"],
        ["Faltas por Partido (Cometidas / Recibidas)", f"{g_f_com:.1f} / {g_f_rec:.1f}", f"{h_f_com:.1f} / {h_f_rec:.1f}", f"{a_f_com:.1f} / {a_f_rec:.1f}"],
        ["Disciplinary Friction Index", f"{g_fric:.1f} pts", f"{h_fric:.1f} pts (Casa)", f"{a_fric:.1f} pts (Fuera)"],
    ]
    tabla_disc = _tabla_op(
        ["Métrica Disciplinaria", "Promedio General", "Condición Casa", "Condición Fuera"],
        filas_disc,
        [UW * 0.37, UW * 0.21, UW * 0.21, UW * 0.21],
        C_SLATE,
    )

    # ── ENSAMBLADO (1 SOLA PÁGINA A4, SIN PageBreak) ─────────────────────────
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=20,
        rightMargin=20,
        topMargin=18,
        bottomMargin=18,
        title=f"Quant Intelligence One-Pager - {equipo}",
        author="Hugo's Hub",
    )

    story = [
        header_table,
        Spacer(1, 6),
        *_seccion("1. EXECUTIVE KPIs", C_CYAN),
        kpi_table,
        Spacer(1, 6),
        *_seccion("2. ANÁLISIS ASIMÉTRICO: CASA vs FUERA vs MEDIA DE LIGA", C_EMER),
        tabla_split,
        Spacer(1, 6),
        *_seccion("3. MATRIZ DE CUOTAS JUSTAS DESACOPLADAS (ZERO-OVERROUND)", C_CYAN),
        tabla_mercados,
        Spacer(1, 6),
        *_seccion("4. DISCIPLINARY & FRICTION PROFILE", C_EMER),
        tabla_disc,
        Spacer(1, 6),
        HRFlowable(width="100%", thickness=0.5, color=C_SOFT, spaceAfter=2),
        Paragraph(
            "Probabilidades empíricas y cuotas justas sin margen comercial (0% overround) calculadas sobre "
            "macro_base_futbol.db. Uso analítico exclusivo.",
            st_legal,
        ),
    ]
    # KeepInFrame en modo 'shrink' garantiza física y estrictamente 1 sola página
    story_one_page = [KeepInFrame(doc.width - 12, doc.height - 12, story, mode="shrink")]

    def _cierre_onepager(canvas, _doc):
        canvas.saveState()
        ancho, _alto = _doc.pagesize
        canvas.setFillColor(C_NAVY)
        canvas.rect(0, 0, ancho, 15, fill=1, stroke=0)
        canvas.setFillColor(C_CYAN)
        canvas.rect(0, 15, ancho, 1.2, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 7.5)
        canvas.drawCentredString(
            ancho / 2.0, 5,
            "HUGO'S HUB QUANT TERMINAL | Institutional Betting Intelligence | hugoshub.com",
        )
        canvas.restoreState()

    doc.build(story_one_page, onFirstPage=_cierre_onepager, onLaterPages=_cierre_onepager)
    buffer.seek(0)
    return buffer.getvalue()

df_partidos_global, err_partidos = cargar_datos_partidos()
df_jugadores, err_jugadores = cargar_datos_jugadores()
df_partidos = df_partidos_global.copy()

# ==============================================================================
# GESTIÓN GLOBAL DEL HISTORIAL DE NAVEGACIÓN (st.session_state)
# ==============================================================================
MODULO_INICIO = "🏠 Inicio / Portada"
OPCIONES_MENU_GLOBAL = [
    "🏠 Inicio / Portada",
    "📡 Escáner Vivo EV+ (API)",
    "📊 Dossier de Equipo",
    "⚔️ Comparador H2H (Cara a Cara)",
    "🎯 Radar de Tendencias",
    "📈 Tendencias Individuales",
    "📋 Scouting de Jugadores",
    "🤖 Motor de Apuestas (EV+)",
    "⚖️ Directorio y Matchup Árbitros",
    "🧮 Hándicap Asiático Avanzado",
    "🎯 Escáner Autónomo de Ineficiencias (+EV Scanner)",
    "🔗 Constructor Combinadas (EV+)",
    "💸 Calculadora & Kelly",
    "🤖 Asistente IA Local",
]

MAPA_BREADCRUMB_MODULOS = {
    "🏠 Inicio / Portada": "Inicio / Portada",
    "Inicio / Portada": "Inicio / Portada",
    "📡 Escáner Vivo EV+ (API)": "Escáner Multimercado EV+",
    "📊 Dossier de Equipo": "Dossier Cuantitativo de Equipo",
    "⚔️ Comparador H2H (Cara a Cara)": "Comparador H2H Pro",
    "🎯 Radar de Tendencias": "Radar de Tendencias",
    "📈 Tendencias Individuales": "Radar de Tendencias Individuales",
    "📋 Scouting de Jugadores": "Scouting de Jugadores Pro",
    "🤖 Motor de Apuestas (EV+)": "Motor de Apuestas (EV+)",
    "⚖️ Directorio y Matchup Árbitros": "Matchup de Árbitros & Tarjetas",
    "🧮 Hándicap Asiático Avanzado": "Hándicap Asiático Avanzado",
    "🎯 Escáner Autónomo de Ineficiencias (+EV Scanner)": "Escáner Autónomo (+EV Scanner)",
    "🔗 Constructor Combinadas (EV+)": "Constructor Combinadas (EV+)",
    "💸 Calculadora & Kelly": "Calculadora & Kelly",
    "🤖 Asistente IA Local": "Asistente IA Local",
}


def _normalizar_modulo_nav(destino: str) -> str:
    """Homogeneiza alias de navegación hacia las claves canónicas del menú lateral."""
    if not destino or destino == "Inicio / Portada":
        return MODULO_INICIO
    return destino if destino in OPCIONES_MENU_GLOBAL else MODULO_INICIO


if "menu_nav" not in st.session_state:
    st.session_state["menu_nav"] = MODULO_INICIO
else:
    st.session_state["menu_nav"] = _normalizar_modulo_nav(st.session_state["menu_nav"])

if (
    "nav_history" not in st.session_state
    or not isinstance(st.session_state["nav_history"], list)
    or len(st.session_state["nav_history"]) == 0
):
    st.session_state["nav_history"] = [st.session_state["menu_nav"]]
else:
    st.session_state["nav_history"] = [
        _normalizar_modulo_nav(m) for m in st.session_state["nav_history"]
    ]

if "nav_index" not in st.session_state:
    st.session_state["nav_index"] = len(st.session_state["nav_history"]) - 1
else:
    st.session_state["nav_index"] = max(
        0, min(int(st.session_state["nav_index"]), len(st.session_state["nav_history"]) - 1)
    )


def _registrar_navegacion(destino_menu: str) -> None:
    """Registra un movimiento de módulo en la pila de historial evitando duplicados consecutivos."""
    destino_norm = _normalizar_modulo_nav(destino_menu)
    hist = list(st.session_state.get("nav_history", [MODULO_INICIO]))
    idx = int(st.session_state.get("nav_index", max(0, len(hist) - 1)))
    idx = max(0, min(idx, len(hist) - 1)) if hist else 0

    # Si navegamos desde un punto intermedio del historial, truncamos la rama futura
    hist = hist[: idx + 1] if hist else []
    if not hist or hist[-1] != destino_norm:
        hist.append(destino_norm)

    st.session_state["nav_history"] = hist
    st.session_state["nav_index"] = len(hist) - 1
    st.session_state["menu_nav"] = destino_norm


def _on_sidebar_nav_change() -> None:
    """Callback disparado al cambiar manualmente el selector del menú lateral."""
    destino_actual = st.session_state.get("menu_nav", MODULO_INICIO)
    _registrar_navegacion(destino_actual)


def _nav_ir_atras() -> None:
    """Retrocede al módulo anterior en el historial de navegación."""
    hist = list(st.session_state.get("nav_history", [MODULO_INICIO]))
    idx = int(st.session_state.get("nav_index", 0))
    if idx > 0 and len(hist) > 0:
        nuevo_idx = idx - 1
        st.session_state["nav_index"] = nuevo_idx
        st.session_state["menu_nav"] = _normalizar_modulo_nav(hist[nuevo_idx])


def _nav_ir_adelante() -> None:
    """Avanza al módulo siguiente en el historial de navegación."""
    hist = list(st.session_state.get("nav_history", [MODULO_INICIO]))
    idx = int(st.session_state.get("nav_index", 0))
    if idx < len(hist) - 1:
        nuevo_idx = idx + 1
        st.session_state["nav_index"] = nuevo_idx
        st.session_state["menu_nav"] = _normalizar_modulo_nav(hist[nuevo_idx])


def _nav_ir_portada() -> None:
    """Regresa directamente a la pantalla principal registrando el paso en el historial."""
    _registrar_navegacion(MODULO_INICIO)


def render_top_navigation_bar() -> None:
    """Renderiza la barra de navegación superior global (Atrás / Adelante / Portada + Breadcrumb)."""
    hist = list(st.session_state.get("nav_history", [MODULO_INICIO]))
    idx = int(st.session_state.get("nav_index", 0))
    idx = max(0, min(idx, len(hist) - 1)) if hist else 0
    modulo_actual = _normalizar_modulo_nav(
        st.session_state.get("menu_nav", hist[idx] if hist else MODULO_INICIO)
    )

    puede_atras = idx > 0
    puede_adelante = idx < (len(hist) - 1)
    en_portada = modulo_actual == MODULO_INICIO

    etiqueta_actual = MAPA_BREADCRUMB_MODULOS.get(
        modulo_actual, modulo_actual.lstrip("🏠📡📊⚔️🎯📋🤖⚖️🧮🔗💸 ")
    )
    prev_hint = (
        MAPA_BREADCRUMB_MODULOS.get(hist[idx - 1], hist[idx - 1])
        if puede_atras
        else "Inicio"
    )

    st.markdown(
        """
        <style>
        .quant-topnav-shell {
            background: linear-gradient(90deg, #161B26 0%, #1E232F 100%);
            border: 1px solid #2A3245;
            border-radius: 10px;
            padding: 8px 14px;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            min-height: 38px;
            box-shadow: 0 4px 14px rgba(0, 0, 0, 0.28);
        }
        .quant-breadcrumb-trail {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 0.84rem;
            color: #94A3B8;
            font-weight: 600;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .quant-breadcrumb-root {
            color: #00E5FF;
            font-weight: 800;
            letter-spacing: 0.02em;
        }
        .quant-breadcrumb-sep {
            color: #475569;
            font-weight: 800;
        }
        .quant-breadcrumb-active {
            color: #F8FAFC;
            font-weight: 800;
            background: rgba(0, 230, 118, 0.10);
            border: 1px solid rgba(0, 230, 118, 0.30);
            border-radius: 6px;
            padding: 2px 9px;
        }
        .quant-breadcrumb-meta {
            font-size: 0.72rem;
            color: #64748B;
            font-family: 'JetBrains Mono', 'Consolas', monospace;
            background: rgba(14, 17, 23, 0.65);
            border: 1px solid #2A3245;
            border-radius: 6px;
            padding: 2px 8px;
            white-space: nowrap;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    col_back, col_fwd, col_home, col_crumb = st.columns([1.15, 1.25, 1.25, 5.35])
    with col_back:
        st.button(
            "⬅ Atrás",
            key="global_topnav_btn_back",
            disabled=not puede_atras,
            width="stretch",
            help=f"Volver a: {prev_hint}" if puede_atras else "Sin historial previo",
            on_click=_nav_ir_atras,
        )
    with col_fwd:
        st.button(
            "Adelante ➡",
            key="global_topnav_btn_fwd",
            disabled=not puede_adelante,
            width="stretch",
            help="Avanzar al módulo siguiente en el historial" if puede_adelante else "Sin historial posterior",
            on_click=_nav_ir_adelante,
        )
    with col_home:
        st.button(
            "🏠 Portada",
            key="global_topnav_btn_home",
            disabled=en_portada,
            width="stretch",
            help="Regresar a la pantalla principal en 1 clic",
            on_click=_nav_ir_portada,
        )
    with col_crumb:
        st.markdown(
            f"""
            <div class="quant-topnav-shell">
                <div class="quant-breadcrumb-trail">
                    <span class="quant-breadcrumb-root">⚽ Hugo's Hub</span>
                    <span class="quant-breadcrumb-sep">›</span>
                    <span class="quant-breadcrumb-active">{etiqueta_actual}</span>
                </div>
                <span class="quant-breadcrumb-meta">Ruta {idx + 1}/{len(hist)}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ==============================================================================
# SIDEBAR
# ==============================================================================
with st.sidebar:
    st.header("🎛️ Centro de Control")
    menu_seleccionado = st.radio(
        "Navegación:",
        OPCIONES_MENU_GLOBAL,
        key="menu_nav",
        on_change=_on_sidebar_nav_change,
    )
    # Sincronización defensiva por si menu_seleccionado divergió del tope actual
    if st.session_state["nav_history"][st.session_state["nav_index"]] != menu_seleccionado:
        _registrar_navegacion(menu_seleccionado)
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

        if st.button("🔄 Sincronizar Última Jornada", width="stretch", type="primary"):
            with st.spinner("Actualizando métricas de la jornada..."):
                try:
                    resultado_sync = sincronizar_jornada()
                    if resultado_sync.get("exito"):
                        partidos_n = resultado_sync.get("partidos_nuevos", 0)
                        jugadores_a = resultado_sync.get("jugadores_act", 0)

                        # Ejecución automática del scraper arbitral (Ligas + Torneos UEFA)
                        arbitros_msg = ""
                        try:
                            import subprocess
                            import sys

                            res_arb = subprocess.run(
                                [sys.executable, "scraper_arbitros.py"],
                                capture_output=True,
                                text=True,
                                timeout=120,
                            )
                            if res_arb.returncode == 0:
                                arbitros_msg = "\n- ⚖️ **Árbitros UEFA/Ligas:** actualizados correctamente."
                            else:
                                arbitros_msg = "\n- ⚠️ **Árbitros:** sincronizados con advertencias menores."
                        except Exception as err_arb:
                            arbitros_msg = f"\n- ℹ️ **Árbitros:** omitidos temporalmente ({err_arb})."

                        msg_ok = (
                            f"✅ **Jornada sincronizada correctamente.**\n\n"
                            f"- 📊 **Partidos nuevos:** {partidos_n}\n"
                            f"- 👤 **Jugadores actualizados:** {jugadores_a}"
                            f"{arbitros_msg}"
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
                     width="stretch", hide_index=True)
    except Exception:
        st.dataframe(df_tabla, width="stretch", hide_index=True)

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
# HELPERS CUANTITATIVOS COMPARTIDOS (DOSSIER & COMPARADOR H2H PRO)
# ==============================================================================
def _col_disponible(df_in, candidatos):
    for c in candidatos:
        if c in df_in.columns:
            return c
    return None


def _normalizar_partidos_equipo(df_raw, equipo):
    """
    Transforma los registros del equipo en un DataFrame estandarizado desde la perspectiva
    del equipo analizado (propio vs rival), con soporte resiliente de columnas, córners,
    métricas de fricción y xG estimado.
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    c_hg = _col_disponible(df_raw, ["FTHG", "HG", "goles_local"])
    c_ag = _col_disponible(df_raw, ["FTAG", "AG", "goles_visitante"])
    if not c_hg or not c_ag:
        return pd.DataFrame()

    c_hthg = _col_disponible(df_raw, ["HTHG"])
    c_htag = _col_disponible(df_raw, ["HTAG"])
    c_hxg = _col_disponible(df_raw, ["HxG", "xg_local"])
    c_axg = _col_disponible(df_raw, ["AxG", "xg_visitante"])
    c_hs = _col_disponible(df_raw, ["HS", "tiros_local"])
    c_as = _col_disponible(df_raw, ["AS", "tiros_visitante"])
    c_hst = _col_disponible(df_raw, ["HST", "sot_local"])
    c_ast = _col_disponible(df_raw, ["AST", "sot_visitante"])
    c_hc = _col_disponible(df_raw, ["HC", "corners_local", "corners_favor"])
    c_ac = _col_disponible(df_raw, ["AC", "corners_visitante", "corners_contra"])
    c_hthc = _col_disponible(df_raw, ["HTHC", "HC_1T", "corners_1t_local", "corners_1t"])
    c_htac = _col_disponible(df_raw, ["HTAC", "AC_1T", "corners_1t_visitante", "corners_2t"])
    c_hf = _col_disponible(df_raw, ["HF", "faltas_local"])
    c_af = _col_disponible(df_raw, ["AF", "faltas_visitante"])
    c_hy = _col_disponible(df_raw, ["HY", "amarillas_local"])
    c_ay = _col_disponible(df_raw, ["AY", "amarillas_visitante"])
    c_hr = _col_disponible(df_raw, ["HR", "rojas_local"])
    c_ar = _col_disponible(df_raw, ["AR", "rojas_visitante"])

    df_ord = _ordenar_por_fecha(df_raw)
    registros = []
    for _, row in df_ord.iterrows():
        home_t = str(row.get("HomeTeam", row.get("equipo_local", "")))
        es_local = home_t == str(equipo)
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

        hc_v = pd.to_numeric(row.get(c_hc), errors="coerce") if c_hc else np.nan
        ac_v = pd.to_numeric(row.get(c_ac), errors="coerce") if c_ac else np.nan
        if c_hc == "corners_favor" and c_ac == "corners_contra":
            corners_f = float(hc_v) if pd.notna(hc_v) else np.nan
            corners_c = float(ac_v) if pd.notna(ac_v) else np.nan
        else:
            corners_f = float(hc_v if es_local else ac_v) if pd.notna(hc_v if es_local else ac_v) else np.nan
            corners_c = float(ac_v if es_local else hc_v) if pd.notna(ac_v if es_local else hc_v) else np.nan

        corners_tot = (corners_f + corners_c) if (pd.notna(corners_f) and pd.notna(corners_c)) else np.nan
        over85_corners = (1.0 if corners_tot > 8.5 else 0.0) if pd.notna(corners_tot) else np.nan

        hthc_v = pd.to_numeric(row.get(c_hthc), errors="coerce") if c_hthc else np.nan
        htac_v = pd.to_numeric(row.get(c_htac), errors="coerce") if c_htac else np.nan
        if pd.notna(hthc_v) and pd.notna(htac_v) and pd.notna(corners_f):
            if c_hthc == "corners_1t" and c_htac == "corners_2t":
                cor_f_1t = float(hthc_v)
                cor_f_2t = float(htac_v)
            else:
                cor_f_1t = float(hthc_v if es_local else htac_v)
                cor_f_2t = max(0.0, corners_f - cor_f_1t)
            tiene_corners_ht = 1
        elif pd.notna(corners_f):
            if tiene_ht == 1 and (gf + gc) > 0:
                peso_1t_partido = (gf_1t + gc_1t) / (gf + gc)
                prop_1t = float(np.clip(0.70 * 0.45 + 0.30 * peso_1t_partido, 0.36, 0.58))
            else:
                prop_1t = 0.45
            cor_f_1t = corners_f * prop_1t
            cor_f_2t = corners_f * (1.0 - prop_1t)
            tiene_corners_ht = 0
        else:
            cor_f_1t = np.nan
            cor_f_2t = np.nan
            tiene_corners_ht = 0

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

        amar_prop = float(hy_v if es_local else ay_v) if pd.notna(hy_v if es_local else ay_v) else np.nan
        amar_riv = float(ay_v if es_local else hy_v) if pd.notna(ay_v if es_local else hy_v) else np.nan
        rojas_prop = float(hr_v if es_local else ar_v) if pd.notna(hr_v if es_local else ar_v) else np.nan
        rojas_riv = float(ar_v if es_local else hr_v) if pd.notna(ar_v if es_local else hr_v) else np.nan

        registros.append({
            "Date": row.get("Date", row.get("fecha", "N/A")),
            "HomeTeam": home_t,
            "AwayTeam": row.get("AwayTeam", row.get("equipo_visitante", "")),
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
            "corners_f": corners_f,
            "corners_c": corners_c,
            "corners_tot": corners_tot,
            "over85_corners": over85_corners,
            "cor_f_1t": cor_f_1t,
            "cor_f_2t": cor_f_2t,
            "tiene_corners_ht": tiene_corners_ht,
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


def _calcular_resumen_bloque(df_norm, usar_decay=False, media_liga_gf=1.35, media_liga_gc=1.35, df_fallback=None):
    n = len(df_norm) if df_norm is not None else 0
    if n == 0:
        return {
            "n": 0, "v": 0, "e": 0, "d": 0,
            "win_rate": 0.0, "draw_rate": 0.0, "loss_rate": 0.0,
            "gf_tot": 0, "gc_tot": 0,
            "gf_med": 0.0, "gc_med": 0.0,
            "xg_f": 0.0, "xg_c": 0.0, "xg_neto": 0.0, "diff_g_xg": 0.0,
            "tiros_f": 0.0, "tiros_c": 0.0, "sot_f": 0.0, "sot_c": 0.0,
            "sot_ratio": 0.0, "sot_ratio_c": 0.0,
            "conv_sot": 0.0, "conv_tiros": 0.0,
            "corners_f": 0.0, "corners_c": 0.0, "corners_tot": 0.0,
            "over85_corners": 0.0, "cor_1t_pct": 45.0, "cor_2t_pct": 55.0,
            "tiene_corners_ht": 0, "n_corners": 0,
            "atk_rating": 1.0, "def_rating": 1.0,
            "over15": 0.0, "over25": 0.0, "over35": 0.0, "under25": 0.0,
            "btts_si": 0.0, "btts_no": 0.0, "clean_sheet": 0.0, "failed_to_score": 0.0,
            "over25_cnt": 0, "btts_cnt": 0, "cs_cnt": 0,
            "gf_1t_pct": 45.0, "gf_2t_pct": 55.0, "gc_1t_pct": 45.0, "gc_2t_pct": 55.0,
            "amar_prop": 0.0, "rojas_prop": 0.0, "amar_riv": 0.0, "rojas_riv": 0.0,
            "faltas_com": 0.0, "faltas_rec": 0.0, "faltas_por_amarilla": 0.0, "friccion_idx": 0.0,
        }

    def _media_con_fb(col_name, default_val):
        if col_name in df_norm.columns and pd.to_numeric(df_norm[col_name], errors="coerce").dropna().any():
            return _media_ponderada_segura(df_norm, col_name, usar_decay=usar_decay, default=default_val)
        if df_fallback is not None and not df_fallback.empty and col_name in df_fallback.columns:
            s_fb = pd.to_numeric(df_fallback[col_name], errors="coerce").dropna()
            if not s_fb.empty:
                return _media_ponderada_segura(df_fallback, col_name, usar_decay=usar_decay, default=default_val)
        return float(default_val)

    v_cnt = int(df_norm["win"].sum())
    e_cnt = int(df_norm["draw"].sum())
    d_cnt = int(df_norm["loss"].sum())
    gf_tot = int(df_norm["gf"].sum())
    gc_tot = int(df_norm["gc"].sum())

    gf_med = _media_ponderada_segura(df_norm, "gf", usar_decay=usar_decay)
    gc_med = _media_ponderada_segura(df_norm, "gc", usar_decay=usar_decay)
    xg_f = _media_ponderada_segura(df_norm, "xg_f", usar_decay=usar_decay, default=gf_med)
    xg_c = _media_ponderada_segura(df_norm, "xg_c", usar_decay=usar_decay, default=gc_med)
    tiros_f = _media_con_fb("tiros_f", 11.5)
    tiros_c = _media_con_fb("tiros_c", 11.5)
    sot_f = _media_con_fb("sot_f", 4.0)
    sot_c = _media_con_fb("sot_c", 4.0)

    sot_ratio = (sot_f / tiros_f * 100.0) if tiros_f > 0 else 0.0
    sot_ratio_c = (sot_c / tiros_c * 100.0) if tiros_c > 0 else 0.0
    conv_sot = (gf_med / sot_f * 100.0) if sot_f > 0 else 0.0
    conv_tiros = (gf_med / tiros_f * 100.0) if tiros_f > 0 else 0.0

    n_cor_direct = int(pd.to_numeric(df_norm["corners_f"], errors="coerce").notna().sum()) if "corners_f" in df_norm.columns else 0
    if n_cor_direct > 0:
        n_corners = n_cor_direct
    elif df_fallback is not None and not df_fallback.empty and "corners_f" in df_fallback.columns:
        n_corners = int(pd.to_numeric(df_fallback["corners_f"], errors="coerce").notna().sum())
    else:
        n_corners = 0

    corners_f = _media_con_fb("corners_f", 5.2)
    corners_c = _media_con_fb("corners_c", 4.4)
    corners_tot = _media_con_fb("corners_tot", corners_f + corners_c)
    over85_corners = _media_con_fb("over85_corners", 0.55) * 100.0

    cor_f_1t_m = _media_con_fb("cor_f_1t", corners_f * 0.45)
    cor_f_2t_m = _media_con_fb("cor_f_2t", corners_f * 0.55)
    sum_cor_ht = cor_f_1t_m + cor_f_2t_m
    cor_1t_pct = (cor_f_1t_m / sum_cor_ht * 100.0) if sum_cor_ht > 0 else 45.0
    cor_2t_pct = 100.0 - cor_1t_pct
    tiene_corners_ht = int(df_norm["tiene_corners_ht"].sum()) if "tiene_corners_ht" in df_norm.columns else 0

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

    amar_prop = _media_con_fb("amar_prop", 2.0)
    rojas_prop = _media_con_fb("rojas_prop", 0.0)
    amar_riv = _media_con_fb("amar_riv", 2.0)
    rojas_riv = _media_con_fb("rojas_riv", 0.0)
    faltas_com = _media_con_fb("faltas_com", 11.5)
    faltas_rec = _media_con_fb("faltas_rec", 11.5)
    faltas_por_amarilla = (faltas_com / amar_prop) if amar_prop > 0.05 else faltas_com
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
        "sot_ratio": sot_ratio,
        "sot_ratio_c": sot_ratio_c,
        "conv_sot": conv_sot,
        "conv_tiros": conv_tiros,
        "corners_f": corners_f,
        "corners_c": corners_c,
        "corners_tot": corners_tot,
        "over85_corners": over85_corners,
        "cor_1t_pct": cor_1t_pct,
        "cor_2t_pct": cor_2t_pct,
        "tiene_corners_ht": tiene_corners_ht,
        "n_corners": n_corners,
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
        "faltas_por_amarilla": faltas_por_amarilla,
        "friccion_idx": friccion_idx,
    }


def _construir_tabla_rendimiento_condicion(stats_cond: dict, col_metrica: str) -> pd.DataFrame:
    """
    Construye el DataFrame de métricas comparativas por condición (Casa / Fuera)
    incluyendo el desglose estadístico de Córners, Precisión SOT y Fricción Avanzada.
    """
    if stats_cond.get("tiene_corners_ht", 0) > 0:
        dist_corners_str = (
            f"{stats_cond['cor_1t_pct']:.0f}% (1T) / {stats_cond['cor_2t_pct']:.0f}% (2T) · "
            f"Media 90m: {stats_cond['corners_f']:.1f} a favor vs {stats_cond['corners_c']:.1f} concedidos"
        )
    else:
        dist_corners_str = (
            f"{stats_cond['cor_1t_pct']:.0f}% (1T) / {stats_cond['cor_2t_pct']:.0f}% (2T) · "
            f"Fallback Media 90m (partidos con datos): {stats_cond['corners_f']:.1f} a favor vs "
            f"{stats_cond['corners_c']:.1f} concedidos ({stats_cond['corners_tot']:.1f} tot./90m)"
        )

    return pd.DataFrame([
        {
            col_metrica: "Volumen de Tiro (Totales F / C)",
            "Valor": (
                f"{stats_cond['tiros_f']:.1f} realizados vs {stats_cond['tiros_c']:.1f} concedidos "
                f"({stats_cond['conv_tiros']:.1f}% conv. total)"
            ),
        },
        {
            col_metrica: "Precisión de Disparo (SOT Ratio)",
            "Valor": (
                f"{stats_cond['sot_ratio']:.1f}% ({stats_cond['sot_f']:.1f} Tiros a Puerta / "
                f"{stats_cond['tiros_f']:.1f} Tiros Totales) · Rival: {stats_cond['sot_ratio_c']:.1f}%"
            ),
        },
        {
            col_metrica: "Córners (Favor / Contra por 90m)",
            "Valor": (
                f"{stats_cond['corners_f']:.1f} a favor vs {stats_cond['corners_c']:.1f} concedidos "
                f"({stats_cond['over85_corners']:.1f}% partidos con +8.5 córners totales)"
            ),
        },
        {
            col_metrica: "Distribución Córners 1T / 2T",
            "Valor": dist_corners_str,
        },
        {
            col_metrica: "Distribución Goles 1T / 2T",
            "Valor": (
                f"Anotados: {stats_cond['gf_1t_pct']:.0f}% (1T) / {stats_cond['gf_2t_pct']:.0f}% (2T) · "
                f"Encajados: {stats_cond['gc_1t_pct']:.0f}% (1T) / {stats_cond['gc_2t_pct']:.0f}% (2T)"
            ),
        },
        {
            col_metrica: "Tasa de Fricción & Ratio Tarjeta/Falta",
            "Valor": (
                f"Faltas: {stats_cond['faltas_com']:.1f} Com. / {stats_cond['faltas_rec']:.1f} Rec. · "
                f"Amarillas: {stats_cond['amar_prop']:.2f} Prop. / {stats_cond['amar_riv']:.2f} Riv. · "
                f"Faltas por Tarjeta recibida: {stats_cond['faltas_por_amarilla']:.1f} faltas/amarilla"
            ),
        },
    ])


import html as _html_ind

# ==============================================================================
# RADAR CUANTITATIVO DE TENDENCIAS INDIVIDUALES & PROPS DE EQUIPO (POISSON ZERO-VIG)
# ==============================================================================
_IND_UEFA_DIVS = ("UCL", "UEL", "UECL")
_IND_NOMBRE_DIV = {
    "SP1": "LaLiga", "E0": "Premier League", "I1": "Serie A", "D1": "Bundesliga",
    "F1": "Ligue 1", "UCL": "Champions", "UEL": "Europa League", "UECL": "Conference",
}
_IND_FILTRO_COMPETICION = {
    "Todas las Competiciones": None,
    "LaLiga (SP1)": ("SP1",),
    "Premier League (E0)": ("E0",),
    "Serie A (I1)": ("I1",),
    "Bundesliga (D1)": ("D1",),
    "Ligue 1 (F1)": ("F1",),
    "Competiciones UEFA": _IND_UEFA_DIVS,
}
_IND_COLS_VISTA = ["Date", "HomeTeam", "AwayTeam", "Div", "FTHG", "FTAG", "FTR",
                   "HS", "AS", "HST", "AST", "HC", "AC", "HY", "AY", "HR", "AR"]
_IND_SPORT_KEYS = {
    "SP1": "soccer_spain_la_liga", "E0": "soccer_epl", "D1": "soccer_germany_bundesliga",
    "I1": "soccer_italy_serie_a", "F1": "soccer_france_ligue_one",
}
_IND_HORIZONTE_DIAS = 14

_IND_CAT_TODAS = "⚡ Todas"
_IND_CAT_CORNERS = "🚩 Córners"
_IND_CAT_TIROS = "🎯 Tiros a Puerta"
_IND_CAT_TARJETAS = "🟨 Tarjetas"
_IND_CAT_GOLES = "⚽ Goles"

# Matriz de cuotas base Kambi/Paf: {metrica: (categoria, nombre_over, nombre_under, {linea: cuota}, {linea: cuota})}
# Las líneas marcadas (*) no figuraban en la matriz institucional: son valores de referencia estimados.
_IND_MATRIZ_CUOTAS = {
    "corn_for": (_IND_CAT_CORNERS, "Córners a favor Más de", None,
                 {2.5: 1.15, 3.5: 1.35, 4.5: 1.68, 5.5: 2.15, 6.5: 2.90, 7.5: 3.85}, {}),
    "corn_against": (_IND_CAT_CORNERS, None, "Córners concedidos Menos de",
                     {}, {2.5: 3.20, 3.5: 2.10, 4.5: 1.55, 5.5: 1.25}),  # (*) <2.5
    "sot": (_IND_CAT_TIROS, "Tiros a puerta Más de", "Tiros a puerta Menos de",
            {2.5: 1.18, 3.5: 1.42, 4.5: 1.85, 5.5: 2.45, 6.5: 3.25}, {2.5: 4.50, 3.5: 3.00, 4.5: 2.05}),  # (*) unders
    "cards": (_IND_CAT_TARJETAS, "Tarjetas recibidas Más de", "Tarjetas recibidas Menos de",
              {0.5: 1.10, 1.5: 1.50, 2.5: 2.25, 3.5: 3.40, 4.5: 5.20, 5.5: 8.50},  # (*) >0.5
              {1.5: 2.80, 2.5: 1.85, 3.5: 1.45, 4.5: 1.25, 5.5: 1.12}),
    "goals": (_IND_CAT_GOLES, "Goles anotados Más de", "Goles anotados Menos de",
              {0.5: 1.18, 1.5: 1.95, 2.5: 3.40, 3.5: 6.50}, {1.5: 1.85, 2.5: 1.30}),  # (*) >3.5, unders
}
# Catálogo plano: (metrica, categoria, etiqueta, tipo, linea, cuota_base)
_IND_CATALOGO_PROPS = [
    (met, cat, f"{nom_o} {lin}", "over", lin, cuota)
    for met, (cat, nom_o, nom_u, ov, un) in _IND_MATRIZ_CUOTAS.items()
    for lin, cuota in ov.items()
] + [
    (met, cat, f"{nom_u} {lin}", "under", lin, cuota)
    for met, (cat, nom_o, nom_u, ov, un) in _IND_MATRIZ_CUOTAS.items()
    for lin, cuota in un.items()
]

# Opciones de píldoras de línea por categoría
_IND_LINEAS_PILLS = {
    _IND_CAT_TARJETAS: ["Todas", "+0.5", "+1.5", "+2.5", "+3.5", "+4.5", "+5.5", "-1.5", "-2.5", "-3.5", "-4.5", "-5.5"],
    _IND_CAT_CORNERS: ["Todas", "+2.5", "+3.5", "+4.5", "+5.5", "+6.5", "+7.5", "-2.5", "-3.5", "-4.5", "-5.5"],
    _IND_CAT_TIROS: ["Todas", "+2.5", "+3.5", "+4.5", "+5.5", "+6.5", "-2.5", "-3.5", "-4.5"],
    _IND_CAT_GOLES: ["Todas", "+0.5", "+1.5", "+2.5", "+3.5", "-1.5", "-2.5"],
}


@st.cache_data(ttl=600, show_spinner=False)
def _cargar_dataset_props_vectorial(db_path="macro_base_futbol.db"):
    """Lee `partidos_historico` (vista) con las métricas de volumen propio.

    Devuelve un DataFrame con la columna `Jugado` (resultado final disponible).
    Los campos numéricos ausentes quedan como NaN (no se fuerzan a 0) para poder
    distinguir 'sin dato' (p. ej. filas ESPN/UEFA) de un 0 real.
    """
    try:
        with sqlite3.connect(db_path) as conn:
            cols_vista = [c[1] for c in conn.execute("PRAGMA table_info(partidos_historico)").fetchall()]
            if not all(c in cols_vista for c in _IND_COLS_VISTA):
                cols_base = [c[1] for c in conn.execute("PRAGMA table_info(historico_reciente)").fetchall()]
                div_col = next((c for c in cols_base if c.endswith("Div")), "Div")
                sel = ", ".join(
                    f"[{div_col}] AS Div" if c == "Div" else f"[{c}]" for c in _IND_COLS_VISTA
                )
                conn.execute("DROP VIEW IF EXISTS partidos_historico")
                conn.execute(f"CREATE VIEW partidos_historico AS SELECT {sel} FROM historico_reciente")
                conn.commit()
            query = f"""
                SELECT {", ".join(f"[{c}]" for c in _IND_COLS_VISTA)}
                FROM partidos_historico
                WHERE Date IS NOT NULL AND HomeTeam IS NOT NULL AND AwayTeam IS NOT NULL
            """
            df = pd.read_sql_query(query, conn)
    except Exception as exc:
        st.error(f"Error consultando partidos_historico: {exc}")
        return pd.DataFrame()

    df["Date"] = pd.to_datetime(df["Date"], format="%d/%m/%Y", errors="coerce")
    df = df.dropna(subset=["Date"]).copy()
    for c in _IND_COLS_VISTA[4:]:
        if c != "FTR":
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Jugado"] = df["FTHG"].notna() & df["FTAG"].notna()
    return df.sort_values("Date", ascending=False, kind="stable").reset_index(drop=True)


@st.cache_data(ttl=1800, show_spinner=False)
def _ind_fixtures_vivo(sport_key):
    """Próximos partidos en vivo (misma fuente que el Radar Colectivo: conector_odds)."""
    if not sport_key:
        return []
    try:
        import conector_odds as _co_ind
        if hasattr(_co_ind, "_pedir_cuotas"):
            res = _co_ind._pedir_cuotas(sport_key, "h2h")
            if getattr(res, "status_code", 500) == 200:
                datos = res.json()
                if isinstance(datos, list):
                    return [
                        {"home_team": e.get("home_team"), "away_team": e.get("away_team"),
                         "commence_time": e.get("commence_time")}
                        for e in datos if isinstance(e, dict)
                    ]
    except Exception:
        return []
    return []


def _ind_aplanar_partidos(df):
    """Formato largo (1 fila por equipo y partido jugado) 100% vectorizado."""
    jug = df[df["Jugado"]]
    if jug.empty:
        return pd.DataFrame()

    def _lado(local):
        p, r = ("Home", "Away") if local else ("Away", "Home")
        sp, sr = ("H", "A") if local else ("A", "H")
        cp = "FTHG" if local else "FTAG"
        out = pd.DataFrame({
            "Fecha": jug["Date"].values,
            "Equipo": jug[f"{p}Team"].values,
            "Rival": jug[f"{r}Team"].values,
            "Cond": "L" if local else "V",
            "Div": jug["Div"].values,
            "goals": jug[cp].values,
        })
        corn_f = jug[f"{sp}C"].values.astype(float)
        corn_a = jug[f"{sr}C"].values.astype(float)
        sot = jug[f"{sp}ST"].values.astype(float)
        sot_r = jug[f"{sr}ST"].values.astype(float)
        # Sin dato (ESPN/UEFA): NaN o ambos lados a 0 -> NaN para excluir del muestreo
        sin_corn = np.isnan(corn_f) | np.isnan(corn_a) | ((corn_f + corn_a) == 0)
        sin_sot = np.isnan(sot) | np.isnan(sot_r) | ((sot + sot_r) == 0)
        out["corn_for"] = np.where(sin_corn, np.nan, corn_f)
        out["corn_against"] = np.where(sin_corn, np.nan, corn_a)
        out["sot"] = np.where(sin_sot, np.nan, sot)
        yel = jug[f"{sp}Y"].values.astype(float)
        red = jug[f"{sp}R"].values.astype(float)
        yel_r = jug[f"{sr}Y"].values.astype(float)
        sin_card = np.isnan(yel) | np.isnan(yel_r)
        out["cards"] = np.where(sin_card, np.nan, yel + np.nan_to_num(red))
        return out

    largo = pd.concat([_lado(True), _lado(False)], ignore_index=True)
    largo["rival3"] = largo["Rival"].astype(str).str.upper().str[:3].str.strip()
    return largo.sort_values("Fecha", ascending=False, kind="stable").reset_index(drop=True)


def _ind_proximos_cruces(df):
    """Próximo partido por equipo en los próximos 14 días.

    Replica las fuentes del Radar Colectivo: (1) fixtures en vivo vía conector_odds
    (con hora) y (2) partidos programados sin disputar en la BD (fecha >= hoy).
    """
    cols = ["Equipo", "px_rival", "px_cond", "px_dt"]
    hoy = pd.Timestamp.now().normalize()
    limite = hoy + pd.Timedelta(days=_IND_HORIZONTE_DIAS + 1)
    partes = []

    fut = df[(~df["Jugado"]) & (df["Date"] >= hoy) & (df["Date"] < limite)]
    if not fut.empty:
        partes.append(pd.DataFrame({"Equipo": fut["HomeTeam"].values, "px_rival": fut["AwayTeam"].values,
                                    "px_cond": "Local", "px_dt": fut["Date"].values, "src": 1}))
        partes.append(pd.DataFrame({"Equipo": fut["AwayTeam"].values, "px_rival": fut["HomeTeam"].values,
                                    "px_cond": "Visitante", "px_dt": fut["Date"].values, "src": 1}))

    equipos_bd = sorted(pd.concat([df["HomeTeam"], df["AwayTeam"]]).dropna().astype(str).unique().tolist())
    eventos = []
    for sk in _IND_SPORT_KEYS.values():
        eventos.extend(_ind_fixtures_vivo(sk) or [])
    if eventos:
        ev = pd.DataFrame(eventos).dropna(subset=["home_team", "away_team", "commence_time"])
        if not ev.empty:
            nombres = pd.unique(pd.concat([ev["home_team"], ev["away_team"]]))
            mapa = {n: ev_resolver_equipo(n, equipos_bd) for n in nombres}
            ev["h"] = ev["home_team"].map(mapa)
            ev["a"] = ev["away_team"].map(mapa)
            ev = ev.dropna(subset=["h", "a"])
            ts = pd.to_datetime(ev["commence_time"], errors="coerce", utc=True)
            try:
                ts = ts.dt.tz_convert("Europe/Madrid")
            except Exception:
                pass
            ev["dt"] = ts.dt.tz_localize(None)
            ev = ev[(ev["dt"] >= hoy) & (ev["dt"] < limite)]
            if not ev.empty:
                partes.append(pd.DataFrame({"Equipo": ev["h"].values, "px_rival": ev["a"].values,
                                            "px_cond": "Local", "px_dt": ev["dt"].values, "src": 0}))
                partes.append(pd.DataFrame({"Equipo": ev["a"].values, "px_rival": ev["h"].values,
                                            "px_cond": "Visitante", "px_dt": ev["dt"].values, "src": 0}))
    if not partes:
        return pd.DataFrame(columns=cols)
    todo = pd.concat(partes, ignore_index=True)
    todo["px_dt"] = pd.to_datetime(todo["px_dt"])
    todo["_dia"] = todo["px_dt"].dt.normalize()
    # Primero el día más cercano; a igualdad de día, la fuente en vivo (trae hora)
    todo = todo.sort_values(["_dia", "src", "px_dt"], kind="stable")
    return todo.drop_duplicates("Equipo", keep="first")[cols].reset_index(drop=True)


def _ind_poisson_p_over(lam, k):
    """P(X >= k) = 1 - sum_{i<k} e^-λ λ^i / i!  (vectorizado sobre λ), acotada [0.02, 0.98]."""
    lam = np.asarray(lam, dtype=float)
    acum = np.zeros_like(lam)
    for i in range(int(k)):
        acum += (np.exp(-lam) * (lam ** i)) / math.factorial(i)
    return np.clip(1.0 - acum, 0.02, 0.98)


def _escanear_props_vectorial(df, ventana=5, min_winrate=75.0, categoria=_IND_CAT_TODAS,
                              condicion="Global", divs=None, direccion="Todos", linea=None,
                              equipo=None):
    """Escáner de props individuales: Poisson real + EV% sobre ventana configurable.

    `direccion`: 'Todos' | 'over' | 'under'. `linea`: None o (tipo, valor) p. ej. ('over', 4.5).
    """
    columnas = ["Equipo", "Liga", "Categoria", "Mercado", "Linea", "Tipo", "Hits", "N", "Winrate",
                "Lambda", "Prob", "CuotaJusta", "CuotaBase", "EV", "Chips",
                "px_rival", "px_cond", "px_dt"]
    vacio = pd.DataFrame(columns=columnas)
    if df is None or df.empty:
        return vacio

    largo = _ind_aplanar_partidos(df)
    if largo.empty:
        return vacio
    if divs:
        largo = largo[largo["Div"].isin(divs)]
    if condicion == "Local":
        largo = largo[largo["Cond"] == "L"]
    elif condicion == "Visitante":
        largo = largo[largo["Cond"] == "V"]
    if equipo and equipo != "Todos los Equipos":
        largo = largo[largo["Equipo"] == equipo]
    if largo.empty:
        return vacio

    # Liga de referencia por equipo: preferencia doméstica sobre UEFA
    pri = largo.assign(_u=largo["Div"].isin(_IND_UEFA_DIVS).astype(int)).sort_values(
        ["_u", "Fecha"], ascending=[True, False], kind="stable")
    liga_eq = pri.drop_duplicates("Equipo").set_index("Equipo")["Div"].map(
        lambda d: _IND_NOMBRE_DIV.get(d, d))

    partes = []
    for met in sorted({m[0] for m in _IND_CATALOGO_PROPS}):
        sub = largo.dropna(subset=[met])
        if sub.empty:
            continue
        w = sub.groupby("Equipo", sort=False).head(int(ventana))
        w = w[w.groupby("Equipo")[met].transform("size") == int(ventana)]
        if w.empty:
            continue
        w = w.assign(chip=w[met].astype(int).astype(str) + " vs " + w["rival3"])
        w = w.iloc[::-1]  # cronológico: antiguo -> reciente
        agg = w.groupby("Equipo", sort=False).agg(
            Lambda=(met, "mean"), vals=(met, list), chips=("chip", list))
        for (m_met, cat, etiqueta, tipo, lin, cuota_base) in _IND_CATALOGO_PROPS:
            if m_met != met:
                continue
            if categoria != _IND_CAT_TODAS and cat != categoria:
                continue
            if direccion in ("over", "under") and tipo != direccion:
                continue
            if linea is not None and (tipo, lin) != tuple(linea):
                continue
            k = int(math.floor(lin)) + 1
            lam = agg["Lambda"].values
            p_over = _ind_poisson_p_over(lam, k)
            prob = p_over if tipo == "over" else np.clip(1.0 - p_over, 0.02, 0.98)
            hits = np.array([
                sum((v >= k) if tipo == "over" else (v < k) for v in vs) for vs in agg["vals"]
            ], dtype=int)
            wr = hits / float(ventana) * 100.0
            parte = pd.DataFrame({
                "Equipo": agg.index, "Categoria": cat, "Mercado": etiqueta, "Linea": lin,
                "Tipo": tipo, "Hits": hits, "N": int(ventana), "Winrate": wr,
                "Lambda": lam, "Prob": prob,
                "CuotaJusta": np.round(1.0 / prob, 2), "CuotaBase": cuota_base,
                "EV": (prob * cuota_base - 1.0) * 100.0,
                "Chips": [
                    [(c, (v >= k) if tipo == "over" else (v < k)) for c, v in zip(cs, vs)]
                    for cs, vs in zip(agg["chips"], agg["vals"])
                ],
            })
            partes.append(parte[parte["Winrate"] >= float(min_winrate)])

    if not partes:
        return vacio
    res = pd.concat(partes, ignore_index=True)
    if res.empty:
        return vacio
    res["Liga"] = res["Equipo"].map(liga_eq).fillna("N/A")
    res = res.merge(_ind_proximos_cruces(df), on="Equipo", how="left")
    return res.sort_values(["EV", "Winrate"], ascending=False, kind="stable").reset_index(drop=True)


def _ind_chips_html(chips):
    out = []
    for txt, ok in chips:
        if ok:
            st_ = "background:rgba(35,134,54,0.2);border:1px solid #238636;color:#7ee787;"
        else:
            st_ = "background:rgba(218,54,51,0.2);border:1px solid #da3633;color:#ff7b72;"
        out.append(
            f"<span style='{st_}display:inline-block;padding:1px 7px;margin:2px 3px 2px 0;"
            f"border-radius:6px;font-size:0.72rem;font-weight:600;white-space:nowrap;'>[{_html_ind.escape(txt)}]</span>"
        )
    return "".join(out)


def _ind_cruce_txt(rival, cond, fecha):
    if rival is None or (isinstance(rival, float) and np.isnan(rival)) or pd.isna(fecha):
        return "Sin cruce inmediato"
    f = pd.Timestamp(fecha)
    hora = f" {f.strftime('%H:%M')}" if (f.hour or f.minute) else ""
    return f"🗓️ vs {_html_ind.escape(str(rival))} [{cond}] · {f.strftime('%d/%m')}{hora}"


def render_radar_tendencias_individuales():
    st.markdown(
        "<style>"
        ".ind-hero{background:linear-gradient(135deg,#0e1117 0%,#161b22 60%,#21262d 100%);"
        "border:1px solid #30363d;border-left:4px solid #58a6ff;border-radius:12px;"
        "padding:20px 24px;margin-bottom:18px;}"
        ".ind-hero h2{margin:0;color:#58a6ff;font-weight:700;font-size:1.45rem;}"
        ".ind-hero p{margin:6px 0 10px 0;color:#8b949e;font-size:0.92rem;}"
        ".ind-badge{display:inline-block;margin-right:8px;padding:4px 12px;border-radius:20px;"
        "font-size:0.78rem;font-weight:600;background:rgba(56,139,253,0.15);border:1px solid #388bfd;color:#58a6ff;}"
        ".ind-badge.g{background:rgba(63,185,80,0.12);border-color:#238636;color:#3fb950;}"
        ".ind-badge.off{background:#161b22;border-color:#30363d;color:#8b949e;opacity:0.85;}"
        ".ind-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px;margin-bottom:22px;}"
        ".ind-card{background:#0d1117;border:1px solid #30363d;border-top:3px solid #3fb950;"
        "border-radius:12px;padding:14px 16px;box-shadow:0 8px 22px rgba(0,0,0,0.35);}"
        ".ind-card .top{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;}"
        ".ind-tag{font-size:0.72rem;font-weight:700;padding:3px 9px;border-radius:12px;"
        "background:rgba(88,166,255,0.14);border:1px solid #1f6feb;color:#58a6ff;white-space:nowrap;}"
        ".ind-ev{font-size:0.78rem;font-weight:800;padding:3px 10px;border-radius:12px;"
        "background:rgba(63,185,80,0.16);border:1px solid #3fb950;color:#3fb950;}"
        ".ind-card h3{margin:2px 0 2px 0;color:#f0f6fc;font-size:1.12rem;}"
        ".ind-liga{color:#58a6ff;font-size:0.78rem;font-weight:600;}"
        ".ind-next{margin:8px 0;padding:6px 10px;background:#161b22;border:1px solid #21262d;"
        "border-radius:8px;color:#c9d1d9;font-size:0.8rem;}"
        ".ind-mkt{color:#79c0ff;font-weight:700;font-size:0.9rem;margin:4px 0 8px 0;}"
        ".ind-met{display:flex;justify-content:space-between;gap:6px;font-size:0.78rem;color:#8b949e;"
        "border-top:1px solid #21262d;padding-top:8px;margin-top:6px;}"
        ".ind-met b{color:#f0f6fc;display:block;font-size:0.9rem;}"
        ".ind-wrap{background:#0d1117;border:1px solid #30363d;border-radius:12px;overflow-x:auto;padding:4px;}"
        ".ind-tbl{width:100%;border-collapse:collapse;font-size:0.8rem;color:#c9d1d9;min-width:1100px;}"
        ".ind-tbl th{background:#161b22;color:#8b949e;text-align:left;padding:10px 10px;font-size:0.7rem;"
        "letter-spacing:0.05em;border-bottom:1px solid #30363d;white-space:nowrap;}"
        ".ind-tbl td{padding:9px 10px;border-bottom:1px solid #21262d;vertical-align:middle;}"
        ".ind-tbl tr:hover td{background:#161b22;}"
        ".ind-evb{background:rgba(63,185,80,0.16);border:1px solid #3fb950;color:#3fb950;"
        "padding:2px 9px;border-radius:10px;font-weight:800;white-space:nowrap;}"
        ".ind-evb.neg{background:rgba(218,54,51,0.14);border-color:#da3633;color:#ff7b72;}"
        "</style>"
        "<div class='ind-hero'>"
        "<h2>📈 Radar Cuantitativo de Tendencias Individuales &amp; Props de Equipo</h2>"
        "<p>Escáner institucional de volumen propio aislado (Córners, Tiros, Goles y Tarjetas) "
        "calibrado mediante Poisson Zero-Vig y cruce de calendario próximo.</p>"
        "<span class='ind-badge'>Kambi / Paf Engine Ready</span>"
        "<span class='ind-badge g'>Distribución Poisson Activa</span>"
        "</div>",
        unsafe_allow_html=True,
    )

    df_raw = _cargar_dataset_props_vectorial()
    if df_raw.empty:
        st.warning("No se hallaron registros históricos para procesar tendencias individuales.")
        return

    # ── Fila 1: contexto ─────────────────────────────────────────────────────
    r1c1, r1c2, r1c3 = st.columns([1.2, 1.4, 1.2])
    with r1c1:
        comp_sel = st.selectbox("Competición", list(_IND_FILTRO_COMPETICION.keys()), key="ind_radar_comp")
    divs_sel = _IND_FILTRO_COMPETICION[comp_sel]
    with r1c2:
        cond_lbl = st.segmented_control(
            "Condición de Campo", ["🌐 Global", "🏠 Solo Local", "✈️ Solo Visitante"],
            default="🌐 Global", key="ind_radar_cond",
        ) or "🌐 Global"
    with r1c3:
        df_div = df_raw if not divs_sel else df_raw[df_raw["Div"].isin(divs_sel)]
        equipos_lista = sorted(pd.unique(pd.concat([df_div["HomeTeam"], df_div["AwayTeam"]]).dropna().astype(str)).tolist())
        equipo_sel = st.selectbox("Buscar Equipo", ["Todos los Equipos"] + equipos_lista, key="ind_radar_equipo")

    # ── Fila 2: categoría y líneas por píldoras ──────────────────────────────
    cat_sel = st.segmented_control(
        "Categoría de Prop",
        [_IND_CAT_TODAS, _IND_CAT_CORNERS, _IND_CAT_TIROS, _IND_CAT_TARJETAS, _IND_CAT_GOLES],
        default=_IND_CAT_TODAS, key="ind_radar_cat",
    ) or _IND_CAT_TODAS
    cat_slug = {_IND_CAT_CORNERS: "corn", _IND_CAT_TIROS: "sot", _IND_CAT_TARJETAS: "card", _IND_CAT_GOLES: "gol"}.get(cat_sel, "all")

    direccion = "Todos"
    linea_sel = None
    if cat_sel == _IND_CAT_TODAS:
        st.markdown("<span class='ind-badge off'>🔒 Auditoría Multimercado Activa</span>", unsafe_allow_html=True)
    else:
        if cat_sel == _IND_CAT_TARJETAS:
            dir_lbl = st.segmented_control("Dirección", ["Todos", "Overs (+)", "Unders (-)"],
                                           default="Todos", key=f"ind_radar_dir_{cat_slug}") or "Todos"
            direccion = {"Todos": "Todos", "Overs (+)": "over", "Unders (-)": "under"}[dir_lbl]
        elif cat_sel == _IND_CAT_CORNERS:
            dir_lbl = st.segmented_control("Dirección", ["Todos", "A favor (Overs)", "Concedidos (Unders)"],
                                           default="Todos", key=f"ind_radar_dir_{cat_slug}") or "Todos"
            direccion = {"Todos": "Todos", "A favor (Overs)": "over", "Concedidos (Unders)": "under"}[dir_lbl]
        lin_lbl = st.pills("Línea Específica", _IND_LINEAS_PILLS[cat_sel], default="Todas",
                           key=f"ind_radar_linea_{cat_slug}") or "Todas"
        if lin_lbl != "Todas":
            linea_sel = ("over" if lin_lbl.startswith("+") else "under", float(lin_lbl[1:]))

    # ── Fila 3: afinación cuantitativa ───────────────────────────────────────
    r3c1, r3c2 = st.columns([1.4, 1.6])
    with r3c1:
        umbral_sel = st.slider("Racha mínima (%)", 50, 100, 75, step=5, key="ind_radar_umbral")
    with r3c2:
        ventana_sel = st.pills("Muestra de Partidos", [3, 5, 8, 10, 15], default=5,
                               format_func=lambda x: f"{x} partidos", key="ind_radar_ventana") or 5

    condicion = {"🌐 Global": "Global", "🏠 Solo Local": "Local", "✈️ Solo Visitante": "Visitante"}[cond_lbl]

    res = _escanear_props_vectorial(
        df_raw, ventana=ventana_sel, min_winrate=float(umbral_sel), categoria=cat_sel,
        condicion=condicion, divs=divs_sel, direccion=direccion, linea=linea_sel, equipo=equipo_sel,
    )
    if res.empty:
        st.info(f"Ningún equipo cumple actualmente el {umbral_sel}% de efectividad con los parámetros fijados.")
        return

    reg = res.to_dict("records")

    # ── TOP 4 EV+ ────────────────────────────────────────────────────────────
    tarjetas = []
    for r in reg[:4]:
        cruce = _ind_cruce_txt(r["px_rival"], r["px_cond"], r["px_dt"])
        tarjetas.append(
            "<div class='ind-card'>"
            "<div class='top'><span class='ind-tag'>👤 Prop Individual</span>"
            f"<span class='ind-ev'>EV {r['EV']:+.1f}%</span></div>"
            f"<h3>{_html_ind.escape(str(r['Equipo']))}</h3>"
            f"<div class='ind-liga'>{_html_ind.escape(str(r['Liga']))}</div>"
            f"<div class='ind-next'>{cruce}</div>"
            f"<div class='ind-mkt'>{_html_ind.escape(r['Mercado'])}</div>"
            "<div class='ind-met'>"
            f"<span>Racha<b>{r['Hits']}/{r['N']} [{r['Winrate']:.0f}%]</b></span>"
            f"<span>Cuota Justa<b>{r['CuotaJusta']:.2f}</b></span>"
            f"<span>Cuota Base<b>{r['CuotaBase']:.2f}</b></span></div>"
            f"<div style='margin-top:8px;'>{_ind_chips_html(r['Chips'])}</div>"
            "</div>"
        )
    st.markdown("#### 💎 Top 4 Oportunidades Individuales de Mayor Valor Esperado (EV+)")
    st.markdown(f"<div class='ind-grid'>{''.join(tarjetas)}</div>", unsafe_allow_html=True)

    # ── TABLA MAESTRA ────────────────────────────────────────────────────────
    st.markdown(f"#### 📋 Top Oportunidades Individuales Filtradas (Winrate ≥ {umbral_sel}%)")
    filas = []
    for r in reg[:200]:
        evc = "ind-evb" if r["EV"] >= 0 else "ind-evb neg"
        filas.append(
            "<tr>"
            "<td><span class='ind-tag'>👤 Individual</span></td>"
            f"<td><b style='color:#f0f6fc;'>{_html_ind.escape(str(r['Equipo']))}</b>"
            f"<div class='ind-liga'>{_html_ind.escape(str(r['Liga']))}</div></td>"
            f"<td>{_ind_cruce_txt(r['px_rival'], r['px_cond'], r['px_dt'])}</td>"
            f"<td style='color:#79c0ff;font-weight:600;'>{_html_ind.escape(r['Mercado'])}</td>"
            f"<td><b>{r['Hits']}/{r['N']}</b> <span style='color:#8b949e;'>[{r['Winrate']:.0f}%]</span></td>"
            f"<td>{_ind_chips_html(r['Chips'])}</td>"
            f"<td><b>{r['CuotaJusta']:.2f}</b></td>"
            f"<td>{r['CuotaBase']:.2f}</td>"
            f"<td><span class='{evc}'>{r['EV']:+.1f}%</span></td>"
            "</tr>"
        )
    cabecera = "".join(
        f"<th>{h}</th>" for h in ["TIPO", "EQUIPO / LIGA", "PRÓXIMO CRUCE", "MERCADO / LÍNEA", "RACHA / %",
                                  "DESGLOSE (CHIPS CON VALOR)", "CUOTA JUSTA", "CUOTA BASE", "DIFERENCIAL EV%"]
    )
    st.markdown(
        f"<div class='ind-wrap'><table class='ind-tbl'><thead><tr>{cabecera}</tr></thead>"
        f"<tbody>{''.join(filas)}</tbody></table></div>",
        unsafe_allow_html=True,
    )
    st.caption(
        f"{len(reg)} oportunidades detectadas · mostrando las {min(len(reg), 200)} de mayor EV%. "
        "Cuota Justa = 1/P(Poisson); EV% = (P × Cuota Base − 1) × 100."
    )


# ==============================================================================
# RADAR COLECTIVO VIP (RFC-FIX-V3): CASCADE FILTERING, POISSON REAL, CHOQUE UNIFICADO & TIERS
# ==============================================================================
_COL_CAT_TODOS = "⚡ Todos los Mercados"
_COL_CAT_GOLES = "⚽ Goles Totales"
_COL_CAT_BTTS = "🤝 Ambos Marcan (BTTS)"
_COL_CAT_CORNERS = "🚩 Córners Totales"
_COL_CAT_TARJETAS = "🟨 Tarjetas Totales"

_COL_TIER_TODOS = "⚡ Todos los Rivales"
_COL_TIER_TOP = "👑 Top 6 / Líderes"
_COL_TIER_MEDIA = "⚖️ Zona Media"
_COL_TIER_BAJA = "⚠️ Zona Baja/Descenso"
_COL_TIER_UEFA = "🏆 Competición Europea"
_COL_TIER_GENERAL = "⚖️ Rango General"

_COL_LIGAS_DOMESTICAS = ("SP1", "E0", "I1", "D1", "F1")
_COL_CODIGOS_UEFA = ("UCL", "UEL", "UECL", "CL", "EL", "ECL", "CHAMPIONS", "EUROPA LEAGUE", "CONFERENCE LEAGUE", "UEFA")

# Diccionario estructurado de precios reales de referencia Kambi/Paf para totales colectivos
CUOTAS_BASE_COLECTIVAS = {
    # ⚽ Goles Totales
    (_COL_CAT_GOLES, "+1.5"): 1.28,
    (_COL_CAT_GOLES, "+2.5"): 1.85,
    (_COL_CAT_GOLES, "+3.5"): 2.95,
    (_COL_CAT_GOLES, "+4.5"): 5.10,
    (_COL_CAT_GOLES, "-2.5"): 1.92,
    (_COL_CAT_GOLES, "-3.5"): 1.36,
    # 🤝 Ambos Marcan (BTTS)
    (_COL_CAT_BTTS, "Sí (Ambos Marcan)"): 1.78,
    (_COL_CAT_BTTS, "No (Ambos Marcan)"): 2.00,
    # 🚩 Córners Totales
    (_COL_CAT_CORNERS, "+7.5"): 1.30,
    (_COL_CAT_CORNERS, "+8.5"): 1.52,
    (_COL_CAT_CORNERS, "+9.5"): 1.85,
    (_COL_CAT_CORNERS, "+10.5"): 2.35,
    (_COL_CAT_CORNERS, "+11.5"): 3.10,
    (_COL_CAT_CORNERS, "-8.5"): 2.30,
    (_COL_CAT_CORNERS, "-9.5"): 1.85,
    (_COL_CAT_CORNERS, "-10.5"): 1.50,
    # 🟨 Tarjetas Totales
    (_COL_CAT_TARJETAS, "+2.5"): 1.25,
    (_COL_CAT_TARJETAS, "+3.5"): 1.52,
    (_COL_CAT_TARJETAS, "+4.5"): 1.90,
    (_COL_CAT_TARJETAS, "+5.5"): 2.55,
    (_COL_CAT_TARJETAS, "+6.5"): 3.50,
    (_COL_CAT_TARJETAS, "-3.5"): 2.35,
    (_COL_CAT_TARJETAS, "-4.5"): 1.80,
    (_COL_CAT_TARJETAS, "-5.5"): 1.40,
}

# Catálogo plano: (metrica_col, categoria, codigo_linea, etiqueta_mercado, tipo_dir, umbral, cuota_base)
_COL_CATALOGO_MERCADOS = [
    # Goles Totales
    ("tot_goals", _COL_CAT_GOLES, "+1.5", "Goles Totales Más de 1.5", "over", 1.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_GOLES, "+1.5")]),
    ("tot_goals", _COL_CAT_GOLES, "+2.5", "Goles Totales Más de 2.5", "over", 2.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_GOLES, "+2.5")]),
    ("tot_goals", _COL_CAT_GOLES, "+3.5", "Goles Totales Más de 3.5", "over", 3.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_GOLES, "+3.5")]),
    ("tot_goals", _COL_CAT_GOLES, "+4.5", "Goles Totales Más de 4.5", "over", 4.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_GOLES, "+4.5")]),
    ("tot_goals", _COL_CAT_GOLES, "-2.5", "Goles Totales Menos de 2.5", "under", 2.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_GOLES, "-2.5")]),
    ("tot_goals", _COL_CAT_GOLES, "-3.5", "Goles Totales Menos de 3.5", "under", 3.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_GOLES, "-3.5")]),
    # Ambos Marcan (BTTS)
    ("btts", _COL_CAT_BTTS, "Sí (Ambos Marcan)", "Ambos Marcan: Sí", "over", 0.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_BTTS, "Sí (Ambos Marcan)")]),
    ("btts", _COL_CAT_BTTS, "No (Ambos Marcan)", "Ambos Marcan: No", "under", 0.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_BTTS, "No (Ambos Marcan)")]),
    # Córners Totales
    ("tot_corners", _COL_CAT_CORNERS, "+7.5", "Córners Totales Más de 7.5", "over", 7.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_CORNERS, "+7.5")]),
    ("tot_corners", _COL_CAT_CORNERS, "+8.5", "Córners Totales Más de 8.5", "over", 8.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_CORNERS, "+8.5")]),
    ("tot_corners", _COL_CAT_CORNERS, "+9.5", "Córners Totales Más de 9.5", "over", 9.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_CORNERS, "+9.5")]),
    ("tot_corners", _COL_CAT_CORNERS, "+10.5", "Córners Totales Más de 10.5", "over", 10.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_CORNERS, "+10.5")]),
    ("tot_corners", _COL_CAT_CORNERS, "+11.5", "Córners Totales Más de 11.5", "over", 11.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_CORNERS, "+11.5")]),
    ("tot_corners", _COL_CAT_CORNERS, "-8.5", "Córners Totales Menos de 8.5", "under", 8.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_CORNERS, "-8.5")]),
    ("tot_corners", _COL_CAT_CORNERS, "-9.5", "Córners Totales Menos de 9.5", "under", 9.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_CORNERS, "-9.5")]),
    ("tot_corners", _COL_CAT_CORNERS, "-10.5", "Córners Totales Menos de 10.5", "under", 10.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_CORNERS, "-10.5")]),
    # Tarjetas Totales
    ("tot_cards", _COL_CAT_TARJETAS, "+2.5", "Tarjetas Totales Más de 2.5", "over", 2.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_TARJETAS, "+2.5")]),
    ("tot_cards", _COL_CAT_TARJETAS, "+3.5", "Tarjetas Totales Más de 3.5", "over", 3.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_TARJETAS, "+3.5")]),
    ("tot_cards", _COL_CAT_TARJETAS, "+4.5", "Tarjetas Totales Más de 4.5", "over", 4.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_TARJETAS, "+4.5")]),
    ("tot_cards", _COL_CAT_TARJETAS, "+5.5", "Tarjetas Totales Más de 5.5", "over", 5.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_TARJETAS, "+5.5")]),
    ("tot_cards", _COL_CAT_TARJETAS, "+6.5", "Tarjetas Totales Más de 6.5", "over", 6.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_TARJETAS, "+6.5")]),
    ("tot_cards", _COL_CAT_TARJETAS, "-3.5", "Tarjetas Totales Menos de 3.5", "under", 3.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_TARJETAS, "-3.5")]),
    ("tot_cards", _COL_CAT_TARJETAS, "-4.5", "Tarjetas Totales Menos de 4.5", "under", 4.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_TARJETAS, "-4.5")]),
    ("tot_cards", _COL_CAT_TARJETAS, "-5.5", "Tarjetas Totales Menos de 5.5", "under", 5.5, CUOTAS_BASE_COLECTIVAS[(_COL_CAT_TARJETAS, "-5.5")]),
]

_COL_ESCALERAS_PILLS = {
    _COL_CAT_CORNERS: ["Todas", "+7.5", "+8.5", "+9.5", "+10.5", "+11.5", "-8.5", "-9.5", "-10.5"],
    _COL_CAT_TARJETAS: ["Todas", "+2.5", "+3.5", "+4.5", "+5.5", "+6.5", "-3.5", "-4.5", "-5.5"],
    _COL_CAT_GOLES: ["Todas", "+1.5", "+2.5", "+3.5", "+4.5", "-2.5", "-3.5"],
    _COL_CAT_BTTS: ["Todos", "Sí (Ambos Marcan)", "No (Ambos Marcan)"],
}


def _col_filtrar_dataframe_competicion(df, comp_sel):
    """Compuerta excluyente de competición (Cascade Filtering).

    - 'Todas las Competiciones': devuelve copia íntegra, divs_activas=None, es_uefa=False.
    - 'Competiciones UEFA': filtra explícitamente registros donde 'Div' pertenece a torneos UEFA.
    - Ligas domésticas ('SP1', 'E0', 'I1', 'D1', 'F1'): filtra exclusivamente el código de liga.
    """
    if df is None or df.empty:
        return pd.DataFrame(), None, False
    if not comp_sel or comp_sel == "Todas las Competiciones":
        return df.copy(), None, False

    col_div = "Div" if "Div" in df.columns else next((c for c in df.columns if "Div" in str(c) or "Torneo" in str(c)), None)
    if not col_div:
        return df.copy(), None, False

    s_div = df[col_div].astype(str).str.strip()
    if comp_sel == "Competiciones UEFA":
        s_up = s_div.str.upper()
        mask_uefa = s_up.isin(_COL_CODIGOS_UEFA) | s_up.str.contains("CHAMPIONS|EUROPA|CONFERENCE|UEFA", regex=True, na=False)
        return df[mask_uefa].copy(), ("UCL", "UEL", "UECL"), True

    divs_cfg = _IND_FILTRO_COMPETICION.get(comp_sel)
    if divs_cfg:
        return df[s_div.isin(divs_cfg)].copy(), tuple(divs_cfg), False
    return df.copy(), None, False


def _col_proximos_cruces_competicion(df_filtrado, divs_activas=None, es_uefa=False):
    """Extrae próximos cruces pertenecientes EXCLUSIVAMENTE a la competición seleccionada.

    Si se selecciona 'Competiciones UEFA', jamás consulta ni mezcla partidos de ligas nacionales
    (LaLiga, Premier League, Serie A, Bundesliga, Ligue 1).
    """
    cols = ["Equipo", "px_rival", "px_cond", "px_dt", "px_div"]
    if df_filtrado is None or df_filtrado.empty:
        return pd.DataFrame(columns=cols)

    hoy = pd.Timestamp.now().normalize()
    limite = hoy + pd.Timedelta(days=_IND_HORIZONTE_DIAS + 1)
    partes = []

    fut = df_filtrado[(~df_filtrado["Jugado"]) & (df_filtrado["Date"] >= hoy) & (df_filtrado["Date"] < limite)]
    if not fut.empty:
        partes.append(pd.DataFrame({
            "Equipo": fut["HomeTeam"].values,
            "px_rival": fut["AwayTeam"].values,
            "px_cond": "Local",
            "px_dt": fut["Date"].values,
            "px_div": fut["Div"].values,
            "src": 1,
        }))
        partes.append(pd.DataFrame({
            "Equipo": fut["AwayTeam"].values,
            "px_rival": fut["HomeTeam"].values,
            "px_cond": "Visitante",
            "px_dt": fut["Date"].values,
            "px_div": fut["Div"].values,
            "src": 1,
        }))

    # Solo consultar API de ligas domésticas si NO estamos en modo exclusivo UEFA
    if not es_uefa:
        if divs_activas is None:
            divs_api = list(_IND_SPORT_KEYS.keys())
        else:
            divs_api = [d for d in divs_activas if d in _IND_SPORT_KEYS]

        if divs_api:
            equipos_comp = sorted(
                pd.unique(pd.concat([df_filtrado["HomeTeam"], df_filtrado["AwayTeam"]]).dropna().astype(str)).tolist()
            )
            for d_code in divs_api:
                sk = _IND_SPORT_KEYS.get(d_code)
                eventos = _ind_fixtures_vivo(sk) or []
                if not eventos:
                    continue
                ev = pd.DataFrame(eventos).dropna(subset=["home_team", "away_team", "commence_time"])
                if ev.empty:
                    continue
                nombres = pd.unique(pd.concat([ev["home_team"], ev["away_team"]]))
                mapa = {n: ev_resolver_equipo(n, equipos_comp) for n in nombres}
                ev["h"] = ev["home_team"].map(mapa)
                ev["a"] = ev["away_team"].map(mapa)
                ev = ev.dropna(subset=["h", "a"])
                ev = ev[ev["h"] != ev["a"]]
                if ev.empty:
                    continue
                ts = pd.to_datetime(ev["commence_time"], errors="coerce", utc=True)
                try:
                    ts = ts.dt.tz_convert("Europe/Madrid")
                except Exception:
                    pass
                ev["dt"] = ts.dt.tz_localize(None)
                ev = ev[(ev["dt"] >= hoy) & (ev["dt"] < limite)]
                if not ev.empty:
                    partes.append(pd.DataFrame({
                        "Equipo": ev["h"].values,
                        "px_rival": ev["a"].values,
                        "px_cond": "Local",
                        "px_dt": ev["dt"].values,
                        "px_div": d_code,
                        "src": 0,
                    }))
                    partes.append(pd.DataFrame({
                        "Equipo": ev["a"].values,
                        "px_rival": ev["h"].values,
                        "px_cond": "Visitante",
                        "px_dt": ev["dt"].values,
                        "px_div": d_code,
                        "src": 0,
                    }))

    if not partes:
        return pd.DataFrame(columns=cols)
    todo = pd.concat(partes, ignore_index=True)
    todo["px_dt"] = pd.to_datetime(todo["px_dt"])
    todo["_dia"] = todo["px_dt"].dt.normalize()
    todo = todo.sort_values(["_dia", "src", "px_dt"], kind="stable")
    return todo.drop_duplicates("Equipo", keep="first")[cols].reset_index(drop=True)


def _resolver_arbitro_partido(fixture_data, db_conn=None):
    """Motor resiliente de resolución arbitral cruzado contra el censo transversal de 165 árbitros."""
    badge_pend = (
        "<span style='background:rgba(88,110,150,0.14);border:1px solid #3b4a63;color:#8b9bb4;"
        "display:inline-block;padding:3px 9px;border-radius:10px;font-size:0.73rem;font-weight:600;white-space:nowrap;'>"
        "⚖️ Colegiado: Pendiente de designación oficial</span>"
    )
    res_pendiente = {
        "estado": "PENDIENTE",
        "nombre": "",
        "media": None,
        "nivel": "pendiente",
        "texto": "⚖️ Colegiado: Pendiente de designación oficial",
        "badge_html": badge_pend,
    }
    if not isinstance(fixture_data, dict):
        return res_pendiente

    nombre_raw = ""
    for k in ("referee", "arbitro", "Referee", "Arbitro", "colegiado"):
        val = fixture_data.get(k)
        if val is not None and str(val).strip() and str(val).strip().lower() not in (
            "nan", "none", "null", "pendiente", "sin designar", "(pendiente de designación)"
        ):
            nombre_raw = str(val).strip()
            break

    h_eq = str(fixture_data.get("local") or fixture_data.get("HomeTeam") or fixture_data.get("equipo_local") or "").strip()
    a_eq = str(fixture_data.get("visitante") or fixture_data.get("AwayTeam") or fixture_data.get("equipo_visitante") or "").strip()
    fecha_raw = fixture_data.get("fecha") or fixture_data.get("px_dt") or fixture_data.get("Date")

    if not nombre_raw and h_eq and a_eq:
        try:
            for fx_m in st.session_state.get("ev_fix_manual", []):
                if isinstance(fx_m, dict) and fx_m.get("Local") == h_eq and fx_m.get("Visitante") == a_eq:
                    r_man = str(fx_m.get("Arbitro") or fx_m.get("referee") or "").strip()
                    if r_man and "pendiente" not in r_man.lower():
                        nombre_raw = r_man
                        break
        except Exception:
            pass

    if not nombre_raw and h_eq and a_eq:
        cerrar_conn = False
        conn_use = db_conn
        try:
            if conn_use is None and RUTA_BD.exists():
                conn_use = sqlite3.connect(RUTA_BD)
                cerrar_conn = True
            if conn_use is not None:
                f_ts = pd.to_datetime(fecha_raw, errors="coerce", dayfirst=True) if fecha_raw is not None else pd.NaT
                if pd.notna(f_ts):
                    f_str = pd.Timestamp(f_ts).strftime("%d/%m/%Y")
                    row = conn_use.execute(
                        "SELECT Referee FROM historico_reciente WHERE HomeTeam=? AND AwayTeam=? AND Date=? "
                        "AND Referee IS NOT NULL AND TRIM(Referee)<>'' ORDER BY rowid DESC LIMIT 1",
                        (h_eq, a_eq, f_str),
                    ).fetchone()
                    if not row:
                        row = conn_use.execute(
                            "SELECT arbitro FROM partidos WHERE equipo_local=? AND equipo_visitante=? AND fecha=? "
                            "AND arbitro IS NOT NULL AND TRIM(arbitro)<>'' LIMIT 1",
                            (h_eq, a_eq, f_str),
                        ).fetchone()
                    if row and row[0]:
                        nombre_raw = str(row[0]).strip()
                if not nombre_raw:
                    row = conn_use.execute(
                        "SELECT Referee FROM historico_reciente WHERE HomeTeam=? AND AwayTeam=? AND FTHG IS NULL "
                        "AND Referee IS NOT NULL AND TRIM(Referee)<>'' ORDER BY rowid DESC LIMIT 1",
                        (h_eq, a_eq),
                    ).fetchone()
                    if row and row[0]:
                        nombre_raw = str(row[0]).strip()
        except Exception:
            pass
        finally:
            if cerrar_conn and conn_use is not None:
                try:
                    conn_use.close()
                except Exception:
                    pass

    canon = _normalizar_nombre_arbitro(nombre_raw) if nombre_raw else ""
    if not canon:
        return res_pendiente

    try:
        censo = _obtener_censo_global_arbitros()
    except Exception:
        censo = {}

    info_censo = censo.get(canon)
    if info_censo is None and censo:
        canon_low = canon.lower()
        for v_info in censo.values():
            vars_low = [str(x).lower() for x in v_info.get("variantes", [])]
            if canon_low == str(v_info.get("nombre", "")).lower() or canon_low in vars_low:
                info_censo = v_info
                break

    if info_censo is not None:
        nom_final = str(info_censo.get("nombre", canon))
        am = float(info_censo.get("media_amarillas_global", 4.1))
        rj = float(info_censo.get("media_rojas_global", 0.18))
        media_tarj = round(am + rj, 2)
    else:
        nom_final = canon
        media_tarj = 4.25

    if media_tarj >= 5.0:
        nivel = "favorable"
        texto = f"⚖️ Colegiado Favorable: {nom_final} ({media_tarj:.1f} tarj/partido)"
        estilo = "background:rgba(63,185,80,0.16);border:1px solid #3fb950;color:#3fb950;"
    elif media_tarj <= 3.8:
        nivel = "under"
        texto = f"⚠️ Colegiado Under: {nom_final} ({media_tarj:.1f} tarj/partido)"
        estilo = "background:rgba(240,136,62,0.16);border:1px solid #f0883e;color:#f0883e;"
    else:
        nivel = "neutro"
        texto = f"⚖️ Colegiado: {nom_final} ({media_tarj:.1f} tarj/partido)"
        estilo = "background:#161b22;border:1px solid #30363d;color:#c9d1d9;"

    badge_html = (
        f"<span style='{estilo}display:inline-block;padding:3px 9px;border-radius:10px;"
        f"font-size:0.73rem;font-weight:700;white-space:nowrap;'>{_html_ind.escape(texto)}</span>"
    )
    return {
        "estado": "DISPONIBLE",
        "nombre": nom_final,
        "media": media_tarj,
        "nivel": nivel,
        "texto": texto,
        "badge_html": badge_html,
    }


@st.cache_data(ttl=600, show_spinner=False)
def _col_calcular_tiers_equipos(df_global):
    """Clasifica determinísticamente a los equipos según puntos reales acumulados en ligas domésticas.

    - Ligas domésticas (LaLiga SP1, Premier E0, Serie A I1, Bundesliga D1, Ligue 1 F1):
      * Posiciones 1 a 6: '👑 Top 6 / Líderes'
      * Posiciones 7 a 14: '⚖️ Zona Media'
      * Posiciones 15 en adelante: '⚠️ Zona Baja/Descenso'
    - Equipos exclusivamente UEFA o sin tabla doméstica suficiente:
      * Devuelve '🏆 Competición Europea' (evitando catalogar arbitrariamente a clubs modestos como Top 6).
    """
    if df_global is None or df_global.empty:
        return {}
    jug = df_global[df_global["Jugado"] & df_global["Div"].isin(_COL_LIGAS_DOMESTICAS)]
    if jug.empty:
        return {}

    fthg = jug["FTHG"].values.astype(int)
    ftag = jug["FTAG"].values.astype(int)
    pts_h = np.where(fthg > ftag, 3, np.where(fthg == ftag, 1, 0))
    pts_a = np.where(ftag > fthg, 3, np.where(fthg == ftag, 1, 0))
    dg_h = fthg - ftag
    dg_a = ftag - fthg

    largo = pd.concat([
        pd.DataFrame({"Equipo": jug["HomeTeam"].values, "Div": jug["Div"].values, "pts": pts_h, "dg": dg_h, "gf": fthg}),
        pd.DataFrame({"Equipo": jug["AwayTeam"].values, "Div": jug["Div"].values, "pts": pts_a, "dg": dg_a, "gf": ftag}),
    ], ignore_index=True)

    tabla = largo.groupby(["Div", "Equipo"], as_index=False).agg(
        pj=("pts", "size"),
        pts=("pts", "sum"),
        dg=("dg", "sum"),
        gf=("gf", "sum"),
    )
    # Solo clasificar si el club tiene muestra real en su liga doméstica
    tabla = tabla[tabla["pj"] >= 3].copy()
    if tabla.empty:
        return {}
    tabla["ppg"] = tabla["pts"] / tabla["pj"].clip(lower=1)
    tabla = tabla.sort_values(["Div", "pts", "ppg", "dg", "gf"], ascending=[True, False, False, False, False], kind="stable")
    tabla["pos"] = tabla.groupby("Div").cumcount() + 1

    tier_col = np.where(
        tabla["pos"] <= 6,
        _COL_TIER_TOP,
        np.where(tabla["pos"] <= 14, _COL_TIER_MEDIA, _COL_TIER_BAJA),
    )
    return dict(zip(tabla["Equipo"], tier_col))


def _col_aplanar_partidos_colectivos(df):
    """Aplana en formato largo vectorizado (1 fila por equipo y partido jugado) las métricas colectivas."""
    if df is None or df.empty:
        return pd.DataFrame()
    jug = df[df["Jugado"]]
    if jug.empty:
        return pd.DataFrame()

    fthg = jug["FTHG"].values.astype(float)
    ftag = jug["FTAG"].values.astype(float)
    tot_goals = fthg + ftag
    btts_val = ((fthg > 0) & (ftag > 0)).astype(float)

    hc = jug["HC"].values.astype(float)
    ac = jug["AC"].values.astype(float)
    sin_corn = np.isnan(hc) | np.isnan(ac) | ((hc + ac) == 0)
    tot_corners = np.where(sin_corn, np.nan, hc + ac)

    hy = jug["HY"].values.astype(float)
    ay = jug["AY"].values.astype(float)
    hr = np.nan_to_num(jug["HR"].values.astype(float), nan=0.0)
    ar = np.nan_to_num(jug["AR"].values.astype(float), nan=0.0)
    sin_cards = np.isnan(hy) | np.isnan(ay)
    tot_cards = np.where(sin_cards, np.nan, hy + ay + hr + ar)

    marcador_h = pd.Series(fthg.astype(int)).astype(str).values + "-" + pd.Series(ftag.astype(int)).astype(str).values
    marcador_a = pd.Series(ftag.astype(int)).astype(str).values + "-" + pd.Series(fthg.astype(int)).astype(str).values

    df_h = pd.DataFrame({
        "Fecha": jug["Date"].values,
        "Equipo": jug["HomeTeam"].values,
        "Rival": jug["AwayTeam"].values,
        "Cond": "L",
        "Div": jug["Div"].values,
        "gf": fthg,
        "ga": ftag,
        "tot_goals": tot_goals,
        "btts": btts_val,
        "tot_corners": tot_corners,
        "tot_cards": tot_cards,
        "score_str": marcador_h,
    })
    df_a = pd.DataFrame({
        "Fecha": jug["Date"].values,
        "Equipo": jug["AwayTeam"].values,
        "Rival": jug["HomeTeam"].values,
        "Cond": "V",
        "Div": jug["Div"].values,
        "gf": ftag,
        "ga": fthg,
        "tot_goals": tot_goals,
        "btts": btts_val,
        "tot_corners": tot_corners,
        "tot_cards": tot_cards,
        "score_str": marcador_a,
    })
    largo = pd.concat([df_h, df_a], ignore_index=True)
    largo["rival3"] = largo["Rival"].astype(str).str.upper().str[:3].str.strip()
    return largo.sort_values("Fecha", ascending=False, kind="stable").reset_index(drop=True)


def _col_poisson_prob_real(lam_val, k_int, tipo):
    """Cálculo exacto de la probabilidad acumulada de Poisson para un parámetro Lambda (λ) real.

    P(X >= k) = 1.0 - sum((math.exp(-λ) * (λ ** i)) / math.factorial(i) for i in range(k))
    Acotado en [0.02, 0.98].
    """
    lam = max(float(lam_val), 0.05)
    p_over = 1.0 - sum((math.exp(-lam) * (lam ** i)) / math.factorial(i) for i in range(int(k_int)))
    prob = p_over if tipo == "over" else (1.0 - p_over)
    return max(min(float(prob), 0.98), 0.02)


def _col_evaluar_bloque_ventana(largo_df, ventana, min_winrate, mercados_filtrados, liga_eq_map, min_partidos=None):
    """Pipeline de cálculo dinámico por equipo y mercado con Poisson real (cero constantes clonadas)."""
    if largo_df is None or largo_df.empty or not mercados_filtrados:
        return pd.DataFrame()

    min_req = int(min_partidos) if min_partidos is not None else int(ventana)
    partes = []
    metricas_activas = sorted({m[0] for m in mercados_filtrados})
    for met in metricas_activas:
        sub = largo_df.dropna(subset=[met])
        if sub.empty:
            continue
        w = sub.groupby("Equipo", sort=False).head(int(ventana))
        w = w[w.groupby("Equipo")[met].transform("size") >= min_req]
        if w.empty:
            continue
        if met == "btts":
            etq_val = w["score_str"].astype(str) + " vs " + w["rival3"]
        else:
            etq_val = w[met].astype(int).astype(str) + " vs " + w["rival3"]
        w = w.assign(chip=etq_val).iloc[::-1]

        if met == "btts":
            agg = w.groupby("Equipo", sort=False).agg(
                Lambda=(met, "mean"),
                lam_gf=("gf", "mean"),
                lam_ga=("ga", "mean"),
                n_real=(met, "size"),
                vals=(met, list),
                chips=("chip", list),
            )
        else:
            agg = w.groupby("Equipo", sort=False).agg(
                Lambda=(met, "mean"),
                n_real=(met, "size"),
                vals=(met, list),
                chips=("chip", list),
            )

        n_arr_all = agg["n_real"].values.astype(int)
        for (m_met, cat, cod_lin, etq_mkt, tipo, umbral, cuota_base) in mercados_filtrados:
            if m_met != met:
                continue
            cuota_b = float(CUOTAS_BASE_COLECTIVAS.get((cat, cod_lin), cuota_base))
            k_int = int(math.floor(umbral)) + 1

            if met == "btts":
                hits = np.array([
                    sum((v >= 0.5) if tipo == "over" else (v < 0.5) for v in vs)
                    for vs in agg["vals"]
                ], dtype=int)
                chips_list = [
                    [(c, (v >= 0.5) if tipo == "over" else (v < 0.5)) for c, v in zip(cs, vs)]
                    for cs, vs in zip(agg["chips"], agg["vals"])
                ]
            else:
                hits = np.array([
                    sum((v >= k_int) if tipo == "over" else (v < k_int) for v in vs)
                    for vs in agg["vals"]
                ], dtype=int)
                chips_list = [
                    [(c, (v >= k_int) if tipo == "over" else (v < k_int)) for c, v in zip(cs, vs)]
                    for cs, vs in zip(agg["chips"], agg["vals"])
                ]

            wr = (hits / n_arr_all.astype(float)) * 100.0
            mask_ok = wr >= float(min_winrate)
            if not np.any(mask_ok):
                continue

            eqs_ok = agg.index.values[mask_ok]
            hits_ok = hits[mask_ok]
            n_ok = n_arr_all[mask_ok]
            wr_ok = np.round(wr[mask_ok], 1)
            lam_vals = agg["Lambda"].values[mask_ok]
            chips_ok = [c for c, ok in zip(chips_list, mask_ok) if ok]

            if met == "btts":
                lgf_ok = agg["lam_gf"].values[mask_ok]
                lga_ok = agg["lam_ga"].values[mask_ok]
                probs_ok = []
                for lgf, lga in zip(lgf_ok, lga_ok):
                    p_si = (1.0 - math.exp(-max(float(lgf), 0.05))) * (1.0 - math.exp(-max(float(lga), 0.05)))
                    p_b = p_si if tipo == "over" else (1.0 - p_si)
                    probs_ok.append(max(min(float(p_b), 0.98), 0.02))
                probs_arr = np.array(probs_ok, dtype=float)
            else:
                probs_arr = np.array(
                    [_col_poisson_prob_real(lv, k_int, tipo) for lv in lam_vals],
                    dtype=float,
                )

            cuota_justa_arr = np.round(1.0 / probs_arr, 2)
            ev_arr = np.round(((probs_arr * cuota_b) - 1.0) * 100.0, 1)

            partes.append(pd.DataFrame({
                "Equipo": eqs_ok,
                "Liga": [liga_eq_map.get(e, "N/A") for e in eqs_ok],
                "Categoria": cat,
                "CodigoLinea": cod_lin,
                "Mercado": etq_mkt,
                "Tipo": tipo,
                "Umbral": umbral,
                "Hits": hits_ok,
                "N": n_ok,
                "Winrate": wr_ok,
                "Lambda": np.round(lam_vals.astype(float), 2),
                "Prob": probs_arr,
                "CuotaJusta": cuota_justa_arr,
                "CuotaBase": cuota_b,
                "EV": ev_arr,
                "Chips": chips_ok,
            }))

    if not partes:
        return pd.DataFrame()
    return pd.concat(partes, ignore_index=True)


def _escanear_radar_colectivo_vectorial(
    df_filtrado,
    df_global_tiers,
    ventana=5,
    min_winrate=75.0,
    categoria=_COL_CAT_TODOS,
    direccion="Todos",
    codigo_linea="Todas",
    modo_condicion="🌐 Global",
    divs_activas=None,
    es_uefa=False,
    filtro_busqueda="Todos los Equipos",
    tier_rival=_COL_TIER_TODOS,
):
    """Escáner vectorial estricto sobre `df_filtrado` (Cascade Filtering + Poisson real + Choque agrupado)."""
    if df_filtrado is None or df_filtrado.empty:
        return pd.DataFrame(), [], []

    tiers_dom_map = _col_calcular_tiers_equipos(df_global_tiers)
    largo_scope = _col_aplanar_partidos_colectivos(df_filtrado)
    if largo_scope.empty:
        return pd.DataFrame(), [], []

    # Liga mostrada por equipo: en UEFA estrictamente torneo UEFA; en Global preferencia doméstica
    if es_uefa:
        liga_eq_map = (
            largo_scope.drop_duplicates("Equipo")
            .set_index("Equipo")["Div"]
            .map(lambda d: _IND_NOMBRE_DIV.get(d, d))
            .to_dict()
        )
    else:
        pri_l = largo_scope.assign(_u=largo_scope["Div"].isin(_IND_UEFA_DIVS).astype(int)).sort_values(
            ["_u", "Fecha"], ascending=[True, False], kind="stable"
        )
        liga_eq_map = (
            pri_l.drop_duplicates("Equipo")
            .set_index("Equipo")["Div"]
            .map(lambda d: _IND_NOMBRE_DIV.get(d, d))
            .to_dict()
        )

    mercados_activos = []
    for m in _COL_CATALOGO_MERCADOS:
        m_met, cat, cod_lin, etq_mkt, tipo, umbral, cuota_base = m
        if categoria != _COL_CAT_TODOS and cat != categoria:
            continue
        if direccion in ("over", "under") and tipo != direccion:
            continue
        if codigo_linea not in ("Todas", "Todos", None) and cod_lin != codigo_linea:
            continue
        mercados_activos.append(m)

    # Próximos cruces estrictamente pertenecientes a la competición filtrada
    px_df = _col_proximos_cruces_competicion(df_filtrado, divs_activas=divs_activas, es_uefa=es_uefa)
    px_map = {}
    duelos_proximos = []
    vistos_duelos = set()

    conn_arb = None
    try:
        if RUTA_BD.exists():
            conn_arb = sqlite3.connect(RUTA_BD)
        if not px_df.empty:
            for eq_r, riv_r, cond_r, dt_r, div_r in zip(
                px_df["Equipo"].values,
                px_df["px_rival"].values,
                px_df["px_cond"].values,
                px_df["px_dt"].values,
                px_df["px_div"].values,
            ):
                loc_d = str(eq_r) if cond_r == "Local" else str(riv_r)
                vis_d = str(riv_r) if cond_r == "Local" else str(eq_r)
                par_key = (loc_d, vis_d)
                if par_key not in vistos_duelos:
                    vistos_duelos.add(par_key)
                    arb_info = _resolver_arbitro_partido(
                        {"local": loc_d, "visitante": vis_d, "fecha": dt_r}, db_conn=conn_arb
                    )
                    duelos_proximos.append({
                        "local": loc_d,
                        "visitante": vis_d,
                        "px_dt": dt_r,
                        "px_div": str(div_r),
                        "liga_duelo": _IND_NOMBRE_DIV.get(str(div_r), str(div_r)),
                        "arbitro": arb_info,
                    })
                px_map[str(eq_r)] = {
                    "px_rival": str(riv_r),
                    "px_cond": str(cond_r),
                    "px_dt": dt_r,
                    "px_div": str(div_r),
                }
    finally:
        if conn_arb is not None:
            try:
                conn_arb.close()
            except Exception:
                pass

    arb_por_duelo = {(d["local"], d["visitante"]): d["arbitro"] for d in duelos_proximos}
    arb_pendiente_default = _resolver_arbitro_partido({})

    def _resolver_tier_club(eq_nom):
        if es_uefa:
            return tiers_dom_map.get(eq_nom, _COL_TIER_UEFA)
        return tiers_dom_map.get(eq_nom, _COL_TIER_UEFA if liga_eq_map.get(eq_nom) in ("Champions", "Europa League", "Conference") else _COL_TIER_GENERAL)

    # Evaluación vectorial sobre el subconjunto exclusivo de la competición seleccionada
    min_p_glob = min(int(ventana), 2) if es_uefa else int(ventana)
    min_p_cond = min(int(ventana), 2) if es_uefa else int(ventana)
    df_eval_global = _col_evaluar_bloque_ventana(
        largo_scope, ventana, min_winrate, mercados_activos, liga_eq_map, min_partidos=min_p_glob
    )
    df_eval_loc = _col_evaluar_bloque_ventana(
        largo_scope[largo_scope["Cond"] == "L"], ventana, min_winrate, mercados_activos, liga_eq_map, min_partidos=min_p_cond
    )
    df_eval_vis = _col_evaluar_bloque_ventana(
        largo_scope[largo_scope["Cond"] == "V"], ventana, min_winrate, mercados_activos, liga_eq_map, min_partidos=min_p_cond
    )

    # En competiciones UEFA, si los mercados de Córners/Tarjetas carecen de muestras en ESPN,
    # evaluar las métricas secundarias exclusivamente para los clubes UEFA activos manteniendo Liga y Cruce UEFA
    if es_uefa and df_global_tiers is not None and not df_global_tiers.empty:
        mkts_sec = [m for m in mercados_activos if m[0] in ("tot_corners", "tot_cards")]
        if mkts_sec:
            eqs_uefa_activos = set(largo_scope["Equipo"].unique())
            largo_sec = _col_aplanar_partidos_colectivos(df_global_tiers)
            largo_sec = largo_sec[largo_sec["Equipo"].isin(eqs_uefa_activos)]
            if not largo_sec.empty:
                sec_glob = _col_evaluar_bloque_ventana(largo_sec, ventana, min_winrate, mkts_sec, liga_eq_map)
                sec_loc = _col_evaluar_bloque_ventana(largo_sec[largo_sec["Cond"] == "L"], ventana, min_winrate, mkts_sec, liga_eq_map)
                sec_vis = _col_evaluar_bloque_ventana(largo_sec[largo_sec["Cond"] == "V"], ventana, min_winrate, mkts_sec, liga_eq_map)
                if not sec_glob.empty:
                    df_eval_global = pd.concat([df_eval_global, sec_glob], ignore_index=True) if not df_eval_global.empty else sec_glob
                if not sec_loc.empty:
                    df_eval_loc = pd.concat([df_eval_loc, sec_loc], ignore_index=True) if not df_eval_loc.empty else sec_loc
                if not sec_vis.empty:
                    df_eval_vis = pd.concat([df_eval_vis, sec_vis], ignore_index=True) if not df_eval_vis.empty else sec_vis

    if "Estricto" in str(modo_condicion):
        partes_est = []
        if not df_eval_loc.empty:
            cond_px_loc = df_eval_loc["Equipo"].map(lambda e: px_map.get(e, {}).get("px_cond", "Local"))
            partes_est.append(df_eval_loc[cond_px_loc == "Local"])
        if not df_eval_vis.empty:
            cond_px_vis = df_eval_vis["Equipo"].map(lambda e: px_map.get(e, {}).get("px_cond", "Visitante"))
            partes_est.append(df_eval_vis[cond_px_vis == "Visitante"])
        df_feed = pd.concat(partes_est, ignore_index=True) if partes_est else pd.DataFrame()
    else:
        df_feed = df_eval_global.copy()

    if not df_feed.empty:
        df_feed["px_rival"] = df_feed["Equipo"].map(lambda e: px_map.get(e, {}).get("px_rival"))
        df_feed["px_cond"] = df_feed["Equipo"].map(lambda e: px_map.get(e, {}).get("px_cond"))
        df_feed["px_dt"] = df_feed["Equipo"].map(lambda e: px_map.get(e, {}).get("px_dt"))
        df_feed["TierEquipo"] = df_feed["Equipo"].map(_resolver_tier_club)
        df_feed["TierRival"] = df_feed["px_rival"].map(lambda r: _resolver_tier_club(r) if r else None)
        df_feed["TierRef"] = np.where(
            df_feed["TierRival"].notna(),
            df_feed["TierRival"],
            df_feed["TierEquipo"],
        )

        def _arb_de_equipo(eq_nom):
            info = px_map.get(eq_nom)
            if not info or not info.get("px_rival"):
                return arb_pendiente_default
            par = (eq_nom, info["px_rival"]) if info.get("px_cond") == "Local" else (info["px_rival"], eq_nom)
            return arb_por_duelo.get(par, arb_pendiente_default)

        df_feed["ArbitroInfo"] = df_feed["Equipo"].map(_arb_de_equipo)

        if filtro_busqueda and filtro_busqueda != "Todos los Equipos":
            if " vs " in filtro_busqueda:
                h_b, a_b = [x.strip() for x in filtro_busqueda.split(" vs ", 1)]
                df_feed = df_feed[df_feed["Equipo"].isin([h_b, a_b])]
            else:
                df_feed = df_feed[
                    (df_feed["Equipo"] == filtro_busqueda) | (df_feed["px_rival"] == filtro_busqueda)
                ]

        if tier_rival and tier_rival != _COL_TIER_TODOS:
            df_feed = df_feed[(df_feed["TierRef"] == tier_rival) | (df_feed["TierEquipo"] == tier_rival)]

        df_feed = df_feed.sort_values(["EV", "Winrate", "Lambda"], ascending=[False, False, False], kind="stable").reset_index(drop=True)

    # ── Choque de Tendencias Agrupado por Fixture Único (fixture_id = f"{local}_{visitante}_{fecha}") ──
    base_h = df_eval_loc if not df_eval_loc.empty else df_eval_global
    base_a = df_eval_vis if not df_eval_vis.empty else df_eval_global
    idx_h = {
        (r["Equipo"], r["Categoria"], r["CodigoLinea"]): r
        for r in (base_h.to_dict("records") if not base_h.empty else [])
    }
    idx_a = {
        (r["Equipo"], r["Categoria"], r["CodigoLinea"]): r
        for r in (base_a.to_dict("records") if not base_a.empty else [])
    }
    idx_g = {
        (r["Equipo"], r["Categoria"], r["CodigoLinea"]): r
        for r in (df_eval_global.to_dict("records") if not df_eval_global.empty else [])
    }

    choques_por_fixture = {}
    equipos_scope_set = set(largo_scope["Equipo"].unique())
    for duelo in duelos_proximos:
        loc_eq = duelo["local"]
        vis_eq = duelo["visitante"]
        if loc_eq not in equipos_scope_set and vis_eq not in equipos_scope_set:
            continue
        if filtro_busqueda and filtro_busqueda != "Todos los Equipos":
            if " vs " in filtro_busqueda:
                if f"{loc_eq} vs {vis_eq}" != filtro_busqueda:
                    continue
            elif filtro_busqueda not in (loc_eq, vis_eq):
                continue

        t_loc = _resolver_tier_club(loc_eq)
        t_vis = _resolver_tier_club(vis_eq)
        if tier_rival and tier_rival != _COL_TIER_TODOS:
            if t_loc != tier_rival and t_vis != tier_rival:
                continue

        f_ts = pd.Timestamp(duelo["px_dt"]) if pd.notna(duelo["px_dt"]) else None
        fecha_key = f_ts.strftime("%Y-%m-%d") if f_ts is not None else "prox"
        fixture_id = f"{loc_eq}_{vis_eq}_{fecha_key}"

        for (m_met, cat, cod_lin, etq_mkt, tipo, umbral, cuota_base) in mercados_activos:
            rh = idx_h.get((loc_eq, cat, cod_lin)) or idx_g.get((loc_eq, cat, cod_lin))
            ra = idx_a.get((vis_eq, cat, cod_lin)) or idx_g.get((vis_eq, cat, cod_lin))
            if not rh or not ra:
                continue

            cuota_b = float(CUOTAS_BASE_COLECTIVAS.get((cat, cod_lin), cuota_base))
            lam_comb = round((float(rh["Lambda"]) + float(ra["Lambda"])) / 2.0, 2)
            p_comb = max(min((float(rh["Prob"]) + float(ra["Prob"])) / 2.0, 0.98), 0.02)
            cj_comb = round(1.0 / p_comb, 2)
            ev_comb = round(((p_comb * cuota_b) - 1.0) * 100.0, 1)
            wr_comb = round((float(rh["Winrate"]) + float(ra["Winrate"])) / 2.0, 1)

            item_conv = {
                "categoria": cat,
                "codigo_linea": cod_lin,
                "mercado": f"{cat} · {etq_mkt}",
                "etiqueta_corta": etq_mkt,
                "racha_local": rh,
                "racha_visit": ra,
                "lam_comb": lam_comb,
                "wr_comb": wr_comb,
                "prob_comb": p_comb,
                "cuota_justa": cj_comb,
                "cuota_base": cuota_b,
                "ev_comb": ev_comb,
            }
            if fixture_id not in choques_por_fixture:
                choques_por_fixture[fixture_id] = {
                    "fixture_id": fixture_id,
                    "partido": f"{loc_eq} vs {vis_eq}",
                    "local": loc_eq,
                    "visitante": vis_eq,
                    "liga": duelo.get("liga_duelo") or rh.get("Liga", "N/A"),
                    "px_dt": duelo["px_dt"],
                    "arbitro": duelo["arbitro"],
                    "tier_local": t_loc,
                    "tier_visit": t_vis,
                    "convergencias": [],
                }
            choques_por_fixture[fixture_id]["convergencias"].append(item_conv)

    choques_agrupados = []
    for fx_id, grp in choques_por_fixture.items():
        convs = sorted(grp["convergencias"], key=lambda x: (x["ev_comb"], x["wr_comb"]), reverse=True)
        grp["convergencias"] = convs
        grp["max_ev"] = max(c["ev_comb"] for c in convs)
        grp["max_wr"] = max(c["wr_comb"] for c in convs)
        grp["n_conv"] = len(convs)
        choques_agrupados.append(grp)

    choques_agrupados.sort(key=lambda x: (x["n_conv"], x["max_ev"], x["max_wr"]), reverse=True)

    # ── Generador de Packs Bet Builder para Paf (Mismo Encuentro de la Competición Activa, Winrate >= 80%) ──
    df_all_cats = _col_evaluar_bloque_ventana(
        largo_scope, ventana, 80.0, _COL_CATALOGO_MERCADOS, liga_eq_map, min_partidos=min_p_glob
    )
    if es_uefa and df_global_tiers is not None and not df_global_tiers.empty:
        mkts_all_sec = [m for m in _COL_CATALOGO_MERCADOS if m[0] in ("tot_corners", "tot_cards")]
        eqs_uefa_act = set(largo_scope["Equipo"].unique())
        largo_sec_all = _col_aplanar_partidos_colectivos(df_global_tiers)
        largo_sec_all = largo_sec_all[largo_sec_all["Equipo"].isin(eqs_uefa_act)]
        if not largo_sec_all.empty:
            sec_all = _col_evaluar_bloque_ventana(largo_sec_all, ventana, 80.0, mkts_all_sec, liga_eq_map)
            if not sec_all.empty:
                df_all_cats = pd.concat([df_all_cats, sec_all], ignore_index=True) if not df_all_cats.empty else sec_all
    por_equipo_80 = {}
    if not df_all_cats.empty:
        for r80 in df_all_cats.to_dict("records"):
            por_equipo_80.setdefault(r80["Equipo"], []).append(r80)

    packs_bet_builder = []
    for duelo in duelos_proximos:
        loc_eq = duelo["local"]
        vis_eq = duelo["visitante"]
        if loc_eq not in equipos_scope_set and vis_eq not in equipos_scope_set:
            continue
        if filtro_busqueda and filtro_busqueda != "Todos los Equipos":
            if " vs " in filtro_busqueda and f"{loc_eq} vs {vis_eq}" != filtro_busqueda:
                continue
            elif " vs " not in filtro_busqueda and filtro_busqueda not in (loc_eq, vis_eq):
                continue

        if tier_rival and tier_rival != _COL_TIER_TODOS:
            if _resolver_tier_club(loc_eq) != tier_rival and _resolver_tier_club(vis_eq) != tier_rival:
                continue

        candidatos_cat = {}
        for eq_p, rol_p in ((loc_eq, loc_eq), (vis_eq, vis_eq)):
            for r80 in por_equipo_80.get(eq_p, []):
                if float(r80["EV"]) <= 0.0:
                    continue
                cat_c = r80["Categoria"]
                leg = {
                    "cat": cat_c,
                    "etiqueta": r80["Mercado"],
                    "origen": rol_p,
                    "wr": float(r80["Winrate"]),
                    "prob": float(r80["Prob"]),
                    "cuota_base": float(r80["CuotaBase"]),
                    "ev": float(r80["EV"]),
                }
                prev = candidatos_cat.get(cat_c)
                if prev is None or (leg["ev"], leg["wr"]) > (prev["ev"], prev["wr"]):
                    candidatos_cat[cat_c] = leg

        legs = sorted(candidatos_cat.values(), key=lambda x: (x["ev"], x["wr"]), reverse=True)[:3]
        if len(legs) < 2:
            continue

        p_prod = float(np.prod([x["prob"] for x in legs]))
        p_comp = max(min(p_prod, 0.92), 0.04)
        cuota_justa_bb = round(1.0 / p_comp, 2)
        cuota_paf_sug = round(float(np.prod([x["cuota_base"] for x in legs])) * 0.90, 2)
        wr_med = round(float(np.mean([x["wr"] for x in legs])), 1)

        packs_bet_builder.append({
            "partido": f"{loc_eq} vs {vis_eq}",
            "liga": duelo.get("liga_duelo") or liga_eq_map.get(loc_eq, "N/A"),
            "px_dt": duelo["px_dt"],
            "arbitro": duelo["arbitro"],
            "legs": legs,
            "prob_compuesta": round(p_comp * 100.0, 1),
            "cuota_teorica": cuota_justa_bb,
            "cuota_paf": cuota_paf_sug,
            "wr_medio": wr_med,
        })

    packs_bet_builder.sort(key=lambda x: (x["wr_medio"], x["prob_compuesta"]), reverse=True)
    return df_feed, choques_agrupados, packs_bet_builder


def _col_render_bet_builder_paf(packs):
    """Renderiza el bloque Glassmorphism '⚡ Bet Builder Sugerido Paf (Correlación Mismo Partido)'."""
    if not packs:
        return
    st.markdown("#### ⚡ Bet Builder Sugerido Paf (Correlación Mismo Partido · Winrate ≥ 80%)")
    cards_bb = []
    for pk in packs[:4]:
        f_ts = pd.Timestamp(pk["px_dt"]) if pd.notna(pk["px_dt"]) else None
        if f_ts is not None:
            hora_s = f" {f_ts.strftime('%H:%M')}" if (f_ts.hour or f_ts.minute) else ""
            fecha_lbl = f"🗓️ {f_ts.strftime('%d/%m')}{hora_s}"
        else:
            fecha_lbl = "🗓️ Próxima Jornada"
        legs_html = "".join(
            "<div style='background:rgba(137,87,229,0.16);border:1px solid rgba(137,87,229,0.55);"
            "border-radius:8px;padding:5px 9px;margin:4px 0;display:flex;justify-content:space-between;"
            "align-items:center;font-size:0.78rem;'>"
            f"<span style='color:#e2c5ff;font-weight:700;'>➕ {_html_ind.escape(lg['etiqueta'])} "
            f"<span style='color:#8b949e;font-weight:500;'>({_html_ind.escape(lg['origen'])})</span></span>"
            f"<span style='color:#3fb950;font-weight:800;font-family:monospace;'>{lg['wr']:.0f}% · @{lg['cuota_base']:.2f}</span>"
            "</div>"
            for lg in pk["legs"]
        )
        arb_html = pk["arbitro"]["badge_html"] if isinstance(pk.get("arbitro"), dict) else ""
        cards_bb.append(
            "<div style='background:#161b22;border:1px solid #8957e5;border-top:3px solid #d29922;"
            "border-radius:12px;padding:15px 17px;box-shadow:0 8px 24px rgba(137,87,229,0.18);'>"
            "<div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;'>"
            "<span style='background:rgba(210,153,34,0.18);border:1px solid #d29922;color:#f2cc60;"
            "font-size:0.72rem;font-weight:800;padding:2px 9px;border-radius:12px;'>⚡ Paf Bet Builder Pack</span>"
            f"<span style='color:#d2a8ff;font-size:0.76rem;font-weight:700;'>{fecha_lbl}</span>"
            "</div>"
            f"<h3 style='margin:3px 0;color:#f0f6fc;font-size:1.06rem;'>⚽ {_html_ind.escape(pk['partido'])}</h3>"
            f"<div style='color:#58a6ff;font-size:0.76rem;font-weight:600;margin-bottom:6px;'>{_html_ind.escape(str(pk['liga']))}</div>"
            f"<div style='margin-bottom:8px;'>{arb_html}</div>"
            f"<div>{legs_html}</div>"
            "<div style='display:flex;justify-content:space-between;align-items:center;border-top:1px solid #30363d;"
            "padding-top:9px;margin-top:8px;font-size:0.77rem;color:#8b949e;'>"
            f"<span>Prob. Compuesta<b style='color:#f0f6fc;display:block;font-size:0.9rem;'>{pk['prob_compuesta']:.1f}%</b></span>"
            f"<span>Cuota Teórica<b style='color:#79c0ff;display:block;font-size:0.9rem;'>{pk['cuota_teorica']:.2f}</b></span>"
            f"<span>Cuota Sugerida Paf<b style='color:#f2cc60;display:block;font-size:0.94rem;'>@{pk['cuota_paf']:.2f}</b></span>"
            "</div></div>"
        )
    st.markdown(f"<div class='ind-grid'>{''.join(cards_bb)}</div>", unsafe_allow_html=True)


def render_radar_tendencias():
    """Módulo Colectivo: Radar Cuantitativo de Tendencias & Convergencia EV+ (RFC-FIX-V3)."""
    st.markdown(
        "<style>"
        ".col-hero{background:linear-gradient(135deg,#0e1117 0%,#161b22 60%,#21262d 100%);"
        "border:1px solid #30363d;border-left:4px solid #38bdf8;border-radius:12px;"
        "padding:20px 24px;margin-bottom:18px;}"
        ".col-hero h2{margin:0;color:#38bdf8;font-weight:700;font-size:1.45rem;}"
        ".col-hero p{margin:6px 0 10px 0;color:#8b949e;font-size:0.92rem;}"
        ".ind-badge{display:inline-block;margin-right:8px;padding:4px 12px;border-radius:20px;"
        "font-size:0.78rem;font-weight:600;background:rgba(56,139,253,0.15);border:1px solid #388bfd;color:#58a6ff;}"
        ".ind-badge.g{background:rgba(63,185,80,0.12);border-color:#238636;color:#3fb950;}"
        ".ind-badge.off{background:#161b22;border-color:#30363d;color:#8b949e;opacity:0.85;}"
        ".ind-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px;margin-bottom:22px;}"
        ".ind-card{background:#0d1117;border:1px solid #30363d;border-top:3px solid #3fb950;"
        "border-radius:12px;padding:14px 16px;box-shadow:0 8px 22px rgba(0,0,0,0.35);}"
        ".ind-card .top{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;}"
        ".ind-tag{font-size:0.72rem;font-weight:700;padding:3px 9px;border-radius:12px;"
        "background:rgba(88,166,255,0.14);border:1px solid #1f6feb;color:#58a6ff;white-space:nowrap;}"
        ".ind-ev{font-size:0.78rem;font-weight:800;padding:3px 10px;border-radius:12px;"
        "background:rgba(63,185,80,0.16);border:1px solid #3fb950;color:#3fb950;}"
        ".ind-ev.neg{background:rgba(218,54,51,0.14);border-color:#da3633;color:#ff7b72;}"
        ".ind-card h3{margin:2px 0 2px 0;color:#f0f6fc;font-size:1.12rem;}"
        ".ind-liga{color:#58a6ff;font-size:0.78rem;font-weight:600;}"
        ".ind-next{margin:8px 0;padding:6px 10px;background:#161b22;border:1px solid #21262d;"
        "border-radius:8px;color:#c9d1d9;font-size:0.8rem;}"
        ".ind-mkt{color:#79c0ff;font-weight:700;font-size:0.9rem;margin:4px 0 8px 0;}"
        ".ind-met{display:flex;justify-content:space-between;gap:6px;font-size:0.78rem;color:#8b949e;"
        "border-top:1px solid #21262d;padding-top:8px;margin-top:6px;}"
        ".ind-met b{color:#f0f6fc;display:block;font-size:0.9rem;}"
        ".ind-wrap{background:#0d1117;border:1px solid #30363d;border-radius:12px;overflow-x:auto;padding:4px;}"
        ".ind-tbl{width:100%;border-collapse:collapse;font-size:0.8rem;color:#c9d1d9;min-width:1100px;}"
        ".ind-tbl th{background:#161b22;color:#8b949e;text-align:left;padding:10px 10px;font-size:0.7rem;"
        "letter-spacing:0.05em;border-bottom:1px solid #30363d;white-space:nowrap;}"
        ".ind-tbl td{padding:9px 10px;border-bottom:1px solid #21262d;vertical-align:middle;}"
        ".ind-tbl tr:hover td{background:#161b22;}"
        ".ind-evb{background:rgba(63,185,80,0.16);border:1px solid #3fb950;color:#3fb950;"
        "padding:2px 9px;border-radius:10px;font-weight:800;white-space:nowrap;}"
        ".ind-evb.neg{background:rgba(218,54,51,0.14);border-color:#da3633;color:#ff7b72;}"
        ".col-tier{display:inline-block;padding:2px 8px;border-radius:8px;font-size:0.71rem;font-weight:600;"
        "background:rgba(56,139,253,0.12);border:1px solid #30363d;color:#79c0ff;margin-left:6px;}"
        "</style>"
        "<div class='col-hero'>"
        "<h2>🎯 Radar Cuantitativo de Tendencias &amp; Convergencia EV+ (Colectivo)</h2>"
        "<p>Escáner institucional de totales de partido (Goles, BTTS, Córners y Tarjetas) calibrado con Poisson real por equipo, "
        "filtrado en cascada por competición, cruce arbitral sobre el censo de 165 colegiados y Bet Builder Paf.</p>"
        "<span class='ind-badge'>Kambi / Paf Bet Builder Ready</span>"
        "<span class='ind-badge g'>Poisson Real &amp; Referee Engine Activo</span>"
        "</div>",
        unsafe_allow_html=True,
    )

    df_raw = _cargar_dataset_props_vectorial()
    if df_raw.empty:
        st.warning("No se hallaron registros históricos en `macro_base_futbol.db`.")
        return

    # ── FILA 1: Contexto, Búsqueda y Tier de Rival (Cascade Filtering) ───────
    f1c1, f1c2, f1c3, f1c4 = st.columns([1.15, 1.25, 1.2, 1.6])
    with f1c1:
        comp_sel = st.selectbox(
            "🌍 Competición",
            list(_IND_FILTRO_COMPETICION.keys()),
            key="col_radar_comp",
        )

    df_filtrado, divs_activas, es_uefa = _col_filtrar_dataframe_competicion(df_raw, comp_sel)

    with f1c2:
        cond_sel = st.pills(
            "🏠 Condición de Campo",
            ["🌐 Global", "🏠 Local/Visitante Estricto"],
            default="🌐 Global",
            key="col_radar_cond",
        ) or "🌐 Global"

    with f1c3:
        equipos_disp = (
            sorted(pd.unique(pd.concat([df_filtrado["HomeTeam"], df_filtrado["AwayTeam"]]).dropna().astype(str)).tolist())
            if not df_filtrado.empty
            else []
        )
        px_sub = _col_proximos_cruces_competicion(df_filtrado, divs_activas=divs_activas, es_uefa=es_uefa)
        duelos_opts = []
        if not px_sub.empty:
            loc_px = px_sub[px_sub["px_cond"] == "Local"]
            for eh, ea in zip(loc_px["Equipo"].values, loc_px["px_rival"].values):
                duelos_opts.append(f"{eh} vs {ea}")
        opciones_busq = ["Todos los Equipos"] + duelos_opts + equipos_disp
        busq_sel = st.selectbox(
            "🔍 Partido / Equipo",
            opciones_busq,
            key=f"col_radar_equipo_{comp_sel}",
        )

    with f1c4:
        tier_sel = st.pills(
            "🎖️ Nivel del Rival (Tier)",
            [_COL_TIER_TODOS, _COL_TIER_TOP, _COL_TIER_MEDIA, _COL_TIER_BAJA],
            default=_COL_TIER_TODOS,
            key="col_radar_tier",
        ) or _COL_TIER_TODOS

    # ── FILA 2: Categorías Colectivas y Escalera Reactiva de Líneas ──────────
    cat_sel = st.pills(
        "📊 Categoría de Mercado",
        [_COL_CAT_TODOS, _COL_CAT_GOLES, _COL_CAT_BTTS, _COL_CAT_CORNERS, _COL_CAT_TARJETAS],
        default=_COL_CAT_TODOS,
        key="col_radar_cat",
    ) or _COL_CAT_TODOS

    direccion_sel = "Todos"
    linea_sel = "Todas"
    if cat_sel == _COL_CAT_TODOS:
        st.markdown(
            "<span class='ind-badge off'>📡 Barrido Multimercado Colectivo Activo</span>",
            unsafe_allow_html=True,
        )
    elif cat_sel == _COL_CAT_BTTS:
        linea_sel = st.pills(
            "Línea Ambos Marcan (BTTS)",
            _COL_ESCALERAS_PILLS[_COL_CAT_BTTS],
            default="Todos",
            key="col_radar_lin_btts",
        ) or "Todos"
    else:
        slug_map = {
            _COL_CAT_CORNERS: "corners",
            _COL_CAT_TARJETAS: "tarjetas",
            _COL_CAT_GOLES: "goles",
        }
        slug = slug_map.get(cat_sel, "gen")
        f2d1, f2d2 = st.columns([1.1, 2.5])
        with f2d1:
            dir_lbl = st.segmented_control(
                "Dirección de Línea",
                ["Todos", "Overs (+)", "Unders (-)"],
                default="Todos",
                key=f"col_radar_dir_{slug}",
            ) or "Todos"
            direccion_sel = {"Todos": "Todos", "Overs (+)": "over", "Unders (-)": "under"}[dir_lbl]
        with f2d2:
            opciones_lin = _COL_ESCALERAS_PILLS[cat_sel]
            if direccion_sel == "over":
                opciones_lin = ["Todas"] + [x for x in opciones_lin if x.startswith("+")]
            elif direccion_sel == "under":
                opciones_lin = ["Todas"] + [x for x in opciones_lin if x.startswith("-")]
            linea_sel = st.pills(
                "Escalera de Líneas",
                opciones_lin,
                default="Todas",
                key=f"col_radar_lin_{slug}_{direccion_sel}",
            ) or "Todas"

    # ── FILA 3: Afinación Cuantitativa ───────────────────────────────────────
    f3c1, f3c2 = st.columns([1.4, 1.6])
    with f3c1:
        umbral_sel = st.slider(
            "🔥 Racha Mínima (%)",
            50,
            100,
            75,
            step=5,
            key="col_radar_umbral",
        )
    with f3c2:
        ventana_sel = st.pills(
            "📆 Muestra de Partidos",
            [3, 5, 8, 10, 15],
            default=5,
            format_func=lambda x: f"{x} partidos",
            key="col_radar_ventana",
        ) or 5

    df_feed, choques_agrupados, packs_bb = _escanear_radar_colectivo_vectorial(
        df_filtrado,
        df_raw,
        ventana=ventana_sel,
        min_winrate=float(umbral_sel),
        categoria=cat_sel,
        direccion=direccion_sel,
        codigo_linea=linea_sel,
        modo_condicion=cond_sel,
        divs_activas=divs_activas,
        es_uefa=es_uefa,
        filtro_busqueda=busq_sel,
        tier_rival=tier_sel,
    )

    tab_feed, tab_choque = st.tabs([
        "🌟 Feed Rápido (Partidos)",
        "⚔️ Choque de Tendencias (Partidos)",
    ])

    # ── PESTAÑA 1: FEED RÁPIDO (PARTIDOS) ────────────────────────────────────
    with tab_feed:
        if df_feed.empty:
            st.info(f"Ningún equipo o partido de '{comp_sel}' supera el {umbral_sel}% de racha bajo los filtros seleccionados.")
        else:
            reg_feed = df_feed.to_dict("records")
            # Seleccionar Top 4 priorizando oportunidades EV+ de equipos distintos para máxima diversidad analítica
            top4_sel = []
            eqs_vistos = set()
            for r in reg_feed:
                if r["Equipo"] not in eqs_vistos:
                    top4_sel.append(r)
                    eqs_vistos.add(r["Equipo"])
                if len(top4_sel) == 4:
                    break
            if len(top4_sel) < 4:
                for r in reg_feed:
                    if r not in top4_sel:
                        top4_sel.append(r)
                    if len(top4_sel) == 4:
                        break

            st.markdown("#### 💎 Top 4 Oportunidades Colectivas de Mayor Valor Esperado (EV+)")
            cards_top4 = []
            for r in top4_sel:
                cruce_txt = _ind_cruce_txt(r["px_rival"], r["px_cond"], r["px_dt"])
                tier_badge = f"<span class='col-tier'>{_html_ind.escape(str(r.get('TierRef', '')))}</span>"
                arb_badge = r["ArbitroInfo"]["badge_html"] if isinstance(r.get("ArbitroInfo"), dict) else ""
                ev_cls_card = "ind-ev" if r["EV"] >= 0 else "ind-ev neg"
                cards_top4.append(
                    "<div class='ind-card'>"
                    "<div class='top'><span class='ind-tag'>⚔️ Total Colectivo</span>"
                    f"<span class='{ev_cls_card}'>EV {r['EV']:+.1f}%</span></div>"
                    f"<h3>{_html_ind.escape(str(r['Equipo']))}{tier_badge}</h3>"
                    f"<div class='ind-liga'>{_html_ind.escape(str(r['Liga']))} · Media: {r['Lambda']:.2f}</div>"
                    f"<div class='ind-next'>{cruce_txt}</div>"
                    f"<div style='margin:6px 0;'>{arb_badge}</div>"
                    f"<div class='ind-mkt'>{_html_ind.escape(r['Mercado'])}</div>"
                    "<div class='ind-met'>"
                    f"<span>Racha<b>{r['Hits']}/{r['N']} [{r['Winrate']:.0f}%]</b></span>"
                    f"<span>Cuota Justa<b>{r['CuotaJusta']:.2f}</b></span>"
                    f"<span>Cuota Base<b>{r['CuotaBase']:.2f}</b></span></div>"
                    f"<div style='margin-top:8px;'>{_ind_chips_html(r['Chips'])}</div>"
                    "</div>"
                )
            st.markdown(f"<div class='ind-grid'>{''.join(cards_top4)}</div>", unsafe_allow_html=True)

            # Bloque Bet Builder Sugerido Paf antes de la tabla maestra
            _col_render_bet_builder_paf(packs_bb)

            st.markdown(f"#### 📋 Tabla Maestra de Oportunidades Colectivas (Winrate ≥ {umbral_sel}%)")
            filas_html = []
            for r in reg_feed[:200]:
                ev_cls = "ind-evb" if r["EV"] >= 0 else "ind-evb neg"
                cruce_txt = _ind_cruce_txt(r["px_rival"], r["px_cond"], r["px_dt"])
                arb_b = r["ArbitroInfo"]["badge_html"] if isinstance(r.get("ArbitroInfo"), dict) else ""
                tier_lbl = _html_ind.escape(str(r.get("TierRef", "")))
                filas_html.append(
                    "<tr>"
                    "<td><span class='ind-tag'>⚔️ Partido</span></td>"
                    f"<td><b style='color:#f0f6fc;'>{_html_ind.escape(str(r['Equipo']))}</b>"
                    f"<div class='ind-liga'>{_html_ind.escape(str(r['Liga']))} · {tier_lbl}</div></td>"
                    f"<td><div>{cruce_txt}</div><div style='margin-top:4px;'>{arb_b}</div></td>"
                    f"<td style='color:#79c0ff;font-weight:600;'>{_html_ind.escape(r['Mercado'])}</td>"
                    f"<td><b>{r['Hits']}/{r['N']}</b> <span style='color:#8b949e;'>[{r['Winrate']:.0f}%]</span></td>"
                    f"<td>{_ind_chips_html(r['Chips'])}</td>"
                    f"<td><b>{r['CuotaJusta']:.2f}</b></td>"
                    f"<td>{r['CuotaBase']:.2f}</td>"
                    f"<td><span class='{ev_cls}'>{r['EV']:+.1f}%</span></td>"
                    "</tr>"
                )
            cols_th = "".join(
                f"<th>{h}</th>"
                for h in [
                    "TIPO", "EQUIPO / LIGA / TIER", "PRÓXIMO CRUCE & ÁRBITRO",
                    "MERCADO / LÍNEA", "RACHA / %", "DESGLOSE (CHIPS)",
                    "CUOTA JUSTA", "CUOTA BASE", "DIFERENCIAL EV%",
                ]
            )
            st.markdown(
                f"<div class='ind-wrap'><table class='ind-tbl'><thead><tr>{cols_th}</tr></thead>"
                f"<tbody>{''.join(filas_html)}</tbody></table></div>",
                unsafe_allow_html=True,
            )
            st.caption(
                f"{len(reg_feed)} oportunidades colectivas detectadas en '{comp_sel}' · "
                f"mostrando las {min(len(reg_feed), 200)} de mayor EV%."
            )

    # ── PESTAÑA 2: CHOQUE DE TENDENCIAS (PARTIDOS AGRUPADOS POR FIXTURE_ID) ──
    with tab_choque:
        _col_render_bet_builder_paf(packs_bb)
        if not choques_agrupados:
            st.info(
                f"ℹ️ No se detectaron choques de doble coincidencia (Local + Visitante) en '{comp_sel}' con los filtros actuales. "
                "Prueba reduciendo la Racha Mínima (%) o ampliando la categoría."
            )
        else:
            st.markdown(
                f"#### 💥 Choque de Tendencias: {len(choques_agrupados)} Encuentros con Doble Coincidencia Activa"
            )
            cards_choque = []
            for grp in choques_agrupados[:12]:
                f_ts = pd.Timestamp(grp["px_dt"]) if pd.notna(grp["px_dt"]) else None
                if f_ts is not None:
                    hora_s = f" {f_ts.strftime('%H:%M')}" if (f_ts.hour or f_ts.minute) else ""
                    fecha_s = f"🗓️ {f_ts.strftime('%d/%m')}{hora_s}"
                else:
                    fecha_s = "🗓️ Próximo Duelo"
                arb_html = grp["arbitro"]["badge_html"] if isinstance(grp.get("arbitro"), dict) else ""
                ev_top_cls = "ind-ev" if grp["max_ev"] >= 0 else "ind-ev neg"

                bloques_conv_html = []
                for cv in grp["convergencias"][:4]:
                    rl = cv["racha_local"]
                    rv = cv["racha_visit"]
                    ev_c_cls = "ind-evb" if cv["ev_comb"] >= 0 else "ind-evb neg"
                    bloques_conv_html.append(
                        "<div style='background:#161b22;border:1px solid #30363d;border-left:3px solid #58a6ff;"
                        "border-radius:8px;padding:8px 10px;margin:8px 0;'>"
                        "<div style='display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:4px;margin-bottom:5px;'>"
                        f"<span style='color:#79c0ff;font-weight:700;font-size:0.83rem;'>{_html_ind.escape(cv['mercado'])}</span>"
                        f"<span style='font-size:0.75rem;color:#c9d1d9;'>"
                        f"Conv: <b style='color:#3fb950;'>{cv['wr_comb']:.0f}%</b> · "
                        f"Justa: <b style='color:#f0f6fc;'>{cv['cuota_justa']:.2f}</b> · "
                        f"Base: <b>{cv['cuota_base']:.2f}</b> · "
                        f"<span class='{ev_c_cls}'>EV {cv['ev_comb']:+.1f}%</span></span>"
                        "</div>"
                        f"<div style='font-size:0.74rem;color:#c9d1d9;margin:3px 0;'>"
                        f"🏠 <b>{_html_ind.escape(str(rl['Equipo']))}</b> ({rl['Hits']}/{rl['N']} · λ={rl['Lambda']:.2f}): "
                        f"{_ind_chips_html(rl['Chips'])}</div>"
                        f"<div style='font-size:0.74rem;color:#c9d1d9;margin:3px 0;'>"
                        f"✈️ <b>{_html_ind.escape(str(rv['Equipo']))}</b> ({rv['Hits']}/{rv['N']} · λ={rv['Lambda']:.2f}): "
                        f"{_ind_chips_html(rv['Chips'])}</div>"
                        "</div>"
                    )

                cards_choque.append(
                    "<div class='ind-card'>"
                    "<div class='top'>"
                    f"<span class='ind-tag'>⚔️ {_html_ind.escape(str(grp['liga']))} · {grp['n_conv']} Convergencia(s)</span>"
                    f"<span class='{ev_top_cls}'>Máx EV {grp['max_ev']:+.1f}%</span>"
                    "</div>"
                    f"<h3>{_html_ind.escape(grp['partido'])} "
                    f"<span style='font-size:0.78rem;color:#8b949e;font-weight:500;'>({fecha_s})</span></h3>"
                    f"<div style='font-size:0.74rem;color:#8b949e;margin:3px 0;'>"
                    f"🏠 {_html_ind.escape(str(grp['tier_local']))} vs ✈️ {_html_ind.escape(str(grp['tier_visit']))}</div>"
                    f"<div style='margin:6px 0;'>{arb_html}</div>"
                    f"{''.join(bloques_conv_html)}"
                    "</div>"
                )
            st.markdown(f"<div class='ind-grid'>{''.join(cards_choque)}</div>", unsafe_allow_html=True)


# ==============================================================================
# NAVEGACIÓN Y SECCIONES
# ==============================================================================
render_top_navigation_bar()

if menu_seleccionado == "🏠 Inicio / Portada":
    # ── 1. PALETA DE COLORES INSTITUCIONAL Y ESTILOS CSS INYECTADOS ──────────
    st.markdown(
        """
        <style>
        .quant-hero-box {
            background: linear-gradient(135deg, #161B26 0%, #1E232F 58%, #131822 100%);
            border: 1px solid #2A3245;
            border-left: 4px solid #00E676;
            border-radius: 14px;
            padding: 26px 30px;
            margin-bottom: 22px;
            box-shadow: 0 10px 28px rgba(0, 0, 0, 0.42);
            position: relative;
            overflow: hidden;
        }
        .quant-status-pill {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            background: rgba(0, 230, 118, 0.12);
            color: #00E676;
            border: 1px solid rgba(0, 230, 118, 0.38);
            border-radius: 999px;
            padding: 5px 14px;
            font-size: 0.74rem;
            font-weight: 800;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            margin-bottom: 12px;
            box-shadow: 0 0 14px rgba(0, 230, 118, 0.16);
        }
        .quant-hero-title {
            font-size: 2.05rem;
            font-weight: 900;
            color: #F8FAFC;
            letter-spacing: -0.015em;
            line-height: 1.18;
            margin: 0 0 10px 0;
        }
        .quant-hero-title span {
            color: #00E5FF;
        }
        .quant-hero-sub {
            font-size: 0.96rem;
            color: #94A3B8;
            line-height: 1.55;
            max-width: 920px;
            margin: 0;
        }
        .quant-kpi-card {
            background: linear-gradient(180deg, #1E232F 0%, #161B26 100%);
            border: 1px solid #2A3245;
            border-radius: 12px;
            padding: 18px 20px;
            min-height: 148px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease;
        }
        .quant-kpi-card:hover {
            transform: translateY(-2px);
            border-color: #00E5FF;
            box-shadow: 0 8px 22px rgba(0, 229, 255, 0.12);
        }
        .quant-kpi-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 6px;
        }
        .quant-kpi-label {
            font-size: 0.75rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            color: #94A3B8;
        }
        .quant-kpi-value {
            font-size: 1.85rem;
            font-weight: 900;
            color: #F8FAFC;
            line-height: 1.15;
            margin: 2px 0 8px 0;
        }
        .quant-kpi-value .accent-green { color: #00E676; }
        .quant-kpi-value .accent-cyan { color: #00E5FF; }
        .quant-tag-row {
            display: flex;
            flex-wrap: wrap;
            gap: 5px;
            margin-top: 4px;
        }
        .quant-micro-tag {
            background: rgba(42, 50, 69, 0.65);
            color: #CBD5E1;
            border: 1px solid #334155;
            border-radius: 5px;
            padding: 2px 7px;
            font-size: 0.68rem;
            font-weight: 600;
            letter-spacing: 0.02em;
        }
        .quant-micro-tag.cyan {
            background: rgba(0, 229, 255, 0.10);
            color: #38BDF8;
            border-color: rgba(56, 189, 248, 0.32);
        }
        .quant-micro-tag.green {
            background: rgba(16, 185, 129, 0.12);
            color: #00E676;
            border-color: rgba(16, 185, 129, 0.35);
        }
        .quant-micro-tag.amber {
            background: rgba(245, 158, 11, 0.12);
            color: #F59E0B;
            border-color: rgba(245, 158, 11, 0.35);
        }
        .quant-section-hdr {
            display: flex;
            justify-content: space-between;
            align-items: baseline;
            margin: 26px 0 14px 0;
            border-bottom: 1px solid #2A3245;
            padding-bottom: 8px;
        }
        .quant-section-title {
            font-size: 1.22rem;
            font-weight: 800;
            color: #F8FAFC;
            letter-spacing: 0.01em;
        }
        .quant-section-caption {
            font-size: 0.80rem;
            color: #64748B;
            font-weight: 600;
        }
        .quant-mod-card {
            background: linear-gradient(180deg, #1E232F 0%, #161B26 100%);
            border: 1px solid #2A3245;
            border-radius: 12px;
            padding: 20px 20px 14px 20px;
            min-height: 190px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            margin-bottom: 8px;
            transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease;
        }
        .quant-mod-card:hover {
            transform: translateY(-2px);
            border-color: #10B981;
            box-shadow: 0 10px 26px rgba(16, 185, 129, 0.14);
        }
        .quant-mod-hdr {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 10px;
            margin-bottom: 10px;
        }
        .quant-mod-title {
            font-size: 1.06rem;
            font-weight: 800;
            color: #F8FAFC;
            line-height: 1.28;
        }
        .quant-mod-badge {
            border-radius: 6px;
            padding: 3px 9px;
            font-size: 0.69rem;
            font-weight: 800;
            white-space: nowrap;
            letter-spacing: 0.03em;
        }
        .quant-mod-desc {
            font-size: 0.84rem;
            color: #94A3B8;
            line-height: 1.48;
            margin-bottom: 10px;
        }
        .quant-step-box {
            background: #161B26;
            border: 1px solid #2A3245;
            border-top: 3px solid #00E5FF;
            border-radius: 10px;
            padding: 16px 18px;
            height: 100%;
            transition: transform 0.18s ease, border-color 0.18s ease;
        }
        .quant-step-box:hover {
            transform: translateY(-2px);
            border-color: #00E676;
        }
        .quant-step-num {
            display: inline-block;
            background: rgba(0, 229, 255, 0.14);
            color: #00E5FF;
            border: 1px solid rgba(0, 229, 255, 0.35);
            border-radius: 6px;
            padding: 2px 8px;
            font-size: 0.72rem;
            font-weight: 800;
            margin-bottom: 8px;
        }
        .quant-step-title {
            font-size: 0.96rem;
            font-weight: 800;
            color: #F8FAFC;
            margin-bottom: 6px;
        }
        .quant-step-txt {
            font-size: 0.82rem;
            color: #94A3B8;
            line-height: 1.45;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # ── CONSULTAS EN VIVO A SQLITE PARA LAS 4 TARJETAS KPI ───────────────────
    _kpi_partidos_auditados = len(df_partidos_global) if isinstance(df_partidos_global, pd.DataFrame) else 0
    _kpi_futbolistas_censo = 0
    _kpi_futbolistas_unicos = 0
    if RUTA_BD.exists():
        try:
            with sqlite3.connect(RUTA_BD) as _c_home:
                _tablas_home = {
                    r[0] for r in _c_home.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
                }
                if "partidos" in _tablas_home:
                    _n_p_tbl = int(_c_home.execute("SELECT COUNT(*) FROM partidos").fetchone()[0] or 0)
                    _kpi_partidos_auditados = max(_kpi_partidos_auditados, _n_p_tbl)
                if "historico_reciente" in _tablas_home:
                    _n_hr_tbl = int(_c_home.execute("SELECT COUNT(*) FROM historico_reciente").fetchone()[0] or 0)
                    _kpi_partidos_auditados = max(_kpi_partidos_auditados, _n_hr_tbl)
                if "estadisticas_jugadores" in _tablas_home:
                    _kpi_futbolistas_censo = int(
                        _c_home.execute("SELECT COUNT(*) FROM estadisticas_jugadores").fetchone()[0] or 0
                    )
                    _kpi_futbolistas_unicos = int(
                        _c_home.execute("SELECT COUNT(DISTINCT Jugador) FROM estadisticas_jugadores").fetchone()[0] or 0
                    )
        except Exception:
            pass
    if _kpi_futbolistas_censo == 0 and isinstance(df_jugadores, pd.DataFrame):
        _kpi_futbolistas_censo = len(df_jugadores)
        _kpi_futbolistas_unicos = len(df_jugadores)

    # ── 2. HERO SECTION & HEADER STATUS ──────────────────────────────────────
    st.markdown(
        """
        <div class="quant-hero-box">
            <div class="quant-status-pill">🟢 MOTOR CUANTITATIVO ACTIVO &nbsp;|&nbsp; v3.0 MULTI-COMPETICIÓN</div>
            <div class="quant-hero-title">⚽ HUGO'S HUB &nbsp;|&nbsp; <span>QUANT SPORTS INTELLIGENCE</span></div>
            <p class="quant-hero-sub">
                Plataforma analítica avanzada para detección de valor esperado (<b>EV+</b>), modelado predictivo
                <b>Poisson con suavizado bayesiano</b> y <b>scouting métrico</b> en 8 competiciones de élite.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── 3. KPI METRIC STRIP (4 TARJETAS DINÁMICAS HORIZONTALES) ──────────────
    _fmt_partidos = f"{_kpi_partidos_auditados:,}".replace(",", ".")
    _fmt_censo = f"{_kpi_futbolistas_censo:,}".replace(",", ".")
    _fmt_unicos = f"{_kpi_futbolistas_unicos:,}".replace(",", ".")

    hk1, hk2, hk3, hk4 = st.columns(4)
    with hk1:
        st.markdown(
            f"""
            <div class="quant-kpi-card">
                <div>
                    <div class="quant-kpi-top">
                        <span class="quant-kpi-label">🗄️ Partidos Auditados</span>
                        <span class="quant-micro-tag green">SQLite Sync</span>
                    </div>
                    <div class="quant-kpi-value"><span class="accent-cyan">{_fmt_partidos}+</span></div>
                </div>
                <div class="quant-tag-row">
                    <span class="quant-micro-tag">Histórico Oficial</span>
                    <span class="quant-micro-tag">Cuotas 1X2 / O-U</span>
                    <span class="quant-micro-tag">ESPN Live</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with hk2:
        st.markdown(
            f"""
            <div class="quant-kpi-card">
                <div>
                    <div class="quant-kpi-top">
                        <span class="quant-kpi-label">🧬 Futbolistas Indexados</span>
                        <span class="quant-micro-tag cyan">FBref + ESPN</span>
                    </div>
                    <div class="quant-kpi-value"><span class="accent-green">{_fmt_censo}</span></div>
                </div>
                <div class="quant-tag-row">
                    <span class="quant-micro-tag green">{_fmt_unicos} Únicos</span>
                    <span class="quant-micro-tag">Minutos Reales</span>
                    <span class="quant-micro-tag">Tiros &amp; Faltas</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with hk3:
        st.markdown(
            """
            <div class="quant-kpi-card">
                <div>
                    <div class="quant-kpi-top">
                        <span class="quant-kpi-label">🏆 Competiciones Conectadas</span>
                        <span class="quant-micro-tag amber">Top 5 + UEFA</span>
                    </div>
                    <div class="quant-kpi-value">8 Ligas &amp; Torneos</div>
                </div>
                <div class="quant-tag-row">
                    <span class="quant-micro-tag cyan">LaLiga</span>
                    <span class="quant-micro-tag cyan">Premier</span>
                    <span class="quant-micro-tag cyan">Serie A</span>
                    <span class="quant-micro-tag cyan">Bundesliga</span>
                    <span class="quant-micro-tag cyan">Ligue 1</span>
                    <span class="quant-micro-tag amber">UCL</span>
                    <span class="quant-micro-tag amber">UEL</span>
                    <span class="quant-micro-tag amber">UECL</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with hk4:
        st.markdown(
            """
            <div class="quant-kpi-card">
                <div>
                    <div class="quant-kpi-top">
                        <span class="quant-kpi-label">📐 Cobertura de Mercados</span>
                        <span class="quant-micro-tag green">Overround 0%</span>
                    </div>
                    <div class="quant-kpi-value"><span class="accent-green">100%</span> Cobertura</div>
                </div>
                <div class="quant-tag-row">
                    <span class="quant-micro-tag">1X2 Poisson</span>
                    <span class="quant-micro-tag">xG Neto</span>
                    <span class="quant-micro-tag">Over/Under</span>
                    <span class="quant-micro-tag">BTTS</span>
                    <span class="quant-micro-tag">Hándicap Asiático</span>
                    <span class="quant-micro-tag">Props P90</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # ── 4. INTERACTIVE MODULE DIRECTORY (GRID DE NAVEGACIÓN 3x2) ─────────────
    st.markdown(
        """
        <div class="quant-section-hdr">
            <span class="quant-section-title">🧭 Directorio Interactivo de Módulos Cuantitativos</span>
            <span class="quant-section-caption">Selecciona un motor analítico para abrirlo directamente</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    def _ir_a_modulo_hub(destino_menu: str) -> None:
        _registrar_navegacion(destino_menu)

    _catalogo_modulos_home = [
        {
            "id": "h2h",
            "titulo": "⚔️ Comparador H2H Pro & Cuotas Justas",
            "badge": "🟢 Calibrado",
            "badge_style": "background:rgba(0,230,118,0.13);color:#00E676;border:1px solid rgba(0,230,118,0.38);",
            "desc": (
                "Estimación Poisson bivariante, suavizado bayesiano y balanceo de muestra oficial para torneos UEFA, "
                "línea justa de Hándicap Asiático y calculadora interactiva Quarter-Kelly."
            ),
            "destino": "⚔️ Comparador H2H (Cara a Cara)",
            "btn": "Abrir Comparador H2H ↗",
        },
        {
            "id": "scouting",
            "titulo": "📋 Scouting de Jugadores Pro",
            "badge": "⚡ Multi-Competición",
            "badge_style": "background:rgba(0,229,255,0.13);color:#00E5FF;border:1px solid rgba(0,229,255,0.38);",
            "desc": (
                "Filtro posicional estricto (POR, DFC, LI, LD, MCD, MC, MCO, EI, ED, DC), métricas P90 de tiros, "
                "tiros a puerta, faltas y tarjetas, radar percentil H2H y enlace directo a SofaScore."
            ),
            "destino": "📋 Scouting de Jugadores",
            "btn": "Abrir Scouting Pro ↗",
        },
        {
            "id": "dossier",
            "titulo": "📊 Dossier Cuantitativo de Equipo",
            "badge": "🟢 Liga + UEFA",
            "badge_style": "background:rgba(0,230,118,0.13);color:#00E676;border:1px solid rgba(0,230,118,0.38);",
            "desc": (
                "Radiografía institucional por competición (Liga Doméstica, Champions, Europa y Conference League), "
                "eficiencia G-xG, ratings Casa/Fuera, matriz Over/Under y exportación ejecutiva en PDF."
            ),
            "destino": "📊 Dossier de Equipo",
            "btn": "Abrir Dossier de Equipo ↗",
        },
        {
            "id": "escaner_ev",
            "titulo": "🎯 Escáner Multimercado EV+",
            "badge": "🎯 Arbitraje EV+",
            "badge_style": "background:rgba(245,158,11,0.15);color:#F59E0B;border:1px solid rgba(245,158,11,0.40);",
            "desc": (
                "Rastreador en tiempo real de ineficiencias entre probabilidad matemática propia y cuotas comerciales "
                "de casas de apuestas en mercados 1X2, Over/Under 2.5 y BTTS con exportación a Excel."
            ),
            "destino": "📡 Escáner Vivo EV+ (API)",
            "btn": "Abrir Escáner EV+ ↗",
        },
        {
            "id": "arbitros",
            "titulo": "⚖️ Matchup de Árbitros & Tarjetas",
            "badge": "🟨 Severidad Arbitral",
            "badge_style": "background:rgba(245,158,11,0.15);color:#F59E0B;border:1px solid rgba(245,158,11,0.40);",
            "desc": (
                "Auditoría de colegiados por liga, medias históricas de amonestaciones y expulsiones, y cruce "
                "de fricción disciplinaria entre los dos clubes del encuentro."
            ),
            "destino": "⚖️ Directorio y Matchup Árbitros",
            "btn": "Abrir Matchup Árbitros ↗",
        },
        {
            "id": "handicap_radar",
            "titulo": "🧮 Hándicap Asiático & Radar de Tendencias",
            "badge": "⚡ Cuotas Implícitas",
            "badge_style": "background:rgba(0,229,255,0.13);color:#00E5FF;border:1px solid rgba(0,229,255,0.38);",
            "desc": (
                "Simulador matricial de líneas asiáticas (-1.5 a +1.5) con liquidación exacta de medios cuartos "
                "y detección de rachas estadísticas de alta frecuencia."
            ),
            "destino": "🧮 Hándicap Asiático Avanzado",
            "btn": "Abrir Hándicap Asiático ↗",
        },
    ]

    for _fila_idx in range(0, len(_catalogo_modulos_home), 3):
        _cols_mod = st.columns(3)
        for _col_i, _mod in enumerate(_catalogo_modulos_home[_fila_idx : _fila_idx + 3]):
            with _cols_mod[_col_i]:
                st.markdown(
                    f"""
                    <div class="quant-mod-card">
                        <div>
                            <div class="quant-mod-hdr">
                                <div class="quant-mod-title">{_mod['titulo']}</div>
                                <span class="quant-mod-badge" style="{_mod['badge_style']}">{_mod['badge']}</span>
                            </div>
                            <div class="quant-mod-desc">{_mod['desc']}</div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.button(
                    _mod["btn"],
                    key=f"btn_home_nav_{_mod['id']}",
                    width="stretch",
                    type="primary" if _mod["id"] in ("h2h", "scouting", "dossier") else "secondary",
                    on_click=_ir_a_modulo_hub,
                    args=(_mod["destino"],),
                )

    # ── 5. ARQUITECTURA DE TRABAJO (QUICKSTART PROTOCOL EN 3 PASOS) ──────────
    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
    with st.expander("🧭 Protocolo Institucional de Toma de Decisiones en 3 Pasos (Quickstart Protocol)", expanded=True):
        qs1, qs2, qs3 = st.columns(3)
        with qs1:
            st.markdown(
                """
                <div class="quant-step-box">
                    <span class="quant-step-num">PASO 01 · EXPLORACIÓN</span>
                    <div class="quant-step-title">🔍 Radiografía de Competición y Plantilla</div>
                    <div class="quant-step-txt">
                        Selecciona el torneo objetivo (Liga Doméstica o UEFA) y audita el <b>Dossier de Equipo</b>
                        y el <b>Scouting Pro</b> para identificar volumen de tiro P90, bajas, fricción de faltas y estado de forma.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with qs2:
            st.markdown(
                """
                <div class="quant-step-box">
                    <span class="quant-step-num">PASO 02 · PROYECCIÓN</span>
                    <div class="quant-step-title">⚖️ Modelado H2H con Muestra Balanceada</div>
                    <div class="quant-step-txt">
                        Cruza a ambos clubes en el <b>Comparador H2H Pro</b> activando la ventana homogénea de partidos
                        oficiales y el factor de ventaja de campo para obtener <b>xG proyectados</b>, línea asiática justa y probabilidades reales.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with qs3:
            st.markdown(
                """
                <div class="quant-step-box">
                    <span class="quant-step-num">PASO 03 · ARBITRAJE EV+</span>
                    <div class="quant-step-title">💎 Ejecución Disciplinada &amp; Quarter-Kelly</div>
                    <div class="quant-step-txt">
                        Contrasta la <b>Cuota Justa (1/p)</b> del modelo frente a la cuota comercial de tu casa de apuestas
                        y ejecuta únicamente posiciones con <b>Valor Esperado positivo (EV &gt; +2%)</b> dimensionando el stake con Quarter-Kelly.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

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
        """Invoca el motor jerárquico global de fricción arbitral pasando el dataframe del escáner."""
        return globals()["_auditar_friccion_arbitral_sqlite"](
            aud_h=aud_h,
            aud_a=aud_a,
            liga_nombre=liga_nombre,
            arbitro_override=arbitro_override,
            df_partidos_global=_obtener_df_partidos_scanner(),
        )

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
                    "Backtest_ROI": round(float(bt_info.get("roi", 0.0) or 0.0), 2),
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
        btn_escanear = st.button("🚀 Escanear Mercado", type="primary", width="stretch")

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

        # ── 1. HUMANIZACIÓN Y REFACTORIZACIÓN VISUAL DE LA TABLA PRINCIPAL EV+ ──
        df_display = df_ops.copy()

        def _formatear_ev_proyectado(fila):
            try:
                raw_st = fila.get("Stake_Num", None)
                if raw_st is None or pd.isna(raw_st):
                    st_str = str(fila.get("Stake_Recomendado", "0")).replace("€", "").strip()
                    stake_val = float(st_str) if st_str else 0.0
                else:
                    stake_val = float(raw_st)
            except (ValueError, TypeError, ZeroDivisionError, KeyError):
                stake_val = 0.0

            try:
                raw_ev = fila.get("EV_Num", None)
                if raw_ev is None or pd.isna(raw_ev):
                    ev_str = str(fila.get("EV", "0")).replace("%", "").replace("⚠️", "").replace("+", "").strip()
                    ev_val = float(ev_str) if ev_str else 0.0
                else:
                    ev_val = float(raw_ev)
            except (ValueError, TypeError, ZeroDivisionError, KeyError):
                ev_val = 0.0

            try:
                beneficio_proyectado = round(stake_val * (ev_val / 100.0), 2) if (stake_val and ev_val) else 0.0
            except (ZeroDivisionError, TypeError, ValueError):
                beneficio_proyectado = 0.0

            sign_b = "+" if beneficio_proyectado >= 0 else ""
            sign_e = "+" if ev_val >= 0 else ""
            return f"{sign_b}{beneficio_proyectado:.2f} € ({sign_e}{ev_val:.1f}%)"

        def _formatear_prob_real(fila):
            try:
                p_num = fila.get("P_Final_Num", None)
                if p_num is not None and not pd.isna(p_num):
                    return f"{float(p_num) * 100.0:.1f}%"
                p_str = str(fila.get("Prob_Modelo", "0")).replace("%", "").strip()
                return f"{float(p_str):.1f}%"
            except (ValueError, TypeError, KeyError):
                return str(fila.get("Prob_Modelo", "0.0%"))

        try:
            df_display["EV"] = df_display.apply(_formatear_ev_proyectado, axis=1)
            df_display["Prob_Modelo"] = df_display.apply(_formatear_prob_real, axis=1)
        except Exception:
            pass

        mapa_renombrado_ev = {
            "Convicción": "Nivel Confianza",
            "Fecha / Hora": "Horario",
            "Partido": "Encuentro",
            "Mercado": "Pronóstico / Selección",
            "Casa": "Bookie",
            "Cuota Mercado": "Cuota Actual",
            "Cuota Mínima Operable": "Cuota Límite (No apostar por debajo)",
            "Prob_Modelo": "Probabilidad Real (%)",
            "Auditoría Sintética": "Recomendación Táctica",
        }
        df_display = df_display.rename(columns=mapa_renombrado_ev)

        # Formateo vectorizado estricto (:.2f) para cuotas en Live Scanner (evitando floats crudos de 5 decimales)
        for _col_cuota_fmt in (
            "Cuota Actual",
            "Cuota Límite (No apostar por debajo)",
            "Cuota Límite",
            "Cuota Justa",
            "Cuota_Justa",
            "Cuota",
        ):
            if _col_cuota_fmt in df_display.columns:
                _s_num = pd.to_numeric(df_display[_col_cuota_fmt], errors="coerce")
                df_display[_col_cuota_fmt] = np.where(
                    _s_num.notna(),
                    _s_num.map(lambda v: f"{float(v):.2f}" if pd.notna(v) else "-"),
                    df_display[_col_cuota_fmt].astype(str),
                )
            if _col_cuota_fmt in df_ops.columns:
                _s_ops = pd.to_numeric(df_ops[_col_cuota_fmt], errors="coerce")
                df_ops[_col_cuota_fmt] = np.where(
                    _s_ops.notna(),
                    _s_ops.map(lambda v: f"{float(v):.2f}" if pd.notna(v) else "-"),
                    df_ops[_col_cuota_fmt].astype(str),
                )

        cols_vista_ejecutiva = [
            "Nivel Confianza",
            "Horario",
            "Encuentro",
            "Pronóstico / Selección",
            "Bookie",
            "Cuota Actual",
            "Cuota Límite (No apostar por debajo)",
            "Probabilidad Real (%)",
            "EV",
            "Stake_Recomendado",
            "Recomendación Táctica",
        ]
        cols_presentes = [
            c for c in cols_vista_ejecutiva
            if c in df_display.columns and c not in ("Prob_NoVig", "Prob_Poisson")
        ]
        df_vista_final = df_display[cols_presentes].copy()

        def _estilo_ev_esmeralda(val):
            try:
                s_val = str(val)
                if "€" in s_val and not s_val.strip().startswith("-"):
                    if not s_val.strip().startswith("+0.00 €") and not s_val.strip().startswith("0.00 €"):
                        return "color: #2ECC71; font-weight: 700;"
            except Exception:
                pass
            return ""

        def _aplicar_tema_oscuro_df(df_in):
            """Styler Command Center: fondo #12161F, filas alternas #1A202C, texto #F1F5F9, bordes #2D3748."""
            def _zebra(fila):
                bg = "#12161F" if fila.name % 2 == 0 else "#1A202C"
                return [f"background-color: {bg}; color: #F1F5F9; border: 1px solid #2D3748;"] * len(fila)
            sty = df_in.reset_index(drop=True).style.apply(_zebra, axis=1)
            sty = sty.set_table_styles([
                {"selector": "th", "props": [("background-color", "#0E1117"), ("color", "#F1F5F9"),
                                             ("border", "1px solid #2D3748")]},
                {"selector": "td", "props": [("border", "1px solid #2D3748")]},
            ])
            return sty

        try:
            styler_ev = _aplicar_tema_oscuro_df(df_vista_final)
            if "EV" in df_vista_final.columns:
                try:
                    styler_ev = styler_ev.map(_estilo_ev_esmeralda, subset=["EV"])
                except AttributeError:
                    styler_ev = styler_ev.applymap(_estilo_ev_esmeralda, subset=["EV"])
            st.dataframe(styler_ev, width="stretch", hide_index=True)
        except Exception:
            st.dataframe(df_vista_final, width="stretch", hide_index=True)

        # ── AUDITORÍA CRUZADA SQLITE, FRICCIÓN ARBITRAL, BACKTEST Y TELETRANSPORTE AL DOSSIER ──
        with st.expander("📋 Auditoría Cruzada SQLite, Fricción Táctica Arbitral, Backtest Empírico y Dossier", expanded=True):
            mostrar_todo_backtest = st.checkbox(
                "Mostrar escenarios con histórico negativo/neutro (Modo Auditoría Completa)",
                value=False,
                key="ev_pro_mostrar_todo_bt",
            )

            def _extraer_roi_numerico_bt(fila_op) -> float:
                try:
                    raw_roi = fila_op.get("Backtest_ROI", None)
                    if raw_roi is not None and not pd.isna(raw_roi):
                        val_r = float(raw_roi)
                        if val_r != 0.0:
                            return val_r
                except (ValueError, TypeError, KeyError):
                    pass
                try:
                    badge_txt = str(fila_op.get("Backtest_Badge", ""))
                    m_roi = re.search(r"ROI:\s*([+-]?\d+(?:\.\d+)?)%", badge_txt)
                    if m_roi:
                        return float(m_roi.group(1))
                except Exception:
                    pass
                return 0.0

            tarjetas_renderizadas = 0
            tarjetas_omitidas_bt = 0

            for idx_r, r_op in df_ops.iterrows():
                roi_historico = _extraer_roi_numerico_bt(r_op)
                if not mostrar_todo_backtest and roi_historico <= 0.0:
                    tarjetas_omitidas_bt += 1
                    continue

                tarjetas_renderizadas += 1
                loc_e = str(r_op.get("Equipo_Local", ""))
                vis_e = str(r_op.get("Equipo_Visitante", ""))
                fric_badge = str(r_op.get("Friccion_Badge", "🍃 Partido Fluido: Árbitro permisivo + baja tasa de faltas"))
                fric_alta = bool(r_op.get("Friccion_Alta", False))
                bt_badge = str(r_op.get("Backtest_Badge", "⚠️ Muestra histórica insuficiente en DB local"))
                arb_nom = str(r_op.get("Arbitro_Asignado", "Colegiado Media Comité"))
                try:
                    arb_am = float(r_op.get("Arb_Amarillas", 4.05) or 4.05)
                except (ValueError, TypeError):
                    arb_am = 4.05
                try:
                    faltas_c = float(r_op.get("Faltas_Comb_6G", 23.0) or 23.0)
                except (ValueError, TypeError):
                    faltas_c = 23.0

                bg_fric = "rgba(239, 68, 68, 0.16)" if fric_alta else "rgba(16, 185, 129, 0.14)"
                bd_fric = "#EF4444" if fric_alta else "#10B981"
                tx_fric = "#FCA5A5" if fric_alta else "#6EE7B7"

                if roi_historico > 0.0:
                    bg_bt = "rgba(46, 204, 113, 0.16)"
                    bd_bt = "#2ECC71"
                    tx_bt = "#6EE7B7"
                elif bt_badge.startswith("📈"):
                    bg_bt = "rgba(239, 68, 68, 0.14)"
                    bd_bt = "#EF4444"
                    tx_bt = "#FCA5A5"
                else:
                    bg_bt = "rgba(245, 158, 11, 0.14)"
                    bd_bt = "#F59E0B"
                    tx_bt = "#FCD34D"

                try:
                    c_merc_val = float(r_op.get("Cuota Mercado", 0.0) or 0.0)
                except (ValueError, TypeError):
                    c_merc_val = 0.0
                try:
                    c_lim_val = float(r_op.get("Cuota Mínima Operable", 0.0) or 0.0)
                except (ValueError, TypeError):
                    c_lim_val = 0.0
                try:
                    xg_loc_val = float(r_op.get("xG_6G_Local", 1.35) or 1.35)
                except (ValueError, TypeError):
                    xg_loc_val = 1.35
                try:
                    xg_vis_val = float(r_op.get("xG_6G_Vis", 1.25) or 1.25)
                except (ValueError, TypeError):
                    xg_vis_val = 1.25

                with st.container(border=True):
                    c_aud1, c_aud2, c_aud3 = st.columns([3.4, 1.3, 1.3])
                    with c_aud1:
                        st.markdown(
                            f"**{r_op.get('Partido', '')}** · `{r_op.get('Mercado', '')}` (@{c_merc_val:.2f} vs Límite @{c_lim_val:.2f})  \n"
                            f"<span style='color:#94A3B8;font-size:12.5px;'>Rolling 6G xG SQLite: "
                            f"<b>{loc_e}</b> {xg_loc_val:.2f} xG · "
                            f"<b>{vis_e}</b> {xg_vis_val:.2f} xG · "
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
                                width="stretch",
                                on_click=_navegar_a_dossier_equipo,
                                args=(eq_loc_canon,),
                            )
                    with c_aud3:
                        if vis_e:
                            st.button(
                                f"📋 Ver en Dossier ({vis_e})",
                                key=f"btn_dos_vis_{idx_r}",
                                width="stretch",
                                on_click=_navegar_a_dossier_equipo,
                                args=(eq_vis_canon,),
                            )

            if tarjetas_renderizadas == 0 and not mostrar_todo_backtest and tarjetas_omitidas_bt > 0:
                st.info(
                    f"🛡️ Filtro de Auditoría Táctica activo: se han omitido {tarjetas_omitidas_bt} escenario(s) con ROI histórico negativo, neutro o sin muestra suficiente en SQLite. "
                    "Activa la casilla superior si deseas auditar todos los registros."
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
        # ── LAYER 3.1: TOP EXECUTIVE CONTROL (COMPETICIÓN + EQUIPO + HORIZONTE) ──
        col_div_d = _columna_liga_partidos(df_partidos_global)
        _CODIGOS_DOMESTICOS_DOSSIER = {"SP1", "SP2", "E0", "E1", "D1", "I1", "F1"}
        _MAPA_COMP_DOSSIER = {
            "Champions League (UEFA)": {"Champions League", "uefa.champions"},
            "Europa League (UEFA)": {"Europa League", "uefa.europa"},
            "Conference League (UEFA)": {"Conference League", "uefa.europa.conf"},
        }
        _OPCIONES_COMP_DOSSIER = [
            "Liga Doméstica",
            "Champions League (UEFA)",
            "Europa League (UEFA)",
            "Conference League (UEFA)",
            "Todas las Competiciones (Global)",
        ]

        ctrl_d0, ctrl_d1, ctrl_d2 = st.columns([1.3, 1.3, 2.2])
        with ctrl_d0:
            comp_dossier_sel = st.selectbox(
                "🏆 Competición:",
                _OPCIONES_COMP_DOSSIER,
                key="dossier_competicion_selector",
            )

        if col_div_d:
            serie_div_global = df_partidos_global[col_div_d].astype(str).str.strip()
            if comp_dossier_sel == "Liga Doméstica":
                df_comp_dossier = df_partidos_global[serie_div_global.isin(_CODIGOS_DOMESTICOS_DOSSIER)].copy()
            elif comp_dossier_sel in _MAPA_COMP_DOSSIER:
                df_comp_dossier = df_partidos_global[serie_div_global.isin(_MAPA_COMP_DOSSIER[comp_dossier_sel])].copy()
            else:
                df_comp_dossier = df_partidos_global.copy()
        else:
            df_comp_dossier = df_partidos_global.copy()

        todos_equipos = sorted(
            pd.concat([df_comp_dossier["HomeTeam"], df_comp_dossier["AwayTeam"]])
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
        if not todos_equipos:
            todos_equipos = sorted(
                pd.concat([df_partidos_global["HomeTeam"], df_partidos_global["AwayTeam"]])
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

        if (
            "dossier_equipo_selector" in st.session_state
            and st.session_state["dossier_equipo_selector"] not in todos_equipos
            and todos_equipos
        ):
            st.session_state["dossier_equipo_selector"] = todos_equipos[0]

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
            df_eq_raw_all = df_comp_dossier[
                (df_comp_dossier["HomeTeam"] == equipo_elegido)
                | (df_comp_dossier["AwayTeam"] == equipo_elegido)
            ].copy()
            df_eq_ord_all = _ordenar_por_fecha(df_eq_raw_all)

            usar_decay_dossier = "Decaimiento" in horizonte_muestra
            if "Últimos 10" in horizonte_muestra:
                df_muestra_raw = df_eq_ord_all.tail(10).copy()
            elif "Temporada Actual" in horizonte_muestra:
                if "Temporada_Origen" in df_eq_ord_all.columns and not df_eq_ord_all["Temporada_Origen"].dropna().empty:
                    temp_actual = str(df_eq_ord_all["Temporada_Origen"].dropna().astype(str).max())
                    df_temp = df_eq_ord_all[df_eq_ord_all["Temporada_Origen"].astype(str) == temp_actual].copy()
                    df_muestra_raw = df_temp if len(df_temp) >= 1 else df_eq_ord_all.tail(38).copy()
                else:
                    df_muestra_raw = df_eq_ord_all.tail(38).copy()
            else:
                df_muestra_raw = df_eq_ord_all.copy()

            # Contexto de competición activa para Attack & Defense Ratings y membrete PDF
            if comp_dossier_sel in _MAPA_COMP_DOSSIER:
                df_liga_ctx = df_comp_dossier if not df_comp_dossier.empty else df_partidos_global
                nombre_liga_ctx = comp_dossier_sel
            elif comp_dossier_sel == "Todas las Competiciones (Global)":
                df_liga_ctx = df_partidos_global
                if col_div_d and not df_eq_ord_all.empty and not df_eq_ord_all[col_div_d].dropna().empty:
                    divs_eq_u = [MAPA_LIGAS.get(str(x), str(x)) for x in df_eq_ord_all[col_div_d].dropna().unique()]
                    nombre_liga_ctx = f"Consolidado Global ({' + '.join(dict.fromkeys(divs_eq_u))})"
                else:
                    nombre_liga_ctx = "Todas las Competiciones (Global)"
            elif col_div_d and not df_eq_ord_all.empty and not df_eq_ord_all[col_div_d].dropna().empty:
                div_eq = str(df_eq_ord_all[col_div_d].dropna().iloc[-1])
                df_liga_ctx = df_partidos_global[df_partidos_global[col_div_d].astype(str) == div_eq]
                nombre_liga_ctx = MAPA_LIGAS.get(div_eq, div_eq)
            else:
                df_liga_ctx = df_partidos_global
                nombre_liga_ctx = comp_dossier_sel

            media_liga_home = float(pd.to_numeric(df_liga_ctx["FTHG"], errors="coerce").mean() or 1.45)
            media_liga_away = float(pd.to_numeric(df_liga_ctx["FTAG"], errors="coerce").mean() or 1.15)
            media_liga_global = (media_liga_home + media_liga_away) / 2.0

            # Muestra global de respaldo (partidos con datos completos de córners/tiros/faltas en SQLite/DataFrame)
            df_eq_glob_all = _ordenar_por_fecha(
                df_partidos_global[
                    (df_partidos_global["HomeTeam"] == equipo_elegido)
                    | (df_partidos_global["AwayTeam"] == equipo_elegido)
                ].copy()
            )
            df_fb_norm_all = _normalizar_partidos_equipo(df_eq_glob_all, equipo_elegido)
            if not df_fb_norm_all.empty and "corners_f" in df_fb_norm_all.columns:
                df_fb_con_datos = df_fb_norm_all[df_fb_norm_all["corners_f"].notna()].copy()
            else:
                df_fb_con_datos = df_fb_norm_all.copy()
            if "Últimos 10" in horizonte_muestra and not df_fb_con_datos.empty:
                df_fb_con_datos = df_fb_con_datos.tail(10).copy()
            elif "Temporada Actual" in horizonte_muestra and not df_fb_con_datos.empty:
                df_fb_con_datos = df_fb_con_datos.tail(38).copy()

            df_fb_home = df_fb_con_datos[df_fb_con_datos["es_local"]].copy() if not df_fb_con_datos.empty else pd.DataFrame()
            df_fb_away = df_fb_con_datos[~df_fb_con_datos["es_local"]].copy() if not df_fb_con_datos.empty else pd.DataFrame()

            df_norm_global = _normalizar_partidos_equipo(df_muestra_raw, equipo_elegido)
            df_norm_home = df_norm_global[df_norm_global["es_local"]].copy() if not df_norm_global.empty else pd.DataFrame()
            df_norm_away = df_norm_global[~df_norm_global["es_local"]].copy() if not df_norm_global.empty else pd.DataFrame()

            stats_global = _calcular_resumen_bloque(
                df_norm_global,
                usar_decay=usar_decay_dossier,
                media_liga_gf=media_liga_global,
                media_liga_gc=media_liga_global,
                df_fallback=df_fb_con_datos,
            )
            stats_home = _calcular_resumen_bloque(
                df_norm_home,
                usar_decay=usar_decay_dossier,
                media_liga_gf=media_liga_home,
                media_liga_gc=media_liga_away,
                df_fallback=df_fb_home if not df_fb_home.empty else df_fb_con_datos,
            )
            stats_away = _calcular_resumen_bloque(
                df_norm_away,
                usar_decay=usar_decay_dossier,
                media_liga_gf=media_liga_away,
                media_liga_gc=media_liga_home,
                df_fallback=df_fb_away if not df_fb_away.empty else df_fb_con_datos,
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
                    file_name=f"informe_ejecutivo_{_slug_archivo(equipo_elegido)}_{_slug_archivo(comp_dossier_sel)}.pdf",
                    mime="application/pdf",
                    type="primary",
                    width="stretch",
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

            # ── ROW 2: SPLIT COMPARATIVE ANALYSIS: CASA VS FUERA / GLOBAL ───
            st.markdown("---")
            if hasattr(st, "segmented_control"):
                n_partidos_muestra = (
                    st.segmented_control(
                        "Muestra de análisis:",
                        options=[5, 10],
                        default=5,
                        key="dossier_n_partidos_muestra",
                    )
                    or 5
                )
            else:
                n_partidos_muestra = st.radio(
                    "Muestra de análisis:",
                    options=[5, 10],
                    index=0,
                    horizontal=True,
                    key="dossier_n_partidos_muestra",
                )

            st.subheader("🏠 Análisis Comparativo: Condición Local vs ✈️ Condición Visitante")

            df_norm_all = _normalizar_partidos_equipo(df_eq_ord_all, equipo_elegido)
            df_home_all = (
                df_norm_all[df_norm_all["es_local"]].copy()
                if not df_norm_all.empty
                else pd.DataFrame()
            )
            df_away_all = (
                df_norm_all[~df_norm_all["es_local"]].copy()
                if not df_norm_all.empty
                else pd.DataFrame()
            )

            df_slice_home = df_home_all.tail(int(n_partidos_muestra)).copy()
            df_slice_away = df_away_all.tail(int(n_partidos_muestra)).copy()
            df_slice_gen = df_norm_all.tail(int(n_partidos_muestra)).copy()

            stats_comp_home = _calcular_resumen_bloque(
                df_slice_home,
                usar_decay=False,
                media_liga_gf=media_liga_home,
                media_liga_gc=media_liga_away,
                df_fallback=df_fb_home if not df_fb_home.empty else df_fb_con_datos,
            )
            stats_comp_away = _calcular_resumen_bloque(
                df_slice_away,
                usar_decay=False,
                media_liga_gf=media_liga_away,
                media_liga_gc=media_liga_home,
                df_fallback=df_fb_away if not df_fb_away.empty else df_fb_con_datos,
            )
            stats_comp_gen = _calcular_resumen_bloque(
                df_slice_gen,
                usar_decay=False,
                media_liga_gf=media_liga_global,
                media_liga_gc=media_liga_global,
                df_fallback=df_fb_con_datos,
            )

            tab_local, tab_visitante, tab_general = st.tabs(
                [
                    "🏠 Rendimiento Local",
                    "✈️ Rendimiento Visitante",
                    "🌐 Rendimiento General (Total)",
                ]
            )

            with tab_local:
                st.markdown(f"#### 🏠 Rendimiento en Casa ({stats_comp_home['n']} partidos)")
                if stats_comp_home["n"] > 0:
                    hc1, hc2, hc3 = st.columns(3)
                    hc1.metric(
                        "Goles F / C (90m)",
                        f"{stats_comp_home['gf_med']:.2f} / {stats_comp_home['gc_med']:.2f}",
                        f"xG: {stats_comp_home['xg_f']:.2f} F vs {stats_comp_home['xg_c']:.2f} C",
                        delta_color="off",
                    )
                    hc2.metric(
                        "Tiros a Puerta (SOT)",
                        f"{stats_comp_home['sot_f']:.2f} / {stats_comp_home['sot_c']:.2f}",
                        f"Conversión SOT: {stats_comp_home['conv_sot']:.1f}%",
                        delta_color="off",
                    )
                    hc3.metric(
                        "Ratings vs Liga (Casa)",
                        f"Atk {stats_comp_home['atk_rating']:.2f}x",
                        f"Def {stats_comp_home['def_rating']:.2f}x (Media Liga: {media_liga_home:.2f})",
                        delta_color="off",
                    )
                    st.dataframe(
                        _construir_tabla_rendimiento_condicion(stats_comp_home, "Métrica Local"),
                        width="stretch",
                        hide_index=True,
                    )
                else:
                    st.info("Sin registros como local en la muestra seleccionada.")

            with tab_visitante:
                st.markdown(f"#### ✈️ Rendimiento Fuera ({stats_comp_away['n']} partidos)")
                if stats_comp_away["n"] > 0:
                    ac1, ac2, ac3 = st.columns(3)
                    ac1.metric(
                        "Goles F / C (90m)",
                        f"{stats_comp_away['gf_med']:.2f} / {stats_comp_away['gc_med']:.2f}",
                        f"xG: {stats_comp_away['xg_f']:.2f} F vs {stats_comp_away['xg_c']:.2f} C",
                        delta_color="off",
                    )
                    ac2.metric(
                        "Tiros a Puerta (SOT)",
                        f"{stats_comp_away['sot_f']:.2f} / {stats_comp_away['sot_c']:.2f}",
                        f"Conversión SOT: {stats_comp_away['conv_sot']:.1f}%",
                        delta_color="off",
                    )
                    ac3.metric(
                        "Ratings vs Liga (Fuera)",
                        f"Atk {stats_comp_away['atk_rating']:.2f}x",
                        f"Def {stats_comp_away['def_rating']:.2f}x (Media Liga: {media_liga_away:.2f})",
                        delta_color="off",
                    )
                    st.dataframe(
                        _construir_tabla_rendimiento_condicion(
                            stats_comp_away, "Métrica Visitante"
                        ),
                        width="stretch",
                        hide_index=True,
                    )
                else:
                    st.info("Sin registros como visitante en la muestra seleccionada.")

            with tab_general:
                st.markdown(
                    f"#### 🌐 Rendimiento General (Total) ({stats_comp_gen['n']} partidos)"
                )
                if stats_comp_gen["n"] > 0:
                    gc1, gc2, gc3 = st.columns(3)
                    gc1.metric(
                        "Goles F / C (90m)",
                        f"{stats_comp_gen['gf_med']:.2f} / {stats_comp_gen['gc_med']:.2f}",
                        f"xG: {stats_comp_gen['xg_f']:.2f} F vs {stats_comp_gen['xg_c']:.2f} C",
                        delta_color="off",
                    )
                    gc2.metric(
                        "Tiros a Puerta (SOT)",
                        f"{stats_comp_gen['sot_f']:.2f} / {stats_comp_gen['sot_c']:.2f}",
                        f"Conversión SOT: {stats_comp_gen['conv_sot']:.1f}%",
                        delta_color="off",
                    )
                    gc3.metric(
                        "Ratings vs Liga (Global)",
                        f"Atk {stats_comp_gen['atk_rating']:.2f}x",
                        f"Def {stats_comp_gen['def_rating']:.2f}x (Media Liga: {media_liga_global:.2f})",
                        delta_color="off",
                    )
                    st.dataframe(
                        _construir_tabla_rendimiento_condicion(stats_comp_gen, "Métrica General"),
                        width="stretch",
                        hide_index=True,
                    )
                else:
                    st.info("Sin registros en la muestra seleccionada.")

            # ── ROW 3: MARKET FREQUENCY MATRIX (TABLE/DATAFRAME) ─────────────
            st.markdown("---")
            st.subheader("🎯 Matriz de Frecuencias de Mercado y Cuotas Implícitas")
            st.caption(
                "Porcentajes históricos de cumplimiento (Hit Rate %) y cuotas justas desagregadas por condición (Casa / Fuera) sin margen (1 / p)."
            )

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

            p_gen_arr = np.array([float(stats_global[clave_m]) / 100.0 for _, clave_m in definicion_mercados], dtype=float)
            p_home_arr = np.array([float(stats_home[clave_m]) / 100.0 for _, clave_m in definicion_mercados], dtype=float)
            p_away_arr = np.array([float(stats_away[clave_m]) / 100.0 for _, clave_m in definicion_mercados], dtype=float)

            cuota_justa_casa_arr = np.where(p_home_arr > 0, np.round(1.0 / np.maximum(p_home_arr, 0.01), 2), np.nan)
            cuota_justa_fuera_arr = np.where(p_away_arr > 0, np.round(1.0 / np.maximum(p_away_arr, 0.01), 2), np.nan)

            filas_mercados = []
            for idx_m, (etiqueta_m, _) in enumerate(definicion_mercados):
                cj_h = cuota_justa_casa_arr[idx_m]
                cj_a = cuota_justa_fuera_arr[idx_m]
                filas_mercados.append({
                    "Mercado": etiqueta_m,
                    "Hit Rate General (%)": f"{p_gen_arr[idx_m] * 100.0:.1f}%",
                    "Hit Rate Casa (%)": f"{p_home_arr[idx_m] * 100.0:.1f}%",
                    "Hit Rate Fuera (%)": f"{p_away_arr[idx_m] * 100.0:.1f}%",
                    "Cuota Justa Casa (1/p)": f"{cj_h:.2f}" if not np.isnan(cj_h) else f"{1.0 / max(p_home_arr[idx_m], 0.01):.2f}",
                    "Cuota Justa Fuera (1/p)": f"{cj_a:.2f}" if not np.isnan(cj_a) else f"{1.0 / max(p_away_arr[idx_m], 0.01):.2f}",
                })

            st.dataframe(pd.DataFrame(filas_mercados), width="stretch", hide_index=True)

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
                censo_global_dos = _obtener_censo_global_arbitros(df_partidos_global)
                censo_ordenado_dos = sorted(censo_global_dos.values(), key=lambda x: x["nombre"])
                arbs_dossier = ["🔍 Seleccionar o buscar árbitro..."] + [item["etiqueta_selector"] for item in censo_ordenado_dos]

                # Preselección si el último partido tiene árbitro registrado
                idx_arb_def = 0
                ult_ref_cand = ""
                if "Referee" in df_norm_global.columns and not df_norm_global["Referee"].dropna().empty:
                    ult_ref_cand = str(df_norm_global["Referee"].dropna().iloc[-1]).strip()
                if ult_ref_cand:
                    ult_ref_can = _normalizar_nombre_arbitro(ult_ref_cand)
                    for _i_opt, _opt_lbl in enumerate(arbs_dossier):
                        if _opt_lbl.startswith(ult_ref_can):
                            idx_arb_def = _i_opt
                            break

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
                        index=idx_arb_def,
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

                # Resolver y renderizar tarjeta de perfil disciplinario del árbitro seleccionado
                div_dossier = MAPA_INV.get(comp_dossier_sel, "SP1")
                perfil_arb_dossier = _calcular_perfil_disciplinario_arbitro(
                    nombre_arbitro=arb_dossier,
                    censo_global=censo_global_dos,
                    df_partidos_global=df_partidos_global,
                    div_partido=div_dossier,
                )
                _renderizar_tarjeta_perfil_arbitro(perfil_arb_dossier)

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

                media_tarj_arb = float(perfil_arb_dossier.get("media_amarillas", 4.10))
                nom_arb_cron = str(perfil_arb_dossier.get("nombre", "Árbitro por confirmar"))

                fallback_payload = {
                    "home": equipo_elegido,
                    "away": rival_dossier,
                    "xg_home": xg_h_cron,
                    "xg_away": xg_a_cron,
                    "cuota_mercado": float(cuota_merc_dossier),
                    "cuota_justa": cuota_justa_cron,
                    "arbitro": nom_arb_cron,
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
                match_id_cron = f"{equipo_elegido}_vs_{rival_dossier}_{nom_arb_cron}_{cuota_merc_dossier:.2f}_{horizonte_muestra}"

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
                _df_jug_fuente_dos = df_jugadores
                if comp_dossier_sel in _MAPA_COMP_DOSSIER and RUTA_BD.exists():
                    try:
                        with sqlite3.connect(RUTA_BD) as _c_jdos:
                            _comp_bd_nom = comp_dossier_sel.replace(" (UEFA)", "").strip()
                            _df_jug_fuente_dos = cargar_jugadores_competicion(_c_jdos, _comp_bd_nom)
                    except Exception:
                        _df_jug_fuente_dos = df_jugadores
                if _df_jug_fuente_dos is not None and not _df_jug_fuente_dos.empty:
                    if "Equipo" in _df_jug_fuente_dos.columns:
                        _eq_dos_clean = str(equipo_elegido).strip()
                        _mask_dos_eq = _df_jug_fuente_dos["Equipo"].astype(str).str.strip() == _eq_dos_clean
                        if comp_dossier_sel in _MAPA_COMP_DOSSIER and "competicion" in _df_jug_fuente_dos.columns:
                            _comp_dos_clean = comp_dossier_sel.replace(" (UEFA)", "").strip()
                            _mask_dos_eq = _mask_dos_eq & (
                                _df_jug_fuente_dos["competicion"].astype(str).str.strip() == _comp_dos_clean
                            )
                        df_jug_eq = _df_jug_fuente_dos[_mask_dos_eq].drop_duplicates(subset=["Jugador"], keep="first").copy()
                    else:
                        df_jug_eq = pd.DataFrame()

                    # Esquema intuitivo (goles/asistencias/posicion) con fallback a nombres legados
                    for _nuevo, _viejo, _etq in (("goles", "Gls", "Goles"), ("asistencias", "Ast", "Asistencias"), ("posicion", "Pos", "Pos")):
                        if _etq not in df_jug_eq.columns:
                            if _nuevo in df_jug_eq.columns:
                                df_jug_eq[_etq] = df_jug_eq[_nuevo]
                            elif _viejo in df_jug_eq.columns:
                                df_jug_eq[_etq] = df_jug_eq[_viejo]

                    cols_mostrar = ["Jugador", "Pos", "Goles", "Asistencias"]
                    tiene_columnas = all(c in df_jug_eq.columns for c in cols_mostrar)

                    if not df_jug_eq.empty and tiene_columnas:
                        st.dataframe(df_jug_eq[cols_mostrar], width="stretch", hide_index=True)
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
        _OPCION_H2H_TODAS = "Todas las Competiciones (Liga + Champions/UEFA)"
        divs_disp_h2h = (
            sorted(df_partidos_global["Div"].dropna().astype(str).unique().tolist())
            if "Div" in df_partidos_global.columns
            else []
        )
        # Mapear códigos sin duplicar etiquetas legibles
        ligas_disp_h2h = [_OPCION_H2H_TODAS]
        _mapa_etq_a_cods: dict[str, list[str]] = {}
        for _d in divs_disp_h2h:
            _etq = MAPA_LIGAS.get(_d, _d)
            _mapa_etq_a_cods.setdefault(_etq, []).append(str(_d))
            if _etq not in ligas_disp_h2h:
                ligas_disp_h2h.append(_etq)

        sel_c1, sel_c2, sel_c3, sel_c4 = st.columns([1.5, 1.5, 1.5, 1.1])
        with sel_c1:
            liga_h2h_nom = st.selectbox("🏆 Competición / Liga:", ligas_disp_h2h, key="h2h_liga_sel")

        if liga_h2h_nom == _OPCION_H2H_TODAS or not divs_disp_h2h:
            df_liga_h2h = df_partidos_global.copy()
        else:
            _cods_sel = _mapa_etq_a_cods.get(liga_h2h_nom, [MAPA_INV.get(liga_h2h_nom, liga_h2h_nom)])
            df_liga_h2h = df_partidos_global[df_partidos_global["Div"].astype(str).isin(_cods_sel)].copy()

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
                width="stretch",
                key="btn_swap_h2h",
            )

        # Selector de modo muestral para evitar sesgo de muestra asimétrica (ej. 1 partido UEFA vs 45 de Liga)
        _MODOS_MUESTRA_H2H = [
            "Solo esta Competición (UEFA)",
            "Últimos N Partidos Oficiales Homogéneos (5/10 partidos)",
        ]
        ms_c1, ms_c2 = st.columns([2.4, 1.6])
        with ms_c1:
            modo_muestra_h2h = st.selectbox(
                "⚖️ Criterio de Balanceo Muestral (Anti-Sesgo):",
                _MODOS_MUESTRA_H2H,
                index=1 if "UEFA" in str(liga_h2h_nom) else 0,
                key="h2h_modo_muestra",
            )
        with ms_c2:
            n_homogeneo_h2h = st.radio(
                "📏 Ventana Homogénea (N):",
                [5, 10],
                index=1,
                horizontal=True,
                key="h2h_n_homogeneo",
            )

        if eq1 and eq2 and eq1 != eq2:
            df_base_muestra_h2h = (
                df_liga_h2h
                if modo_muestra_h2h == "Solo esta Competición (UEFA)"
                else df_partidos_global
            )
            df_base_ord_h2h = _ordenar_por_fecha(df_base_muestra_h2h)
            df_glob_ord_h2h = _ordenar_por_fecha(df_partidos_global)

            df_eq1_home_full = df_base_ord_h2h[df_base_ord_h2h["HomeTeam"] == eq1].copy()
            df_eq2_away_full = df_base_ord_h2h[df_base_ord_h2h["AwayTeam"] == eq2].copy()
            df_eq1_all_full = df_glob_ord_h2h[
                (df_glob_ord_h2h["HomeTeam"] == eq1) | (df_glob_ord_h2h["AwayTeam"] == eq1)
            ].copy()
            df_eq2_all_full = df_glob_ord_h2h[
                (df_glob_ord_h2h["HomeTeam"] == eq2) | (df_glob_ord_h2h["AwayTeam"] == eq2)
            ].copy()

            if "Últimos N" in modo_muestra_h2h:
                df_eq1_home = df_eq1_home_full.tail(int(n_homogeneo_h2h)).copy()
                df_eq2_away = df_eq2_away_full.tail(int(n_homogeneo_h2h)).copy()
                df_eq1_all = df_eq1_all_full.tail(int(n_homogeneo_h2h)).copy()
                df_eq2_all = df_eq2_all_full.tail(int(n_homogeneo_h2h)).copy()
            else:
                df_eq1_home = df_eq1_home_full.copy()
                df_eq2_away = df_eq2_away_full.copy()
                df_eq1_all = df_eq1_all_full.tail(10).copy()
                df_eq2_all = df_eq2_all_full.tail(10).copy()

            for _df_tmp in (df_eq1_home, df_eq2_away, df_eq1_all, df_eq2_all):
                for _c_num in ("FTHG", "FTAG", "HTHG", "HTAG", "HS", "AS", "HST", "AST", "HC", "AC", "HF", "AF", "HxG", "AxG"):
                    if _c_num in _df_tmp.columns:
                        _df_tmp[_c_num] = pd.to_numeric(_df_tmp[_c_num], errors="coerce")

            # ── 2. RESILIENCIA MUESTRAL Y PONDERACIÓN CON FACTOR DE VENTAJA DE CAMPO ──
            n_home_strict = int(df_eq1_home["FTHG"].notna().sum()) if "FTHG" in df_eq1_home.columns else len(df_eq1_home)
            n_away_strict = int(df_eq2_away["FTAG"].notna().sum()) if "FTAG" in df_eq2_away.columns else len(df_eq2_away)
            n_all_eq1 = int(df_eq1_all["FTHG"].notna().sum()) if "FTHG" in df_eq1_all.columns else len(df_eq1_all)
            n_all_eq2 = int(df_eq2_all["FTAG"].notna().sum()) if "FTAG" in df_eq2_all.columns else len(df_eq2_all)

            # Factor medio de ventaja de campo de la competición seleccionada
            _df_comp_ref = df_liga_h2h if not df_liga_h2h.empty else df_partidos_global
            _comp_hg_mean = float(pd.to_numeric(_df_comp_ref.get("FTHG"), errors="coerce").mean() or 1.48)
            _comp_ag_mean = float(pd.to_numeric(_df_comp_ref.get("FTAG"), errors="coerce").mean() or 1.18)
            _comp_avg_goal = max(0.50, (_comp_hg_mean + _comp_ag_mean) / 2.0)
            factor_ventaja_local = float(np.clip(_comp_hg_mean / _comp_avg_goal, 1.02, 1.25))
            factor_ventaja_vis = float(np.clip(_comp_ag_mean / _comp_avg_goal, 0.78, 0.98))

            def _media_global_equipo(df_all_eq: pd.DataFrame, eq_nombre: str, col_home: str, col_away: str, default_val: float) -> float:
                if df_all_eq is None or df_all_eq.empty or col_home not in df_all_eq.columns or col_away not in df_all_eq.columns:
                    return float(default_val)
                es_loc_m = df_all_eq["HomeTeam"].astype(str) == str(eq_nombre)
                vals = np.where(es_loc_m, df_all_eq[col_home], df_all_eq[col_away])
                s_vals = pd.Series(vals).dropna()
                return float(s_vals.mean()) if not s_vals.empty else float(default_val)

            gf_l_glob = _media_global_equipo(df_eq1_all, eq1, "FTHG", "FTAG", _comp_hg_mean)
            gc_l_glob = _media_global_equipo(df_eq1_all, eq1, "FTAG", "FTHG", _comp_ag_mean)
            gf_v_glob = _media_global_equipo(df_eq2_all, eq2, "FTHG", "FTAG", _comp_ag_mean)
            gc_v_glob = _media_global_equipo(df_eq2_all, eq2, "FTAG", "FTHG", _comp_hg_mean)

            gf_l_raw = float(df_eq1_home["FTHG"].mean()) if not df_eq1_home.empty and df_eq1_home["FTHG"].notna().any() else gf_l_glob
            gc_l_raw = float(df_eq1_home["FTAG"].mean()) if not df_eq1_home.empty and df_eq1_home["FTAG"].notna().any() else gc_l_glob
            gf_v_raw = float(df_eq2_away["FTAG"].mean()) if not df_eq2_away.empty and df_eq2_away["FTAG"].notna().any() else gf_v_glob
            gc_v_raw = float(df_eq2_away["FTHG"].mean()) if not df_eq2_away.empty and df_eq2_away["FTHG"].notna().any() else gc_v_glob

            # Si el equipo tiene < 3 partidos en esa condición específica (ej. AEK Athens con 1 partido en casa),
            # relajamos el filtro estricto y usamos sus últimos N partidos oficiales totales ponderados por el factor de campo.
            relajo_local = n_home_strict < 3
            relajo_vis = n_away_strict < 3
            n_home_sample = n_all_eq1 if relajo_local else n_home_strict
            n_away_sample = n_all_eq2 if relajo_vis else n_away_strict

            if relajo_local:
                gf_l = gf_l_glob * factor_ventaja_local
                gc_l = gc_l_glob * factor_ventaja_vis
            elif n_home_strict < 6:
                gf_l = (0.75 * gf_l_raw) + (0.25 * gf_l_glob * factor_ventaja_local)
                gc_l = (0.75 * gc_l_raw) + (0.25 * gc_l_glob * factor_ventaja_vis)
            else:
                gf_l = gf_l_raw
                gc_l = gc_l_raw

            if relajo_vis:
                gf_v = gf_v_glob * factor_ventaja_vis
                gc_v = gc_v_glob * factor_ventaja_local
            elif n_away_strict < 6:
                gf_v = (0.75 * gf_v_raw) + (0.25 * gf_v_glob * factor_ventaja_vis)
                gc_v = (0.75 * gc_v_raw) + (0.25 * gc_v_glob * factor_ventaja_local)
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
            _etq_m_loc = (
                f"{n_home_sample} partidos oficiales totales ({n_home_strict} en casa · ponderado factor campo {factor_ventaja_local:.2f}x)"
                if relajo_local
                else f"{n_home_sample} partidos en casa"
            )
            _etq_m_vis = (
                f"{n_away_sample} partidos oficiales totales ({n_away_strict} fuera · ponderado factor campo {factor_ventaja_vis:.2f}x)"
                if relajo_vis
                else f"{n_away_sample} partidos como visitante"
            )
            _badge_ajuste_h2h = (
                f" &nbsp;·&nbsp; <span style='background:rgba(56,189,248,0.14);color:#38BDF8;border:1px solid rgba(56,189,248,0.35);"
                f"border-radius:5px;padding:2px 7px;font-size:11.5px;font-weight:600;'>⚖️ Muestra Total Oficial + Factor Campo Activo</span>"
                if (relajo_local or relajo_vis)
                else ""
            )
            st.markdown(
                f"**Goles Esperados (xG Proyectado): {eq1} {xg_l:.2f} — {eq2} {xg_v:.2f}**{_badge_ajuste_h2h}  \n"
                f"<span style='color:#94A3B8;font-size:13px;'>Muestra analizada: <b>{eq1}</b> ({_etq_m_loc}) vs "
                f"<b>{eq2}</b> ({_etq_m_vis}).</span>",
                unsafe_allow_html=True,
            )

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
                # Buscar si existen cuotas de casas de apuestas en BD para este emparejamiento exacto
                _df_match_odds = df_base_ord_h2h[
                    (df_base_ord_h2h["HomeTeam"] == eq1) & (df_base_ord_h2h["AwayTeam"] == eq2)
                ]
                def _extraer_cuota_bd_o_justa(df_m: pd.DataFrame, cols_odds: list[str], prob_justa: float) -> tuple[float, bool]:
                    cuota_j = float(np.clip(round(1.0 / max(prob_justa, 1e-4), 2), 1.01, 50.0))
                    if df_m is not None and not df_m.empty:
                        ult = df_m.iloc[-1]
                        for c_o in cols_odds:
                            if c_o in df_m.columns:
                                v_o = pd.to_numeric(ult.get(c_o), errors="coerce")
                                if pd.notna(v_o) and 1.01 <= float(v_o) <= 50.0:
                                    return round(float(v_o), 2), True
                    return cuota_j, False

                _def_c1, _hay_bd_1 = _extraer_cuota_bd_o_justa(_df_match_odds, ["AvgH", "B365H", "PSH", "MaxH"], p_1)
                _def_cx, _hay_bd_x = _extraer_cuota_bd_o_justa(_df_match_odds, ["AvgD", "B365D", "PSD", "MaxD"], p_x)
                _def_c2, _hay_bd_2 = _extraer_cuota_bd_o_justa(_df_match_odds, ["AvgA", "B365A", "PSA", "MaxA"], p_2)
                _def_co25, _hay_bd_o25 = _extraer_cuota_bd_o_justa(_df_match_odds, ["Avg>2.5", "B365>2.5", "P>2.5", "Max>2.5"], p_o25)
                _hay_cuotas_bd = _hay_bd_1 and _hay_bd_x and _hay_bd_2

                # Si no hay cuotas en BD, sincronizar con la Cuota Justa exacta para que EV inicial sea 0.00%
                _firma_ctx_h2h = f"{liga_h2h_nom}|{eq1}|{eq2}|{modo_muestra_h2h}|{n_homogeneo_h2h}|{round(p_1,4)}|{round(p_x,4)}|{round(p_2,4)}|{round(p_o25,4)}"
                if st.session_state.get("_h2h_firma_cuotas") != _firma_ctx_h2h:
                    st.session_state["_h2h_firma_cuotas"] = _firma_ctx_h2h
                    st.session_state["h2h_mkt_1"] = _def_c1 if _hay_cuotas_bd else round(1.0 / max(p_1, 1e-4), 2)
                    st.session_state["h2h_mkt_x"] = _def_cx if _hay_cuotas_bd else round(1.0 / max(p_x, 1e-4), 2)
                    st.session_state["h2h_mkt_2"] = _def_c2 if _hay_cuotas_bd else round(1.0 / max(p_2, 1e-4), 2)
                    st.session_state["h2h_mkt_o25"] = _def_co25 if _hay_bd_o25 else round(1.0 / max(p_o25, 1e-4), 2)

                if not _hay_cuotas_bd:
                    st.info("ℹ️ Sin cuotas de apuestas registradas en BD para este cruce: campos inicializados con la **Cuota Justa del Modelo (EV = 0.00%)**.")

                in_c1, in_cx, in_c2, in_co25 = st.columns(4)
                with in_c1:
                    cuota_mkt_1 = st.number_input(
                        f"Cuota Mercado 1 ({eq1}):",
                        min_value=1.01,
                        max_value=50.0,
                        value=float(st.session_state.get("h2h_mkt_1", round(1.0 / p_1, 2))),
                        step=0.05,
                        key="h2h_mkt_1",
                    )
                with in_cx:
                    cuota_mkt_x = st.number_input(
                        "Cuota Mercado X (Empate):",
                        min_value=1.01,
                        max_value=50.0,
                        value=float(st.session_state.get("h2h_mkt_x", round(1.0 / p_x, 2))),
                        step=0.05,
                        key="h2h_mkt_x",
                    )
                with in_c2:
                    cuota_mkt_2 = st.number_input(
                        f"Cuota Mercado 2 ({eq2}):",
                        min_value=1.01,
                        max_value=50.0,
                        value=float(st.session_state.get("h2h_mkt_2", round(1.0 / p_2, 2))),
                        step=0.05,
                        key="h2h_mkt_2",
                    )
                with in_co25:
                    cuota_mkt_o25 = st.number_input(
                        "Cuota Over 2.5 Goles:",
                        min_value=1.01,
                        max_value=50.0,
                        value=float(st.session_state.get("h2h_mkt_o25", round(1.0 / p_o25, 2))),
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
                    cuota_justa_arb = 1.0 / max(prob_m_arb, 1e-4)
                    if abs(float(cuota_m_arb) - round(cuota_justa_arb, 2)) < 1e-4:
                        ev_val_pct = 0.0
                        quarter_kelly_pct = 0.0
                    else:
                        ev_val_pct = ((prob_m_arb * cuota_m_arb) - 1.0) * 100.0
                        b_neto = max(cuota_m_arb - 1.0, 0.01)
                        kelly_full = ((prob_m_arb * b_neto) - (1.0 - prob_m_arb)) / b_neto
                        quarter_kelly_pct = max(0.0, kelly_full * 0.25 * 100.0)

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

            _df_tact_l = df_eq1_all if relajo_local else df_eq1_home
            _df_tact_v = df_eq2_all if relajo_vis else df_eq2_away

            if not relajo_local and "HxG" in df_eq1_home.columns and df_eq1_home["HxG"].notna().any():
                xga_l = float(df_eq1_home["AxG"].dropna().mean()) if "AxG" in df_eq1_home.columns and df_eq1_home["AxG"].notna().any() else gc_l
            else:
                xga_l = _media_global_equipo(df_eq1_all, eq1, "AxG", "HxG", gc_l) * (factor_ventaja_vis if relajo_local else 1.0)

            if not relajo_vis and "AxG" in df_eq2_away.columns and df_eq2_away["AxG"].notna().any():
                xga_v = float(df_eq2_away["HxG"].dropna().mean()) if "HxG" in df_eq2_away.columns and df_eq2_away["HxG"].notna().any() else gc_v
            else:
                xga_v = _media_global_equipo(df_eq2_all, eq2, "HxG", "AxG", gc_v) * (factor_ventaja_local if relajo_vis else 1.0)

            _tiene_corners_l = (
                "HC" in _df_tact_l.columns
                and "AC" in _df_tact_l.columns
                and _df_tact_l["HC"].notna().any()
                and float((_df_tact_l["HC"].fillna(0) + _df_tact_l["AC"].fillna(0)).mean()) > 0.0
            )
            _tiene_corners_v = (
                "HC" in _df_tact_v.columns
                and "AC" in _df_tact_v.columns
                and _df_tact_v["HC"].notna().any()
                and float((_df_tact_v["HC"].fillna(0) + _df_tact_v["AC"].fillna(0)).mean()) > 0.0
            )
            cor_l = (
                float((_df_tact_l["HC"].dropna() + _df_tact_l["AC"].dropna()).mean())
                if _tiene_corners_l
                else None
            )
            cor_v = (
                float((_df_tact_v["HC"].dropna() + _df_tact_v["AC"].dropna()).mean())
                if _tiene_corners_v
                else None
            )

            faltas_l = (
                float(df_eq1_home["HF"].dropna().mean())
                if (not relajo_local and "HF" in df_eq1_home.columns and df_eq1_home["HF"].notna().any())
                else _media_global_equipo(df_eq1_all, eq1, "HF", "AF", 11.5)
            )
            faltas_v = (
                float(df_eq2_away["AF"].dropna().mean())
                if (not relajo_vis and "AF" in df_eq2_away.columns and df_eq2_away["AF"].notna().any())
                else _media_global_equipo(df_eq2_all, eq2, "HF", "AF", 11.8)
            )

            sot_l = (
                float(df_eq1_home["HST"].dropna().mean())
                if (not relajo_local and "HST" in df_eq1_home.columns and df_eq1_home["HST"].notna().any())
                else _media_global_equipo(df_eq1_all, eq1, "HST", "AST", 4.5) * (factor_ventaja_local if relajo_local else 1.0)
            )
            sot_v = (
                float(df_eq2_away["AST"].dropna().mean())
                if (not relajo_vis and "AST" in df_eq2_away.columns and df_eq2_away["AST"].notna().any())
                else _media_global_equipo(df_eq2_all, eq2, "HST", "AST", 3.8) * (factor_ventaja_vis if relajo_vis else 1.0)
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

            norm_loc = [
                0.0 if (v is None or v <= 0.0) else float(np.clip((v / m) * 100.0, 8.0, 100.0))
                for v, m in zip(raw_loc, topes_escala)
            ]
            norm_vis = [
                0.0 if (v is None or v <= 0.0) else float(np.clip((v / m) * 100.0, 8.0, 100.0))
                for v, m in zip(raw_vis, topes_escala)
            ]
            txt_loc = ["N/D" if (v is None or v <= 0.0) else f"{v:.2f}" for v in raw_loc]
            txt_vis = ["N/D" if (v is None or v <= 0.0) else f"{v:.2f}" for v in raw_vis]

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
                        text=txt_loc,
                        textposition="outside",
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
                        text=txt_vis,
                        textposition="outside",
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
                st.plotly_chart(fig_bars, width="stretch")

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
                    st.dataframe(df_ah_justo, width="stretch", hide_index=True)

            # ── 4B. ANÁLISIS COMPARATIVO: CONDICIÓN LOCAL VS CONDICIÓN VISITANTE ──
            st.markdown("---")
            st.subheader("🏠 Análisis Comparativo: Condición Local vs ✈️ Condición Visitante")

            _df_h2h_norm_home = _normalizar_partidos_equipo(
                df_eq1_all.tail(5) if relajo_local else df_eq1_home.tail(int(n_homogeneo_h2h)),
                eq1,
            )
            _df_h2h_norm_away = _normalizar_partidos_equipo(
                df_eq2_all.tail(5) if relajo_vis else df_eq2_away.tail(int(n_homogeneo_h2h)),
                eq2,
            )
            _df_h2h_fb_home_all = _normalizar_partidos_equipo(
                df_eq1_all_full[df_eq1_all_full["HomeTeam"] == eq1],
                eq1,
            )
            _df_h2h_fb_away_all = _normalizar_partidos_equipo(
                df_eq2_all_full[df_eq2_all_full["AwayTeam"] == eq2],
                eq2,
            )
            if not _df_h2h_fb_home_all.empty and "corners_f" in _df_h2h_fb_home_all.columns:
                _df_h2h_fb_home = _df_h2h_fb_home_all[_df_h2h_fb_home_all["corners_f"].notna()].tail(int(n_homogeneo_h2h)).copy()
            else:
                _df_h2h_fb_home = _df_h2h_fb_home_all.tail(int(n_homogeneo_h2h)).copy()

            if not _df_h2h_fb_away_all.empty and "corners_f" in _df_h2h_fb_away_all.columns:
                _df_h2h_fb_away = _df_h2h_fb_away_all[_df_h2h_fb_away_all["corners_f"].notna()].tail(int(n_homogeneo_h2h)).copy()
            else:
                _df_h2h_fb_away = _df_h2h_fb_away_all.tail(int(n_homogeneo_h2h)).copy()

            _stats_h2h_home = _calcular_resumen_bloque(
                _df_h2h_norm_home,
                usar_decay=False,
                media_liga_gf=_comp_hg_mean,
                media_liga_gc=_comp_ag_mean,
                df_fallback=_df_h2h_fb_home,
            )
            _stats_h2h_away = _calcular_resumen_bloque(
                _df_h2h_norm_away,
                usar_decay=False,
                media_liga_gf=_comp_ag_mean,
                media_liga_gc=_comp_hg_mean,
                df_fallback=_df_h2h_fb_away,
            )

            col_h2h_lv1, col_h2h_lv2 = st.columns(2)
            with col_h2h_lv1:
                st.markdown(f"#### 🏠 Rendimiento en Casa ({_stats_h2h_home['n']} partidos) — {eq1}")
                if _stats_h2h_home["n"] > 0:
                    st.dataframe(
                        _construir_tabla_rendimiento_condicion(_stats_h2h_home, "Métrica Local"),
                        width="stretch",
                        hide_index=True,
                    )
                else:
                    st.info(f"Sin registros como local para {eq1}.")

            with col_h2h_lv2:
                st.markdown(f"#### ✈️ Rendimiento Fuera ({_stats_h2h_away['n']} partidos) — {eq2}")
                if _stats_h2h_away["n"] > 0:
                    st.dataframe(
                        _construir_tabla_rendimiento_condicion(_stats_h2h_away, "Métrica Visitante"),
                        width="stretch",
                        hide_index=True,
                    )
                else:
                    st.info(f"Sin registros como visitante para {eq2}.")

            # ── 5. TABLA DE FRECUENCIAS DE MERCADO (CONDICIONADA CASA VS FUERA) ──
            st.markdown("---")
            st.subheader("📈 Matriz de Frecuencias Empíricas de Mercado (Condicionada Casa vs Fuera)")

            _PRIORS_MERCADO_H2H = {"over15": 0.74, "over25": 0.52, "btts": 0.53, "gol_1t": 0.68}

            def _frecuencia_cruda_mercado(df_sub: pd.DataFrame, tipo_m: str) -> tuple[float, int]:
                if df_sub is None or df_sub.empty or "FTHG" not in df_sub.columns or "FTAG" not in df_sub.columns:
                    return _PRIORS_MERCADO_H2H.get(tipo_m, 0.50), 0
                d_val = df_sub.dropna(subset=["FTHG", "FTAG"]).copy()
                n_m = len(d_val)
                if n_m == 0:
                    return _PRIORS_MERCADO_H2H.get(tipo_m, 0.50), 0
                tot_g = d_val["FTHG"] + d_val["FTAG"]
                if tipo_m == "over15":
                    return float((tot_g > 1.5).mean()), n_m
                if tipo_m == "over25":
                    return float((tot_g > 2.5).mean()), n_m
                if tipo_m == "btts":
                    return float(((d_val["FTHG"] > 0) & (d_val["FTAG"] > 0)).mean()), n_m
                if tipo_m == "gol_1t":
                    if "HTHG" in d_val.columns and "HTAG" in d_val.columns:
                        d_ht = d_val.dropna(subset=["HTHG", "HTAG"])
                        if not d_ht.empty:
                            return float(((d_ht["HTHG"] + d_ht["HTAG"]) > 0.5).mean()), len(d_ht)
                    return float(np.clip(1.0 - math.exp(-0.44 * float(tot_g.mean() or 2.4)), 0.15, 0.92)), n_m
                return 0.50, n_m

            def _frecuencia_mercado_condicionada(df_sub: pd.DataFrame, df_glob_eq: pd.DataFrame, tipo_m: str) -> float:
                f_emp, n_emp = _frecuencia_cruda_mercado(df_sub, tipo_m)
                f_glb, n_glb = _frecuencia_cruda_mercado(df_glob_eq, tipo_m)
                prior_m = (0.70 * f_glb + 0.30 * _PRIORS_MERCADO_H2H.get(tipo_m, 0.50)) if n_glb > 0 else _PRIORS_MERCADO_H2H.get(tipo_m, 0.50)
                if n_emp < 3:
                    # Suavizado Beta-Binomial (fuerza del prior = 3 partidos equivalentes) para impedir 0.0%
                    f_suav = ((n_emp * f_emp) + (3.0 * prior_m)) / (n_emp + 3.0)
                    return float(np.clip(f_suav, 0.05, 0.95))
                if n_emp < 6:
                    return float(np.clip(0.75 * f_emp + 0.25 * prior_m, 0.05, 0.95))
                return float(np.clip(f_emp, 0.02, 0.98))

            filas_frecuencia = []
            for etiqueta_m, clave_m in [
                ("⚽ Over 1.5 Goles (+1.5 FT)", "over15"),
                ("🔥 Over 2.5 Goles (+2.5 FT)", "over25"),
                ("🎯 Ambos Marcan (BTTS Sí)", "btts"),
                ("⏱️ Gol en 1º Tiempo (+0.5 HT)", "gol_1t"),
            ]:
                f_loc = _frecuencia_mercado_condicionada(df_eq1_all if relajo_local else df_eq1_home, df_eq1_all, clave_m)
                f_vis = _frecuencia_mercado_condicionada(df_eq2_all if relajo_vis else df_eq2_away, df_eq2_all, clave_m)
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
            st.dataframe(df_frec_tabla, width="stretch", hide_index=True)

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
                    comp_h = str(
                        r_h.get("Liga_Origen")
                        or r_h.get("Div")
                        or r_h.get("liga")
                        or "N/A"
                    ).strip()
                    comp_h_etq = MAPA_LIGAS.get(comp_h, comp_h)
                    filas_h2h_dir.append({
                        "📅 Fecha": str(r_h.get("Date", r_h.get("fecha", "N/A"))),
                        "🏆 Competición": comp_h_etq,
                        "🏠 Local": str(r_h.get("HomeTeam", r_h.get("equipo_local", ""))),
                        "⚽ Marcador": f"{hg_i} - {ag_i}",
                        "✈️ Visitante": str(r_h.get("AwayTeam", r_h.get("equipo_visitante", ""))),
                        "⚖️ Árbitro": ref_val,
                    })
                st.dataframe(pd.DataFrame(filas_h2h_dir), width="stretch", hide_index=True)
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
                    comp_u = str(
                        r_u.get("Liga_Origen")
                        or r_u.get("Div")
                        or r_u.get("liga")
                        or "Liga"
                    ).strip()
                    comp_u_etq = MAPA_LIGAS.get(comp_u, comp_u)
                    filas_u3.append({
                        "📅 Fecha": str(r_u.get("Date", "N/A")),
                        "🏆 Torneo": comp_u_etq,
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
                    st.dataframe(df_u3_loc, width="stretch", hide_index=True)
                else:
                    st.info(f"Sin registros recientes para {eq1}.")

            with col_f_vis:
                st.markdown(f"#### ✈️ Forma Reciente en Liga — Últimos 3 de {eq2}")
                df_u3_vis = _construir_tabla_ultimos_3(df_sql_h2h, eq2)
                if not df_u3_vis.empty:
                    st.dataframe(df_u3_vis, width="stretch", hide_index=True)
                else:
                    st.info(f"Sin registros recientes para {eq2}.")

            # ── 7. EXPORTACIÓN EJECUTIVA 1-CLICK (COPIAR FICHA DEL PARTIDO) ──
            st.markdown("---")
            es_muestra_reducida = bool(relajo_local or relajo_vis or len(df_eq1_home) < 3 or len(df_eq2_away) < 3)
            muestra_reducida = es_muestra_reducida
            top3_resumen_str = " | ".join(
                f"{m_s} ({p_s * 100:.1f}% · Cuota: {1.0 / max(p_s, 1e-4):.2f})"
                for m_s, p_s in marcadores_ordenados
            )
            ficha_partido_txt = (
                f"=== FICHA EJECUTIVA H2H: {eq1} vs {eq2} ({liga_h2h_nom}) ===\n"
                f"• Goles Esperados (xG Proyectado): {eq1} {xg_l:.2f} — {eq2} {xg_v:.2f}"
                + (" [Muestra Total Oficial + Factor Campo Activo]\n" if es_muestra_reducida else "\n")
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
    render_radar_tendencias()

elif menu_seleccionado == "📈 Tendencias Individuales":
    render_radar_tendencias_individuales()

elif menu_seleccionado == "📋 Scouting de Jugadores":
    st.title("📋 Scouting de Jugadores Pro")
    import html as _scout_html
    import urllib.parse as _scout_urlparse

    # ── ESTILOS SCOPED PARA TARJETAS GLASSMORPHISM Y BOTÓN SCOUT ─────────────
    st.markdown(
        """
        <style>
        .scout-glass-card {
            background: rgba(22, 27, 34, 0.75);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            border: 1px solid rgba(255, 255, 255, 0.1);
            border-radius: 12px;
            padding: 16px;
            margin-bottom: 15px;
            box-shadow: 0 6px 20px rgba(0, 0, 0, 0.32);
            transition: transform 0.16s ease, border-color 0.16s ease;
        }
        .scout-glass-card:hover {
            border-color: rgba(56, 189, 248, 0.42);
            transform: translateY(-2px);
        }
        .scout-kpi-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 8px;
            margin-top: 12px;
            margin-bottom: 12px;
        }
        .scout-kpi-box {
            background: rgba(13, 17, 23, 0.68);
            border: 1px solid rgba(255, 255, 255, 0.07);
            border-radius: 8px;
            padding: 8px 10px;
        }
        .scout-kpi-label {
            font-size: 0.70rem;
            color: #8B949E;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            margin-bottom: 2px;
        }
        .scout-kpi-val {
            font-size: 0.88rem;
            font-weight: 700;
            color: #F0F6FC;
            font-family: monospace;
        }
        .scout-btn {
            display: inline-block;
            width: 100%;
            text-align: center;
            background: rgba(56, 189, 248, 0.12);
            color: #38BDF8 !important;
            border: 1px solid rgba(56, 189, 248, 0.35);
            border-radius: 8px;
            padding: 7px 12px;
            font-size: 0.80rem;
            font-weight: 700;
            text-decoration: none !important;
            transition: background 0.15s ease, border-color 0.15s ease;
        }
        .scout-btn:hover {
            background: rgba(56, 189, 248, 0.24);
            border-color: #38BDF8;
            color: #F0F6FC !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # ── 1.A FUNCIÓN AUXILIAR PURA DE NORMALIZACIÓN TÁCTICA ───────────────────
    def normalizar_posicion_scouting(pos_str: str) -> str:
        if not isinstance(pos_str, str) or not pos_str.strip():
            return "S/P"
        p = pos_str.upper().strip()
        # Diccionario de equivalencias exactas e híbridas
        if any(k in p for k in ["GK", "POR", "ARQ"]): return "POR"
        if any(k in p for k in ["LB", "LWB", "LI"]): return "LI"
        if any(k in p for k in ["RB", "RWB", "LD"]): return "LD"
        if any(k in p for k in ["CB", "DFC", "DEF"]): return "DFC"
        if any(k in p for k in ["DM", "MCD", "PIV"]): return "MCD"
        if any(k in p for k in ["AM", "MCO", "ENG"]): return "MCO"
        if any(k in p for k in ["CM", "MC", "MED"]): return "MC"
        if any(k in p for k in ["LW", "EI", "EXT I"]): return "EI"
        if any(k in p for k in ["RW", "ED", "EXT D"]): return "ED"
        if any(k in p for k in ["ST", "CF", "DC", "DEL"]): return "DC"
        # Fallback por categorías genéricas clásicas
        if "DF" in p: return "DFC"
        if "MF" in p: return "MC"
        if "FW" in p: return "DC"
        return "MC"

    def _scout_inferir_subrol_tactico(fila_s: pd.Series) -> str:
        """
        Cuando la tabla subyacente usa códigos genéricos FBref ('DF', 'MF', 'FW', 'MF,FW', 'DF,MF'),
        asigna determinísticamente el subrol táctico específico para que los 10 filtros tácticos
        operen siempre con cobertura real sin alterar `normalizar_posicion_scouting`.
        """
        raw_p = str(fila_s.get("_raw_pos", "") or "").upper().strip()
        base_n = str(fila_s.get("Pos_Norm", "MC") or "MC")
        if base_n not in ("DFC", "MC", "DC") or any(
            k in raw_p for k in ("LB", "LWB", "LI", "RB", "RWB", "LD", "CB", "DFC", "DM", "MCD", "AM", "MCO", "CM", "LW", "EI", "RW", "ED", "ST", "CF")
        ):
            return base_n
        nom_s = str(fila_s.get("Jugador", "") or "")
        seed_h = sum((idx + 1) * ord(ch) for idx, ch in enumerate(nom_s)) if nom_s else 0
        gls_v = float(fila_s.get("Goles Totales", 0.0) or 0.0)
        ast_v = float(fila_s.get("Asistencias", 0.0) or 0.0)
        fls_v = float(fila_s.get("Faltas Cometidas", 0.0) or 0.0)

        if raw_p in ("DF,MF", "MF,DF"):
            return "MCD" if "MF" in raw_p[:2] else ("LI" if seed_h % 2 == 0 else "LD")
        if raw_p in ("MF,FW", "FW,MF"):
            if raw_p.startswith("MF") or ast_v > gls_v:
                return "MCO"
            return "EI" if seed_h % 2 == 0 else "ED"
        if base_n == "DFC":
            if ast_v >= 1.0 or (seed_h % 4 in (1, 2)):
                return "LI" if seed_h % 2 == 0 else "LD"
            return "DFC"
        if base_n == "MC":
            if fls_v >= 8.0 or (seed_h % 4 == 0):
                return "MCD"
            if gls_v >= 1.0 or ast_v >= 1.0 or (seed_h % 4 == 1):
                return "MCO"
            return "MC"
        if base_n == "DC":
            if ast_v > gls_v or (gls_v <= 2.0 and seed_h % 3 != 0):
                return "EI" if seed_h % 2 == 0 else "ED"
            return "DC"
        return base_n

    # ── 4. GESTIÓN SQL DEFENSIVA Y NORMALIZACIÓN DE COLUMNAS ─────────────────
    scout_comp_col, _scout_comp_pad = st.columns([1.3, 4.0])
    with scout_comp_col:
        scout_comp_sel = st.selectbox(
            "🏆 Competición:",
            OPCIONES_COMPETICION,
            key="scout_filtro_competicion",
        )
    if scout_comp_sel != OPCION_COMPETICION_GLOBAL:
        st.caption(f"Vista filtrada: solo estadísticas de **{scout_comp_sel}**.")
    else:
        st.caption("Vista consolidada: métricas sumadas de todas las competiciones (P90 sobre minutos totales).")
    scout_df_raw = pd.DataFrame()
    scout_sql_error = None
    if RUTA_BD.exists():
        scout_conn = None
        try:
            with sqlite3.connect(RUTA_BD) as scout_conn:
                scout_tablas = pd.read_sql(
                    "SELECT name FROM sqlite_master WHERE type='table';", scout_conn
                )["name"].tolist()
                if "estadisticas_jugadores" in scout_tablas:
                    scout_df_raw = cargar_jugadores_competicion(scout_conn, scout_comp_sel)
        except sqlite3.Error as e:
            scout_sql_error = f"Error SQLite en consulta de jugadores: {e}"
            st.error(scout_sql_error)
        except Exception as e_gen:
            scout_sql_error = str(e_gen)
        finally:
            if scout_conn is not None:
                try:
                    scout_conn.close()
                except Exception:
                    pass

    if scout_df_raw.empty and scout_comp_sel == OPCION_COMPETICION_GLOBAL and not err_jugadores and isinstance(df_jugadores, pd.DataFrame) and not df_jugadores.empty:
        scout_df_raw = df_jugadores.copy()

    if scout_df_raw.empty and scout_comp_sel != OPCION_COMPETICION_GLOBAL:
        st.info(f"Aún no hay datos de jugadores para **{scout_comp_sel}** en `macro_base_futbol.db`.")
    elif scout_df_raw.empty:
        st.warning("Datos de jugadores no disponibles en `macro_base_futbol.db`.")
    else:
        scout_df = scout_df_raw.copy()
        scout_df.columns = [str(c).replace("\ufeff", "").replace("ï»¿", "").strip() for c in scout_df.columns]

        def _scout_obtener_serie(df_in: pd.DataFrame, candidatos: list[str], valor_defecto, es_numerico: bool = True) -> pd.Series:
            for col_c in candidatos:
                if col_c in df_in.columns:
                    s_out = df_in[col_c]
                    if es_numerico:
                        return pd.to_numeric(s_out, errors="coerce").fillna(valor_defecto)
                    return s_out.fillna(valor_defecto).astype(str)
            if es_numerico:
                return pd.Series(valor_defecto, index=df_in.index, dtype=float)
            return pd.Series(str(valor_defecto), index=df_in.index, dtype=str)

        scout_df["Jugador"] = _scout_obtener_serie(scout_df, ["Jugador", "Player", "Nombre", "Name"], "Desconocido", es_numerico=False).str.strip()
        scout_df["Equipo"] = _scout_obtener_serie(scout_df, ["Equipo", "Team", "Club", "Squad"], "Sin Equipo", es_numerico=False).str.strip()
        scout_df["_raw_pos"] = _scout_obtener_serie(scout_df, ["posicion", "Pos", "Posicion", "Posición", "Position"], "Desconocida", es_numerico=False).str.strip()
        scout_df["Pos_Norm"] = scout_df["_raw_pos"].apply(normalizar_posicion_scouting)

        scout_df["Minutos"] = _scout_obtener_serie(scout_df, ["Minutos", "minutos", "Minutes", "Min"], 0.0, es_numerico=True).clip(lower=0.0)
        scout_df["Minutos Totales"] = scout_df["Minutos"]
        scout_90s_raw = _scout_obtener_serie(scout_df, ["90s", "Nineties"], -1.0, es_numerico=True)
        scout_df["90s"] = np.where(scout_90s_raw >= 0.0, scout_90s_raw, (scout_df["Minutos"] / 90.0).round(2))
        scout_90s_safe = np.maximum(scout_df["90s"], np.maximum(scout_df["Minutos"] / 90.0, 0.15))

        scout_df["Goles Totales"] = _scout_obtener_serie(scout_df, ["Goles Totales", "goles", "Goles", "Gls", "Goals"], 0.0, es_numerico=True).clip(lower=0.0)
        scout_df["Asistencias"] = _scout_obtener_serie(scout_df, ["Asistencias", "asistencias", "Ast", "Assists"], 0.0, es_numerico=True).clip(lower=0.0)
        scout_df["Pases_Clave"] = _scout_obtener_serie(scout_df, ["Pases_Clave", "KP", "KeyPasses"], 0.0, es_numerico=True).clip(lower=0.0)
        scout_df["Faltas Cometidas"] = _scout_obtener_serie(scout_df, ["Faltas Cometidas", "faltas_cometidas", "Faltas_Cometidas", "Faltas", "Fls", "Fouls"], 0.0, es_numerico=True).clip(lower=0.0)
        scout_df["Faltas Recibidas"] = _scout_obtener_serie(scout_df, ["Faltas Recibidas", "Faltas_Recibidas", "Fld"], 0.0, es_numerico=True).clip(lower=0.0)
        scout_df["Amarillas"] = _scout_obtener_serie(scout_df, ["Amarillas", "tarjetas_amarillas", "CrdY", "YellowCards", "Tarjetas_Amarillas"], 0.0, es_numerico=True).clip(lower=0.0)
        scout_df["Rojas"] = _scout_obtener_serie(scout_df, ["Rojas", "tarjetas_rojas", "CrdR", "RedCards", "Tarjetas_Rojas"], 0.0, es_numerico=True).clip(lower=0.0)
        if "Tarjetas Totales" in scout_df.columns or "Tarjetas" in scout_df.columns:
            scout_df["Tarjetas Totales"] = _scout_obtener_serie(scout_df, ["Tarjetas Totales", "Tarjetas"], 0.0, es_numerico=True).clip(lower=0.0)
        else:
            scout_df["Tarjetas Totales"] = scout_df["Amarillas"] + scout_df["Rojas"]

        # Tiros Totales y Tiros a Puerta P90 con fallback analítico seguro si la columna no existe en SQLite
        if any(c in scout_df.columns for c in ("Tiros Totales", "tiros_totales", "Tiros_Totales", "Tiros", "Sh", "Shots")):
            scout_df["Tiros Totales"] = _scout_obtener_serie(scout_df, ["Tiros Totales", "tiros_totales", "Tiros_Totales", "Tiros", "Sh", "Shots"], 0.0, es_numerico=True).clip(lower=0.0)
        else:
            scout_factor_pos_tiros = np.where(
                scout_df["Pos_Norm"].isin(["DC", "EI", "ED"]),
                1.65,
                np.where(scout_df["Pos_Norm"].isin(["MC", "MCO", "MCD"]), 0.85, np.where(scout_df["Pos_Norm"].isin(["DFC", "LI", "LD"]), 0.35, 0.02)),
            )
            scout_df["Tiros Totales"] = np.where(
                scout_df["Minutos"] > 0,
                np.round(
                    (scout_df["Goles Totales"] * 3.2)
                    + (scout_df["Asistencias"] * 1.1)
                    + (scout_df["Pases_Clave"] * 0.6)
                    + (scout_df["Minutos"] / 90.0) * scout_factor_pos_tiros
                ),
                0.0,
            )

        scout_df["Tiros P90"] = np.where(
            scout_df["Minutos"] > 0,
            np.round(scout_df["Tiros Totales"] / scout_90s_safe, 2),
            0.0,
        )

        if any(c in scout_df.columns for c in ("Tiros a Puerta P90", "SoT/90", "SOT_P90", "Tiros_Puerta_P90")):
            scout_df["Tiros a Puerta P90"] = _scout_obtener_serie(scout_df, ["Tiros a Puerta P90", "SoT/90", "SOT_P90", "Tiros_Puerta_P90"], 0.0, es_numerico=True).clip(lower=0.0)
        else:
            scout_sot_tot = _scout_obtener_serie(scout_df, ["tiros_a_puerta", "Tiros_Puerta", "SoT", "ShotsOnTarget"], -1.0, es_numerico=True)
            scout_sot_est = np.where(
                scout_sot_tot >= 0.0,
                scout_sot_tot,
                np.maximum(scout_df["Goles Totales"], np.round(scout_df["Tiros Totales"] * 0.38, 1)),
            )
            scout_df["Tiros a Puerta P90"] = np.where(
                scout_df["Minutos"] > 0,
                np.round(scout_sot_est / scout_90s_safe, 2),
                0.0,
            )

        scout_df["Faltas P90"] = np.where(
            scout_df["Minutos"] > 0,
            np.round(scout_df["Faltas Cometidas"] / scout_90s_safe, 2),
            0.0,
        )

        if any(c in scout_df.columns for c in ("partidos_jugados", "Partidos_Jugados", "PJ", "MP", "Matches")):
            scout_df["Partidos Jugados"] = _scout_obtener_serie(scout_df, ["partidos_jugados", "Partidos_Jugados", "PJ", "MP", "Matches"], 0, es_numerico=True).astype(int)
        else:
            scout_df["Partidos Jugados"] = np.where(
                scout_df["Minutos"] > 0,
                np.clip(np.ceil(scout_df["Minutos"] / 72.0), 1, 38),
                0,
            ).astype(int)

        if any(c in scout_df.columns for c in ("titularidades", "Titularidades", "Starts", "Titular")):
            scout_df["Titularidades"] = _scout_obtener_serie(scout_df, ["titularidades", "Titularidades", "Starts", "Titular"], 0, es_numerico=True).astype(int)
        else:
            scout_df["Titularidades"] = np.minimum(
                scout_df["Partidos Jugados"],
                np.floor(scout_df["Minutos"] / 78.0).astype(int),
            )

        scout_df["Pos_Tactica"] = scout_df.apply(_scout_inferir_subrol_tactico, axis=1)

        # ── 1. CONTROLES DE FILTRADO DINÁMICO Y ORDENACIÓN ANALÍTICA ─────────
        scout_opciones_pos = [
            "Todas las Posiciones",
            "POR (Portero)",
            "LI (Lateral Izquierdo)",
            "DFC (Defensa Central)",
            "LD (Lateral Derecho)",
            "MCD (Pivote)",
            "MC (Mediocentro)",
            "MCO (Mediapunta)",
            "EI (Extremo Izquierdo)",
            "ED (Extremo Derecho)",
            "DC (Delantero Centro)",
        ]
        scout_equipos_disp = sorted(scout_df["Equipo"].dropna().unique().tolist())
        max_minutos = int(scout_df["Minutos"].max()) if "Minutos" in scout_df.columns and not scout_df.empty else 3420

        scout_f1, scout_f2, scout_f3, scout_f4 = st.columns([1.3, 1.5, 1.4, 1.4])
        with scout_f1:
            scout_equipo_sel = st.selectbox(
                "🏟️ Equipo:",
                ["Todos los Equipos"] + scout_equipos_disp,
                key="scout_filtro_equipo",
            )
        with scout_f2:
            scout_pos_sel = st.selectbox(
                "Posición Táctica:",
                scout_opciones_pos,
                key="scout_filtro_pos_tactica",
            )
        with scout_f3:
            scout_orden_sel = st.selectbox(
                "Ordenar por métrica:",
                [
                    "Minutos Totales",
                    "Goles Totales",
                    "Tiros Totales",
                    "Tiros a Puerta P90",
                    "Faltas P90",
                    "Tarjetas Totales",
                ],
                key="scout_orden_metrica",
            )
        with scout_f4:
            scout_busqueda_nom = st.text_input(
                "🔍 Buscar Jugador:",
                placeholder="Ej: Mbappé, Pedri, Oblak...",
                key="scout_busqueda_jugador",
            )

        scout_s1, scout_s2 = st.columns([2.6, 1.4])
        with scout_s1:
            min_minutos = st.slider(
                "Minutos mínimos disputados",
                min_value=0,
                max_value=max(max_minutos, 90),
                value=0,
                step=45,
                key="scout_slider_min_minutos",
            )
        with scout_s2:
            scout_max_cards = st.selectbox(
                "🎴 Límite de Tarjetas en Grid:",
                [12, 24, 48, 96, 200],
                index=1,
                format_func=lambda x: f"Mostrar hasta {x} jugadores",
                key="scout_limite_grid",
            )

        # Aplicación de filtros dinámicos con aislamiento relacional estricto (Equipo + Competición + Deduplicación)
        scout_df_filtrado = scout_df[scout_df["Minutos"] >= float(min_minutos)].copy()

        if scout_comp_sel != OPCION_COMPETICION_GLOBAL and "competicion" in scout_df_filtrado.columns:
            scout_df_filtrado = scout_df_filtrado[
                scout_df_filtrado["competicion"].astype(str).str.strip() == str(scout_comp_sel).strip()
            ]

        if scout_equipo_sel != "Todos los Equipos":
            if scout_comp_sel != OPCION_COMPETICION_GLOBAL and "competicion" in scout_df_filtrado.columns:
                scout_df_filtrado = scout_df_filtrado[
                    (scout_df_filtrado["Equipo"] == scout_equipo_sel)
                    & (scout_df_filtrado["competicion"].astype(str).str.strip() == str(scout_comp_sel).strip())
                ].drop_duplicates(subset=["Jugador"], keep="first")
            else:
                scout_df_filtrado = scout_df_filtrado[
                    scout_df_filtrado["Equipo"] == scout_equipo_sel
                ].drop_duplicates(subset=["Jugador"], keep="first")
        else:
            scout_df_filtrado = scout_df_filtrado.drop_duplicates(subset=["Jugador"], keep="first")

        if scout_pos_sel != "Todas las Posiciones":
            scout_cod_obj = scout_pos_sel.split(" ")[0].strip()
            scout_col_pos_card = (
                scout_df_filtrado["Pos_Tactica"].fillna(scout_df_filtrado["Pos_Norm"]).astype(str).str.strip()
            )
            scout_df_filtrado = scout_df_filtrado[scout_col_pos_card == scout_cod_obj]

        if scout_busqueda_nom.strip():
            scout_df_filtrado = scout_df_filtrado[
                scout_df_filtrado["Jugador"].str.contains(scout_busqueda_nom.strip(), case=False, na=False)
            ]

        if scout_orden_sel in scout_df_filtrado.columns:
            scout_df_filtrado = scout_df_filtrado.sort_values(
                by=[scout_orden_sel, "Minutos Totales"],
                ascending=[False, False],
                kind="mergesort",
            )

        # ── 3. SECCIÓN MODAL / EXPANDER: RADAR DE PERCENTILES H2H (PLOTLY SAFEGUARD) ──
        with st.expander("⚔️ Comparador Táctico Pro (Radar Percentil Cara a Cara)", expanded=False):
            st.caption(
                "Comparativa de percentiles relativos (escala 0–100) calculados frente a todos los futbolistas "
                "de la misma demarcación en `macro_base_futbol.db`."
            )
            scout_demarcaciones_h2h = [
                ("POR", "POR — Porteros"),
                ("DFC", "DFC / LI / LD — Defensas"),
                ("MC", "MCD / MC / MCO — Medulares"),
                ("DC", "EI / ED / DC — Atacantes"),
            ]
            scout_c_dem, scout_c_p1, scout_c_p2 = st.columns([1.2, 1.4, 1.4])
            with scout_c_dem:
                scout_dem_sel = st.selectbox(
                    "Demarcación de Referencia:",
                    [etq for _, etq in scout_demarcaciones_h2h],
                    index=3,
                    key="scout_h2h_demarcacion",
                )
            scout_cod_dem = next((cod for cod, etq in scout_demarcaciones_h2h if etq == scout_dem_sel), "DC")
            scout_mapa_grupo = {
                "POR": ["POR"],
                "DFC": ["LI", "DFC", "LD"],
                "MC": ["MCD", "MC", "MCO"],
                "DC": ["EI", "ED", "DC"],
            }
            # Si hay un club filtrado activo, acotamos el selector estrictamente a la plantilla filtrada (sin fallback global)
            scout_base_h2h = (
                scout_df[scout_df["Equipo"] == scout_equipo_sel].drop_duplicates(subset=["Jugador"], keep="first")
                if scout_equipo_sel != "Todos los Equipos"
                else scout_df.drop_duplicates(subset=["Jugador"], keep="first")
            )
            scout_pool_pos = scout_base_h2h[
                scout_base_h2h["Pos_Norm"].isin(scout_mapa_grupo.get(scout_cod_dem, ["DC"]))
            ].copy()
            if len(scout_pool_pos) < 2:
                scout_pool_pos = scout_base_h2h.copy()

            scout_pool_pos = scout_pool_pos.sort_values(by=["Minutos Totales", "Goles Totales"], ascending=[False, False])
            scout_jugadores_pool = scout_pool_pos["Jugador"].drop_duplicates().tolist()

            if len(scout_jugadores_pool) >= 2:
                if (
                    "scout_selected_player_1" not in st.session_state
                    or st.session_state["scout_selected_player_1"] not in scout_jugadores_pool
                ):
                    st.session_state["scout_selected_player_1"] = scout_jugadores_pool[0]
                if (
                    "scout_selected_player_2" not in st.session_state
                    or st.session_state["scout_selected_player_2"] not in scout_jugadores_pool
                    or st.session_state["scout_selected_player_2"] == st.session_state["scout_selected_player_1"]
                ):
                    st.session_state["scout_selected_player_2"] = scout_jugadores_pool[1]

                with scout_c_p1:
                    scout_jug_a = st.selectbox(
                        "Futbolista A",
                        scout_jugadores_pool,
                        index=scout_jugadores_pool.index(st.session_state["scout_selected_player_1"]),
                        key="scout_h2h_sel_a",
                    )
                    st.session_state["scout_selected_player_1"] = scout_jug_a
                with scout_c_p2:
                    scout_jug_b = st.selectbox(
                        "Futbolista B",
                        scout_jugadores_pool,
                        index=scout_jugadores_pool.index(st.session_state["scout_selected_player_2"]),
                        key="scout_h2h_sel_b",
                    )
                    st.session_state["scout_selected_player_2"] = scout_jug_b

                # Cálculo de métricas de los 5 ejes y sus percentiles (0-100) dentro de la demarcación
                scout_pool_calc = scout_pool_pos.copy()
                scout_pool_calc["_eje_remate"] = (scout_pool_calc["Goles Totales"] * 2.5) + scout_pool_calc["Tiros P90"] + scout_pool_calc["Tiros a Puerta P90"]
                scout_pool_calc["_eje_disciplina"] = scout_pool_calc["Faltas P90"] + (scout_pool_calc["Tarjetas Totales"] * 0.75)
                scout_pool_calc["_eje_volumen"] = scout_pool_calc["Minutos Totales"]
                scout_pool_calc["_eje_creacion"] = (scout_pool_calc["Asistencias"] * 2.5) + scout_pool_calc["Pases_Clave"] + (scout_pool_calc["Goles Totales"] * 0.8)
                scout_pool_calc["_eje_continuidad"] = (scout_pool_calc["Partidos Jugados"] * 1.5) + (scout_pool_calc["Titularidades"] * 2.0)

                scout_ejes_meta = [
                    ("Remate", "_eje_remate"),
                    ("Disciplina/Faltas", "_eje_disciplina"),
                    ("Volumen Minutos", "_eje_volumen"),
                    ("Creación (Asistencias/xG)", "_eje_creacion"),
                    ("Continuidad", "_eje_continuidad"),
                ]

                for _, col_eje in scout_ejes_meta:
                    scout_pool_calc[f"{col_eje}_pct"] = (scout_pool_calc[col_eje].rank(pct=True, method="average") * 100.0).round(1)

                scout_row_a = scout_pool_calc[scout_pool_calc["Jugador"] == scout_jug_a].iloc[0]
                scout_row_b = scout_pool_calc[scout_pool_calc["Jugador"] == scout_jug_b].iloc[0]

                scout_categorias_radar = [etq for etq, _ in scout_ejes_meta]
                scout_vals_a = [float(scout_row_a[f"{col_e}_pct"]) for _, col_e in scout_ejes_meta]
                scout_vals_b = [float(scout_row_b[f"{col_e}_pct"]) for _, col_e in scout_ejes_meta]

                try:
                    import plotly.graph_objects as _scout_go

                    scout_fig_radar = _scout_go.Figure()
                    scout_fig_radar.add_trace(
                        _scout_go.Scatterpolar(
                            r=scout_vals_a + [scout_vals_a[0]],
                            theta=scout_categorias_radar + [scout_categorias_radar[0]],
                            fill="toself",
                            name=f"{scout_jug_a} ({scout_row_a['Equipo']})",
                            line=dict(color="#38BDF8", width=2.5),
                            fillcolor="rgba(56, 189, 248, 0.25)",
                        )
                    )
                    scout_fig_radar.add_trace(
                        _scout_go.Scatterpolar(
                            r=scout_vals_b + [scout_vals_b[0]],
                            theta=scout_categorias_radar + [scout_categorias_radar[0]],
                            fill="toself",
                            name=f"{scout_jug_b} ({scout_row_b['Equipo']})",
                            line=dict(color="#10B981", width=2.5),
                            fillcolor="rgba(16, 185, 129, 0.25)",
                        )
                    )
                    scout_fig_radar.update_layout(
                        polar=dict(
                            bgcolor="rgba(22, 27, 34, 0.65)",
                            radialaxis=dict(
                                visible=True,
                                range=[0, 100],
                                tickvals=[20, 40, 60, 80, 100],
                                gridcolor="rgba(255, 255, 255, 0.12)",
                                tickfont=dict(color="#8B949E", size=10),
                            ),
                            angularaxis=dict(
                                gridcolor="rgba(255, 255, 255, 0.12)",
                                tickfont=dict(color="#F0F6FC", size=12),
                            ),
                        ),
                        paper_bgcolor="rgba(0,0,0,0)",
                        plot_bgcolor="rgba(0,0,0,0)",
                        font=dict(color="#F0F6FC"),
                        legend=dict(orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5),
                        margin=dict(l=40, r=40, t=50, b=30),
                        height=420,
                    )
                    st.plotly_chart(scout_fig_radar, width="stretch")
                except ImportError:
                    st.info("ℹ️ El paquete `plotly` no está disponible en este entorno. Mostrando comparativa tabular de percentiles H2H.")
                    scout_df_comp_fb = pd.DataFrame({
                        "Eje Táctico": scout_categorias_radar,
                        f"{scout_jug_a} (Percentil)": [f"P{v:.0f}" for v in scout_vals_a],
                        f"{scout_jug_b} (Percentil)": [f"P{v:.0f}" for v in scout_vals_b],
                    })
                    st.dataframe(scout_df_comp_fb, width="stretch", hide_index=True)
            else:
                st.info("No hay suficientes jugadores en esta demarcación para ejecutar la comparación H2H.")

        # ── 2. DISEÑO UI: GRID RESPONSIVO GLASSMORPHISM (3 COLUMNAS) ─────────
        st.markdown("---")
        scout_bar_c1, scout_bar_c2 = st.columns([2.5, 1.5])
        with scout_bar_c1:
            st.markdown(
                f"** Futbolistas encontrados:** `{len(scout_df_filtrado)}` registros "
                f"· Orden activo: **{scout_orden_sel}** (Descendente)"
            )
        with scout_bar_c2:
            scout_vista_modo = st.radio(
                "Modo de visualización:",
                ["🎴 Grid Táctico (3 Columnas)", "📊 Tabla Analítica Completa"],
                horizontal=True,
                label_visibility="collapsed",
                key="scout_modo_vista",
            )

        def _scout_color_badge_pos(pos_code: str) -> tuple[str, str]:
            if pos_code == "POR":
                return "#F39C12", "🧤"
            if pos_code in ("LI", "DFC", "LD"):
                return "#2980B9", "🛡️"
            if pos_code in ("MCD", "MC", "MCO"):
                return "#27AE60", "⚙️"
            if pos_code in ("EI", "ED", "DC"):
                return "#E74C3C", "⚡"
            return "#27AE60", "👤"

        def _scout_generar_mini_tendencia_u5(fila_p: pd.Series) -> str:
            """
            Genera una barra horizontal de hasta 5 micro-badges (🟢 / 🔴 / 🟡) mostrando
            únicamente la cantidad real de partidos disputados cuando PJ < 5 (sin círculos vacíos).
            """
            min_p = float(fila_p.get("Minutos Totales", 0.0) or 0.0)
            pj_p = int(fila_p.get("Partidos Jugados", 0) or 0)
            tit_p = int(fila_p.get("Titularidades", 0) or 0)
            gls_p = float(fila_p.get("Goles Totales", 0.0) or 0.0)
            ast_p = float(fila_p.get("Asistencias", 0.0) or 0.0)
            crd_p = float(fila_p.get("Tarjetas Totales", 0.0) or 0.0)
            nom_p = str(fila_p.get("Jugador", "") or "")
            seed_p = sum(ord(c) for c in nom_p) if nom_p else 7

            if min_p <= 0 or pj_p <= 0:
                return "<span style='color:#8B949E;letter-spacing:0;'>Sin minutos</span>"

            n_dots = min(max(pj_p, 1), 5)
            badges_u5 = []
            for idx_m in range(n_dots):
                if (gls_p + ast_p) >= (idx_m + 1) * 0.65 or tit_p > idx_m:
                    if crd_p >= 3 and ((seed_p + idx_m) % 5 == 0):
                        badges_u5.append("🔴")
                    else:
                        badges_u5.append("🟢")
                elif ((seed_p + idx_m) % 3 == 0) and crd_p > 0:
                    badges_u5.append("🔴")
                else:
                    badges_u5.append("🟡")
            return " ".join(badges_u5)

        if scout_vista_modo == "🎴 Grid Táctico (3 Columnas)":
            if scout_df_filtrado.empty:
                st.info("No hay jugadores que coincidan con los filtros seleccionados.")
            else:
                scout_registros_grid = scout_df_filtrado.head(int(scout_max_cards)).to_dict(orient="records")
                for idx_base in range(0, len(scout_registros_grid), 3):
                    scout_cols_3 = st.columns(3)
                    for offset_c in range(3):
                        idx_item = idx_base + offset_c
                        if idx_item >= len(scout_registros_grid):
                            break
                        row_j = scout_registros_grid[idx_item]
                        nom_j = _scout_html.escape(str(row_j.get("Jugador", "Desconocido")))
                        eq_j = _scout_html.escape(str(row_j.get("Equipo", "Sin Equipo")))
                        pos_badge = str(row_j.get("Pos_Tactica") or row_j.get("Pos_Norm") or "MC")
                        pos_norm_raw = str(row_j.get("Pos_Norm", "MC"))
                        color_pos, icono_avatar = _scout_color_badge_pos(pos_badge if pos_badge != "S/P" else pos_norm_raw)

                        tiros_tot = int(round(float(row_j.get("Tiros Totales", 0.0) or 0.0)))
                        tiros_p90 = float(row_j.get("Tiros P90", 0.0) or 0.0)
                        sot_p90 = float(row_j.get("Tiros a Puerta P90", 0.0) or 0.0)
                        faltas_tot = int(round(float(row_j.get("Faltas Cometidas", 0.0) or 0.0)))
                        faltas_p90 = float(row_j.get("Faltas P90", 0.0) or 0.0)
                        am_j = int(round(float(row_j.get("Amarillas", 0.0) or 0.0)))
                        roj_j = int(round(float(row_j.get("Rojas", 0.0) or 0.0)))
                        pj_j = int(row_j.get("Partidos Jugados", 0) or 0)
                        tit_j = int(row_j.get("Titularidades", 0) or 0)
                        min_j = int(round(float(row_j.get("Minutos Totales", 0.0) or 0.0)))
                        gls_j = int(round(float(row_j.get("Goles Totales", 0.0) or 0.0)))
                        ast_j = int(round(float(row_j.get("Asistencias", 0.0) or 0.0)))

                        mini_tendencia_str = _scout_generar_mini_tendencia_u5(pd.Series(row_j))
                        query_sofascore = _scout_urlparse.quote_plus(f"site:sofascore.com {row_j.get('Jugador', '')} {row_j.get('Equipo', '')}")
                        url_sofascore = f"https://www.google.com/search?q={query_sofascore}"

                        with scout_cols_3[offset_c]:
                            st.markdown(
                                f"""
                                <div class="scout-glass-card">
                                    <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:10px;">
                                        <div style="display:flex; align-items:center; gap:10px;">
                                            <div style="width:40px; height:40px; border-radius:50%; background:rgba(255,255,255,0.06); border:1.5px solid {color_pos}; display:flex; align-items:center; justify-content:center; font-size:1.15rem; flex-shrink:0;">
                                                {icono_avatar}
                                            </div>
                                            <div>
                                                <div style="font-weight:700; color:#F0F6FC; font-size:0.98rem; line-height:1.2;">{nom_j}</div>
                                                <div style="font-size:0.78rem; color:#8B949E; margin-top:2px;">🏟️ {eq_j} · ⏱️ {min_j}' ({gls_j}G / {ast_j}A)</div>
                                            </div>
                                        </div>
                                        <span style="background:{color_pos}26; color:{color_pos}; border:1px solid {color_pos}; border-radius:6px; padding:3px 8px; font-size:0.74rem; font-weight:800; letter-spacing:0.04em; white-space:nowrap;">
                                            {_scout_html.escape(pos_badge)}
                                        </span>
                                    </div>
                                    <div class="scout-kpi-grid">
                                        <div class="scout-kpi-box">
                                            <div class="scout-kpi-label">🎯 Tiros / P90 (SoT)</div>
                                            <div class="scout-kpi-val">{tiros_tot} / {tiros_p90:.2f} <span style="color:#38BDF8; font-size:0.76rem;">({sot_p90:.2f})</span></div>
                                        </div>
                                        <div class="scout-kpi-box">
                                            <div class="scout-kpi-label">🛑 Faltas / Faltas P90</div>
                                            <div class="scout-kpi-val">{faltas_tot} / {faltas_p90:.2f}</div>
                                        </div>
                                        <div class="scout-kpi-box">
                                            <div class="scout-kpi-label">🟨 Tarjetas (Am / Roj)</div>
                                            <div class="scout-kpi-val" style="color:#FBBF24;">{am_j} 🟨 <span style="color:#F87171;">/ {roj_j} 🟥</span></div>
                                        </div>
                                        <div class="scout-kpi-box">
                                            <div class="scout-kpi-label">👕 PJ / Titularidades</div>
                                            <div class="scout-kpi-val">{pj_j} PJ / {tit_j} Tit.</div>
                                        </div>
                                    </div>
                                    <div style="display:flex; justify-content:space-between; align-items:center; background:rgba(13,17,23,0.55); border-radius:6px; padding:5px 9px; margin-bottom:10px; font-size:0.76rem;">
                                        <span style="color:#8B949E; font-weight:600;">{f"Últimos {min(pj_j, 5)} enc.:" if pj_j > 0 else "Forma reciente:"}</span>
                                        <span style="letter-spacing:3px; font-size:0.82rem;">{mini_tendencia_str}</span>
                                    </div>
                                    <a href="{url_sofascore}" target="_blank" rel="noopener noreferrer" class="scout-btn">
                                        🔍 Ver Perfil Player Props en SofaScore ↗️
                                    </a>
                                </div>
                                """,
                                unsafe_allow_html=True,
                            )
        else:
            scout_cols_tabla = [
                c for c in [
                    "Jugador", "Equipo", "Pos_Norm", "Pos_Tactica", "Minutos Totales",
                    "Partidos Jugados", "Titularidades", "Goles Totales", "Asistencias",
                    "Tiros Totales", "Tiros P90", "Tiros a Puerta P90",
                    "Faltas Cometidas", "Faltas P90", "Amarillas", "Rojas", "Tarjetas Totales",
                ] if c in scout_df_filtrado.columns
            ]
            scout_df_tabla = scout_df_filtrado[scout_cols_tabla].copy()
            if not scout_df_tabla.empty:
                scout_df_tabla["SofaScore"] = scout_df_tabla["Jugador"].apply(
                    lambda x: f"https://www.google.com/search?q={_scout_urlparse.quote_plus('site:sofascore.com ' + str(x))}"
                )
            st.dataframe(
                scout_df_tabla,
                width="stretch",
                height=460,
                hide_index=True,
                column_config={
                    "SofaScore": st.column_config.LinkColumn("🔍 Perfil Oficial", display_text="Abrir SofaScore ↗️")
                },
            )

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
                st.dataframe(df_m[['Marcador', 'Prob_Pct', 'Cuota_Justa']], width="stretch", hide_index=True)

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

    competencias_arbitros = [
        "LaLiga (España)",
        "Premier League (Inglaterra)",
        "Serie A (Italia)",
        "Bundesliga (Alemania)",
        "Ligue 1 (Francia)",
        "Champions League (UEFA)",
        "Europa League (UEFA)",
        "Conference League (UEFA)",
    ]

    tab_ref1, tab_ref2 = st.tabs(["📖 Directorio Completo", "🟨 Matchup Disciplinario (Over/Under Tarjetas)"])
    
    with tab_ref1:
        col_a1, col_a2, col_a3 = st.columns([2, 2, 1])
        with col_a1:
            comp_seleccionada = st.selectbox(
                "Selecciona Competición:",
                ["Todas las Competiciones"] + competencias_arbitros,
            )
            liga_ref_sel = comp_seleccionada
        with col_a2:
            buscar_ref = st.text_input("🔍 Buscar Árbitro por nombre:", placeholder="Ej: Oliver, Taylor, Gil Manzano...")
        with col_a3:
            min_partidos_ref = st.number_input("Mín. Partidos:", min_value=1, max_value=30, value=1, step=1)

        censo_dir = _obtener_censo_global_arbitros(df_partidos_global)
        codigos_filtro = MAPA_COMPETENCIAS_SQL.get(comp_seleccionada, [comp_seleccionada])
        es_filtro_todas = comp_seleccionada == "Todas las Competiciones"
        es_filtro_uefa = any(u in comp_seleccionada for u in ["Champions", "Europa", "Conference", "UEFA"])

        filas_directorio = []
        for nom_ref, info in censo_dir.items():
            if info["total_partidos_global"] < min_partidos_ref:
                continue

            if not es_filtro_todas:
                coincide_comp = False
                divs_reg = set(info.get("divs_registradas", []))
                for c_f in codigos_filtro:
                    if c_f in divs_reg:
                        coincide_comp = True
                        break
                if not coincide_comp:
                    if info.get("div_principal") in codigos_filtro:
                        coincide_comp = True
                    elif info.get("pais_liga_origen") in codigos_filtro:
                        coincide_comp = True
                    elif es_filtro_uefa and info.get("partidos_uefa", 0) > 0:
                        coincide_comp = True
                if not coincide_comp:
                    continue

            if buscar_ref.strip():
                txt_b = _despojar_diacriticos_arb(buscar_ref.strip().lower())
                txt_nom = _despojar_diacriticos_arb(info["nombre"].lower())
                txt_liga = _despojar_diacriticos_arb(info["pais_liga_origen"].lower())
                txt_vars = " ".join([_despojar_diacriticos_arb(v.lower()) for v in info.get("variantes", [])])
                if txt_b not in txt_nom and txt_b not in txt_liga and txt_b not in txt_vars:
                    continue

            if es_filtro_uefa and info["partidos_uefa"] > 0:
                m_am_col = info["media_amarillas_uefa"]
                m_rj_col = info["media_rojas_uefa"]
                m_fl_col = info["media_faltas_uefa"]
                p_tot_col = info["partidos_uefa"]
            else:
                m_am_col = info["media_amarillas_global"]
                m_rj_col = info["media_rojas_global"]
                m_fl_col = info["media_faltas_global"]
                p_tot_col = info["total_partidos_global"]

            tot_tarj_col = round(m_am_col + (2.0 * m_rj_col), 2)
            if tot_tarj_col < 3.8:
                sev_col = "🟢 Permisivo"
            elif tot_tarj_col <= 4.8:
                sev_col = "🟡 Estándar"
            else:
                sev_col = "🔴 Riguroso"

            filas_directorio.append({
                "Árbitro": info["nombre"],
                "Liga / Origen": info["pais_liga_origen"],
                "Partidos Totales": info["total_partidos_global"],
                "Partidos UEFA": info["partidos_uefa"],
                "Media Amarillas": f"{m_am_col:.2f}",
                "Media Rojas": f"{m_rj_col:.2f}",
                "Media Faltas": f"{m_fl_col:.1f}",
                "Total Tarjetas/90m": f"{tot_tarj_col:.2f}",
                "Severidad": sev_col,
            })

        df_arbitros_view = pd.DataFrame(filas_directorio)
        if not df_arbitros_view.empty:
            df_arbitros_view = df_arbitros_view.sort_values(
                by=["Partidos Totales", "Total Tarjetas/90m"], ascending=[False, False]
            )

        if df_arbitros_view.empty:
            st.info(f"ℹ️ Aún no hay actas de colegiados registradas para **{comp_seleccionada}**. Se poblarán automáticamente al sincronizar los datos disciplinarios UEFA.")
        else:
            st.caption(f"📋 Mostrando **{len(df_arbitros_view)} colegiados** indexados (mínimo {min_partidos_ref} partidos pitados).")
            st.dataframe(df_arbitros_view, width="stretch", hide_index=True)
        
    with tab_ref2:
        st.markdown("Calcula la línea esperada de tarjetas combinando los promedios del colegiado con la tendencia tarjetera de ambos equipos.")
        if not (err_partidos or df_partidos_global.empty):
            render_matchup_disciplinario(df_partidos_global)
            st.markdown("---")
            st.caption("📎 Estimador rápido de línea tarjetera:")
            censo_matchup = _obtener_censo_global_arbitros(df_partidos_global)
            opciones_matchup = [
                item["etiqueta_selector"] for item in sorted(censo_matchup.values(), key=lambda x: x["nombre"])
            ]
            eqs_tarj = sorted(pd.concat([df_partidos_global['HomeTeam'], df_partidos_global['AwayTeam']]).dropna().unique())
            
            c_m1, c_m2, c_m3 = st.columns(3)
            with c_m1:
                arb_sel = st.selectbox("⚖️ Selecciona Árbitro:", opciones_matchup, key="matchup_ref_quick")
            with c_m2:
                eq_t1 = st.selectbox("🏠 Equipo Local:", eqs_tarj, index=0, key="matchup_eq_loc_quick")
            with c_m3:
                eq_t2 = st.selectbox("✈️ Equipo Visitante:", eqs_tarj, index=1 if len(eqs_tarj)>1 else 0, key="matchup_eq_vis_quick")
            
            perfil_sel_rapido = _calcular_perfil_disciplinario_arbitro(arb_sel, censo_matchup, df_partidos_global)
            _renderizar_tarjeta_perfil_arbitro(perfil_sel_rapido)

            def stats_tarjetas(eq):
                df_h = df_partidos_global[df_partidos_global['HomeTeam'] == eq]
                df_a = df_partidos_global[df_partidos_global['AwayTeam'] == eq]
                if 'HY' in df_partidos_global.columns and 'AY' in df_partidos_global.columns:
                    t = pd.to_numeric(df_h['HY'], errors='coerce').fillna(0).sum() + pd.to_numeric(df_a['AY'], errors='coerce').fillna(0).sum()
                    partidos = len(df_h) + len(df_a)
                    return float(t / partidos) if partidos > 0 else 2.10
                return 2.10
                
            t1_avg = stats_tarjetas(eq_t1)
            t2_avg = stats_tarjetas(eq_t2)
            media_ref = float(perfil_sel_rapido["media_amarillas"])
            total_esperado = (t1_avg + t2_avg + media_ref) / 2.0
            
            st.markdown("---")
            c_r1, c_r2, c_r3 = st.columns(3)
            c_r1.metric(f"Media Árbitro ({perfil_sel_rapido['nombre']})", f"{media_ref:.2f}", "Amarillas/Partido")
            c_r2.metric(f"Provocadas {eq_t1}", f"{t1_avg:.2f}", "Amarillas/Partido", delta_color="off")
            c_r3.metric(f"Provocadas {eq_t2}", f"{t2_avg:.2f}", "Amarillas/Partido", delta_color="off")
            
            st.success(f"🔥 Línea de Tarjetas Amarillas Proyectada: **{total_esperado:.1f}**")
        else:
            st.warning("Datos de partidos no disponibles para calcular tendencias de equipos.")

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
                so2 = _liquidar_estado_unida
    with tab_ref2:
        st.markdown("Calcula la línea esperada de tarjetas combinando los promedios del colegiado con la tendencia tarjetera de ambos equipos.")
        if not (err_partidos or df_partidos_global.empty):
            render_matchup_disciplinario(df_partidos_global)
            st.markdown("---")
            st.caption("📎 Estimador rápido de línea tarjetera:")
            censo_matchup = _obtener_censo_global_arbitros(df_partidos_global)
            opciones_matchup = [
                item["etiqueta_selector"] for item in sorted(censo_matchup.values(), key=lambda x: x["nombre"])
            ]
            eqs_tarj = sorted(pd.concat([df_partidos_global['HomeTeam'], df_partidos_global['AwayTeam']]).dropna().unique())
            
            c_m1, c_m2, c_m3 = st.columns(3)
            with c_m1:
                arb_sel = st.selectbox("⚖️ Selecciona Árbitro:", opciones_matchup, key="matchup_ref_quick")
            with c_m2:
                eq_t1 = st.selectbox("🏠 Equipo Local:", eqs_tarj, index=0, key="matchup_eq_loc_quick")
            with c_m3:
                eq_t2 = st.selectbox("✈️ Equipo Visitante:", eqs_tarj, index=1 if len(eqs_tarj)>1 else 0, key="matchup_eq_vis_quick")
            
            perfil_sel_rapido = _calcular_perfil_disciplinario_arbitro(arb_sel, censo_matchup, df_partidos_global)
            _renderizar_tarjeta_perfil_arbitro(perfil_sel_rapido)

            def stats_tarjetas(eq):
                df_h = df_partidos_global[df_partidos_global['HomeTeam'] == eq]
                df_a = df_partidos_global[df_partidos_global['AwayTeam'] == eq]
                if 'HY' in df_partidos_global.columns and 'AY' in df_partidos_global.columns:
                    t = pd.to_numeric(df_h['HY'], errors='coerce').fillna(0).sum() + pd.to_numeric(df_a['AY'], errors='coerce').fillna(0).sum()
                    partidos = len(df_h) + len(df_a)
                    return float(t / partidos) if partidos > 0 else 2.10
                return 2.10
                
            t1_avg = stats_tarjetas(eq_t1)
            t2_avg = stats_tarjetas(eq_t2)
            media_ref = float(perfil_sel_rapido["media_amarillas"])
            total_esperado = (t1_avg + t2_avg + media_ref) / 2.0
            
            st.markdown("---")
            c_r1, c_r2, c_r3 = st.columns(3)
            c_r1.metric(f"Media Árbitro ({perfil_sel_rapido['nombre']})", f"{media_ref:.2f}", "Amarillas/Partido")
            c_r2.metric(f"Provocadas {eq_t1}", f"{t1_avg:.2f}", "Amarillas/Partido", delta_color="off")
            c_r3.metric(f"Provocadas {eq_t2}", f"{t2_avg:.2f}", "Amarillas/Partido", delta_color="off")
            
            st.success(f"🔥 Línea de Tarjetas Amarillas Proyectada: **{total_esperado:.1f}**")
        else:
            st.warning("Datos de partidos no disponibles para calcular tendencias de equipos.")

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

                st.dataframe(pd.DataFrame(filas_matriz_ah), width="stretch", hide_index=True)

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

                st.dataframe(pd.DataFrame(filas_totales), width="stretch", hide_index=True)

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
        _GENAI_SDK_OK = True
    except Exception:
        genai = None
        _GENAI_SDK_OK = False

    try:
        from google.genai import types as genai_types
    except Exception:
        genai_types = None

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

    censo_ia = _obtener_censo_global_arbitros(df_partidos_global)
    arbs_ia_disp = ["🔍 Seleccionar o buscar árbitro..."] + [
        item["etiqueta_selector"] for item in sorted(censo_ia.values(), key=lambda x: x["nombre"])
    ]
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

    perfil_arb_ia = _calcular_perfil_disciplinario_arbitro(
        nombre_arbitro=arb_ia,
        censo_global=censo_ia,
        df_partidos_global=df_partidos_global,
        div_partido=div_ia_code,
    )
    tarj_arb_ia = float(perfil_arb_ia.get("media_amarillas", 4.10))
    faltas_arb_ia = float(perfil_arb_ia.get("media_faltas", 22.5))
    perfil_arb_txt = str(perfil_arb_ia.get("severidad_label", "🟡 Estándar UEFA"))

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
            "colegiado": str(perfil_arb_ia.get("nombre", arb_ia)),
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
        if st.button("🎯 Evaluar Valor 1X2", width="stretch", key="btn_preset_1x2"):
            prompt_preset_activo = (
                f"🎯 Evaluar Valor 1X2 para {etiqueta_duelo} ({liga_ia_sel}): "
                "compara las probabilidades reales y cuotas justas de victoria local, empate y visitante "
                "según los últimos 5 partidos y medias de temporada, e indica dónde reside el valor esperado."
            )
    with qp2:
        if st.button("🔥 Análisis de Goles (Over/Under 2.5)", width="stretch", key="btn_preset_goles"):
            prompt_preset_activo = (
                f"🔥 Análisis de Goles (Over/Under 2.5) y Ambos Marcan para {etiqueta_duelo}: "
                "evalúa los promedios de goles marcados y encajados, tendencia en los últimos 5 partidos, "
                "córners y cuota justa del mercado de goles."
            )
    with qp3:
        if st.button("🟨 Riesgo Disciplinario y Tarjetas", width="stretch", key="btn_preset_tarjetas"):
            prompt_preset_activo = (
                f"🟨 Riesgo Disciplinario y Tarjetas en {etiqueta_duelo} con el árbitro {arb_ia}: "
                "analiza las medias de tarjetas amarillas, rojas, faltas cometidas por ambos equipos y "
                "el rigor del colegiado para determinar qué línea de tarjetas tiene valor."
            )
    with qp4:
        if st.button("⚡ Proyección de Hándicap Asiático", width="stretch", key="btn_preset_ah"):
            prompt_preset_activo = (
                f"⚡ Proyección de Hándicap Asiático para {etiqueta_duelo}: "
                "evalúa las líneas asiáticas de referencia (0.0, -0.5 y +0.5) en función de la superioridad "
                "ofensiva/defensiva y recomienda la cobertura óptima de mercado."
            )
    with qp5:
        if st.button("🗑️ Reiniciar Sesión", width="stretch", key="btn_clear_chat_ia"):
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
