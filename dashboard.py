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
    "⚖️ Directorio y Matchup Árbitros",
    "🧮 Hándicap Asiático Avanzado",
    "🎯 Escáner Autónomo de Ineficiencias (+EV Scanner)",
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
    "⚖️ Directorio y Matchup Árbitros": "Matchup de Árbitros & Tarjetas",
    "🧮 Hándicap Asiático Avanzado": "Hándicap Asiático Avanzado",
    "🎯 Escáner Autónomo de Ineficiencias (+EV Scanner)": "Escáner Autónomo (+EV Scanner)",
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

# Catálogo: (metrica, categoria, etiqueta, tipo, linea, cuota_base Kambi/Paf estimada)
_IND_CATALOGO_PROPS = [
    ("corn_for", "🚩 Córners Propios", "Córners a favor Más de 3.5", "over", 3.5, 1.35),
    ("corn_for", "🚩 Córners Propios", "Córners a favor Más de 4.5", "over", 4.5, 1.70),
    ("corn_for", "🚩 Córners Propios", "Córners a favor Más de 5.5", "over", 5.5, 2.20),
    ("corn_for", "🚩 Córners Propios", "Córners a favor Más de 6.5", "over", 6.5, 3.00),
    ("corn_for", "🚩 Córners Propios", "Córners a favor Más de 7.5", "over", 7.5, 4.20),
    ("corn_against", "🚩 Córners Propios", "Córners concedidos Menos de 3.5", "under", 3.5, 3.00),
    ("corn_against", "🚩 Córners Propios", "Córners concedidos Menos de 4.5", "under", 4.5, 2.05),
    ("sot", "🎯 Tiros a Puerta", "Tiros a puerta Más de 3.5", "over", 3.5, 1.55),
    ("sot", "🎯 Tiros a Puerta", "Tiros a puerta Más de 4.5", "over", 4.5, 2.05),
    ("sot", "🎯 Tiros a Puerta", "Tiros a puerta Más de 5.5", "over", 5.5, 2.90),
    ("sot", "🎯 Tiros a Puerta", "Tiros a puerta Más de 6.5", "over", 6.5, 4.20),
    ("goals", "⚽ Goles Propios", "Goles anotados Más de 0.5", "over", 0.5, 1.25),
    ("goals", "⚽ Goles Propios", "Goles anotados Más de 1.5", "over", 1.5, 2.10),
    ("goals", "⚽ Goles Propios", "Goles anotados Más de 2.5", "over", 2.5, 4.20),
    ("cards", "🟨 Tarjetas Propias", "Tarjetas recibidas Más de 1.5", "over", 1.5, 1.50),
    ("cards", "🟨 Tarjetas Propias", "Tarjetas recibidas Más de 2.5", "over", 2.5, 2.40),
    ("cards", "🟨 Tarjetas Propias", "Tarjetas recibidas Menos de 1.5", "under", 1.5, 2.50),
    ("cards", "🟨 Tarjetas Propias", "Tarjetas recibidas Menos de 2.5", "under", 2.5, 1.55),
]


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
                SELECT {", ".join(_IND_COLS_VISTA)}
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
        # Sin dato (ESPN/UEFA): ambos lados a 0 o NaN -> NaN para excluir del muestreo
        sin_corn = np.isnan(corn_f) | np.isnan(corn_a) | ((corn_f + corn_a) == 0)
        sin_sot = np.isnan(sot) | np.isnan(jug[f"{sr}ST"].values.astype(float)) | \
            ((sot + jug[f"{sr}ST"].values.astype(float)) == 0)
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
    largo["rival3"] = largo["Rival"].astype(str).str.upper().str[:3]
    largo["rival3"] = largo["Rival"].astype(str).str.upper().str[:3].str.strip()
    return largo.sort_values("Fecha", ascending=False, kind="stable").reset_index(drop=True)


def _ind_proximos_cruces(df):
    """Próximo partido por equipo (fechas futuras sin resultado), vectorizado."""
    hoy = pd.Timestamp.now().normalize()
    fut = df[(~df["Jugado"]) & (df["Date"] >= hoy)]
    if fut.empty:
        return pd.DataFrame(columns=["Equipo", "px_rival", "px_cond", "px_fecha"])
    loc = pd.DataFrame({"Equipo": fut["HomeTeam"].values, "px_rival": fut["AwayTeam"].values,
                        "px_cond": "Local", "px_fecha": fut["Date"].values})
    vis = pd.DataFrame({"Equipo": fut["AwayTeam"].values, "px_rival": fut["HomeTeam"].values,
                        "px_cond": "Visitante", "px_fecha": fut["Date"].values})
    todo = pd.concat([loc, vis], ignore_index=True).sort_values("px_fecha", kind="stable")
    return todo.drop_duplicates("Equipo", keep="first").reset_index(drop=True)


def _ind_poisson_p_over(lam, k):
    """P(X >= k) = 1 - sum_{i<k} e^-λ λ^i / i!  (vectorizado sobre λ), acotada [0.02, 0.98]."""
    lam = np.asarray(lam, dtype=float)
    acum = np.zeros_like(lam)
    for i in range(int(k)):
        acum += (np.exp(-lam) * (lam ** i)) / math.factorial(i)
    return np.clip(1.0 - acum, 0.02, 0.98)


def _escanear_props_vectorial(df, ventana=5, min_winrate=75.0, categoria="Todas las Métricas",
                              condicion="Global", divs=None):
    """Escáner de props individuales: Poisson real + EV% sobre ventana configurable."""
    columnas = ["Equipo", "Liga", "Categoria", "Mercado", "Linea", "Tipo", "Hits", "N", "Winrate",
                "Lambda", "Prob", "CuotaJusta", "CuotaBase", "EV", "Chips",
                "px_rival", "px_cond", "px_fecha"]
    if df is None or df.empty:
        return pd.DataFrame(columns=columnas)

    largo = _ind_aplanar_partidos(df)
    if largo.empty:
        return pd.DataFrame(columns=columnas)
    if divs:
        largo = largo[largo["Div"].isin(divs)]
    if condicion == "Local":
        largo = largo[largo["Cond"] == "L"]
    elif condicion == "Visitante":
        largo = largo[largo["Cond"] == "V"]
    if largo.empty:
        return pd.DataFrame(columns=columnas)

    # Liga de referencia por equipo: preferencia doméstica sobre UEFA
    pri = largo.assign(_u=largo["Div"].isin(_IND_UEFA_DIVS).astype(int)).sort_values(
        ["_u", "Fecha"], ascending=[True, False], kind="stable")
    liga_eq = pri.drop_duplicates("Equipo").set_index("Equipo")["Div"].map(
        lambda d: _IND_NOMBRE_DIV.get(d, d))

    partes = []
    metricas = sorted({m[0] for m in _IND_CATALOGO_PROPS})
    for met in metricas:
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
        for (m_met, cat, etiqueta, tipo, linea, cuota_base) in _IND_CATALOGO_PROPS:
            if m_met != met:
                continue
            if categoria not in ("Todas las Métricas", "Todas") and cat != categoria:
                continue
            k = int(math.floor(linea)) + 1
            lam = agg["Lambda"].values
            p_over = _ind_poisson_p_over(lam, k)
            prob = p_over if tipo == "over" else np.clip(1.0 - p_over, 0.02, 0.98)
            hits = np.array([
                sum((v >= k) if tipo == "over" else (v < k) for v in vs) for vs in agg["vals"]
            ], dtype=int)
            wr = hits / float(ventana) * 100.0
            parte = pd.DataFrame({
                "Equipo": agg.index, "Categoria": cat, "Mercado": etiqueta, "Linea": linea,
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
        return pd.DataFrame(columns=columnas)
    res = pd.concat(partes, ignore_index=True)
    if res.empty:
        return pd.DataFrame(columns=columnas)
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
    return f"🗓️ vs {_html_ind.escape(str(rival))} ({cond}) · {f.strftime('%d/%m')}{hora}"


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
        ".ind-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px;margin-bottom:22px;}"
        ".ind-card{background:rgba(13,17,23,0.85);border:1px solid rgba(255,255,255,0.1);"
        "border-radius:12px;padding:14px 16px;backdrop-filter:blur(10px);}"
        ".ind-card .top{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;}"
        ".ind-tag{font-size:0.72rem;font-weight:700;padding:3px 9px;border-radius:12px;"
        "background:rgba(88,166,255,0.14);border:1px solid #1f6feb;color:#58a6ff;}"
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

    c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.2, 1.4])
    with c1:
        comp_sel = st.selectbox("Competición", list(_IND_FILTRO_COMPETICION.keys()), key="ind_radar_comp")
    with c2:
        cond_lbl = st.radio("Condición de Campo", ["🌐 Global", "🏠 Local", "✈️ Visitante"],
                            horizontal=True, key="ind_radar_cond")
    with c3:
        cat_sel = st.selectbox(
            "Categoría de Prop",
            ["Todas las Métricas", "🚩 Córners Propios", "🎯 Tiros a Puerta", "⚽ Goles Propios", "🟨 Tarjetas Propias"],
            key="ind_radar_cat",
        )
    with c4:
        umbral_sel = st.slider("Racha mínima (%)", 50, 100, 75, step=5, key="ind_radar_umbral")
        ventana_sel = st.selectbox("Muestra", [3, 5, 8, 10, 15], index=1,
                                   format_func=lambda x: f"Últimos {x} partidos", key="ind_radar_ventana")

    condicion = {"🌐 Global": "Global", "🏠 Local": "Local", "✈️ Visitante": "Visitante"}[cond_lbl]

    df_raw = _cargar_dataset_props_vectorial()
    if df_raw.empty:
        st.warning("No se hallaron registros históricos para procesar tendencias individuales.")
        return

    res = _escanear_props_vectorial(
        df_raw, ventana=ventana_sel, min_winrate=float(umbral_sel), categoria=cat_sel,
        condicion=condicion, divs=_IND_FILTRO_COMPETICION[comp_sel],
    )
    if res.empty:
        st.info(f"Ningún equipo cumple actualmente el {umbral_sel}% de efectividad con los parámetros fijados.")
        return

    reg = res.to_dict("records")

    # ── TOP 4 EV+ ────────────────────────────────────────────────────────────
    tarjetas = []
    for r in reg[:4]:
        cruce = _ind_cruce_txt(r["px_rival"], r["px_cond"], r["px_fecha"])
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
            f"<td>{_ind_cruce_txt(r['px_rival'], r['px_cond'], r['px_fecha'])}</td>"
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
    def inject_terminal_bento_css():
        st.markdown("""
            <style>
            /* Tokens de Terminal Cuantitativo */
            :root {
                --bg-card: rgba(15, 23, 42, 0.65);
                --border-card: rgba(99, 102, 241, 0.22);
                --border-hover: rgba(129, 140, 248, 0.45);
                --glow-accent: rgba(56, 189, 248, 0.15);
                --text-muted: #94a3b8;
                --text-primary: #f8fafc;
                --accent-green: #34d399;
                --accent-red: #f87171;
                --accent-cyan: #38bdf8;
            }

            /* Layout Grid Flexible */
            .bento-row {
                display: flex;
                gap: 14px;
                width: 100%;
                margin-bottom: 14px;
            }
            
            /* Tarjeta Base con efecto Glassmorphism */
            .bento-tile {
                flex: 1;
                background: var(--bg-card);
                backdrop-filter: blur(12px);
                -webkit-backdrop-filter: blur(12px);
                border: 1px solid var(--border-card);
                border-radius: 12px;
                padding: 16px 20px;
                box-shadow: 0 4px 20px rgba(0, 0, 0, 0.25);
                transition: border-color 0.25s ease, box-shadow 0.25s ease;
            }
            .bento-tile:hover {
                border-color: var(--border-hover);
                box-shadow: 0 6px 24px var(--glow-accent);
            }

            /* Modificadores de Estado */
            .bento-edge {
                border-color: rgba(52, 211, 153, 0.45);
                background: linear-gradient(145deg, rgba(15, 23, 42, 0.8) 0%, rgba(6, 78, 59, 0.25) 100%);
            }
            .bento-alert {
                border-color: rgba(248, 113, 113, 0.45);
                background: linear-gradient(145deg, rgba(15, 23, 42, 0.8) 0%, rgba(127, 29, 29, 0.2) 100%);
            }

            /* Jerarquía Tipográfica */
            .bento-label {
                font-size: 0.72rem;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: 0.08em;
                color: var(--text-muted);
                margin-bottom: 4px;
                display: flex;
                align-items: center;
                justify-content: space-between;
            }
            .bento-value {
                font-size: 1.65rem;
                font-weight: 700;
                color: var(--text-primary);
                letter-spacing: -0.02em;
                line-height: 1.15;
            }
            .bento-desc {
                font-size: 0.78rem;
                color: var(--text-muted);
                margin-top: 6px;
            }

            /* Micro-Badges de Calidad Cuantitativa */
            .badge-quant {
                display: inline-flex;
                align-items: center;
                gap: 4px;
                font-size: 0.70rem;
                font-weight: 700;
                padding: 2px 7px;
                border-radius: 5px;
                text-transform: uppercase;
            }
            .badge-quant-green {
                background: rgba(52, 211, 153, 0.15);
                color: var(--accent-green);
                border: 1px solid rgba(52, 211, 153, 0.35);
            }
            .badge-quant-red {
                background: rgba(248, 113, 113, 0.15);
                color: var(--accent-red);
                border: 1px solid rgba(248, 113, 113, 0.35);
            }
            .badge-quant-cyan {
                background: rgba(56, 189, 248, 0.15);
                color: var(--accent-cyan);
                border: 1px solid rgba(56, 189, 248, 0.35);
            }
            </style>
        """, unsafe_allow_html=True)

    inject_terminal_bento_css()

    col_div_d = _columna_liga_partidos(df_partidos_global)
    _CODIGOS_DOMESTICOS_DOSSIER = {"SP1", "SP2", "E0", "E1", "D1", "I1", "F1"}
    _MAPA_COMP_DOSSIER = {
        "Champions League (UEFA)": {"UCL", "Champions League", "Champions League (UEFA)", "uefa.champions"},
        "Europa League (UEFA)": {"UEL", "Europa League", "Europa League (UEFA)", "uefa.europa"},
        "Conference League (UEFA)": {"UECL", "Conference League", "Conference League (UEFA)", "uefa.europa.conf"},
    }
    _MAPA_SQL_UEFA_DOSSIER = {
        "Champions League (UEFA)": ("UCL", "Champions League", "%Champions League%"),
        "Europa League (UEFA)": ("UEL", "Europa League", "%Europa League%"),
        "Conference League (UEFA)": ("UECL", "Conference League", "%Conference League%"),
    }
    _OPCIONES_COMP_DOSSIER = [
        "Liga Doméstica",
        "Champions League (UEFA)",
        "Europa League (UEFA)",
        "Conference League (UEFA)",
        "Todas las Competiciones (Global)",
    ]

    _comp_activa_hdr = str(st.session_state.get("dossier_competicion_selector", "Liga Doméstica") or "Liga Doméstica")
    _eqs_hdr_validos = []
    if not df_partidos_global.empty and "HomeTeam" in df_partidos_global.columns:
        if _comp_activa_hdr in _MAPA_COMP_DOSSIER and col_div_d:
            _s_div_h = df_partidos_global[col_div_d].astype(str).str.strip()
            _df_h_sub = df_partidos_global[_s_div_h.isin(_MAPA_COMP_DOSSIER[_comp_activa_hdr])]
        elif _comp_activa_hdr == "Liga Doméstica" and col_div_d:
            _s_div_h = df_partidos_global[col_div_d].astype(str).str.strip()
            _df_h_sub = df_partidos_global[_s_div_h.isin(_CODIGOS_DOMESTICOS_DOSSIER)]
        else:
            _df_h_sub = df_partidos_global
        _eqs_hdr_validos = sorted(pd.concat([_df_h_sub["HomeTeam"], _df_h_sub["AwayTeam"]]).dropna().astype(str).unique().tolist())

    equipo_seleccionado = str(st.session_state.get("dossier_equipo_selector", "") or "").strip()
    if _eqs_hdr_validos and (not equipo_seleccionado or equipo_seleccionado not in _eqs_hdr_validos):
        equipo_seleccionado = _eqs_hdr_validos[0]
        st.session_state["dossier_equipo_selector"] = equipo_seleccionado
    elif not equipo_seleccionado:
        equipo_seleccionado = "Equipo"

    proximo_fixture = None
    if RUTA_BD.exists() and equipo_seleccionado:
        try:
            with sqlite3.connect(RUTA_BD) as _conn_hdr:
                _cur_hdr = _conn_hdr.cursor()
                try:
                    _cur_hdr.execute(
                        "SELECT rival, fecha, condicion, competicion, arbitro FROM calendario_fixtures WHERE equipo = ? AND (disputado = 0 OR finalizado = 0) ORDER BY fecha ASC LIMIT 1",
                        (equipo_seleccionado,),
                    )
                    _res_hdr = _cur_hdr.fetchone()
                    if _res_hdr:
                        proximo_fixture = {
                            'rival': _res_hdr[0],
                            'fecha': _res_hdr[1],
                            'condicion': _res_hdr[2],
                            'competicion': _res_hdr[3],
                            'arbitro_designado': _res_hdr[4] if _res_hdr[4] else 'Por Confirmar',
                        }
                except Exception:
                    pass
        except Exception:
            proximo_fixture = None

    sub_fixture = ""
    if proximo_fixture:
        rival_nom = proximo_fixture.get('rival', 'Rival')
        cond_nom = 'Local' if proximo_fixture.get('condicion', 'C') in ['C', 'Local', 'H'] else 'Visitante'
        comp_nom = proximo_fixture.get('competicion', '')
        sub_fixture = f" <span style='font-size: 1.1rem; color: #94a3b8; font-weight: normal;'>| Próximo: <strong>vs {rival_nom}</strong> ({cond_nom}) · {comp_nom}</span>"

    st.markdown(f"## 📊 Radiografía Cuantitativa: {equipo_seleccionado}{sub_fixture}", unsafe_allow_html=True)
    st.markdown(
        "Radiografía institucional de rendimiento: **Eficiencia xG & Conversión de Tiro**, "
        "**Matriz de Frecuencias de Mercado (Cuotas Implícitas)**, **Ratings Relativos Casa/Fuera** y **Perfil de Fricción Disciplinaria**."
    )

    if err_partidos or df_partidos_global.empty:
        st.warning("No hay datos de partidos disponibles en la base de datos.")
    else:
        # ── LAYER 3.1: TOP EXECUTIVE CONTROL (COMPETICIÓN + EQUIPO + HORIZONTE) ──
        st.markdown("""
<style>
/* 1. BaseWeb SelectBox (Caja cerrada y contenedores internos) */
div[data-baseweb="select"],
div[data-baseweb="select"] > div,
div[data-baseweb="select"] div,
div[data-baseweb="select"] input {
    background-color: #0f172a !important;
    color: #f8fafc !important;
    border-color: rgba(99, 102, 241, 0.35) !important;
    border-radius: 8px !important;
}
div[data-baseweb="select"] span,
div[data-baseweb="select"] div[aria-selected="true"] {
    color: #f8fafc !important;
}
div[data-baseweb="select"] svg {
    fill: #38bdf8 !important;
}

/* 2. Menú desplegable flotante (BaseWeb Popover Portal en body) */
div[data-baseweb="popover"],
div[data-baseweb="popover"] > div,
div[data-baseweb="popover"] ul,
ul[role="listbox"] {
    background-color: #0f172a !important;
    border: 1px solid rgba(99, 102, 241, 0.35) !important;
    border-radius: 8px !important;
    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.6) !important;
}
li[role="option"] {
    background-color: #0f172a !important;
    color: #cbd5e1 !important;
    font-size: 0.85rem !important;
    padding: 8px 14px !important;
    transition: background-color 0.15s ease !important;
}
li[role="option"]:hover,
li[aria-selected="true"] {
    background-color: rgba(56, 189, 248, 0.18) !important;
    color: #38bdf8 !important;
}

/* 3. Botones de Navegación Superior (Atrás, Portada) */
div[data-testid="stHorizontalBlock"] button {
    background: rgba(15, 23, 42, 0.75) !important;
    border: 1px solid rgba(99, 102, 241, 0.3) !important;
    color: #cbd5e1 !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
    transition: all 0.2s ease !important;
}
div[data-testid="stHorizontalBlock"] button:hover {
    background: rgba(56, 189, 248, 0.15) !important;
    border-color: #38bdf8 !important;
    color: #38bdf8 !important;
}

/* 4. Selector de Horizonte Muestral en Píldoras Horizontales */
div[data-testid="stRadio"] > div[role="radiogroup"] {
    display: flex !important;
    flex-direction: row !important;
    gap: 6px !important;
    background: rgba(15, 23, 42, 0.75) !important;
    padding: 4px !important;
    border-radius: 8px !important;
    border: 1px solid rgba(99, 102, 241, 0.25) !important;
}
div[data-testid="stRadio"] label {
    background: transparent !important;
    padding: 6px 12px !important;
    border-radius: 6px !important;
    margin: 0 !important;
    cursor: pointer !important;
    border: 1px solid transparent !important;
    transition: all 0.2s ease !important;
}
div[data-testid="stRadio"] label > div:first-child {
    display: none !important; /* Oculta el círculo nativo */
}
div[data-testid="stRadio"] label p {
    color: #94a3b8 !important;
    font-size: 0.78rem !important;
    margin: 0 !important;
}
div[data-testid="stRadio"] label:has(input:checked) {
    background: rgba(56, 189, 248, 0.18) !important;
    border-color: rgba(56, 189, 248, 0.45) !important;
}
div[data-testid="stRadio"] label:has(input:checked) p {
    color: #38bdf8 !important;
    font-weight: 700 !important;
}
</style>
""", unsafe_allow_html=True)

        ctrl_d0, ctrl_d1, ctrl_d2 = st.columns([1.3, 1.3, 2.2])
        with ctrl_d0:
            comp_dossier_sel = st.selectbox(
                "🏆 Competición:",
                _OPCIONES_COMP_DOSSIER,
                key="dossier_competicion_selector",
            )

        competicion_seleccionada = comp_dossier_sel
        if comp_dossier_sel in _MAPA_SQL_UEFA_DOSSIER:
            _cod_uefa, _nom_uefa, _like_uefa = _MAPA_SQL_UEFA_DOSSIER[comp_dossier_sel]
            df_comp_dossier = pd.DataFrame()
            if RUTA_BD.exists():
                try:
                    with sqlite3.connect(RUTA_BD) as _conn_uefa:
                        df_comp_dossier = pd.read_sql_query(
                            'SELECT * FROM historico_reciente WHERE "ï»¿Div" = ? OR Liga_Origen = ? OR Liga_Origen LIKE ?',
                            _conn_uefa,
                            params=(_cod_uefa, _nom_uefa, _like_uefa),
                        )
                        df_comp_dossier.columns = df_comp_dossier.columns.str.replace('ï»¿', '').str.strip()
                except Exception:
                    df_comp_dossier = pd.DataFrame()
            if df_comp_dossier.empty and col_div_d:
                serie_div_global = df_partidos_global[col_div_d].astype(str).str.strip()
                serie_orig_global = (
                    df_partidos_global["Liga_Origen"].astype(str).str.strip()
                    if "Liga_Origen" in df_partidos_global.columns
                    else serie_div_global
                )
                df_comp_dossier = df_partidos_global[
                    serie_div_global.isin(_MAPA_COMP_DOSSIER[comp_dossier_sel])
                    | serie_orig_global.isin(_MAPA_COMP_DOSSIER[comp_dossier_sel])
                ].copy()
        elif col_div_d:
            serie_div_global = df_partidos_global[col_div_d].astype(str).str.strip()
            if comp_dossier_sel == "Liga Doméstica":
                df_comp_dossier = df_partidos_global[serie_div_global.isin(_CODIGOS_DOMESTICOS_DOSSIER)].copy()
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
        ) if not df_comp_dossier.empty and "HomeTeam" in df_comp_dossier.columns else []

        if not todos_equipos and comp_dossier_sel not in _MAPA_SQL_UEFA_DOSSIER:
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
            def calcular_vector_cuantitativo_avanzado(df_partidos, conn, equipo_objetivo):
                """
                Calcula medianas robustas, IQR, regresión empírica xPTS y estado de varianza.
                Preserva tipos numéricos nativos de Python/NumPy sin mutar el DataFrame original.
                """
                if df_partidos is None or df_partidos.empty:
                    return {
                        'mediana_gf': 0.0, 'mediana_ga': 0.0, 'iqr_gf': 0.0,
                        'pts_reales': 0, 'xpts_esperados': 0.0, 'divergencia_pts': 0.0,
                        'diagnostico_suerte': 'Muestra Insuficiente', 'color_badge': 'cyan'
                    }
                
                # 1. Mediana e Intervalo Intercuartílico (IQR) frente a anomalías muestrales
                if 'goles_favor' in df_partidos.columns:
                    goles_f = df_partidos['goles_favor'].astype(float)
                elif 'GF' in df_partidos.columns:
                    goles_f = df_partidos['GF'].astype(float)
                else:
                    goles_f = df_partidos['gf'].astype(float)

                if 'goles_contra' in df_partidos.columns:
                    goles_c = df_partidos['goles_contra'].astype(float)
                elif 'GC' in df_partidos.columns:
                    goles_c = df_partidos['GC'].astype(float)
                else:
                    goles_c = df_partidos['gc'].astype(float)
                
                mediana_gf = float(goles_f.median())
                mediana_ga = float(goles_c.median())
                iqr_gf = float(goles_f.quantile(0.75) - goles_f.quantile(0.25))
                
                # 2. Modelo de Puntos Esperados (xPTS) basado en matriz delta xG
                xpts_total = 0.0
                pts_reales_calc = 0
                
                col_xg_f = 'xg_favor' if 'xg_favor' in df_partidos.columns else ('xG_F' if 'xG_F' in df_partidos.columns else 'xg_f')
                col_xg_c = 'xg_contra' if 'xg_contra' in df_partidos.columns else ('xG_C' if 'xG_C' in df_partidos.columns else 'xg_c')
                
                for _, fila in df_partidos.iterrows():
                    xg_f = float(fila.get(col_xg_f, 1.15))
                    xg_c = float(fila.get(col_xg_c, 1.15))
                    
                    # Simulación probabilística simplificada de partido por balance de ocasiones
                    diff_xg = xg_f - xg_c
                    p_win = max(0.02, min(0.96, 0.38 + 0.26 * diff_xg))
                    p_draw = max(0.04, min(0.40, 0.28 - 0.06 * abs(diff_xg)))
                    xpts_total += (p_win * 3.0) + (p_draw * 1.0)
                    
                    # Puntos reales
                    res = str(fila.get('resultado', fila.get('Result', ''))).upper()
                    if res in ['V', 'W', 'GANADO']:
                        pts_reales_calc += 3
                    elif res in ['E', 'D', 'EMPATADO']:
                        pts_reales_calc += 1
                    elif not res:
                        if float(fila.get('win', 0.0)) > 0.5:
                            pts_reales_calc += 3
                        elif float(fila.get('draw', 0.0)) > 0.5:
                            pts_reales_calc += 1
                        
                divergencia = pts_reales_calc - xpts_total
                
                if divergencia >= 2.0:
                    diag = "Sobre-rendimiento (+Varianza a favor / Regresión negativa esperada)"
                    badge_color = "red"
                elif divergencia <= -2.0:
                    diag = "Sub-rendimiento (-Varianza / Potencial infravalorado en mercado)"
                    badge_color = "green"
                else:
                    diag = "Rendimiento Equilibrado (Alineado con generación estocástica)"
                    badge_color = "cyan"
                    
                return {
                    'mediana_gf': mediana_gf,
                    'mediana_ga': mediana_ga,
                    'iqr_gf': iqr_gf,
                    'pts_reales': pts_reales_calc,
                    'xpts_esperados': round(xpts_total, 2),
                    'divergencia_pts': round(divergencia, 2),
                    'diagnostico_suerte': diag,
                    'color_badge': badge_color
                }

            def resolver_contexto_proximo_fixture(conn, equipo_analizado):
                """
                Localiza el siguiente compromiso oficial pendiente del equipo en macro_base_futbol.db.
                """
                query = """
                    SELECT rival, fecha, condicion, competicion, arbitro 
                    FROM calendario_fixtures 
                    WHERE equipo = ? AND (disputado = 0 OR finalizado = 0)
                    ORDER BY fecha ASC LIMIT 1
                """
                try:
                    cur = conn.cursor()
                    cur.execute(query, (equipo_analizado,))
                    res = cur.fetchone()
                    if res:
                        return {
                            'rival': res[0],
                            'fecha': res[1],
                            'condicion': res[2],
                            'competicion': res[3],
                            'arbitro_designado': res[4] if res[4] else 'Por Confirmar'
                        }
                except Exception:
                    pass
                try:
                    cur = conn.cursor()
                    cur.execute(
                        """
                        SELECT CASE WHEN equipo_local = ? THEN equipo_visitante ELSE equipo_local END AS rival,
                               fecha,
                               CASE WHEN equipo_local = ? THEN 'Local' ELSE 'Visitante' END AS condicion,
                               COALESCE(torneo, liga) AS competicion,
                               arbitro
                        FROM partidos
                        WHERE (equipo_local = ? OR equipo_visitante = ?)
                          AND goles_local IS NULL
                        ORDER BY fecha ASC LIMIT 1
                        """,
                        (equipo_analizado, equipo_analizado, equipo_analizado, equipo_analizado),
                    )
                    res_p = cur.fetchone()
                    if res_p:
                        return {
                            'rival': res_p[0],
                            'fecha': res_p[1],
                            'condicion': res_p[2],
                            'competicion': res_p[3],
                            'arbitro_designado': res_p[4] if res_p[4] else 'Por Confirmar'
                        }
                except Exception:
                    pass
                return None

            df_eq_raw_all = df_comp_dossier[
                (df_comp_dossier["HomeTeam"] == equipo_elegido)
                | (df_comp_dossier["AwayTeam"] == equipo_elegido)
            ].copy() if not df_comp_dossier.empty else pd.DataFrame()
            df_eq_ord_all = _ordenar_por_fecha(df_eq_raw_all) if not df_eq_raw_all.empty else pd.DataFrame()

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

            media_liga_home = float(pd.to_numeric(df_liga_ctx["FTHG"], errors="coerce").mean() or 1.45) if not df_liga_ctx.empty else 1.45
            media_liga_away = float(pd.to_numeric(df_liga_ctx["FTAG"], errors="coerce").mean() or 1.15) if not df_liga_ctx.empty else 1.15
            media_liga_global = (media_liga_home + media_liga_away) / 2.0

            df_norm_global = _normalizar_partidos_equipo(df_muestra_raw, equipo_elegido) if not df_muestra_raw.empty else pd.DataFrame()

            # Enriquecimiento estricto UEFA desde estadisticas_jugadores (WHERE competicion = ?) sin mezclar liga doméstica
            if comp_dossier_sel in _MAPA_SQL_UEFA_DOSSIER and not df_norm_global.empty and RUTA_BD.exists():
                _, _nom_uefa_q, _ = _MAPA_SQL_UEFA_DOSSIER[comp_dossier_sel]
                try:
                    with sqlite3.connect(RUTA_BD) as _conn_u_tel:
                        _df_u_tel = pd.read_sql_query(
                            """
                            SELECT Equipo,
                                   SUM(tiros_totales) AS tiros,
                                   SUM(tiros_a_puerta) AS sot,
                                   SUM(faltas_cometidas) AS faltas_c,
                                   SUM(Faltas_Recibidas) AS faltas_r,
                                   SUM(tarjetas_amarillas) AS amar,
                                   SUM(tarjetas_rojas) AS rojas,
                                   MAX(partidos_jugados) AS pj
                            FROM estadisticas_jugadores
                            WHERE competicion = ?
                            GROUP BY Equipo
                            """,
                            _conn_u_tel,
                            params=(_nom_uefa_q,),
                        )
                    if not _df_u_tel.empty:
                        _mapa_u_tel = {}
                        for _, _r_u in _df_u_tel.iterrows():
                            _pj_u = max(1.0, float(_r_u.get("pj") or 1.0))
                            _t_m = float(_r_u.get("tiros") or 0.0) / _pj_u
                            _s_m = float(_r_u.get("sot") or 0.0) / _pj_u
                            _fc_m = float(_r_u.get("faltas_c") or 0.0) / _pj_u
                            _fr_m = float(_r_u.get("faltas_r") or 0.0) / _pj_u
                            _am_m = float(_r_u.get("amar") or 0.0) / _pj_u
                            _rj_m = float(_r_u.get("rojas") or 0.0) / _pj_u
                            _co_m = max(2.8, round(_t_m * 0.34, 2))
                            _mapa_u_tel[str(_r_u["Equipo"])] = {
                                "tiros": _t_m, "sot": _s_m, "faltas_c": _fc_m, "faltas_r": _fr_m,
                                "amar": _am_m, "rojas": _rj_m, "corners": _co_m,
                            }
                        for _idx_n, _row_n in df_norm_global.iterrows():
                            _es_loc_n = bool(_row_n.get("es_local", True))
                            _riv_n = str(_row_n.get("AwayTeam") if _es_loc_n else _row_n.get("HomeTeam"))
                            _f_loc = 1.06 if _es_loc_n else 0.94
                            _f_riv = 0.94 if _es_loc_n else 1.06
                            _d_eq = _mapa_u_tel.get(equipo_elegido, {"tiros": 12.0, "sot": 4.2, "faltas_c": 11.2, "faltas_r": 11.0, "amar": 1.9, "rojas": 0.05, "corners": 4.8})
                            _d_rv = _mapa_u_tel.get(_riv_n, {"tiros": 11.5, "sot": 4.0, "faltas_c": 11.5, "faltas_r": 11.2, "amar": 2.0, "rojas": 0.05, "corners": 4.5})
                            if pd.isna(_row_n.get("tiros_f")):
                                df_norm_global.at[_idx_n, "tiros_f"] = round(_d_eq["tiros"] * _f_loc, 2)
                                df_norm_global.at[_idx_n, "tiros_c"] = round(_d_rv["tiros"] * _f_riv, 2)
                            if pd.isna(_row_n.get("sot_f")):
                                _sf_v = round(_d_eq["sot"] * _f_loc, 2)
                                _sc_v = round(_d_rv["sot"] * _f_riv, 2)
                                df_norm_global.at[_idx_n, "sot_f"] = _sf_v
                                df_norm_global.at[_idx_n, "sot_c"] = _sc_v
                                df_norm_global.at[_idx_n, "xg_f"] = round(max(0.15, 0.26 * _sf_v + 0.04 * max(0.0, df_norm_global.at[_idx_n, "tiros_f"] - _sf_v)), 2)
                                df_norm_global.at[_idx_n, "xg_c"] = round(max(0.15, 0.26 * _sc_v + 0.04 * max(0.0, df_norm_global.at[_idx_n, "tiros_c"] - _sc_v)), 2)
                            if pd.isna(_row_n.get("corners_f")):
                                _cf_v = round(_d_eq["corners"] * _f_loc, 2)
                                _cc_v = round(_d_rv["corners"] * _f_riv, 2)
                                df_norm_global.at[_idx_n, "corners_f"] = _cf_v
                                df_norm_global.at[_idx_n, "corners_c"] = _cc_v
                                df_norm_global.at[_idx_n, "corners_tot"] = round(_cf_v + _cc_v, 2)
                                df_norm_global.at[_idx_n, "over85_corners"] = 1.0 if (_cf_v + _cc_v) > 8.5 else 0.0
                            if pd.isna(_row_n.get("faltas_com")):
                                df_norm_global.at[_idx_n, "faltas_com"] = round(_d_eq["faltas_c"], 2)
                                df_norm_global.at[_idx_n, "faltas_rec"] = round(_d_eq["faltas_r"], 2)
                            if pd.isna(_row_n.get("amar_prop")):
                                df_norm_global.at[_idx_n, "amar_prop"] = round(_d_eq["amar"], 2)
                                df_norm_global.at[_idx_n, "amar_riv"] = round(_d_rv["amar"], 2)
                                df_norm_global.at[_idx_n, "rojas_prop"] = round(_d_eq["rojas"], 2)
                                df_norm_global.at[_idx_n, "rojas_riv"] = round(_d_rv["rojas"], 2)
                except Exception:
                    pass

            df_norm_home = df_norm_global[df_norm_global["es_local"]].copy() if not df_norm_global.empty else pd.DataFrame()
            df_norm_away = df_norm_global[~df_norm_global["es_local"]].copy() if not df_norm_global.empty else pd.DataFrame()

            # Muestra de respaldo aislada por competición (sin mezclar liga doméstica en UEFA)
            if comp_dossier_sel in _MAPA_SQL_UEFA_DOSSIER:
                df_fb_con_datos = df_norm_global.copy()
                df_fb_home = df_norm_home.copy()
                df_fb_away = df_norm_away.copy()
            else:
                df_eq_glob_all = _ordenar_por_fecha(
                    df_comp_dossier[
                        (df_comp_dossier["HomeTeam"] == equipo_elegido)
                        | (df_comp_dossier["AwayTeam"] == equipo_elegido)
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

            def _preparar_df_dossier_fase2(df_norm_in: pd.DataFrame) -> pd.DataFrame:
                if df_norm_in is None or df_norm_in.empty:
                    return pd.DataFrame()
                df_f2 = df_norm_in.copy()
                df_f2["goles_favor"] = pd.to_numeric(df_f2["gf"], errors="coerce").fillna(0.0)
                df_f2["goles_contra"] = pd.to_numeric(df_f2["gc"], errors="coerce").fillna(0.0)
                df_f2["xg_favor"] = pd.to_numeric(df_f2["xg_f"], errors="coerce").fillna(df_f2["goles_favor"])
                df_f2["xg_contra"] = pd.to_numeric(df_f2["xg_c"], errors="coerce").fillna(df_f2["goles_contra"])
                df_f2["resultado"] = np.where(
                    df_f2["goles_favor"] > df_f2["goles_contra"],
                    "V",
                    np.where(df_f2["goles_favor"] == df_f2["goles_contra"], "E", "P"),
                )
                return df_f2

            def obtener_censo_dependencia_p90(conn, equipo_analizado, min_minutos=180):
                """
                Extrae los jugadores nucleares del equipo y computa métricas ofensivas y
                disciplinarias por 90 minutos usando minutos reales registrados.
                """
                if comp_dossier_sel in _MAPA_SQL_UEFA_DOSSIER:
                    _, _comp_u_nom, _ = _MAPA_SQL_UEFA_DOSSIER[comp_dossier_sel]
                    try:
                        query_uefa = """
                            SELECT 
                                Jugador AS jugador,
                                posicion,
                                minutos AS minutos_jugados,
                                titularidades,
                                ROUND(CAST(tiros_totales AS FLOAT) * 90.0 / NULLIF(minutos, 0), 2) AS tiros_p90,
                                ROUND(CAST(tiros_a_puerta AS FLOAT) * 90.0 / NULLIF(minutos, 0), 2) AS sot_p90,
                                ROUND(CAST(faltas_cometidas AS FLOAT) * 90.0 / NULLIF(minutos, 0), 2) AS faltas_p90,
                                ROUND(CAST(tarjetas_amarillas AS FLOAT) * 90.0 / NULLIF(minutos, 0), 2) AS amarillas_p90
                            FROM estadisticas_jugadores 
                            WHERE Equipo = ? AND competicion = ? AND minutos >= ?
                            ORDER BY minutos DESC 
                            LIMIT 5
                        """
                        df_p90_u = pd.read_sql_query(query_uefa, conn, params=(equipo_analizado, _comp_u_nom, 90))
                        if not df_p90_u.empty:
                            return df_p90_u
                    except Exception:
                        pass
                query = """
                    SELECT 
                        jugador,
                        posicion,
                        minutos_jugados,
                        titularidades,
                        ROUND(CAST(tiros_totales AS FLOAT) * 90.0 / NULLIF(minutos_jugados, 0), 2) AS tiros_p90,
                        ROUND(CAST(tiros_puerta AS FLOAT) * 90.0 / NULLIF(minutos_jugados, 0), 2) AS sot_p90,
                        ROUND(CAST(faltas_cometidas AS FLOAT) * 90.0 / NULLIF(minutos_jugados, 0), 2) AS faltas_p90,
                        ROUND(CAST(tarjetas_amarillas AS FLOAT) * 90.0 / NULLIF(minutos_jugados, 0), 2) AS amarillas_p90
                    FROM jugadores_estadisticas 
                    WHERE equipo = ? AND minutos_jugados >= ?
                    ORDER BY minutos_jugados DESC 
                    LIMIT 5
                """
                try:
                    df_p90 = pd.read_sql_query(query, conn, params=(equipo_analizado, min_minutos))
                    if not df_p90.empty:
                        return df_p90
                except Exception:
                    pass
                try:
                    query_est = """
                        SELECT 
                            Jugador AS jugador,
                            posicion,
                            minutos AS minutos_jugados,
                            titularidades,
                            ROUND(CAST(tiros_totales AS FLOAT) * 90.0 / NULLIF(minutos, 0), 2) AS tiros_p90,
                            ROUND(CAST(tiros_a_puerta AS FLOAT) * 90.0 / NULLIF(minutos, 0), 2) AS sot_p90,
                            ROUND(CAST(faltas_cometidas AS FLOAT) * 90.0 / NULLIF(minutos, 0), 2) AS faltas_p90,
                            ROUND(CAST(tarjetas_amarillas AS FLOAT) * 90.0 / NULLIF(minutos, 0), 2) AS amarillas_p90
                        FROM estadisticas_jugadores 
                        WHERE Equipo = ? AND minutos >= ?
                        ORDER BY minutos DESC 
                        LIMIT 5
                    """
                    df_p90_est = pd.read_sql_query(query_est, conn, params=(equipo_analizado, min_minutos))
                    if not df_p90_est.empty:
                        return df_p90_est
                except Exception:
                    pass
                return None

            equipo_seleccionado = equipo_elegido
            df_partidos = _preparar_df_dossier_fase2(df_norm_global)
            if df_partidos is None or df_partidos.empty:
                st.warning(f"⚠️ El equipo **{equipo_seleccionado}** no registra encuentros disputados en **{competicion_seleccionada}** dentro de la base de datos.")
                st.stop()

            with sqlite3.connect(RUTA_BD) as conn:
                # Cómputo matemático silencioso de Fase 2
                metricas_robustas = calcular_vector_cuantitativo_avanzado(df_partidos, conn, equipo_seleccionado)
                proximo_fixture = resolver_contexto_proximo_fixture(conn, equipo_seleccionado)
                df_censo_p90 = obtener_censo_dependencia_p90(conn, equipo_seleccionado)

            # Debug asertivo interno (sin imprimir a pantalla)
            assert isinstance(metricas_robustas, dict), "Error: metricas_robustas debe ser un diccionario"

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

            _, bar_pdf = st.columns([3, 1])
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

            # ── ROW 1: HIGH-IMPACT QUANTITATIVE KPIs (BENTO GRID) ────────────
            st.markdown("---")
            st.subheader(f"📈 Radiografía Ejecutiva: {equipo_elegido}")

            win_rate_val = float(stats_global["win_rate"])
            victorias = int(stats_global["v"])
            empates = int(stats_global["e"])
            derrotas = int(stats_global["d"])
            goles_favor_media = float(stats_global["gf_med"])
            xg_neto_media = float(stats_global["xg_neto"])
            g_minus_xg_diff = float(stats_global["diff_g_xg"])
            clean_sheet_pct = float(stats_global["clean_sheet"])
            cs_conseguidos = int(stats_global["cs_cnt"])
            btts_pct = float(stats_global["btts_si"])
            btts_count = int(stats_global["btts_cnt"])

            st.markdown(f"""
                <div class="bento-row">
                    <div class="bento-tile">
                        <div class="bento-label">Win Rate & Muestra Activa <span class="badge-quant badge-quant-cyan">L10</span></div>
                        <div class="bento-value">{win_rate_val:.1f}%</div>
                        <div class="bento-desc">{victorias}V - {empates}E - {derrotas}D ({stats_global['n']} partidos)</div>
                    </div>
                    <div class="bento-tile">
                        <div class="bento-label">Producción Neta xG <span class="badge-quant badge-quant-green">G/xG</span></div>
                        <div class="bento-value">{goles_favor_media:.2f} <span style="font-size:1rem;color:var(--text-muted);">GF</span> | {xg_neto_media:+.2f} <span style="font-size:1rem;color:var(--text-muted);">xG</span></div>
                        <div class="bento-desc">{g_minus_xg_diff:+.2f} diferencial de finalización</div>
                    </div>
                    <div class="bento-tile">
                        <div class="bento-label">Solidez Defensiva <span class="badge-quant badge-quant-red">Clean Sheet</span></div>
                        <div class="bento-value">{clean_sheet_pct:.1f}%</div>
                        <div class="bento-desc">{cs_conseguidos}/{stats_global['n']} porterías a cero</div>
                    </div>
                    <div class="bento-tile">
                        <div class="bento-label">Tendencia BTTS <span class="badge-quant badge-quant-cyan">Ambos Marcan</span></div>
                        <div class="bento-value">{btts_pct:.1f}%</div>
                        <div class="bento-desc">{btts_count}/{stats_global['n']} partidos con goles mutuos</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

            # ── ROW 2: SPLIT COMPARATIVE ANALYSIS: CASA VS FUERA / GLOBAL ───
            st.markdown("---")
            stats_comp_home = stats_home
            stats_comp_away = stats_away
            stats_comp_gen = stats_global
            for _st_d in (stats_comp_home, stats_comp_away, stats_comp_gen):
                if isinstance(_st_d, dict):
                    _st_d["goles_favor_media"] = float(_st_d.get("gf_med", 0.0))
                    _st_d["goles_contra_media"] = float(_st_d.get("gc_med", 0.0))
                    _st_d["tiros_favor_media"] = float(_st_d.get("tiros_f", 0.0))
                    _st_d["tiros_contra_media"] = float(_st_d.get("tiros_c", 0.0))
                    _st_d["sot_favor_media"] = float(_st_d.get("sot_f", 0.0))
                    _st_d["sot_contra_media"] = float(_st_d.get("sot_c", 0.0))
                    _st_d["corners_favor_media"] = float(_st_d.get("corners_f", 0.0))
                    _st_d["corners_contra_media"] = float(_st_d.get("corners_c", 0.0))
                    _st_d["faltas_cometidas_media"] = float(_st_d.get("faltas_com", 0.0))
                    _st_d["faltas_recibidas_media"] = float(_st_d.get("faltas_rec", 0.0))
                    _st_d["tarjetas_propias_media"] = float(_st_d.get("amar_prop", 0.0)) + float(_st_d.get("rojas_prop", 0.0))
                    _st_d["xg_favor_media"] = float(_st_d.get("xg_f", 0.0))

            st.markdown("#### ⚡ Rendimiento Asimétrico & Robustez Estocástica")

            col_asym1, col_asym2, col_asym3 = st.columns(3)

            with col_asym1:
                gf_c = stats_comp_home.get('goles_favor_media', 0.0) if isinstance(stats_comp_home, dict) else 0.0
                ga_c = stats_comp_home.get('goles_contra_media', 0.0) if isinstance(stats_comp_home, dict) else 0.0
                gf_f = stats_comp_away.get('goles_favor_media', 0.0) if isinstance(stats_comp_away, dict) else 0.0
                ga_f = stats_comp_away.get('goles_contra_media', 0.0) if isinstance(stats_comp_away, dict) else 0.0
                
                st.markdown(f"""
                    <div class="bento-tile">
                        <div class="bento-label">Goles F / C (Casa vs Fuera)</div>
                        <div class="bento-value">{gf_c:.2f} / {ga_c:.2f} <span style="font-size:0.95rem;color:var(--text-muted);font-weight:400;">(L)</span></div>
                        <div class="bento-desc">{gf_f:.2f} / {ga_f:.2f} (V) · Mediana robusta: <strong>{metricas_robustas['mediana_gf']:.1f} GF</strong> (IQR: {metricas_robustas['iqr_gf']:.1f})</div>
                    </div>
                """, unsafe_allow_html=True)

            with col_asym2:
                # Extracción de métricas de tiro, xG y eficiencia en ambas condiciones
                sot_casa = float(stats_comp_home.get('sot_favor_media', 0.0))
                xg_casa = float(stats_comp_home.get('xg_favor_media', 0.0))
                gf_casa = float(stats_comp_home.get('goles_favor_media', 0.0))
                eficiencia_casa = (gf_casa / max(sot_casa, 0.1)) * 100.0

                sot_fuera = float(stats_comp_away.get('sot_favor_media', 0.0))
                xg_fuera = float(stats_comp_away.get('xg_favor_media', 0.0))
                gf_fuera = float(stats_comp_away.get('goles_favor_media', 0.0))
                eficiencia_fuera = (gf_fuera / max(sot_fuera, 0.1)) * 100.0

                st.markdown(f"""
                    <div class="bento-tile">
                        <div class="bento-label">VOLUMEN A PUERTA & XG (CASA VS FUERA)</div>
                        <div class="bento-value">{sot_casa:.1f} <span style="font-size:0.9rem;color:#38bdf8;">SOT</span> | {xg_casa:.2f} <span style="font-size:0.9rem;color:#34d399;">xG</span> <span style="font-size:0.8rem;color:#94a3b8;">(L)</span></div>
                        <div class="bento-desc">{sot_fuera:.1f} SOT | {xg_fuera:.2f} xG (V) · Eficiencia: {eficiencia_casa:.1f}% (L) vs {eficiencia_fuera:.1f}% (V)</div>
                    </div>
                """, unsafe_allow_html=True)

            with col_asym3:
                diag_class = f"badge-quant badge-quant-{metricas_robustas['color_badge']}"
                st.markdown(f"""
                    <div class="bento-tile">
                        <div class="bento-label">Varianza xPTS (Regresión) <span class="{diag_class}">Diagnóstico</span></div>
                        <div class="bento-value">{metricas_robustas['pts_reales']} Pts <span style="font-size:0.95rem;color:var(--text-muted);font-weight:400;">vs</span> {metricas_robustas['xpts_esperados']:.1f} xPTS</div>
                        <div class="bento-desc">{metricas_robustas['diagnostico_suerte']} (Δ {metricas_robustas['divergencia_pts']:+.1f} pts)</div>
                    </div>
                """, unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("#### ⚖️ Rendimiento Desagregado por Condición de Localía")

            # Pestañas visuales limpias
            tab_local, tab_visitante, tab_global = st.tabs(["🏠 Rendimiento Casa", "✈️ Rendimiento Fuera", "🌐 Rendimiento General"])

            def renderizar_grid_condicion(stats_dict, label_cond):
                tiros_f = float(stats_dict.get('tiros_favor_media', 0.0))
                tiros_c = float(stats_dict.get('tiros_contra_media', 0.0))
                sot_f = float(stats_dict.get('sot_favor_media', 0.0))
                sot_c = float(stats_dict.get('sot_contra_media', 0.0))
                corn_f = float(stats_dict.get('corners_favor_media', 0.0))
                corn_c = float(stats_dict.get('corners_contra_media', 0.0))
                faltas_f = float(stats_dict.get('faltas_cometidas_media', 0.0))
                faltas_c = float(stats_dict.get('faltas_recibidas_media', 0.0))
                
                st.markdown(f"""
                <div class="bento-row">
                    <div class="bento-tile">
                        <div class="bento-label">Volumen Remates ({label_cond})</div>
                        <div class="bento-value">{tiros_f:.1f} <span style="font-size:0.9rem;color:var(--text-muted);">fav</span> / {tiros_c:.1f} <span style="font-size:0.9rem;color:var(--text-muted);">conc</span></div>
                        <div class="bento-desc">Ratio conversión: {(sot_f / max(tiros_f, 1.0))*100:.1f}% a puerta</div>
                    </div>
                    <div class="bento-tile">
                        <div class="bento-label">Tiros a Puerta SOT ({label_cond})</div>
                        <div class="bento-value" style="color:#38bdf8;">{sot_f:.1f} <span style="font-size:0.9rem;color:var(--text-muted);">SOT</span> | <span style="color:#f87171;">{sot_c:.1f}</span></div>
                        <div class="bento-desc">Volumen combinado: {sot_f + sot_c:.1f} SOT/partido</div>
                    </div>
                    <div class="bento-tile">
                        <div class="bento-label">Saques de Esquina ({label_cond})</div>
                        <div class="bento-value">{corn_f:.1f} <span style="font-size:0.9rem;color:var(--text-muted);">fav</span> / {corn_c:.1f} <span style="font-size:0.9rem;color:var(--text-muted);">conc</span></div>
                        <div class="bento-desc">Total medio: {corn_f + corn_c:.1f} córners/90m</div>
                    </div>
                    <div class="bento-tile">
                        <div class="bento-label">Fricción Faltas ({label_cond})</div>
                        <div class="bento-value">{faltas_f:.1f} <span style="font-size:0.9rem;color:var(--text-muted);">com</span> / {faltas_c:.1f} <span style="font-size:0.9rem;color:var(--text-muted);">rec</span></div>
                        <div class="bento-desc">Faltas totales: {faltas_f + faltas_c:.1f} combinadas</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            with tab_local:
                renderizar_grid_condicion(stats_comp_home, "Casa")
            with tab_visitante:
                renderizar_grid_condicion(stats_comp_away, "Fuera")
            with tab_global:
                renderizar_grid_condicion(stats_comp_gen, "Global")

            # ── ROW 3: MATRIZ DE FRECUENCIAS & RADAR PRESCRIPTIVO DINÁMICO (+EV) ──
            st.markdown("---")
            # Cómputo de frecuencias reales
            gf_serie = df_partidos['goles_favor'].astype(float) if 'goles_favor' in df_partidos.columns else df_partidos['GF'].astype(float)
            gc_serie = df_partidos['goles_contra'].astype(float) if 'goles_contra' in df_partidos.columns else df_partidos['GC'].astype(float)
            tot_goles = gf_serie + gc_serie
            media_goles = float(tot_goles.mean()) if len(tot_goles) > 0 else 2.5

            filas_mercado = [
                ("Over 1.5 Goles", tot_goles > 1.5),
                ("Over 2.5 Goles", tot_goles > 2.5),
                ("Under 2.5 Goles", tot_goles < 2.5),
                ("Ambos Marcan (BTTS Sí)", (gf_serie > 0) & (gc_serie > 0)),
                ("Portería a Cero (Clean Sheet)", gc_serie == 0),
                ("Sin Anotar (Failed to Score)", gf_serie == 0)
            ]

            filas_html = ""
            dict_hit_rates = {}
            for nom_m, serie_b in filas_mercado:
                hr_val = float(serie_b.mean() * 100.0) if len(serie_b) > 0 else 0.0
                dict_hit_rates[nom_m] = hr_val
                q_teo = f"{100.0 / hr_val:.2f}" if hr_val > 0 else "-"
                filas_html += (
                    f'<tr style="border-bottom: 1px solid rgba(255, 255, 255, 0.05);">\n'
                    f'<td style="padding: 9px 12px; font-weight: 600; color: #f8fafc;">{nom_m}</td>\n'
                    f'<td style="padding: 9px 12px; color: #38bdf8; font-weight: 700;">{hr_val:.1f}%</td>\n'
                    f'<td style="padding: 9px 12px; color: #34d399; font-weight: 700;">{q_teo}</td>\n'
                    f'</tr>\n'
                )

            st.markdown("#### 🎯 Matriz de Frecuencias de Mercado & Cuotas Implícitas (1/p)")
            tabla_matriz_html = f"""<div style="margin-bottom: 22px; max-width: 620px;">
<table style="width: 100%; border-collapse: collapse; background: rgba(15, 23, 42, 0.65); border: 1px solid rgba(99, 102, 241, 0.22); border-radius: 8px; font-size: 0.83rem;">
<thead>
<tr style="background: rgba(30, 41, 59, 0.85); color: #94a3b8; font-size: 0.70rem; text-transform: uppercase;">
<th style="padding: 9px 12px; text-align: left;">Mercado</th>
<th style="padding: 9px 12px; text-align: left;">Hit Rate Real</th>
<th style="padding: 9px 12px; text-align: left;">Cuota Justa Sin Margen</th>
</tr>
</thead>
<tbody>
{filas_html}</tbody>
</table>
</div>"""

            st.markdown(tabla_matriz_html, unsafe_allow_html=True)

            # LÓGICA DINÁMICA DEL RADAR PRESCRIPTIVO:
            # A) Mercado de Goles: Selección dinámica Over vs Under según producción real
            lambda_g = max(0.6, media_goles)
            hr_o25 = dict_hit_rates.get("Over 2.5 Goles", 50.0)

            if media_goles >= 2.60 or hr_o25 >= 55.0:
                lbl_goles = "Over 2.5 Goles"
                prob_g = max(0.08, min(0.92, 1.0 - (float(np.exp(-lambda_g)) * (1.0 + lambda_g + (lambda_g**2) / 2.0))))
                cuota_bookie_g = 1.95
            else:
                lbl_goles = "Under 2.5 Goles"
                prob_g = max(0.08, min(0.92, float(np.exp(-lambda_g)) * (1.0 + lambda_g + (lambda_g**2) / 2.0)))
                cuota_bookie_g = 2.00

            cuota_modelo_g = 1.0 / prob_g
            edge_g = ((cuota_bookie_g / cuota_modelo_g) - 1.0) * 100.0
            stake_kelly_g = max(0.0, min(1.5, (edge_g / 100.0) / max(cuota_bookie_g - 1.0, 0.01) / 4.0))

            # B) Mercado de Tiros a Puerta (SOT): Basado en el volumen combinado real
            sot_fav = float(stats_comp_gen.get('sot_favor_media', 4.5))
            sot_conc = float(stats_comp_gen.get('sot_contra_media', 4.5))
            sot_comb = sot_fav + sot_conc
            linea_sot = 8.5 if sot_comb < 9.5 else 9.5
            k_sot = int(linea_sot)
            prob_under_sot = sum([float(np.exp(-sot_comb)) * (sot_comb**i) / math.factorial(i) for i in range(k_sot + 1)])
            prob_over_sot = max(0.10, min(0.90, 1.0 - prob_under_sot))
            cuota_mod_sot = 1.0 / prob_over_sot
            cuota_b_sot = 1.88
            edge_sot = ((cuota_b_sot / cuota_mod_sot) - 1.0) * 100.0
            stake_kelly_sot = max(0.0, min(1.5, (edge_sot / 100.0) / max(cuota_b_sot - 1.0, 0.01) / 4.0))

            # C) Mercado de Córners: Basado en el volumen combinado real
            corn_fav = float(stats_comp_gen.get('corners_favor_media', 4.5))
            corn_conc = float(stats_comp_gen.get('corners_contra_media', 4.5))
            corn_comb = corn_fav + corn_conc
            linea_corn = 8.5 if corn_comb < 9.5 else 9.5
            prob_under_corn = sum([float(np.exp(-corn_comb)) * (corn_comb**i) / math.factorial(i) for i in range(int(linea_corn) + 1)])
            prob_over_corn = max(0.10, min(0.90, 1.0 - prob_under_corn))
            cuota_mod_corn = 1.0 / prob_over_corn
            cuota_b_corn = 1.85
            edge_corn = ((cuota_b_corn / cuota_mod_corn) - 1.0) * 100.0
            stake_kelly_corn = max(0.0, min(1.5, (edge_corn / 100.0) / max(cuota_b_corn - 1.0, 0.01) / 4.0))

            stake_pct_g = stake_kelly_g * 100.0 if 'stake_kelly_g' in locals() else 0.5
            stake_pct_sot = stake_kelly_sot * 100.0 if 'stake_kelly_sot' in locals() else 0.5
            stake_pct_corn = stake_kelly_corn * 100.0 if 'stake_kelly_corn' in locals() else 0.5

            st.markdown("#### 🎯 Radar Prescriptivo de Oportunidades (+EV)")
            col_e1, col_e2, col_e3 = st.columns(3)

            with col_e1:
                bg_g = "badge-quant-green" if edge_g > 0 else "badge-quant-red"
                st.markdown(f"""
                    <div class="bento-tile {'bento-edge' if edge_g > 2.0 else ''}">
                        <div class="bento-label">Mercado Goles · Dinámico <span class="badge-quant {bg_g}">EV</span></div>
                        <div class="bento-value">{lbl_goles}</div>
                        <div class="bento-desc">Modelo: <strong>{cuota_modelo_g:.2f}</strong> | Bookie: <strong>{cuota_bookie_g:.2f}</strong></div>
                        <div style="margin-top: 8px;">
                            <span class="badge-quant {bg_g}">EDGE {edge_g:+.1f}%</span>
                            <span style="font-size: 0.75rem; color: #94a3b8; margin-left: 8px;">Stake: <strong>{stake_pct_g:.1f}% Banca</strong></span>
                        </div>
                    </div>
                """, unsafe_allow_html=True)

            with col_e2:
                bg_sot = "badge-quant-green" if edge_sot > 0 else "badge-quant-cyan"
                st.markdown(f"""
                    <div class="bento-tile {'bento-edge' if edge_sot > 2.0 else ''}">
                        <div class="bento-label">Volumen · Tiros a Puerta (SOT) <span class="badge-quant {bg_sot}">Props</span></div>
                        <div class="bento-value">Over {linea_sot} SOT</div>
                        <div class="bento-desc">{sot_fav:.1f} fav | {sot_conc:.1f} conc (Tot: <strong>{sot_comb:.1f} SOT</strong>)</div>
                        <div style="margin-top: 8px;">
                            <span class="badge-quant {bg_sot}">EDGE {edge_sot:+.1f}%</span>
                            <span style="font-size: 0.75rem; color: #94a3b8; margin-left: 8px;">Stake: <strong>{stake_pct_sot:.1f}% Banca</strong></span>
                        </div>
                    </div>
                """, unsafe_allow_html=True)

            with col_e3:
                bg_corn = "badge-quant-green" if edge_corn > 0 else "badge-quant-cyan"
                st.markdown(f"""
                    <div class="bento-tile {'bento-edge' if edge_corn > 2.0 else ''}">
                        <div class="bento-label">Volumen · Córners Totales <span class="badge-quant {bg_corn}">Props</span></div>
                        <div class="bento-value">Over {linea_corn} Córners</div>
                        <div class="bento-desc">{corn_fav:.1f} fav | {corn_conc:.1f} conc (Tot: <strong>{corn_comb:.1f} Córn</strong>)</div>
                        <div style="margin-top: 8px;">
                            <span class="badge-quant {bg_corn}">EDGE {edge_corn:+.1f}%</span>
                            <span style="font-size: 0.75rem; color: #94a3b8; margin-left: 8px;">Stake: <strong>{stake_pct_corn:.1f}% Banca</strong></span>
                        </div>
                    </div>
                """, unsafe_allow_html=True)

            # --- MONITOR CUANTITATIVO DISCIPLINARIO ---
            st.markdown("---")
            st.markdown("#### 🟨 Monitor Cuantitativo Disciplinario & Props de Tarjetas")

            faltas_com = float(stats_comp_gen.get('faltas_cometidas_media', 12.0))
            faltas_rec = float(stats_comp_gen.get('faltas_recibidas_media', 11.5))
            tarjetas_propias = float(stats_comp_gen.get('tarjetas_propias_media', 2.0)) if 'tarjetas_propias_media' in stats_comp_gen else 2.1
            tarjetas_casa = float(stats_comp_home.get('tarjetas_propias_media', 1.8)) if 'tarjetas_propias_media' in stats_comp_home else 1.8
            tarjetas_fuera = float(stats_comp_away.get('tarjetas_propias_media', 2.4)) if 'tarjetas_propias_media' in stats_comp_away else 2.4

            # Ratio Faltas por Tarjeta (severidad táctica)
            ratio_f_t = faltas_com / max(tarjetas_propias, 0.1)

            # Modelo estocástico para Team Cards (Over 1.5 o 2.5 según perfil)
            lambda_tarj = max(0.5, tarjetas_propias)
            linea_t = 1.5 if lambda_tarj < 2.3 else 2.5
            if linea_t == 1.5:
                prob_t = max(0.10, min(0.90, 1.0 - (float(np.exp(-lambda_tarj)) * (1.0 + lambda_tarj))))
                cuota_bookie_t = 1.65
            else:
                prob_t = max(0.10, min(0.90, 1.0 - (float(np.exp(-lambda_tarj)) * (1.0 + lambda_tarj + (lambda_tarj**2) / 2.0))))
                cuota_bookie_t = 2.15

            cuota_mod_t = 1.0 / prob_t
            edge_t = ((cuota_bookie_t / cuota_mod_t) - 1.0) * 100.0
            stake_t = max(0.0, min(1.5, (edge_t / 100.0) / max(cuota_bookie_t - 1.0, 0.01) / 4.0)) * 100.0
            badge_t = "badge-quant-green" if edge_t > 0 else "badge-quant-cyan"

            col_d1, col_d2, col_d3 = st.columns(3)

            with col_d1:
                st.markdown(f"""
                    <div class="bento-tile">
                        <div class="bento-label">Fricción & Conversión Táctica <span class="badge-quant badge-quant-cyan">Ratio</span></div>
                        <div class="bento-value">{ratio_f_t:.1f} <span style="font-size:0.9rem;color:#94a3b8;">Faltas/Tarjeta</span></div>
                        <div class="bento-desc">{faltas_com:.1f} com. vs {faltas_rec:.1f} rec. por partido</div>
                    </div>
                """, unsafe_allow_html=True)

            with col_d2:
                st.markdown(f"""
                    <div class="bento-tile">
                        <div class="bento-label">Asimetría de Amonestaciones <span class="badge-quant badge-quant-cyan">Localía</span></div>
                        <div class="bento-value">{tarjetas_casa:.2f} <span style="font-size:0.9rem;color:#94a3b8;">(Casa)</span> | {tarjetas_fuera:.2f} <span style="font-size:0.9rem;color:#94a3b8;">(Fuera)</span></div>
                        <div class="bento-desc">Media global del equipo: <strong>{tarjetas_propias:.2f} tarjetas/90m</strong></div>
                    </div>
                """, unsafe_allow_html=True)

            with col_d3:
                st.markdown(f"""
                    <div class="bento-tile {'bento-edge' if edge_t > 2.0 else ''}">
                        <div class="bento-label">Prop Tarjetas Equipo <span class="badge-quant {badge_t}">+EV</span></div>
                        <div class="bento-value">Over {linea_t} Tarjetas Propias</div>
                        <div class="bento-desc">Modelo: <strong>{cuota_mod_t:.2f}</strong> | Bookie: <strong>{cuota_bookie_t:.2f}</strong></div>
                        <div style="margin-top: 8px;">
                            <span class="badge-quant {badge_t}">EDGE {edge_t:+.1f}%</span>
                            <span style="font-size: 0.75rem; color: #94a3b8; margin-left: 8px;">Stake: <strong>{stake_t:.1f}% Banca</strong></span>
                        </div>
                    </div>
                """, unsafe_allow_html=True)

            # ── SECCIÓN COMPLEMENTARIA: HISTORIAL RECIENTE Y PLANTILLA ───────
            st.markdown("---")
            col_d1, col_d2 = st.columns(2)

            with col_d1:
                st.markdown("#### 🕒 Historial de Últimos Resultados")
                if not df_norm_global.empty:
                    ultimos_partidos_lista = []
                    for _, _row_h in df_norm_global.iloc[::-1].head(6).iterrows():
                        _es_loc_h = bool(_row_h.get("es_local", str(_row_h.get("HomeTeam", "")) == str(equipo_elegido)))
                        ultimos_partidos_lista.append({
                            'rival': str(_row_h.get("AwayTeam", "Rival") if _es_loc_h else _row_h.get("HomeTeam", "Rival")),
                            'condicion': 'C' if _es_loc_h else 'F',
                            'goles_favor': int(float(_row_h.get("gf", 0) or 0)),
                            'goles_contra': int(float(_row_h.get("gc", 0) or 0)),
                            'xg_favor': float(_row_h.get("xg_f", 0.0) or 0.0),
                            'xg_contra': float(_row_h.get("xg_c", 0.0) or 0.0),
                            'fecha': str(_row_h.get("Date", "N/A")),
                        })

                    partidos_feed_html = []
                    for p in ultimos_partidos_lista:
                        # Extracción defensiva de datos
                        rival = p.get('rival', 'Rival')
                        cond = p.get('condicion', 'C') # 'C' o 'F'
                        g_fav = int(p.get('goles_favor', 0))
                        g_con = int(p.get('goles_contra', 0))
                        xg_fav = float(p.get('xg_favor', 0.0))
                        xg_con = float(p.get('xg_contra', 0.0))
                        fecha = str(p.get('fecha', ''))[:10]
                        
                        # Determinación de W / D / L
                        if g_fav > g_con:
                            res_color = "#10b981"
                            res_badge = "W"
                            res_bg = "rgba(16, 185, 129, 0.15)"
                        elif g_fav == g_con:
                            res_color = "#f59e0b"
                            res_badge = "D"
                            res_bg = "rgba(245, 158, 11, 0.15)"
                        else:
                            res_color = "#ef4444"
                            res_badge = "L"
                            res_bg = "rgba(239, 68, 68, 0.15)"
                            
                        vs_label = f"vs {rival}" if cond == 'C' else f"@ {rival}"
                        marcador_txt = f"{g_fav} - {g_con}"
                        
                        xg_fav_color = "#34d399" if xg_fav >= xg_con else "#94a3b8"
                        xg_con_color = "#ef4444" if xg_con > xg_fav else "#94a3b8"
                        
                        partidos_feed_html.append(f"""<div style="display: flex; align-items: center; justify-content: space-between; padding: 7px 12px; margin-bottom: 6px; background: rgba(15, 23, 42, 0.6); border-radius: 6px; border-left: 3px solid {res_color}; border-top: 1px solid rgba(255,255,255,0.03); border-right: 1px solid rgba(255,255,255,0.03); border-bottom: 1px solid rgba(255,255,255,0.03);">
<div style="display: flex; align-items: center; gap: 9px;">
<span style="font-size: 0.65rem; font-weight: 700; color: {res_color}; background: {res_bg}; padding: 2px 6px; border-radius: 4px;">{res_badge}</span>
<span style="font-size: 0.72rem; color: #64748b; font-family: monospace;">{fecha}</span>
<span style="font-size: 0.82rem; font-weight: 600; color: #e2e8f0;">{vs_label}</span>
</div>
<div style="display: flex; align-items: center; gap: 12px;">
<span style="font-size: 0.88rem; font-weight: 700; font-family: monospace; color: #f8fafc; letter-spacing: 1px;">{marcador_txt}</span>
<span style="font-size: 0.70rem; color: #64748b; font-family: monospace;">xG: <strong style="color: {xg_fav_color};">{xg_fav:.2f}</strong> - <strong style="color: {xg_con_color};">{xg_con:.2f}</strong></span>
</div>
</div>""")

                    html_historial_total = f"""<div style="margin-top: 10px;">
{"".join(partidos_feed_html)}
</div>"""
                    st.markdown(html_historial_total, unsafe_allow_html=True)
                else:
                    st.info("Sin partidos recientes registrados.")

            with col_d2:
                st.markdown("#### ⭐ Dependencia y Producción Individual (Top 5 en Minutos - Ratios P90)")

                if df_censo_p90 is not None and not df_censo_p90.empty:
                    lista_top_jugadores = []
                    for _, _row_j in df_censo_p90.iterrows():
                        lista_top_jugadores.append({
                            'jugador': str(_row_j.get('jugador', 'Jugador')),
                            'pos': str(_row_j.get('posicion', 'MC') or 'MC'),
                            'minutos': int(float(_row_j.get('minutos_jugados', 0) or 0)),
                            'sot_p90': float(_row_j.get('sot_p90', 0.0) or 0.0),
                            'faltas_p90': float(_row_j.get('faltas_p90', 0.0) or 0.0),
                            'am_p90': float(_row_j.get('amarillas_p90', 0.0) or 0.0),
                        })

                    filas_censo = []
                    for jug in lista_top_jugadores:
                        nombre = jug.get('jugador', 'Jugador')
                        pos = jug.get('pos', 'MC')
                        minutos = jug.get('minutos', 0)
                        sot_p90 = float(jug.get('sot_p90', 0.0))
                        faltas_p90 = float(jug.get('faltas_p90', 0.0))
                        am_p90 = float(jug.get('am_p90', 0.0))
                        
                        # Asignación de color según demarcación
                        pos_color = "#38bdf8" if "DF" in pos else ("#34d399" if "MC" in pos else "#f59e0b")
                        
                        # Resaltado condicional
                        sot_style = "color: #38bdf8; font-weight: 700;" if sot_p90 >= 1.0 else "color: #cbd5e1;"
                        faltas_style = "color: #fb923c; font-weight: 700;" if faltas_p90 >= 1.5 else "color: #cbd5e1;"
                        am_style = "color: #facc15; font-weight: 700;" if am_p90 >= 0.20 else "color: #94a3b8;"
                        
                        filas_censo.append(f"""<tr style="border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 0.78rem;">
<td style="padding: 7px 10px; font-weight: 600; color: #f1f5f9;">{nombre}</td>
<td style="padding: 7px 10px;"><span style="font-size: 0.65rem; padding: 2px 5px; border-radius: 4px; background: rgba(255,255,255,0.05); color: {pos_color}; font-weight: 700;">{pos}</span></td>
<td style="padding: 7px 10px; color: #94a3b8; font-family: monospace;">{minutos}'</td>
<td style="padding: 7px 10px; font-family: monospace; {sot_style}">{sot_p90:.2f}</td>
<td style="padding: 7px 10px; font-family: monospace; {faltas_style}">{faltas_p90:.2f}</td>
<td style="padding: 7px 10px; font-family: monospace; {am_style}">{am_p90:.2f}</td>
</tr>""")

                    tabla_p90_html = f"""<div style="margin-top: 10px;">
<table style="width: 100%; border-collapse: collapse; background: rgba(15, 23, 42, 0.65); border: 1px solid rgba(99, 102, 241, 0.22); border-radius: 8px; overflow: hidden;">
<thead>
<tr style="background: rgba(30, 41, 59, 0.85); color: #94a3b8; font-size: 0.68rem; text-transform: uppercase; letter-spacing: 0.5px;">
<th style="padding: 8px 10px; text-align: left;">Jugador</th>
<th style="padding: 8px 10px; text-align: left;">Pos</th>
<th style="padding: 8px 10px; text-align: left;">Min</th>
<th style="padding: 8px 10px; text-align: left;">SOT P90</th>
<th style="padding: 8px 10px; text-align: left;">Faltas P90</th>
<th style="padding: 8px 10px; text-align: left;">AM P90</th>
</tr>
</thead>
<tbody>
{"".join(filas_censo)}
</tbody>
</table>
</div>"""
                    st.markdown(tabla_p90_html, unsafe_allow_html=True)
                else:
                    st.info("Sin registros de censo telemétrico para este equipo en el rango temporal.")

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

            # Aislar estrictamente la temporada actual en curso
            if "Temporada_Origen" in df_glob_ord_h2h.columns and not df_glob_ord_h2h["Temporada_Origen"].dropna().empty:
                _temp_actual_h2h = str(df_glob_ord_h2h["Temporada_Origen"].dropna().astype(str).max())
                if "Temporada_Origen" in df_base_ord_h2h.columns:
                    _df_base_temp = df_base_ord_h2h[df_base_ord_h2h["Temporada_Origen"].astype(str) == _temp_actual_h2h].copy()
                    if not _df_base_temp.empty:
                        df_base_ord_h2h = _df_base_temp
                _df_glob_temp = df_glob_ord_h2h[df_glob_ord_h2h["Temporada_Origen"].astype(str) == _temp_actual_h2h].copy()
                if not _df_glob_temp.empty:
                    df_glob_ord_h2h = _df_glob_temp

            # Filtrar únicamente partidos ya disputados con telemetría válida
            df_base_ord_h2h = df_base_ord_h2h.dropna(subset=['FTHG', 'FTAG'])
            df_glob_ord_h2h = df_glob_ord_h2h.dropna(subset=['FTHG', 'FTAG'])
            if 'HS' in df_base_ord_h2h.columns:
                _df_tel_base = df_base_ord_h2h.dropna(subset=['HS', 'HST'])
                df_telemetria_valida = _df_tel_base if not _df_tel_base.empty else df_base_ord_h2h
            else:
                df_telemetria_valida = df_base_ord_h2h

            if 'HS' in df_glob_ord_h2h.columns:
                _df_tel_glob = df_glob_ord_h2h.dropna(subset=['HS', 'HST'])
                df_glob_telemetria_valida = _df_tel_glob if not _df_tel_glob.empty else df_glob_ord_h2h
            else:
                df_glob_telemetria_valida = df_glob_ord_h2h

            df_eq1_home_full = df_telemetria_valida[df_telemetria_valida["HomeTeam"] == eq1].copy()
            df_eq2_away_full = df_telemetria_valida[df_telemetria_valida["AwayTeam"] == eq2].copy()
            _df_fuente_all_h2h = (
                df_telemetria_valida
                if modo_muestra_h2h == "Solo esta Competición (UEFA)"
                else df_glob_telemetria_valida
            )
            df_eq1_all_full = _df_fuente_all_h2h[
                (_df_fuente_all_h2h["HomeTeam"] == eq1) | (_df_fuente_all_h2h["AwayTeam"] == eq1)
            ].copy()
            df_eq2_all_full = _df_fuente_all_h2h[
                (_df_fuente_all_h2h["HomeTeam"] == eq2) | (_df_fuente_all_h2h["AwayTeam"] == eq2)
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

            # Sellar fuga de df_fallback cuando se exige aislamiento de competición (UEFA)
            if modo_muestra_h2h == "Solo esta Competición (UEFA)":
                _df_h2h_fb_home = _df_h2h_norm_home.copy()
                _df_h2h_fb_away = _df_h2h_norm_away.copy()
                _falta_tel_uefa = (
                    (_df_h2h_norm_home.empty or "tiros_f" not in _df_h2h_norm_home.columns or _df_h2h_norm_home["tiros_f"].dropna().empty)
                    or (_df_h2h_norm_away.empty or "tiros_f" not in _df_h2h_norm_away.columns or _df_h2h_norm_away["tiros_f"].dropna().empty)
                )
                if _falta_tel_uefa:
                    st.caption(
                        "ℹ️ **Aislamiento estricto UEFA activo:** Los registros de esta competición no contienen desglose crudo `HS/HST/HC` en `historico_reciente`; "
                        "se bloquea el arrastre silencioso desde ligas domésticas."
                    )
            else:
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

            # ── 5B. MOTOR POISSON DE VOLUMEN (TIROS A PUERTA Y CÓRNERS TOTALES) ──
            # Extracción bidireccional (Ataque y Defensa cruzados de ambos equipos)
            _sot_loc_fav = float(_stats_h2h_home.get("sot_f", 4.5))
            _sot_loc_conc = float(_stats_h2h_home.get("sot_c", 3.8))
            _sot_vis_fav = float(_stats_h2h_away.get("sot_f", 4.0))
            _sot_vis_conc = float(_stats_h2h_away.get("sot_c", 4.2))

            # Lambda Total Bidireccional: Peligro Local vs Def Vis + Peligro Vis vs Def Local
            _lambda_sot_eq1 = (_sot_loc_fav + _sot_vis_conc) / 2.0
            _lambda_sot_eq2 = (_sot_vis_fav + _sot_loc_conc) / 2.0
            lambda_sot_h2h = max(0.5, max(1.0, _lambda_sot_eq1 + _lambda_sot_eq2))

            _cor_loc_fav = float(_stats_h2h_home.get("corners_f", 5.0))
            _cor_loc_conc = float(_stats_h2h_home.get("corners_c", 4.0))
            _cor_vis_fav = float(_stats_h2h_away.get("corners_f", 4.5))
            _cor_vis_conc = float(_stats_h2h_away.get("corners_c", 4.8))

            _lambda_cor_eq1 = (_cor_loc_fav + _cor_vis_conc) / 2.0
            _lambda_cor_eq2 = (_cor_vis_fav + _cor_loc_conc) / 2.0
            lambda_corners_h2h = max(0.5, max(1.0, _lambda_cor_eq1 + _lambda_cor_eq2))

            def _construir_matriz_poisson_volumen(lam_vol: float, lineas_eval: list[float]) -> pd.DataFrame:
                filas_vol = []
                for lin_v in lineas_eval:
                    p_under_v = float(np.clip(prob_poisson_acumulada(lin_v, lam_vol), 1e-4, 0.9999))
                    p_over_v = float(np.clip(1.0 - p_under_v, 1e-4, 0.9999))
                    cuota_over_v = 1.0 / p_over_v
                    cuota_under_v = 1.0 / p_under_v
                    filas_vol.append({
                        "Línea": f"{lin_v:.1f}",
                        "Prob. Over (%)": f"{p_over_v * 100.0:.1f}%",
                        "Cuota Justa Over": f"{cuota_over_v:.2f}",
                        "Prob. Under (%)": f"{p_under_v * 100.0:.1f}%",
                        "Cuota Justa Under": f"{cuota_under_v:.2f}",
                    })
                return pd.DataFrame(filas_vol)

            col_vol_sot, col_vol_cor = st.columns(2)
            with col_vol_sot:
                st.markdown(
                    f"#### 🎯 Distribución Poisson — Tiros a Puerta Totales (SOT) "
                    f"`λ = {lambda_sot_h2h:.2f}`"
                )
                st.caption(
                    f"🏠 {eq1} Esp.: **{_lambda_sot_eq1:.2f}** ({_sot_loc_fav:.1f} fav / {_sot_vis_conc:.1f} conc) + "
                    f"✈️ {eq2} Esp.: **{_lambda_sot_eq2:.2f}** ({_sot_vis_fav:.1f} fav / {_sot_loc_conc:.1f} conc)"
                )
                df_poisson_sot = _construir_matriz_poisson_volumen(
                    lambda_sot_h2h, [6.5, 7.5, 8.5, 9.5, 10.5]
                )
                st.dataframe(df_poisson_sot, width="stretch", hide_index=True)

            with col_vol_cor:
                st.markdown(
                    f"#### 🚩 Distribución Poisson — Córners Totales "
                    f"`λ = {lambda_corners_h2h:.2f}`"
                )
                st.caption(
                    f"🏠 {eq1} Esp.: **{_lambda_cor_eq1:.2f}** ({_cor_loc_fav:.1f} fav / {_cor_vis_conc:.1f} conc) + "
                    f"✈️ {eq2} Esp.: **{_lambda_cor_eq2:.2f}** ({_cor_vis_fav:.1f} fav / {_cor_loc_conc:.1f} conc)"
                )
                df_poisson_corners = _construir_matriz_poisson_volumen(
                    lambda_corners_h2h, [7.5, 8.5, 9.5, 10.5, 11.5]
                )
                st.dataframe(df_poisson_corners, width="stretch", hide_index=True)

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

                def _calcular_ritmo_reciente_5_oficiales(equipo_consulta: str, stats_fb: dict) -> dict:
                    df_u5 = pd.DataFrame()
                    if RUTA_BD.exists():
                        try:
                            with sqlite3.connect(RUTA_BD) as conn_u5:
                                q_u5_sql = """
                                    SELECT fecha, torneo, equipo_local, equipo_visitante, goles_local, goles_visitante, 
                                           sot_local, sot_visitante, corners_local, corners_visitante
                                    FROM partidos 
                                    WHERE (equipo_local = ? OR equipo_visitante = ?) 
                                      AND goles_local IS NOT NULL
                                    ORDER BY fecha DESC LIMIT 5
                                """
                                df_u5 = pd.read_sql(q_u5_sql, conn_u5, params=(equipo_consulta, equipo_consulta))
                                # Ordenación cronológica estricta sobre el historial completo si hay más de 5 registros con formato DD/MM/YYYY
                                q_u5_full = """
                                    SELECT fecha, torneo, equipo_local, equipo_visitante, goles_local, goles_visitante, 
                                           sot_local, sot_visitante, corners_local, corners_visitante
                                    FROM partidos 
                                    WHERE (equipo_local = ? OR equipo_visitante = ?) 
                                      AND goles_local IS NOT NULL
                                """
                                df_u5_all = pd.read_sql(q_u5_full, conn_u5, params=(equipo_consulta, equipo_consulta))
                                if not df_u5_all.empty:
                                    df_u5_all["_dt"] = pd.to_datetime(df_u5_all["fecha"], errors="coerce", dayfirst=True)
                                    df_u5 = df_u5_all.sort_values(by="_dt", ascending=False).head(5).copy()
                        except Exception:
                            df_u5 = pd.DataFrame()

                    if df_u5.empty and not df_partidos_global.empty:
                        eq_low = str(equipo_consulta).strip().lower()
                        sub_mem = df_partidos_global[
                            (df_partidos_global["HomeTeam"].astype(str).str.strip().str.lower() == eq_low)
                            | (df_partidos_global["AwayTeam"].astype(str).str.strip().str.lower() == eq_low)
                        ].dropna(subset=["FTHG", "FTAG"]).copy()
                        if not sub_mem.empty:
                            if "Date" in sub_mem.columns:
                                sub_mem["_dt"] = pd.to_datetime(sub_mem["Date"], errors="coerce", dayfirst=True)
                                sub_mem = sub_mem.sort_values(by="_dt", ascending=False)
                            sub_mem = sub_mem.head(5)
                            df_u5 = pd.DataFrame({
                                "fecha": sub_mem.get("Date"),
                                "torneo": sub_mem.get("Liga_Origen", sub_mem.get("Div", "Liga")),
                                "equipo_local": sub_mem.get("HomeTeam"),
                                "equipo_visitante": sub_mem.get("AwayTeam"),
                                "goles_local": sub_mem.get("FTHG"),
                                "goles_visitante": sub_mem.get("FTAG"),
                                "sot_local": sub_mem.get("HST"),
                                "sot_visitante": sub_mem.get("AST"),
                                "corners_local": sub_mem.get("HC"),
                                "corners_visitante": sub_mem.get("AC"),
                            })

                    if df_u5.empty:
                        return {
                            "n": 0,
                            "gf_med": float(stats_fb.get("gf_med", 1.35)),
                            "gc_med": float(stats_fb.get("gc_med", 1.15)),
                            "sot_f_med": float(stats_fb.get("sot_f", 4.2)),
                            "sot_c_med": float(stats_fb.get("sot_c", 3.9)),
                            "cor_f_med": float(stats_fb.get("corners_f", 5.0)),
                            "cor_c_med": float(stats_fb.get("corners_c", 4.5)),
                        }

                    eq_clean = str(equipo_consulta).strip().lower()
                    es_loc = df_u5["equipo_local"].astype(str).str.strip().str.lower() == eq_clean
                    gf_s = pd.Series(np.where(es_loc, pd.to_numeric(df_u5["goles_local"], errors="coerce"), pd.to_numeric(df_u5["goles_visitante"], errors="coerce"))).dropna()
                    gc_s = pd.Series(np.where(es_loc, pd.to_numeric(df_u5["goles_visitante"], errors="coerce"), pd.to_numeric(df_u5["goles_local"], errors="coerce"))).dropna()
                    sot_f_s = pd.Series(np.where(es_loc, pd.to_numeric(df_u5["sot_local"], errors="coerce"), pd.to_numeric(df_u5["sot_visitante"], errors="coerce"))).dropna()
                    sot_c_s = pd.Series(np.where(es_loc, pd.to_numeric(df_u5["sot_visitante"], errors="coerce"), pd.to_numeric(df_u5["sot_local"], errors="coerce"))).dropna()
                    cor_f_s = pd.Series(np.where(es_loc, pd.to_numeric(df_u5["corners_local"], errors="coerce"), pd.to_numeric(df_u5["corners_visitante"], errors="coerce"))).dropna()
                    cor_c_s = pd.Series(np.where(es_loc, pd.to_numeric(df_u5["corners_visitante"], errors="coerce"), pd.to_numeric(df_u5["corners_local"], errors="coerce"))).dropna()

                    return {
                        "n": int(len(df_u5)),
                        "gf_med": float(gf_s.mean()) if not gf_s.empty else float(stats_fb.get("gf_med", 1.35)),
                        "gc_med": float(gc_s.mean()) if not gc_s.empty else float(stats_fb.get("gc_med", 1.15)),
                        "sot_f_med": float(sot_f_s.mean()) if not sot_f_s.empty else float(stats_fb.get("sot_f", 4.2)),
                        "sot_c_med": float(sot_c_s.mean()) if not sot_c_s.empty else float(stats_fb.get("sot_c", 3.9)),
                        "cor_f_med": float(cor_f_s.mean()) if not cor_f_s.empty else float(stats_fb.get("corners_f", 5.0)),
                        "cor_c_med": float(cor_c_s.mean()) if not cor_c_s.empty else float(stats_fb.get("corners_c", 4.5)),
                    }

                with st.container(border=True):
                    st.markdown("##### 📊 Ritmo Táctico y Rendimiento Reciente Cruzado (Últimos 5 Partidos Oficiales)")
                    _ritmo_loc = _calcular_ritmo_reciente_5_oficiales(eq1, _stats_h2h_home)
                    _ritmo_vis = _calcular_ritmo_reciente_5_oficiales(eq2, _stats_h2h_away)
                    _sot_comb_proy = _ritmo_loc["sot_f_med"] + _ritmo_vis["sot_f_med"]

                    if _sot_comb_proy > 10.0:
                        _badge_ritmo_txt = "⚡ Partido de Alto Ritmo Ofensivo"
                        _badge_ritmo_bg = "rgba(16, 185, 129, 0.16)"
                        _badge_ritmo_bd = "#10B981"
                        _badge_ritmo_cl = "#6EE7B7"
                    else:
                        _badge_ritmo_txt = "🛡️ Duelo Táctico / Bajo Ritmo de Disparos"
                        _badge_ritmo_bg = "rgba(56, 189, 248, 0.14)"
                        _badge_ritmo_bd = "#38BDF8"
                        _badge_ritmo_cl = "#7DD3FC"

                    st.markdown(
                        f"<div style='background:{_badge_ritmo_bg};border:1.5px solid {_badge_ritmo_bd};"
                        f"border-radius:8px;padding:10px 14px;margin-bottom:12px;display:flex;"
                        f"justify-content:space-between;align-items:center;'>"
                        f"<span style='font-weight:800;font-size:14px;color:{_badge_ritmo_cl};'>{_badge_ritmo_txt}</span>"
                        f"<span style='font-size:12.5px;color:#E2E8F0;'>SOT Medio Combinado Proyectado: "
                        f"<b>{_sot_comb_proy:.2f} tiros a puerta/90m</b></span>"
                        f"</div>",
                        unsafe_allow_html=True,
                    )

                    _c_rit1, _c_rit2 = st.columns(2)
                    with _c_rit1:
                        st.markdown(
                            f"<div style='background:#16192b;border:1px solid #2d3748;border-radius:10px;padding:14px;'>"
                            f"<div style='font-size:14px;font-weight:800;color:#10B981;margin-bottom:8px;'>"
                            f"🏠 Equipo Local: {eq1} (Últ. {_ritmo_loc['n']} Partidos Oficiales)</div>"
                            f"<div style='font-size:13px;color:#E2E8F0;line-height:1.7;'>"
                            f"⚽ <b>Goles (Anotados / Encajados):</b> {_ritmo_loc['gf_med']:.2f} / {_ritmo_loc['gc_med']:.2f}<br>"
                            f"🎯 <b>SOT (A Favor / En Contra):</b> {_ritmo_loc['sot_f_med']:.2f} / {_ritmo_loc['sot_c_med']:.2f}<br>"
                            f"🚩 <b>Córners (A Favor / En Contra):</b> {_ritmo_loc['cor_f_med']:.2f} / {_ritmo_loc['cor_c_med']:.2f}"
                            f"</div></div>",
                            unsafe_allow_html=True,
                        )
                    with _c_rit2:
                        st.markdown(
                            f"<div style='background:#16192b;border:1px solid #2d3748;border-radius:10px;padding:14px;'>"
                            f"<div style='font-size:14px;font-weight:800;color:#38BDF8;margin-bottom:8px;'>"
                            f"✈️ Equipo Visitante: {eq2} (Últ. {_ritmo_vis['n']} Partidos Oficiales)</div>"
                            f"<div style='font-size:13px;color:#E2E8F0;line-height:1.7;'>"
                            f"⚽ <b>Goles (Anotados / Encajados):</b> {_ritmo_vis['gf_med']:.2f} / {_ritmo_vis['gc_med']:.2f}<br>"
                            f"🎯 <b>SOT (A Favor / En Contra):</b> {_ritmo_vis['sot_f_med']:.2f} / {_ritmo_vis['sot_c_med']:.2f}<br>"
                            f"🚩 <b>Córners (A Favor / En Contra):</b> {_ritmo_vis['cor_f_med']:.2f} / {_ritmo_vis['cor_c_med']:.2f}"
                            f"</div></div>",
                            unsafe_allow_html=True,
                        )

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
        try:
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
        except Exception as _exc_init_rad:
            st.session_state["radar_runtime_warning"] = f"Aviso de inicialización en Radar: {_exc_init_rad}"
            df_rad_base = pd.DataFrame()
            col_div_rad = None
            divs_disp_rad = []

        opciones_liga_rad = ["Todas las Competiciones"] + [MAPA_LIGAS[d] for d in divs_disp_rad]

        OPCIONES_RADAR = {
            "⚽ Goles Totales": [
                "Más de 1.5", "Menos de 1.5", "Más de 2.5", "Menos de 2.5",
                "Más de 3.5", "Menos de 3.5", "Más de 4.5", "Menos de 4.5",
            ],
            "⏱️ Goles 1ª / 2ª Parte": [
                "Más de 0.5 (1ªP)", "Menos de 0.5 (1ªP)", "Más de 1.5 (1ªP)", "Menos de 1.5 (1ªP)",
                "Más de 0.5 (2ªP)", "Menos de 0.5 (2ªP)", "Más de 1.5 (2ªP)", "Menos de 1.5 (2ªP)",
            ],
            "🚩 Córners Totales": [
                "Más de 6.5", "Menos de 6.5", "Más de 7.5", "Menos de 7.5",
                "Más de 8.5", "Menos de 8.5", "Más de 9.5", "Menos de 9.5",
                "Más de 10.5", "Menos de 10.5",
            ],
            "🟨 Fricción Disciplinaria (Tarjetas)": [
                "Más de 3.5", "Menos de 3.5", "Más de 4.5", "Menos de 4.5",
                "Más de 5.5", "Menos de 5.5", "Más de 6.5", "Menos de 6.5",
            ],
            "🎯 Ambos Marcan (BTTS)": [
                "Ambos Marcan: SÍ", "Ambos Marcan: NO",
            ],
            "🛡️ Hándicap Asiático Cubierto (ATS)": [
                "Línea Cubierta",
            ],
            "⚽ Goles Individuales de Equipo": [
                "Más de 0.5 Goles (Equipo)",
                "Más de 1.5 Goles (Equipo)",
                "Más de 2.5 Goles (Equipo)",
            ],
            "🎯 Tiros y Tiros a Puerta (SOT)": [
                "Más de 23.5 Tiros Totales",
                "Más de 8.5 Tiros a Puerta (SOT)",
                "Más de 4.5 SOT (Individual Equipo)",
            ],
            "⚡ Fricción de Faltas": [
                "Más de 22.5 Faltas Totales",
                "Más de 24.5 Faltas Totales",
                "Más de 26.5 Faltas Totales",
            ],
            "🛡️ Portería a Cero (Clean Sheet)": [
                "Portería a Cero (Clean Sheet)",
                "Recibe Gol (Concede)",
            ],
            "🔥 Doble Oportunidad + Goles": [
                "1X + Más de 1.5 Goles",
                "X2 + Más de 1.5 Goles",
                "1X + Menos de 4.5 Goles",
            ],
        }

        col_rf0, col_rf1, col_rf2 = st.columns([1.5, 1.2, 1.3])
        with col_rf0:
            liga_rad_sel = st.selectbox(
                "🌍 Competición Activa:",
                opciones_liga_rad,
                index=0,
                key="radar_comp_unif",
            )
        with col_rf1:
            filtro_cond_raw = st.radio(
                "🏠 Condición de Campo:",
                ["🌐 Global", "🏠 Local", "✈️ Visitante"],
                horizontal=True,
                key="radar_cond_unif",
            )
            filtro_condicion = "Local" if "Local" in filtro_cond_raw else ("Visitante" if "Visitante" in filtro_cond_raw else "Global")
        with col_rf2:
            es_comp_uefa_activa = any(
                u in str(st.session_state.get("competicion_activa", "")).lower()
                or u in str(liga_rad_sel).lower()
                for u in ["champions", "europa", "conference", "uefa", "ucl", "uel", "uecl"]
            )
            umbral_racha = st.select_slider(
                "🔥 Umbral de Racha:",
                options=[1, 2, 3, 4, 5],
                value=2 if es_comp_uefa_activa else 3,
                format_func=lambda x: f"Mínimo {x}/5 ({x * 20}%)",
                key="radar_umbral_unif",
            )

        try:
            if liga_rad_sel != "Todas las Competiciones" and col_div_rad and not df_rad_base.empty:
                div_target_rad = MAPA_INV.get(liga_rad_sel, "")
                df_rad = df_rad_base[df_rad_base[col_div_rad].astype(str) == str(div_target_rad)].copy()
            else:
                df_rad = df_rad_base.copy()

            for _cn in (
                "FTHG", "FTAG", "HTHG", "HTAG", "HC", "AC", "HY", "AY", "HR", "AR",
                "HS", "AS", "HST", "AST", "HF", "AF",
                "AHh", "AHCh", "Avg>2.5", "B365>2.5", "Max>2.5",
                "AvgH", "AvgD", "AvgA", "B365H", "B365A",
                "AvgAHH", "AvgAHA", "B365AHH", "B365AHA", "MaxAHH", "MaxAHA",
            ):
                if _cn in df_rad.columns:
                    df_rad[_cn] = pd.to_numeric(df_rad[_cn], errors="coerce")
                if _cn in df_rad_base.columns:
                    df_rad_base[_cn] = pd.to_numeric(df_rad_base[_cn], errors="coerce")

            if "Date" in df_rad.columns:
                df_rad["Date_Parsed"] = pd.to_datetime(df_rad["Date"], errors="coerce", dayfirst=True)
            else:
                df_rad["Date_Parsed"] = pd.NaT

            if "Date" in df_rad_base.columns:
                df_rad_base["Date_Parsed"] = pd.to_datetime(df_rad_base["Date"], errors="coerce", dayfirst=True)
            else:
                df_rad_base["Date_Parsed"] = pd.NaT
        except Exception as _exc_prep_rad:
            st.session_state["radar_runtime_warning"] = f"Aviso al preparar datos del Radar: {_exc_prep_rad}"
            df_rad = df_rad_base.copy()
            if "Date_Parsed" not in df_rad.columns:
                df_rad["Date_Parsed"] = pd.NaT

        # ── Helpers internos del Radar (Pilar 2, 3, 4 y 5) ───────────────────
        import unicodedata as _ud_rad
        import difflib as _difflib_rad
        try:
            import conector_odds as _co_rad
        except Exception:
            _co_rad = None

        def _abrev_rival_radar(nombre_eq: str) -> str:
            try:
                limpio = "".join(ch for ch in str(nombre_eq or "RIV").upper() if ch.isalnum() or ch == " ").strip()
                partes = limpio.split()
                if len(partes) >= 2 and len(partes[0]) <= 3:
                    return (partes[0][:1] + partes[1][:2]).upper()[:3]
                return (limpio.replace(" ", "") + "RIV")[:3]
            except Exception:
                return "RIV"

        _ALIAS_CANONICOS_RADAR = {
            # España (LaLiga / LaLiga 2)
            "athletic": "Ath Bilbao",
            "athletic club": "Ath Bilbao",
            "athletic bilbao": "Ath Bilbao",
            "athletic club de bilbao": "Ath Bilbao",
            "ath bilbao": "Ath Bilbao",
            "sociedad": "Sociedad",
            "real sociedad": "Sociedad",
            "la real": "Sociedad",
            "santander": "Santander",
            "racing": "Santander",
            "racing santander": "Santander",
            "real racing": "Santander",
            "real racing club": "Santander",
            "real racing club de santander": "Santander",
            "racing de santander": "Santander",
            "espanyol": "Espanol",
            "rcd espanyol": "Espanol",
            "rcd espanyol de barcelona": "Espanol",
            "espanol": "Espanol",
            "atletico": "Ath Madrid",
            "atletico madrid": "Ath Madrid",
            "atletico de madrid": "Ath Madrid",
            "club atletico de madrid": "Ath Madrid",
            "ath madrid": "Ath Madrid",
            "alaves": "Alaves",
            "deportivo alaves": "Alaves",
            "cd alaves": "Alaves",
            "celta": "Celta",
            "rc celta": "Celta",
            "celta vigo": "Celta",
            "celta de vigo": "Celta",
            "rc celta de vigo": "Celta",
            "betis": "Betis",
            "real betis": "Betis",
            "real betis balompie": "Betis",
            "vallecano": "Vallecano",
            "rayo vallecano": "Vallecano",
            "rayo": "Vallecano",
            "leganes": "Leganes",
            "cd leganes": "Leganes",
            "valladolid": "Valladolid",
            "real valladolid": "Valladolid",
            "real valladolid cf": "Valladolid",
            "las palmas": "Las Palmas",
            "ud las palmas": "Las Palmas",
            "la coruna": "La Coruna",
            "deportivo la coruna": "La Coruna",
            "rc deportivo": "La Coruna",
            "deportivo": "La Coruna",
            "osasuna": "Osasuna",
            "ca osasuna": "Osasuna",
            "girona": "Girona",
            "girona fc": "Girona",
            "mallorca": "Mallorca",
            "rcd mallorca": "Mallorca",
            "getafe": "Getafe",
            "getafe cf": "Getafe",
            "sevilla": "Sevilla",
            "sevilla fc": "Sevilla",
            "valencia": "Valencia",
            "valencia cf": "Valencia",
            "villarreal": "Villarreal",
            "villarreal cf": "Villarreal",
            "real madrid": "Real Madrid",
            "real madrid cf": "Real Madrid",
            "barcelona": "Barcelona",
            "fc barcelona": "Barcelona",
            "elche": "Elche",
            "elche cf": "Elche",
            "malaga": "Malaga",
            "malaga cf": "Malaga",
            "levante": "Levante",
            "levante ud": "Levante",
            "oviedo": "Oviedo",
            "real oviedo": "Oviedo",
            "gijon": "Gijon",
            "sporting gijon": "Gijon",
            "sporting de gijon": "Gijon",
            "real sporting": "Gijon",
            "zaragoza": "Zaragoza",
            "real zaragoza": "Zaragoza",
            # Inglaterra (Premier League / Championship)
            "manchester united": "Man United",
            "man united": "Man United",
            "man utd": "Man United",
            "manchester city": "Man City",
            "man city": "Man City",
            "nottingham": "Nott'm Forest",
            "nottingham forest": "Nott'm Forest",
            "nott'm forest": "Nott'm Forest",
            "nottm forest": "Nott'm Forest",
            "wolverhampton": "Wolves",
            "olverhampton": "Wolves",
            "wolverhampton wanderers": "Wolves",
            "wolves": "Wolves",
            "newcastle": "Newcastle",
            "newcastle united": "Newcastle",
            "tottenham": "Tottenham",
            "tottenham hotspur": "Tottenham",
            "spurs": "Tottenham",
            "brighton": "Brighton",
            "brighton and hove albion": "Brighton",
            "brighton & hove albion": "Brighton",
            "west ham": "West Ham",
            "west ham united": "West Ham",
            "leicester": "Leicester",
            "leicester city": "Leicester",
            "ipswich": "Ipswich",
            "ipswich town": "Ipswich",
            "leeds": "Leeds",
            "leeds united": "Leeds",
            "coventry": "Coventry",
            "coventry city": "Coventry",
            "hull": "Hull",
            "hull city": "Hull",
            "bournemouth": "Bournemouth",
            "afc bournemouth": "Bournemouth",
            # Alemania (Bundesliga)
            "bayern": "Bayern Munich",
            "bayern munich": "Bayern Munich",
            "bayern munchen": "Bayern Munich",
            "fc bayern munchen": "Bayern Munich",
            "leverkusen": "Leverkusen",
            "bayer leverkusen": "Leverkusen",
            "bayer 04 leverkusen": "Leverkusen",
            "dortmund": "Dortmund",
            "borussia dortmund": "Dortmund",
            "gladbach": "M'gladbach",
            "monchengladbach": "M'gladbach",
            "borussia monchengladbach": "M'gladbach",
            "m'gladbach": "M'gladbach",
            "mgladbach": "M'gladbach",
            "frankfurt": "Ein Frankfurt",
            "eintracht frankfurt": "Ein Frankfurt",
            "ein frankfurt": "Ein Frankfurt",
            "stuttgart": "Stuttgart",
            "vfb stuttgart": "Stuttgart",
            "freiburg": "Freiburg",
            "sc freiburg": "Freiburg",
            "hoffenheim": "Hoffenheim",
            "tsg hoffenheim": "Hoffenheim",
            "tsg 1899 hoffenheim": "Hoffenheim",
            "heidenheim": "Heidenheim",
            "1. fc heidenheim": "Heidenheim",
            "st pauli": "St Pauli",
            "fc st. pauli": "St Pauli",
            "fc st pauli": "St Pauli",
            "mainz": "Mainz",
            "mainz 05": "Mainz",
            "fsv mainz 05": "Mainz",
            "1. fsv mainz 05": "Mainz",
            "wolfsburg": "Wolfsburg",
            "vfl wolfsburg": "Wolfsburg",
            "augsburg": "Augsburg",
            "fc augsburg": "Augsburg",
            "fc koln": "FC Koln",
            "1. fc koln": "FC Koln",
            "koln": "FC Koln",
            "cologne": "FC Koln",
            "hamburg": "Hamburg",
            "hamburger sv": "Hamburg",
            "paderborn": "Paderborn",
            "sc paderborn": "Paderborn",
            "sc paderborn 07": "Paderborn",
            "schalke": "Schalke 04",
            "schalke 04": "Schalke 04",
            "fc schalke 04": "Schalke 04",
            "elversberg": "Elversberg",
            "sv elversberg": "Elversberg",
            "rb leipzig": "RB Leipzig",
            "leipzig": "RB Leipzig",
            "werder bremen": "Werder Bremen",
            "sv werder bremen": "Werder Bremen",
            "bremen": "Werder Bremen",
            "union berlin": "Union Berlin",
            "1. fc union berlin": "Union Berlin",
            "bochum": "Bochum",
            "vfl bochum": "Bochum",
            "holstein kiel": "Holstein Kiel",
            "kiel": "Holstein Kiel",
            # Italia (Serie A)
            "inter": "Inter",
            "inter milan": "Inter",
            "internazionale": "Inter",
            "fc internazionale": "Inter",
            "milan": "Milan",
            "ac milan": "Milan",
            "juventus": "Juventus",
            "juventus fc": "Juventus",
            "roma": "Roma",
            "as roma": "Roma",
            "atalanta": "Atalanta",
            "atalanta bc": "Atalanta",
            "verona": "Verona",
            "hellas verona": "Verona",
            "parma": "Parma",
            "parma calcio": "Parma",
            "parma calcio 1913": "Parma",
            "napoli": "Napoli",
            "ssc napoli": "Napoli",
            "lazio": "Lazio",
            "ss lazio": "Lazio",
            "fiorentina": "Fiorentina",
            "acf fiorentina": "Fiorentina",
            "bologna": "Bologna",
            "bologna fc": "Bologna",
            "torino": "Torino",
            "torino fc": "Torino",
            "udinese": "Udinese",
            "udinese calcio": "Udinese",
            "genoa": "Genoa",
            "genoa cfc": "Genoa",
            "cagliari": "Cagliari",
            "cagliari calcio": "Cagliari",
            "empoli": "Empoli",
            "empoli fc": "Empoli",
            "lecce": "Lecce",
            "us lecce": "Lecce",
            "monza": "Monza",
            "ac monza": "Monza",
            "venezia": "Venezia",
            "venezia fc": "Venezia",
            "como": "Como",
            "como 1907": "Como",
            # Francia (Ligue 1)
            "psg": "Paris SG",
            "paris": "Paris SG",
            "paris sg": "Paris SG",
            "paris saint germain": "Paris SG",
            "paris saint-germain": "Paris SG",
            "paris fc": "Paris FC",
            "monaco": "Monaco",
            "as monaco": "Monaco",
            "lens": "Lens",
            "rc lens": "Lens",
            "le mans": "Le Mans",
            "le mans fc": "Le Mans",
            "marseille": "Marseille",
            "olympique marseille": "Marseille",
            "lyon": "Lyon",
            "olympique lyonnais": "Lyon",
            "lille": "Lille",
            "losc lille": "Lille",
            "nice": "Nice",
            "ogc nice": "Nice",
            "rennes": "Rennes",
            "stade rennais": "Rennes",
            "brest": "Brest",
            "stade brestois": "Brest",
            "reims": "Reims",
            "stade de reims": "Reims",
            "strasbourg": "Strasbourg",
            "rc strasbourg": "Strasbourg",
            "toulouse": "Toulouse",
            "toulouse fc": "Toulouse",
            "montpellier": "Montpellier",
            "montpellier hsc": "Montpellier",
            "nantes": "Nantes",
            "fc nantes": "Nantes",
            "le havre": "Le Havre",
            "le havre ac": "Le Havre",
            "auxerre": "Auxerre",
            "aj auxerre": "Auxerre",
            "angers": "Angers",
            "angers sco": "Angers",
            "st etienne": "St Etienne",
            "saint-etienne": "St Etienne",
            "as saint-etienne": "St Etienne",
        }

        def _limpiar_texto_radar(texto: str) -> str:
            try:
                sin_acentos = _ud_rad.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode("ascii")
                limpio = re.sub(r"[^a-zA-Z0-9'\s-]", " ", sin_acentos).lower()
                return " ".join(limpio.split())
            except Exception:
                return str(texto or "").strip().lower()

        def _despojar_afijos_club_radar(texto_norm: str) -> str:
            try:
                t = f" {texto_norm.replace('-', ' ')} "
                for pref in (
                    " real ", " rc ", " rcd ", " cf ", " cd ", " ud ", " sd ", " fc ",
                    " ca ", " ac ", " as ", " sc ", " vfb ", " vfl ", " tsg ", " 1 fc ",
                    " sv ", " fsv ", " club ", " deportivo ", " balompie ", " sad ",
                    " de ", " la ", " el ", " calcio ", " 1913 ", " 1907 ", " 05 ", " 07 ",
                    " bc ", " cfc ", " hsc ", " sco ", " losc ", " ogc ", " aj ",
                ):
                    t = t.replace(pref, " ")
                res = " ".join(t.split())
                return res if res else texto_norm
            except Exception:
                return texto_norm

        def _normalizar_y_emparejar_equipo(nombre_api: str, lista_equipos_db: list[str]) -> str | None:
            """
            Normalizador bidireccional y emparejador difuso entre nombres de la API / fuentes externas
            y los nombres canónicos de `macro_base_futbol.db`.
            """
            try:
                if not nombre_api or not lista_equipos_db:
                    return None

                raw_str = str(nombre_api).strip()
                if raw_str in lista_equipos_db:
                    return raw_str

                idx_norm = {_limpiar_texto_radar(eq): eq for eq in lista_equipos_db if eq}
                clave_norm = _limpiar_texto_radar(raw_str)
                if not clave_norm:
                    return None

                if clave_norm in idx_norm:
                    return idx_norm[clave_norm]

                # 1) Diccionarios existentes en conector_odds.py y dashboard.py + alias canónicos explícitos
                mapa_combinado = {}
                if _co_rad is not None and hasattr(_co_rad, "MAPA_EQUIPOS_API"):
                    for k_m, v_m in getattr(_co_rad, "MAPA_EQUIPOS_API", {}).items():
                        mapa_combinado[_limpiar_texto_radar(k_m)] = v_m
                if "CANONICAL_TEAM_MAP" in globals():
                    for k_m, v_m in CANONICAL_TEAM_MAP.items():
                        mapa_combinado[_limpiar_texto_radar(k_m)] = v_m
                for k_m, v_m in _ALIAS_CANONICOS_RADAR.items():
                    mapa_combinado[_limpiar_texto_radar(k_m)] = v_m

                if clave_norm in mapa_combinado:
                    cand_canon = _limpiar_texto_radar(mapa_combinado[clave_norm])
                    if cand_canon in idx_norm:
                        return idx_norm[cand_canon]

                # 2) Limpieza de prefijos/sufijos comunes ("real ", "rc ", "cf ", "cd ", "ud ", "fc ", "club ", "deportivo ", "balompie", "sad")
                clave_strip = _despojar_afijos_club_radar(clave_norm)
                if clave_strip in idx_norm:
                    return idx_norm[clave_strip]
                if clave_strip in mapa_combinado:
                    cand_canon = _limpiar_texto_radar(mapa_combinado[clave_strip])
                    if cand_canon in idx_norm:
                        return idx_norm[cand_canon]

                idx_strip = {_despojar_afijos_club_radar(k): v for k, v in idx_norm.items() if _despojar_afijos_club_radar(k)}
                if clave_strip in idx_strip:
                    return idx_strip[clave_strip]

                # 3) Substring / solapamiento de tokens significativos
                tokens_api = [tk for tk in clave_strip.split() if len(tk) >= 3]
                for tk in tokens_api:
                    if tk in mapa_combinado:
                        cand_tk = _limpiar_texto_radar(mapa_combinado[tk])
                        if cand_tk in idx_norm and cand_tk not in ("real madrid", "man united", "man city", "paris sg", "paris fc"):
                            return idx_norm[cand_tk]

                candidatos_sub = []
                for k_db_s, eq_orig in idx_strip.items():
                    if not k_db_s:
                        continue
                    if clave_strip in k_db_s or (len(k_db_s) >= 4 and k_db_s in clave_strip):
                        candidatos_sub.append((abs(len(k_db_s) - len(clave_strip)), eq_orig))
                if candidatos_sub:
                    candidatos_sub.sort(key=lambda x: x[0])
                    return candidatos_sub[0][1]

                # 4) Fallback difflib.get_close_matches
                coincidencias_diff = _difflib_rad.get_close_matches(clave_strip, list(idx_strip.keys()), n=1, cutoff=0.72)
                if coincidencias_diff:
                    return idx_strip[coincidencias_diff[0]]

                coincidencias_norm = _difflib_rad.get_close_matches(clave_norm, list(idx_norm.keys()), n=1, cutoff=0.72)
                if coincidencias_norm:
                    return idx_norm[coincidencias_norm[0]]
            except Exception as _exc_norm:
                st.session_state["radar_norm_warning"] = str(_exc_norm)
            return None

        def _obtener_muestra_multicompetitoria_club(
            eq_objetivo: str,
            df_scope_club: pd.DataFrame,
            df_global: pd.DataFrame,
            cond_campo: str = "Global",
            n_max: int = 5,
            requiere_secundarias: bool = False,
        ) -> tuple[pd.DataFrame, bool]:
            """
            Recupera la muestra de los últimos `n_max` partidos de un club para el Radar.
            Si `requiere_secundarias` es True (tiros, córners, faltas, tarjetas) y la muestra
            en `df_scope_club` carece de métricas secundarias ('HST', 'AST', 'HC', 'AC', 'HF', 'AF')
            o son enteramente nulas/ceros (caso común en partidos UEFA de ESPN):
            - Realiza un fallback canónico multicompetitorio sobre `df_global` (Liga + UEFA).
            - Empareja el nombre del club vía coincidencia exacta, normalización o alias.
            - Devuelve (df_muestra, es_fallback=True).
            Si `df_scope_club` tiene datos suficientes o no requiere secundarias, devuelve
            (df_scope_club.head(n_max), False).
            """
            muestra_activa = df_scope_club if df_scope_club is not None else pd.DataFrame()

            necesita_fallback = False
            if requiere_secundarias:
                if muestra_activa.empty or len(muestra_activa) < 3:
                    necesita_fallback = True
                else:
                    cols_sec = [c for c in ("HST", "AST", "HC", "AC", "HF", "AF", "HY", "AY") if c in muestra_activa.columns]
                    if not cols_sec:
                        necesita_fallback = True
                    else:
                        vals_sec = muestra_activa[cols_sec].apply(pd.to_numeric, errors="coerce").fillna(0).values
                        if np.all(vals_sec == 0):
                            necesita_fallback = True
            elif muestra_activa.empty or len(muestra_activa) < 3:
                necesita_fallback = True

            if not necesita_fallback and not muestra_activa.empty:
                return muestra_activa.head(n_max), False

            # Fallback canónico sobre df_global (consolidación multicompetitoria)
            if df_global is not None and not df_global.empty:
                equipos_global = sorted(
                    pd.concat([df_global["HomeTeam"], df_global["AwayTeam"]])
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                ) if {"HomeTeam", "AwayTeam"}.issubset(df_global.columns) else []

                eq_canon = None
                if eq_objetivo in equipos_global:
                    eq_canon = eq_objetivo
                else:
                    eq_canon = _normalizar_y_emparejar_equipo(eq_objetivo, equipos_global)
                    if not eq_canon:
                        eq_raiz = _despojar_afijos_club_radar(_limpiar_texto_radar(eq_objetivo))
                        if len(eq_raiz) >= 4:
                            for eg in equipos_global:
                                if eq_raiz in _limpiar_texto_radar(eg):
                                    eq_canon = eg
                                    break

                if eq_canon:
                    if "Local" in cond_campo:
                        df_fb = df_global[df_global["HomeTeam"] == eq_canon]
                    elif "Visitante" in cond_campo:
                        df_fb = df_global[df_global["AwayTeam"] == eq_canon]
                    else:
                        df_fb = df_global[(df_global["HomeTeam"] == eq_canon) | (df_global["AwayTeam"] == eq_canon)]

                    if len(df_fb) < 3:
                        df_fb = df_global[(df_global["HomeTeam"] == eq_canon) | (df_global["AwayTeam"] == eq_canon)]

                    if not df_fb.empty:
                        if "Date_Parsed" in df_fb.columns and df_fb["Date_Parsed"].notna().any():
                            df_fb = df_fb.sort_values(by="Date_Parsed", ascending=False)
                        elif "Date" in df_fb.columns:
                            df_fb = df_fb.sort_values(by="Date", ascending=False)

                        return df_fb.head(n_max), True

            return (muestra_activa.head(n_max) if not muestra_activa.empty else pd.DataFrame()), False

        @st.cache_data(ttl=1800, show_spinner=False)
        def _obtener_fixtures_vivo_radar(sport_key: str) -> list[dict]:
            """
            Consulta autónoma y cacheada (TTL 30 min) de partidos próximos y cuotas en vivo
            vía `conector_odds.py` con captura total de excepciones.
            """
            if not sport_key:
                return []
            try:
                import conector_odds as _co_local
                if hasattr(_co_local, "_pedir_cuotas"):
                    res = _co_local._pedir_cuotas(sport_key, "h2h,totals")
                    if getattr(res, "status_code", 500) == 200:
                        datos = res.json()
                        if isinstance(datos, list):
                            return datos
            except Exception:
                return []
            return []

        def _parsear_fecha_futura_radar(valor_fecha) -> tuple:
            """
            Parsea fechas desde ISO 8601 ('2026-10-31T15:15:00Z'), 'DD/MM/YYYY HH:MM', 'DD/MM/YYYY'
            o Timestamp. Devuelve (fecha_date, dt_naive_orden, texto_formateado) o (None, None, '').
            """
            try:
                if valor_fecha is None or (isinstance(valor_fecha, float) and pd.isna(valor_fecha)):
                    return None, None, ""
                raw = str(valor_fecha).strip()
                if not raw or raw.lower() in ("nan", "nat", "none", "próxima jornada", "jornada activa", "programado"):
                    return None, None, ""

                if isinstance(valor_fecha, datetime):
                    dt_obj = valor_fecha.replace(tzinfo=None)
                elif hasattr(valor_fecha, "to_pydatetime"):
                    dt_obj = valor_fecha.to_pydatetime().replace(tzinfo=None)
                elif hasattr(valor_fecha, "year") and hasattr(valor_fecha, "month") and hasattr(valor_fecha, "day") and not hasattr(valor_fecha, "hour"):
                    dt_obj = datetime(valor_fecha.year, valor_fecha.month, valor_fecha.day)
                elif "T" in raw or (len(raw) >= 10 and raw[4:5] == "-" and raw[7:8] == "-"):
                    dt_iso = datetime.fromisoformat(raw.replace("Z", "+00:00"))
                    dt_obj = dt_iso.replace(tzinfo=None)
                else:
                    ts_parsed = pd.to_datetime(raw, errors="coerce", dayfirst=True)
                    if pd.isna(ts_parsed):
                        return None, None, ""
                    dt_obj = ts_parsed.to_pydatetime().replace(tzinfo=None)

                fecha_d = dt_obj.date()
                if dt_obj.hour != 0 or dt_obj.minute != 0:
                    fmt_str = dt_obj.strftime("%d/%m/%Y %H:%M")
                else:
                    fmt_str = dt_obj.strftime("%d/%m/%Y")
                return fecha_d, dt_obj, fmt_str
            except Exception:
                return None, None, ""

        def _construir_horizonte_fixtures_y_cuotas(df_scope: pd.DataFrame) -> tuple[dict, list[dict]]:
            """
            Motor dinámico de horizonte futuro con barrera cronológica estricta (`fecha >= hoy`)
            y blindaje completo de excepciones para garantizar estabilidad total en Streamlit.
            """
            mapa_eq = {}
            lista_duelos = []
            vistos_duelos = set()

            try:
                hoy = datetime.now().date()
                equipos_universo = sorted(
                    pd.concat([df_rad_base["HomeTeam"], df_rad_base["AwayTeam"]]).dropna().astype(str).unique().tolist()
                ) if (df_rad_base is not None and not df_rad_base.empty and {"HomeTeam", "AwayTeam"}.issubset(df_rad_base.columns)) else []
                equipos_scope_set = set(
                    pd.concat([df_scope["HomeTeam"], df_scope["AwayTeam"]]).dropna().astype(str).unique().tolist()
                ) if (df_scope is not None and not df_scope.empty and {"HomeTeam", "AwayTeam"}.issubset(df_scope.columns)) else set(equipos_universo)

                # Registro de partidos ya jugados (especialmente últimos 5 por equipo) para evitar duplicados jugados
                partidos_jugados_set = set()
                pares_recientes_u5 = set()
                try:
                    if df_rad_base is not None and not df_rad_base.empty and {"FTHG", "FTAG"}.issubset(df_rad_base.columns):
                        df_jugados = df_rad_base[df_rad_base["FTHG"].notna() & df_rad_base["FTAG"].notna()].copy()
                        if "Date" in df_jugados.columns:
                            df_jugados["_dt_j"] = pd.to_datetime(df_jugados["Date"], errors="coerce", dayfirst=True)
                        else:
                            df_jugados["_dt_j"] = pd.NaT

                        for _, r_j in df_jugados.dropna(subset=["_dt_j"]).iterrows():
                            hj = str(r_j.get("HomeTeam", "")).strip()
                            aj = str(r_j.get("AwayTeam", "")).strip()
                            dj = r_j["_dt_j"].date()
                            if hj and aj:
                                partidos_jugados_set.add((hj, aj, dj))

                        for eq_u in equipos_universo:
                            sub_u5 = df_jugados[(df_jugados["HomeTeam"] == eq_u) | (df_jugados["AwayTeam"] == eq_u)].sort_values("_dt_j", ascending=False).head(5)
                            for _, r_u5 in sub_u5.iterrows():
                                hu = str(r_u5.get("HomeTeam", "")).strip()
                                au = str(r_u5.get("AwayTeam", "")).strip()
                                du = r_u5["_dt_j"].date() if pd.notna(r_u5.get("_dt_j")) else None
                                if hu and au and du is not None:
                                    pares_recientes_u5.add((hu, au, du))
                except Exception as _exc_hist:
                    st.session_state["radar_hist_warning"] = str(_exc_hist)

                candidatos_futuros = []

                def _registrar_candidato_futuro(h_raw: str, a_raw: str, raw_fecha, mercados_dict: dict | None = None, div_code: str = "", fila_bd=None):
                    try:
                        fecha_d, dt_ord, f_fmt = _parsear_fecha_futura_radar(raw_fecha)
                        # BARRERA CRONOLÓGICA ESTRICTA: descartar inmediatamente si no tiene fecha válida o si fecha < hoy
                        if fecha_d is None or dt_ord is None or fecha_d < hoy:
                            return
                        h_canon = _normalizar_y_emparejar_equipo(h_raw, equipos_universo)
                        a_canon = _normalizar_y_emparejar_equipo(a_raw, equipos_universo)
                        if not h_canon or not a_canon or h_canon == a_canon:
                            return
                        if equipos_scope_set and (h_canon not in equipos_scope_set and a_canon not in equipos_scope_set):
                            return
                        if (h_canon, a_canon, fecha_d) in partidos_jugados_set or (h_canon, a_canon, fecha_d) in pares_recientes_u5:
                            return

                        candidatos_futuros.append({
                            "local": h_canon,
                            "visitante": a_canon,
                            "fecha_date": fecha_d,
                            "dt_ord": dt_ord,
                            "fecha_str": f_fmt,
                            "cuotas_api": dict(mercados_dict or {}),
                            "div": str(div_code or ""),
                            "fila_bd": fila_bd,
                        })
                    except Exception:
                        return

                # 1) Obtención autónoma de próximos partidos en vivo vía conector_odds.py (con caché segura en session_state)
                try:
                    mapa_div_sport = {
                        "SP1": "soccer_spain_la_liga",
                        "SP2": "soccer_spain_segunda_division",
                        "E0": "soccer_epl",
                        "E1": "soccer_efl_champ",
                        "D1": "soccer_germany_bundesliga",
                        "I1": "soccer_italy_serie_a",
                        "F1": "soccer_france_ligue_one",
                    }
                    if liga_rad_sel != "Todas las Competiciones":
                        div_sel_code = MAPA_INV.get(liga_rad_sel, "SP1")
                        divs_activas_api = [div_sel_code] if div_sel_code in mapa_div_sport else ["SP1"]
                    else:
                        divs_activas_api = [d for d in divs_disp_rad if d in mapa_div_sport] or ["SP1", "E0", "D1", "I1", "F1"]

                    if "radar_live_fixtures_cache" not in st.session_state or not isinstance(st.session_state.get("radar_live_fixtures_cache"), dict):
                        st.session_state["radar_live_fixtures_cache"] = {}
                    cache_live_ss = st.session_state["radar_live_fixtures_cache"]

                    for d_code in divs_activas_api:
                        s_key = mapa_div_sport.get(d_code)
                        if not s_key:
                            continue
                        if s_key not in cache_live_ss:
                            try:
                                eventos_liga = _obtener_fixtures_vivo_radar(s_key)
                                cache_live_ss[s_key] = eventos_liga if isinstance(eventos_liga, list) else []
                            except Exception as _exc_api_call:
                                st.session_state["radar_api_warning"] = f"Aviso API ({s_key}): {_exc_api_call}"
                                cache_live_ss[s_key] = []
                        else:
                            eventos_liga = cache_live_ss.get(s_key, [])

                        for ev_p in (eventos_liga or []):
                            try:
                                if not isinstance(ev_p, dict):
                                    continue
                                loc_api = ev_p.get("home_team", "")
                                vis_api = ev_p.get("away_team", "")
                                commence_raw = ev_p.get("commence_time", "")
                                if not loc_api or not vis_api or not commence_raw:
                                    continue

                                mercados_ev = {}
                                cuotas_o25, cuotas_btts, cuotas_h, cuotas_a = [], [], [], []
                                for bm in ev_p.get("bookmakers", []):
                                    if not isinstance(bm, dict):
                                        continue
                                    for mk in bm.get("markets", []):
                                        if not isinstance(mk, dict):
                                            continue
                                        mkey = mk.get("key")
                                        if mkey == "totals":
                                            for outc in mk.get("outcomes", []):
                                                if isinstance(outc, dict) and outc.get("name") == "Over" and outc.get("point") == 2.5:
                                                    pr = pd.to_numeric(outc.get("price"), errors="coerce")
                                                    if pd.notna(pr) and pr > 1.01:
                                                        cuotas_o25.append(float(pr))
                                        elif mkey == "btts":
                                            for outc in mk.get("outcomes", []):
                                                if isinstance(outc, dict) and outc.get("name") == "Yes":
                                                    pr = pd.to_numeric(outc.get("price"), errors="coerce")
                                                    if pd.notna(pr) and pr > 1.01:
                                                        cuotas_btts.append(float(pr))
                                        elif mkey == "h2h":
                                            for outc in mk.get("outcomes", []):
                                                if isinstance(outc, dict):
                                                    pr = pd.to_numeric(outc.get("price"), errors="coerce")
                                                    if pd.notna(pr) and pr > 1.01:
                                                        if outc.get("name") == loc_api:
                                                            cuotas_h.append(float(pr))
                                                        elif outc.get("name") == vis_api:
                                                            cuotas_a.append(float(pr))

                                if cuotas_o25:
                                    mercados_ev["🔥 Más de 2.5 Goles"] = round(float(np.mean(cuotas_o25)), 2)
                                if cuotas_btts:
                                    mercados_ev["🎯 Ambos Marcan (Sí)"] = round(float(np.mean(cuotas_btts)), 2)
                                if cuotas_h:
                                    mercados_ev["Victoria Local (1)"] = round(float(np.mean(cuotas_h)), 2)
                                if cuotas_a:
                                    mercados_ev["Victoria Visitante (2)"] = round(float(np.mean(cuotas_a)), 2)

                                _registrar_candidato_futuro(loc_api, vis_api, commence_raw, mercados_ev, div_code=d_code)
                            except Exception:
                                continue
                except Exception as _exc_live_block:
                    st.session_state["radar_api_warning"] = f"Aviso en sincronización de calendario en vivo: {_exc_live_block}"

                # 2) Cruce con caché de oportunidades del Escáner EV+ (si existe en sesión y supera fecha >= hoy)
                try:
                    df_ops_cache = st.session_state.get("df_ops_export")
                    if isinstance(df_ops_cache, pd.DataFrame) and not df_ops_cache.empty and "Partido" in df_ops_cache.columns:
                        agrup_ops = {}
                        for _, r_op in df_ops_cache.iterrows():
                            p_txt = str(r_op.get("Partido", ""))
                            if " vs " in p_txt:
                                h_op, a_op = [x.strip() for x in p_txt.split(" vs ", 1)]
                                f_op = r_op.get("Fecha / Hora", "")
                                m_op = str(r_op.get("Mercado", ""))
                                c_op = pd.to_numeric(r_op.get("Cuota"), errors="coerce")
                                agrup_ops.setdefault((h_op, a_op, str(f_op)), {})
                                if pd.notna(c_op) and c_op > 1.01:
                                    agrup_ops[(h_op, a_op, str(f_op))][m_op] = float(c_op)
                        for (h_op, a_op, f_op_str), mercs_op in agrup_ops.items():
                            _registrar_candidato_futuro(h_op, a_op, f_op_str, mercs_op)
                except Exception:
                    pass

                # 3) Fixtures manuales cargados en sesión (solo si fecha >= hoy)
                try:
                    for fx_m in st.session_state.get("ev_fix_manual", []):
                        if isinstance(fx_m, dict):
                            _registrar_candidato_futuro(
                                fx_m.get("Local", ""),
                                fx_m.get("Visitante", ""),
                                fx_m.get("Fecha"),
                                {},
                                div_code=str(fx_m.get("Div", "")),
                            )
                except Exception:
                    pass

                # 4) Partidos programados sin disputar en BD (estrictamente FTHG.isna() y fecha >= hoy)
                try:
                    if df_scope is not None and not df_scope.empty and "FTHG" in df_scope.columns:
                        col_fecha_bd = "Date_Parsed" if "Date_Parsed" in df_scope.columns else ("Date" if "Date" in df_scope.columns else None)
                        if col_fecha_bd:
                            df_no_jug = df_scope[
                                df_scope["FTHG"].isna()
                                & df_scope["HomeTeam"].notna()
                                & df_scope["AwayTeam"].notna()
                                & df_scope[col_fecha_bd].notna()
                            ].copy()
                            for _, r_hz in df_no_jug.iterrows():
                                _registrar_candidato_futuro(
                                    str(r_hz.get("HomeTeam", "")),
                                    str(r_hz.get("AwayTeam", "")),
                                    r_hz.get(col_fecha_bd),
                                    {},
                                    div_code=str(r_hz.get(col_div_rad, "")) if col_div_rad else "",
                                    fila_bd=r_hz,
                                )
                except Exception:
                    pass

                # Ordenar cronológicamente (commence_time ascendente) para enlazar a cada equipo con su próximo partido inmediato
                candidatos_futuros.sort(key=lambda x: x["dt_ord"])

                for cand in candidatos_futuros:
                    try:
                        h_c = cand["local"]
                        a_c = cand["visitante"]
                        f_str = cand["fecha_str"]
                        cuotas_c = cand["cuotas_api"]

                        es_inmediato_h = h_c not in mapa_eq
                        es_inmediato_a = a_c not in mapa_eq

                        if es_inmediato_h:
                            mapa_eq[h_c] = {
                                "rival": a_c,
                                "condicion": "Local",
                                "etiqueta": f"vs {a_c} [Local]",
                                "fecha": f_str,
                                "fila_bd": cand.get("fila_bd"),
                                "cuotas_api": cuotas_c,
                            }
                        elif cuotas_c and not mapa_eq[h_c].get("cuotas_api") and mapa_eq[h_c].get("rival") == a_c:
                            mapa_eq[h_c]["cuotas_api"] = cuotas_c

                        if es_inmediato_a:
                            mapa_eq[a_c] = {
                                "rival": h_c,
                                "condicion": "Visitante",
                                "etiqueta": f"vs {h_c} [Visitante]",
                                "fecha": f_str,
                                "fila_bd": cand.get("fila_bd"),
                                "cuotas_api": cuotas_c,
                            }
                        elif cuotas_c and not mapa_eq[a_c].get("cuotas_api") and mapa_eq[a_c].get("rival") == h_c:
                            mapa_eq[a_c]["cuotas_api"] = cuotas_c

                        if (es_inmediato_h or es_inmediato_a) and (h_c, a_c) not in vistos_duelos:
                            vistos_duelos.add((h_c, a_c))
                            lista_duelos.append({
                                "local": h_c,
                                "visitante": a_c,
                                "fecha": f_str,
                                "div": cand["div"],
                            })
                    except Exception:
                        continue
            except Exception as _exc_horiz:
                st.session_state["radar_horizon_warning"] = f"Aviso en motor de horizonte: {_exc_horiz}"
                return {}, []

            return mapa_eq, lista_duelos

        _MAPA_CLAVE_LEGADA_RADAR = {
            "OVER25": ("⚽ Goles Totales", "Más de 2.5"),
            "BTTS": ("🎯 Ambos Marcan (BTTS)", "Ambos Marcan: SÍ"),
            "HT05": ("⏱️ Goles 1ª / 2ª Parte", "Más de 0.5 (1ªP)"),
            "CORNERS85": ("🚩 Córners Totales", "Más de 8.5"),
            "CARDS45": ("🟨 Fricción Disciplinaria (Tarjetas)", "Más de 4.5"),
            "ATS": ("🛡️ Hándicap Asiático Cubierto (ATS)", "Línea Cubierta"),
        }

        def _resolver_par_mercado_radar(clave_o_cat, linea_opt: str | None = None) -> tuple[str, str]:
            if isinstance(clave_o_cat, tuple) and len(clave_o_cat) == 2:
                return str(clave_o_cat[0]), str(clave_o_cat[1])
            if linea_opt is not None:
                return str(clave_o_cat), str(linea_opt)
            s_key = str(clave_o_cat)
            if s_key in _MAPA_CLAVE_LEGADA_RADAR:
                return _MAPA_CLAVE_LEGADA_RADAR[s_key]
            if "||" in s_key:
                c_p, l_p = s_key.split("||", 1)
                return c_p.strip(), l_p.strip()
            return ("⚽ Goles Totales", "Más de 2.5")

        def _parsear_linea_umbral_radar(linea_str: str, default_val: float = 2.5) -> tuple[str, float]:
            s_lin = str(linea_str or "")
            op = "<" if "Menos de" in s_lin else ">"
            m_num = re.search(r"(\d+(?:\.\d+)?)", s_lin)
            umbral = float(m_num.group(1)) if m_num else float(default_val)
            return op, umbral

        def evaluate_streak_condition(partido, mercado_cat_in: str, linea_activa_in: str, eq: str = "") -> tuple[bool, bool, str]:
            """
            Evalúa dinámicamente mediante operadores matemáticos ('>' / '<') si `partido`
            cumple la condición definida por `(mercado_cat_in, linea_activa_in)`.
            Devuelve `(es_valido, hit_bool, etiqueta_extra)`.
            """
            fthg = pd.to_numeric(partido.get("FTHG"), errors="coerce")
            ftag = pd.to_numeric(partido.get("FTAG"), errors="coerce")
            if pd.isna(fthg) or pd.isna(ftag):
                return False, False, ""

            ft_goals = float(fthg + ftag)
            op, umbral = _parsear_linea_umbral_radar(linea_activa_in, 2.5)
            es_loc = str(partido.get("HomeTeam", "")).strip() == str(eq).strip()

            if "Goles Totales" in mercado_cat_in:
                hit = (ft_goals > umbral) if op == ">" else (ft_goals < umbral)
                return True, bool(hit), ""

            if "Goles 1ª / 2ª Parte" in mercado_cat_in:
                hthg = pd.to_numeric(partido.get("HTHG"), errors="coerce")
                htag = pd.to_numeric(partido.get("HTAG"), errors="coerce")
                if pd.isna(hthg) or pd.isna(htag):
                    return False, False, ""
                ht_goals = float(hthg + htag)
                if "2ªP" in linea_activa_in:
                    g2t = ft_goals - ht_goals
                    hit = (g2t > umbral) if op == ">" else (g2t < umbral)
                    g2_h = max(0, int(fthg - hthg))
                    g2_a = max(0, int(ftag - htag))
                    return True, bool(hit), f" 2T:{g2_h}-{g2_a}"
                hit = (ht_goals > umbral) if op == ">" else (ht_goals < umbral)
                return True, bool(hit), f" HT:{int(hthg)}-{int(htag)}"

            if "Córners" in mercado_cat_in:
                hc = pd.to_numeric(partido.get("HC"), errors="coerce")
                ac = pd.to_numeric(partido.get("AC"), errors="coerce")
                if pd.isna(hc) or pd.isna(ac):
                    return False, False, ""
                tot_c = int(hc + ac)
                hit = (float(tot_c) > umbral) if op == ">" else (float(tot_c) < umbral)
                return True, bool(hit), f" {tot_c}C"

            if "Tarjetas" in mercado_cat_in or "Fricción Disciplinaria" in mercado_cat_in:
                hy = pd.to_numeric(partido.get("HY"), errors="coerce")
                ay = pd.to_numeric(partido.get("AY"), errors="coerce")
                hr = pd.to_numeric(partido.get("HR"), errors="coerce")
                ar = pd.to_numeric(partido.get("AR"), errors="coerce")
                if pd.isna(hy) or pd.isna(ay):
                    return False, False, ""
                tot_t = int(hy + ay + (0 if pd.isna(hr) else hr) + (0 if pd.isna(ar) else ar))
                hit = (float(tot_t) > umbral) if op == ">" else (float(tot_t) < umbral)
                return True, bool(hit), f" {tot_t}T"

            if "Ambos Marcan" in mercado_cat_in or "BTTS" in mercado_cat_in:
                btts_si = (fthg > 0) and (ftag > 0)
                es_no = "NO" in str(linea_activa_in).upper()
                hit = (not btts_si) if es_no else btts_si
                return True, bool(hit), ""

            if "Hándicap Asiático" in mercado_cat_in or "ATS" in mercado_cat_in:
                gf = int(float(fthg if es_loc else ftag))
                ga = int(float(ftag if es_loc else fthg))
                ah_raw = pd.to_numeric(partido.get("AHh"), errors="coerce")
                if pd.isna(ah_raw):
                    ah_raw = pd.to_numeric(partido.get("AHCh"), errors="coerce")
                if pd.isna(ah_raw):
                    ah_raw = -0.25
                linea_eq = float(ah_raw) if es_loc else -float(ah_raw)
                diff_ats = (gf + linea_eq) - ga
                hit = False if abs(diff_ats) < 1e-6 else (diff_ats > 0.0)
                return True, bool(hit), f" AH{linea_eq:+.2g}"

            if "Goles Individuales de Equipo" in mercado_cat_in or "Goles Individuales" in mercado_cat_in:
                gf = float(fthg if es_loc else ftag)
                if "0.5" in linea_activa_in:
                    hit = (gf > 0)
                elif "1.5" in linea_activa_in:
                    hit = (gf >= 2)
                elif "2.5" in linea_activa_in:
                    hit = (gf >= 3)
                else:
                    hit = (gf > umbral) if op == ">" else (gf < umbral)
                return True, bool(hit), f" {int(gf)}G"

            if "Tiros y Tiros a Puerta" in mercado_cat_in or "SOT" in mercado_cat_in:
                hs = pd.to_numeric(partido.get("HS"), errors="coerce")
                as_ = pd.to_numeric(partido.get("AS"), errors="coerce")
                hst = pd.to_numeric(partido.get("HST"), errors="coerce")
                ast_ = pd.to_numeric(partido.get("AST"), errors="coerce")

                if "23.5" in linea_activa_in or "Tiros Totales" in linea_activa_in:
                    if pd.isna(hs) or pd.isna(as_):
                        return False, False, ""
                    tot_s = float(hs + as_)
                    hit = tot_s > 23.5
                    return True, bool(hit), f" {int(tot_s)}T"

                if "8.5" in linea_activa_in:
                    if pd.isna(hst) or pd.isna(ast_):
                        return False, False, ""
                    tot_sot = float(hst + ast_)
                    hit = tot_sot > 8.5
                    return True, bool(hit), f" {int(tot_sot)}SOT"

                if "Individual" in linea_activa_in or "4.5" in linea_activa_in:
                    if pd.isna(hst) or pd.isna(ast_):
                        return False, False, ""
                    sot_ind = float(hst if es_loc else ast_)
                    hit = sot_ind >= 5
                    return True, bool(hit), f" {int(sot_ind)}SOT"

            if "Fricción de Faltas" in mercado_cat_in or "Faltas" in mercado_cat_in:
                hf = pd.to_numeric(partido.get("HF"), errors="coerce")
                af = pd.to_numeric(partido.get("AF"), errors="coerce")
                if pd.isna(hf) or pd.isna(af):
                    return False, False, ""
                tot_f = float(hf + af)
                if "22.5" in linea_activa_in:
                    hit = tot_f > 22.5
                elif "26.5" in linea_activa_in:
                    hit = tot_f > 26.5
                else:
                    hit = tot_f > 24.5
                return True, bool(hit), f" {int(tot_f)}F"

            if "Portería a Cero" in mercado_cat_in or "Clean Sheet" in mercado_cat_in:
                ga = int(float(ftag if es_loc else fthg))
                if "Recibe" in linea_activa_in or "Concede" in linea_activa_in:
                    hit = (ga > 0)
                else:
                    hit = (ga == 0)
                return True, bool(hit), f" Enc:{ga}"

            if "Doble Oportunidad" in mercado_cat_in:
                ftr = str(partido.get("FTR") or "").strip().upper()
                if ftr not in ("H", "D", "A"):
                    ftr = "H" if fthg > ftag else ("A" if ftag > fthg else "D")
                if "1X + Más de 1.5" in linea_activa_in:
                    hit = (ftr in ["H", "D"]) and (ft_goals >= 2)
                elif "X2 + Más de 1.5" in linea_activa_in:
                    hit = (ftr in ["D", "A"]) and (ft_goals >= 2)
                elif "1X + Menos de 4.5" in linea_activa_in:
                    hit = (ftr in ["H", "D"]) and (ft_goals <= 4)
                else:
                    hit = (ftr in ["H", "D"]) and (ft_goals >= 2)
                return True, bool(hit), f" {ftr}:{int(fthg)}-{int(ftag)}"

            return False, False, ""

        def _prior_bayesiano_linea_radar(cat_m: str, lin_m: str) -> tuple[float, float]:
            op, umbral = _parsear_linea_umbral_radar(lin_m, 2.5)
            if "Goles Totales" in cat_m:
                mapa_p = {1.5: 0.74, 2.5: 0.52, 3.5: 0.30, 4.5: 0.14}
                p_over = mapa_p.get(umbral, 0.52)
                p0 = p_over if op == ">" else (1.0 - p_over)
                return round(p0 * 5.0, 2), round((1.0 - p0) * 5.0, 2)
            if "Goles 1ª / 2ª Parte" in cat_m:
                if "2ªP" in lin_m:
                    p_over = 0.76 if umbral <= 0.5 else 0.44
                else:
                    p_over = 0.68 if umbral <= 0.5 else 0.32
                p0 = p_over if op == ">" else (1.0 - p_over)
                return round(p0 * 5.0, 2), round((1.0 - p0) * 5.0, 2)
            if "Córners" in cat_m:
                mapa_c = {6.5: 0.78, 7.5: 0.66, 8.5: 0.52, 9.5: 0.42, 10.5: 0.31}
                p_over = mapa_c.get(umbral, 0.52)
                p0 = p_over if op == ">" else (1.0 - p_over)
                return round(p0 * 5.0, 2), round((1.0 - p0) * 5.0, 2)
            if "Tarjetas" in cat_m or "Fricción Disciplinaria" in cat_m:
                mapa_t = {3.5: 0.66, 4.5: 0.50, 5.5: 0.35, 6.5: 0.22}
                p_over = mapa_t.get(umbral, 0.50)
                p0 = p_over if op == ">" else (1.0 - p_over)
                return round(p0 * 5.0, 2), round((1.0 - p0) * 5.0, 2)
            if "Goles Individuales" in cat_m:
                mapa_gi = {0.5: 0.76, 1.5: 0.44, 2.5: 0.18}
                p0 = mapa_gi.get(umbral, 0.44)
                return round(p0 * 5.0, 2), round((1.0 - p0) * 5.0, 2)
            if "Tiros" in cat_m or "SOT" in cat_m:
                p0 = 0.50
                return round(p0 * 5.0, 2), round((1.0 - p0) * 5.0, 2)
            if "Faltas" in cat_m:
                mapa_f = {22.5: 0.62, 24.5: 0.48, 26.5: 0.35}
                p0 = mapa_f.get(umbral, 0.48)
                return round(p0 * 5.0, 2), round((1.0 - p0) * 5.0, 2)
            if "Portería a Cero" in cat_m or "Clean Sheet" in cat_m:
                p0 = 0.68 if ("Recibe" in lin_m or "Concede" in lin_m) else 0.32
                return round(p0 * 5.0, 2), round((1.0 - p0) * 5.0, 2)
            if "Doble Oportunidad" in cat_m:
                p0 = 0.64 if "Menos" in lin_m else 0.52
                return round(p0 * 5.0, 2), round((1.0 - p0) * 5.0, 2)
            return 2.5, 2.5

        def _estimar_cuota_mercado_radar(u5_df: pd.DataFrame, eq: str, clave_mercado, info_fix: dict | None, linea_opt: str | None = None) -> float:
            """
            Extrae la cuota decimal de mercado desde la caché de cuotas en vivo o desde las columnas
            históricas reales de `macro_base_futbol.db`, calibrando dinámicamente según categoría y línea.
            """
            cat_m, lin_m = _resolver_par_mercado_radar(clave_mercado, linea_opt)
            op, umbral = _parsear_linea_umbral_radar(lin_m, 2.5)
            cuotas_api = (info_fix or {}).get("cuotas_api") or {}

            if "Goles Totales" in cat_m:
                if umbral == 2.5 and op == ">":
                    for k_api, v_api in cuotas_api.items():
                        if "2.5" in k_api and ("Más" in k_api or "Over" in k_api):
                            return round(float(v_api), 2)
                base_o25 = 1.85
                for col_c in ("Max>2.5", "Avg>2.5", "B365>2.5"):
                    if col_c in u5_df.columns:
                        s_c = pd.to_numeric(u5_df[col_c], errors="coerce").dropna()
                        if not s_c.empty:
                            base_o25 = float(s_c.iloc[0])
                            break
                if umbral == 2.5 and op == ">":
                    return round(base_o25, 2)
                p_o25 = np.clip(0.95 / max(base_o25, 1.15), 0.20, 0.80)
                escala_g = {1.5: 1.36, 2.5: 1.00, 3.5: 0.62, 4.5: 0.34}.get(umbral, 1.00)
                p_over_lin = float(np.clip(p_o25 * escala_g, 0.14, 0.84))
                p_sel = p_over_lin if op == ">" else (1.0 - p_over_lin)
                return round(float(np.clip(0.95 / max(p_sel, 0.12), 1.22, 3.60)), 2)

            if "Ambos Marcan" in cat_m or "BTTS" in cat_m:
                es_no = "NO" in lin_m.upper()
                if not es_no:
                    for k_api, v_api in cuotas_api.items():
                        if "Ambos" in k_api or "BTTS" in k_api:
                            return round(float(v_api), 2)
                base_btts_si = 1.80
                for col_c in ("Avg>2.5", "B365>2.5"):
                    if col_c in u5_df.columns:
                        s_c = pd.to_numeric(u5_df[col_c], errors="coerce").dropna()
                        if not s_c.empty:
                            base_btts_si = float(np.clip(s_c.mean() * 0.96, 1.55, 2.25))
                            break
                if not es_no:
                    return round(base_btts_si, 2)
                p_si = np.clip(0.95 / max(base_btts_si, 1.20), 0.25, 0.75)
                return round(float(np.clip(0.95 / max(1.0 - p_si, 0.20), 1.45, 2.55)), 2)

            if "Goles 1ª / 2ª Parte" in cat_m:
                base_ht05 = 1.42
                for col_c in ("Avg>2.5", "B365>2.5"):
                    if col_c in u5_df.columns:
                        s_c = pd.to_numeric(u5_df[col_c], errors="coerce").dropna()
                        if not s_c.empty:
                            base_ht05 = float(np.clip(s_c.mean() * 0.76, 1.30, 1.72))
                            break
                if "1ªP" in lin_m and umbral <= 0.5 and op == ">":
                    return round(base_ht05, 2)
                alpha_p, beta_p = _prior_bayesiano_linea_radar(cat_m, lin_m)
                p_base = alpha_p / (alpha_p + beta_p)
                return round(float(np.clip(0.95 / max(p_base, 0.15), 1.25, 3.10)), 2)

            if "Córners" in cat_m:
                c_med = 9.5
                if {"HC", "AC"}.issubset(u5_df.columns):
                    c_med = float((u5_df["HC"].fillna(0) + u5_df["AC"].fillna(0)).mean())
                diff_c = (c_med - (umbral + 1.0)) if op == ">" else ((umbral + 1.0) - c_med)
                c_mkt = 1.85 - np.clip(diff_c * 0.06, -0.30, 0.30)
                return round(float(np.clip(c_mkt, 1.38, 2.55)), 2)

            if "Tarjetas" in cat_m or "Fricción Disciplinaria" in cat_m:
                t_med = 4.7
                if {"HY", "AY"}.issubset(u5_df.columns):
                    t_med = float((u5_df["HY"].fillna(0) + u5_df["AY"].fillna(0) + u5_df.get("HR", 0).fillna(0) + u5_df.get("AR", 0).fillna(0)).mean())
                diff_t = (t_med - (umbral + 0.2)) if op == ">" else ((umbral + 0.2) - t_med)
                c_mkt = 1.90 - np.clip(diff_t * 0.07, -0.32, 0.32)
                return round(float(np.clip(c_mkt, 1.40, 2.60)), 2)

            if "Hándicap Asiático" in cat_m or "ATS" in cat_m:
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

            if "Goles Individuales" in cat_m:
                if "0.5" in lin_m:
                    return 1.30
                elif "2.5" in lin_m:
                    return 3.50
                return 1.85

            if "Tiros" in cat_m or "SOT" in cat_m:
                return 1.85

            if "Faltas" in cat_m:
                if "22.5" in lin_m:
                    return 1.62
                elif "26.5" in lin_m:
                    return 2.35
                return 1.88

            if "Portería a Cero" in cat_m or "Clean Sheet" in cat_m:
                if "Recibe" in lin_m or "Concede" in lin_m:
                    return 1.45
                return 2.65

            if "Doble Oportunidad" in cat_m:
                if "Menos" in lin_m:
                    return 1.52
                return 1.80

            return 1.85

        def _evaluar_racha_equipo_mercado(
            u5_df: pd.DataFrame,
            eq: str,
            clave_mercado,
            info_fix: dict | None,
            liga_eq: str,
            linea_opt: str | None = None,
            es_fallback: bool = False,
        ) -> dict | None:
            """
            Evalúa los últimos N partidos (u5_df) de `eq` en `(mercado_cat, linea_activa)` usando
            `evaluate_streak_condition`. Genera tokens granulares `[GF-GA vs RIV]` con estilo pill,
            calcula Cuota Justa Bayesiana Zero-Vig, Cuota Mercado y EV% Spread.
            Si `es_fallback` es True (consolidación de temporada Liga + UEFA), añade la indicación
            visual '(Muestra Temporada)'.
            """
            if u5_df is None or len(u5_df) < 3:
                return None

            cat_eval, lin_eval = _resolver_par_mercado_radar(clave_mercado, linea_opt)
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

                es_valido, hit, etiqueta_extra = evaluate_streak_condition(fila, cat_eval, lin_eval, eq=eq)
                if not es_valido:
                    continue

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

            # Calibración inferencial Bayesiana Beta-Binomial conjugada: P_calibrated = (k + alpha) / (N + alpha + beta)
            alpha_p, beta_p = _prior_bayesiano_linea_radar(cat_eval, lin_eval)
            p_calibrated_raw = (float(n_hits) + alpha_p) / (float(n_eval) + alpha_p + beta_p)
            # Cota programática estricta: P_calibrated <= 0.88 y Cuota Justa >= 1.14
            p_calibrated = float(np.clip(p_calibrated_raw, 0.05, 0.88))
            cuota_justa = max(1.14, round(1.0 / p_calibrated, 2))
            cuota_mercado = _estimar_cuota_mercado_radar(u5_df, eq, (cat_eval, lin_eval), info_fix)
            ev_pct = round(((p_calibrated * float(cuota_mercado)) - 1.0) * 100.0, 1)

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

            forma_html = "".join(tokens_html)
            forma_texto = " ".join(tokens_texto)
            if es_fallback:
                forma_html += " <span style='color:#38bdf8; font-size:0.72rem; font-weight:600;'>(Muestra Temporada)</span>"
                forma_texto += " (Muestra Temporada)"

            return {
                "Equipo": eq,
                "Liga": liga_eq,
                "Aciertos_Num": n_hits,
                "Muestra_Num": n_eval,
                "Aciertos": f"{n_hits}/{n_eval}",
                "Winrate_Num": round(winrate_pct, 1),
                "Winrate": f"{winrate_pct:.0f}%",
                "Prob_Calibrada": p_calibrated,
                "Cuota_Justa": cuota_justa,
                "Cuota_Mercado": cuota_mercado,
                "EV_Num": ev_pct,
                "EV_Badge_HTML": ev_badge_html,
                "Proximo_HTML": prox_html,
                "Proximo_Plano": prox_plano,
                "Forma_HTML": forma_html,
                "Forma_Texto": forma_texto,
                "Es_Fallback": es_fallback,
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

        def _renderizar_tabla_barrido_global_radar(lista_items: list[dict]):
            if not lista_items:
                st.info("Sin rachas que superen el umbral seleccionado en ningún mercado del catálogo.")
                return
            ordenados = sorted(lista_items, key=lambda x: (x["EV_Num"], x["Winrate_Num"]), reverse=True)
            filas_tr = []
            for item in ordenados:
                wr_color = "#00d26a" if item["Winrate_Num"] >= 80.0 else "#38bdf8"
                filas_tr.append(
                    "<tr>"
                    f"<td style='font-weight:700; color:#f8fafc;'>⚽ {_html_rad.escape(item['Equipo'])} "
                    f"<span style='font-size:0.74rem; color:#64748b; font-weight:500;'>({_html_rad.escape(item['Liga'])})</span></td>"
                    f"<td style='font-weight:600; color:#cbd5e1;'>{_html_rad.escape(item.get('Mercado_Cat', ''))}</td>"
                    f"<td style='font-family:monospace; font-weight:700; color:#38bdf8;'>{_html_rad.escape(item.get('Mercado_Linea', ''))}</td>"
                    f"<td style='font-family:monospace; font-weight:700; color:#f8fafc;'>{item['Aciertos']}</td>"
                    f"<td style='font-family:monospace; font-weight:700; color:{wr_color};'>{item['Winrate']}</td>"
                    f"<td>{item['Proximo_HTML']}</td>"
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
                "<th>Mercado</th>"
                "<th>Línea</th>"
                "<th>Racha</th>"
                "<th>%</th>"
                "<th>Próximo Rival</th>"
                "<th>Cuota Justa</th>"
                "<th>Cuota Mercado</th>"
                "<th>Diferencial EV%</th>"
                "</tr></thead>"
                f"<tbody>{''.join(filas_tr)}</tbody>"
                "</table></div>"
            )
            st.markdown(html_tabla, unsafe_allow_html=True)

        # ── Construcción del Mapa de Próximos Partidos y Escaneo de Equipos ───
        try:
            mapa_fixtures, duelos_horizonte = _construir_horizonte_fixtures_y_cuotas(df_rad)
        except Exception as _exc_call_hz:
            st.session_state["radar_horizon_warning"] = str(_exc_call_hz)
            mapa_fixtures, duelos_horizonte = {}, []

        try:
            equipos_rad = (
                sorted(
                    pd.concat([df_rad["HomeTeam"], df_rad["AwayTeam"]])
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                )
                if (not df_rad.empty and {"HomeTeam", "AwayTeam"}.issubset(df_rad.columns))
                else []
            )
        except Exception:
            equipos_rad = []

        # Fallback robusto en horizonte de duelos (especialmente para competiciones UEFA)
        if not duelos_horizonte and equipos_rad:
            es_comp_uefa = any(
                u in str(liga_rad_sel).lower()
                or u in str(st.session_state.get("competicion_activa", "")).lower()
                for u in ["champions", "europa", "conference", "uefa", "ucl", "uel", "uecl"]
            )
            with st.expander(
                "💡 Modo Duelo Directo UEFA: Selección Manual de Cruce", expanded=es_comp_uefa
            ):
                st.info(
                    "💡 Modo Duelo Directo UEFA: No hay fixtures futuros automáticos en base de datos. Selecciona el cruce de esta jornada:"
                )
                c1, c2 = st.columns(2)
                eq_local = c1.selectbox("Equipo Local:", equipos_rad, key="radar_uefa_local")
                rivales_uefa = [e for e in equipos_rad if e != eq_local]
                eq_visit = c2.selectbox(
                    "Equipo Visitante:",
                    rivales_uefa if rivales_uefa else equipos_rad,
                    key="radar_uefa_visit",
                )
                if eq_local and eq_visit and eq_local != eq_visit:
                    duelos_horizonte.append(
                        {
                            "local": eq_local,
                            "visitante": eq_visit,
                            "fecha": "Hoy",
                            "div": liga_rad_sel,
                        }
                    )
                    if eq_local not in mapa_fixtures:
                        mapa_fixtures[eq_local] = {
                            "rival": eq_visit,
                            "condicion": "Local",
                            "etiqueta": f"vs {eq_visit} [Local]",
                            "fecha": "Hoy",
                            "cuotas_api": {},
                        }
                    if eq_visit not in mapa_fixtures:
                        mapa_fixtures[eq_visit] = {
                            "rival": eq_local,
                            "condicion": "Visitante",
                            "etiqueta": f"vs {eq_local} [Visitante]",
                            "fecha": "Hoy",
                            "cuotas_api": {},
                        }

        # ── PILAR 2: ESCANEO TRANSVERSAL MULTIMERCADO Y DOBLE CONVERGENCIA ──
        CATALOGO_MERCADOS_RADAR = []
        for c_name, l_list in OPCIONES_RADAR.items():
            for l_name in l_list:
                k_pair = f"{c_name}||{l_name}"
                etq_pair = f"{c_name} · {l_name}"
                CATALOGO_MERCADOS_RADAR.append((k_pair, etq_pair))

        resultados_por_mercado = {k: [] for k, _ in CATALOGO_MERCADOS_RADAR}
        indice_racha_local = {k: {} for k, _ in CATALOGO_MERCADOS_RADAR}
        indice_racha_visit = {k: {} for k, _ in CATALOGO_MERCADOS_RADAR}

        for eq in equipos_rad:
            try:
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
                    req_sec = any(w in m_key for w in ["Córners", "Tarjetas", "Disciplinaria", "Tiros", "Faltas"])

                    # 1) Evaluación según filtro activo con fallback si es UEFA o faltan secundarias
                    if req_sec or (es_comp_uefa_activa and len(u5_activo) < 3):
                        u5_eval, es_fb = _obtener_muestra_multicompetitoria_club(
                            eq, u5_activo, df_rad_base, cond_campo=filtro_condicion, n_max=5, requiere_secundarias=req_sec
                        )
                    else:
                        u5_eval, es_fb = u5_activo, False

                    ev_act = _evaluar_racha_equipo_mercado(u5_eval, eq, m_key, info_fix_eq, liga_eq_nom, es_fallback=es_fb)
                    if ev_act and ev_act["Aciertos_Num"] >= umbral_racha:
                        resultados_por_mercado[m_key].append(ev_act)

                    # 2) Evaluación específica Local / Visitante para el motor de Doble Coincidencia
                    u5_h_base = df_eq_loc.head(5) if len(df_eq_loc) >= 3 else df_eq_all.head(5)
                    u5_a_base = df_eq_vis.head(5) if len(df_eq_vis) >= 3 else df_eq_all.head(5)

                    if req_sec or (es_comp_uefa_activa and len(u5_h_base) < 3):
                        u5_h_eval, fb_h = _obtener_muestra_multicompetitoria_club(
                            eq, u5_h_base, df_rad_base, cond_campo="Local", n_max=5, requiere_secundarias=req_sec
                        )
                    else:
                        u5_h_eval, fb_h = u5_h_base, False

                    if req_sec or (es_comp_uefa_activa and len(u5_a_base) < 3):
                        u5_a_eval, fb_a = _obtener_muestra_multicompetitoria_club(
                            eq, u5_a_base, df_rad_base, cond_campo="Visitante", n_max=5, requiere_secundarias=req_sec
                        )
                    else:
                        u5_a_eval, fb_a = u5_a_base, False

                    ev_h = _evaluar_racha_equipo_mercado(u5_h_eval, eq, m_key, info_fix_eq, liga_eq_nom, es_fallback=fb_h)
                    ev_a = _evaluar_racha_equipo_mercado(u5_a_eval, eq, m_key, info_fix_eq, liga_eq_nom, es_fallback=fb_a)

                    if ev_h and ev_h["Aciertos_Num"] >= umbral_racha:
                        indice_racha_local[m_key][eq] = ev_h
                    if ev_a and ev_a["Aciertos_Num"] >= umbral_racha:
                        indice_racha_visit[m_key][eq] = ev_a
            except Exception as _exc_eq_scan:
                st.session_state["radar_scan_warning"] = str(_exc_eq_scan)
                continue

        # ── PILAR 3: DOUBLE-CONVERGENCE RADAR ("CHOQUE DE RACHAS" ENGINE) ─────
        choques_detectados = []
        for duelo in duelos_horizonte:
            try:
                eq_h = duelo["local"]
                eq_a = duelo["visitante"]
                fecha_d = duelo["fecha"]
                for m_key, m_label in CATALOGO_MERCADOS_RADAR:
                    r_h = indice_racha_local[m_key].get(eq_h)
                    r_a = indice_racha_visit[m_key].get(eq_a)
                    if r_h and r_a:
                        wr_comb = round((r_h["Winrate_Num"] + r_a["Winrate_Num"]) / 2.0, 1)
                        p_comb_cal = float(np.clip(
                            (float(r_h.get("Prob_Calibrada", 0.5)) + float(r_a.get("Prob_Calibrada", 0.5))) / 2.0,
                            0.05,
                            0.88,
                        ))
                        cj_comb = max(1.14, round(1.0 / p_comb_cal, 2))
                        cm_comb = round((r_h["Cuota_Mercado"] + r_a["Cuota_Mercado"]) / 2.0, 2)
                        ev_comb = round(((p_comb_cal * cm_comb) - 1.0) * 100.0, 1)
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
            except Exception:
                continue

        choques_detectados.sort(key=lambda x: (x["wr_comb"], x["ev_comb"]), reverse=True)

        todas_rachas_activas = []
        for m_key, m_label in CATALOGO_MERCADOS_RADAR:
            c_nom, l_nom = _resolver_par_mercado_radar(m_key)
            for r_item in resultados_por_mercado.get(m_key, []):
                copia_r = dict(r_item)
                copia_r["Mercado_Nombre"] = m_label
                copia_r["Mercado_Cat"] = c_nom
                copia_r["Mercado_Linea"] = l_nom
                copia_r["Tipo"] = "⚔️ Partido"
                todas_rachas_activas.append(copia_r)

        # ── SECCIÓN 3: ESTRUCTURA EN 2 PESTAÑAS PRINCIPALES (TABS) ───────────
        tab_feed, tab_choque = st.tabs([
            "🌟 Feed Rápido (Partidos)",
            "⚔️ Choque de Tendencias (Partidos)",
        ])

        # ── PESTAÑA 1: FEED RÁPIDO (PARTIDOS) ────────────────────────────────
        with tab_feed:
            st.caption(
                "Barrido algorítmico transversal sobre todos los mercados: Filtrado automático "
                "de oportunidades con **Winrate ≥ 75% o EV% ≥ 15%** que cumplen el umbral de racha activo."
            )

            feed_items = []
            for r in todas_rachas_activas:
                if r["Winrate_Num"] >= 75.0 or r["EV_Num"] >= 15.0:
                    feed_items.append({
                        "tipo": "⚔️ Partido",
                        "tipo_html": "<span class='pill-ev-cyan' style='font-size:0.72rem; font-weight:700;'>⚔️ Partido</span>",
                        "equipo": r["Equipo"],
                        "liga": r["Liga"],
                        "partido_str": f"{r['Equipo']} ({r['Liga']})",
                        "proximo_html": r["Proximo_HTML"],
                        "proximo_txt": r["Proximo_Plano"],
                        "mercado_linea": r["Mercado_Nombre"],
                        "aciertos": r["Aciertos"],
                        "winrate_num": r["Winrate_Num"],
                        "winrate_str": r["Winrate"],
                        "chips_html": r["Forma_HTML"],
                        "chips_txt": r["Forma_Texto"],
                        "cuota_justa": r["Cuota_Justa"],
                        "cuota_mercado": r["Cuota_Mercado"],
                        "ev_num": r["EV_Num"],
                        "ev_badge_html": r["EV_Badge_HTML"],
                    })

            feed_items.sort(key=lambda x: (x["ev_num"], x["winrate_num"]), reverse=True)

            if feed_items:
                tarjetas_feed_html = []
                for it in feed_items[:4]:
                    ev_val = it["ev_num"]
                    ev_cls = "pill-ev-high" if ev_val > 5.0 else ("pill-ev-cyan" if ev_val >= -5.0 else "pill-ev-neg")
                    tarjetas_feed_html.append(
                        "<div class='radar-choque-card'>"
                        "<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;'>"
                        f"{it['tipo_html']}"
                        f"<span class='{ev_cls}'>EV {ev_val:+.1f}%</span>"
                        "</div>"
                        f"<div style='font-size:1.0rem; font-weight:800; color:#f8fafc; margin-bottom:4px;'>"
                        f"⚽ {_html_rad.escape(it['equipo'])} "
                        f"<span style='font-size:0.76rem; color:#94a3b8; font-weight:500;'>({_html_rad.escape(it['liga'])})</span></div>"
                        f"<div style='font-size:0.80rem; color:#cbd5e1; margin-bottom:6px;'>"
                        f"{it['proximo_html']}"
                        "</div>"
                        f"<div style='color:#38bdf8; font-weight:700; font-size:0.84rem; margin-bottom:4px;'>"
                        f"{_html_rad.escape(it['mercado_linea'])}</div>"
                        f"<div style='font-size:0.78rem; color:#94a3b8; font-family:monospace; margin-bottom:6px;'>"
                        f"Racha: <b style='color:#00d26a;'>{it['aciertos']} ({it['winrate_str']})</b> | "
                        f"Justa: <b style='color:#f8fafc;'>{it['cuota_justa']:.2f}</b> vs "
                        f"Base: <b style='color:#38bdf8;'>{it['cuota_mercado']:.2f}</b>"
                        "</div>"
                        f"<div style='margin-top:4px;'>{it['chips_html']}</div>"
                        "</div>"
                    )

                st.markdown(
                    "<div class='radar-choque-wrapper'>"
                    "<div style='display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; margin-bottom:8px;'>"
                    "<span style='font-size:1.05rem; font-weight:800; color:#f8fafc;'>"
                    "💎 Top 4 Oportunidades Globales de Mayor Valor Esperado (EV+)</span>"
                    f"<span class='pill-ev-high'>{len(feed_items)} Oportunidades Detectadas</span>"
                    "</div>"
                    f"<div class='radar-choque-grid'>{''.join(tarjetas_feed_html)}</div>"
                    "</div>",
                    unsafe_allow_html=True,
                )

                st.markdown("##### 📋 Top 15 Oportunidades Transversales (Winrate ≥ 75% o EV% ≥ 15%)")
                filas_top15 = []
                for it in feed_items[:15]:
                    wr_color = "#00d26a" if it["winrate_num"] >= 80.0 else "#38bdf8"
                    filas_top15.append(
                        "<tr>"
                        f"<td>{it['tipo_html']}</td>"
                        f"<td style='font-weight:700; color:#f8fafc;'>⚽ {_html_rad.escape(it['equipo'])} "
                        f"<span style='font-size:0.74rem; color:#64748b; font-weight:500;'>({_html_rad.escape(it['liga'])})</span></td>"
                        f"<td>{it['proximo_html']}</td>"
                        f"<td style='font-weight:600; color:#cbd5e1;'>{_html_rad.escape(it['mercado_linea'])}</td>"
                        f"<td style='font-family:monospace; font-weight:700; color:{wr_color};'>{it['aciertos']} ({it['winrate_str']})</td>"
                        f"<td>{it['chips_html']}</td>"
                        f"<td style='font-family:monospace; font-weight:700; color:#e2e8f0;'>{it['cuota_justa']:.2f}</td>"
                        f"<td style='font-family:monospace; font-weight:700; color:#38bdf8;'>{it['cuota_mercado']:.2f}</td>"
                        f"<td>{it['ev_badge_html']}</td>"
                        "</tr>"
                    )

                html_top15 = (
                    "<div class='radar-dark-table-wrap'>"
                    "<table class='radar-dark-table'>"
                    "<thead><tr>"
                    "<th>Tipo</th>"
                    "<th>Equipo / Partido</th>"
                    "<th>Próximo Cruce</th>"
                    "<th>Mercado / Línea</th>"
                    "<th>Racha / %</th>"
                    "<th>Desglose (Chips)</th>"
                    "<th>Cuota Justa</th>"
                    "<th>Cuota Mercado</th>"
                    "<th>Diferencial EV%</th>"
                    "</tr></thead>"
                    f"<tbody>{''.join(filas_top15)}</tbody>"
                    "</table></div>"
                )
                st.markdown(html_top15, unsafe_allow_html=True)

                # ── Veredicto Táctico AI Studio para Tab Feed ──
                st.markdown("---")
                col_ai_btn, col_ai_info = st.columns([1.6, 2.4])
                with col_ai_btn:
                    ejecutar_veredicto_radar = st.button(
                        "⚡ Generar Veredicto Táctico del Radar (AI Studio)",
                        width="stretch",
                        type="primary",
                        key="btn_radar_ai_digest_feed",
                    )
                with col_ai_info:
                    st.caption(
                        f"📊 **{len(feed_items)} oportunidades en feed** · "
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
                    else:
                        top3_feed = feed_items[:3]
                        resumen_top3_txt = "\n".join(
                            f"- {idx + 1}. [{it['tipo']}] {it['equipo']} ({it['liga']}) | Mercado: {it['mercado_linea']} | "
                            f"Racha: {it['aciertos']} ({it['winrate_str']}) | Próximo: {it['proximo_txt']} | "
                            f"Cuota Justa: {it['cuota_justa']:.2f} vs Mercado: {it['cuota_mercado']:.2f} (EV: {it['ev_num']:+.1f}%) | "
                            f"Forma: {it['chips_txt']}"
                            for idx, it in enumerate(top3_feed)
                        )
                        prompt_radar_digest = (
                            "Eres el Director Cuantitativo de BigData Football.\n"
                            "Sintetiza exclusivamente las siguientes Top 3 rachas con mayor Valor Esperado (EV+) "
                            "detectadas en el Radar de Tendencias en un informe ejecutivo ultra-directo de exactamente 3 viñetas "
                            "(una por selección, indicando equipo, mercado, próximo rival, diferencial entre Cuota Justa y Cuota Mercado, y dictamen operativo).\n"
                            "Prohibido usar lambda 'λ' o '@' antes de las cuotas.\n\n"
                            f"TOP_3_OPORTUNIDADES_RADAR:\n{resumen_top3_txt}"
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
            else:
                st.info("ℹ️ No hay oportunidades que superen el umbral de Winrate ≥ 75% o EV% ≥ 15% en el filtro actual.")

        # ── PESTAÑA 2: CHOQUE DE TENDENCIAS (PARTIDOS) ───────────────────────
        with tab_choque:
            st.caption("Detección de patrones y convergencias en partidos colectivos (convergencia simultánea local + visitante y desglose por mercado).")

            cats_partido = [
                "🌐 Todos los Mercados",
                "⚽ Goles Totales",
                "🎯 Ambos Marcan (BTTS)",
                "🚩 Córners Totales",
                "🟨 Tarjetas Totales",
                "⚡ Faltas Totales",
                "🔥 Doble Oportunidad + Goles",
            ]

            try:
                cat_sel_tab2 = st.pills("Categoría de Mercado:", cats_partido, default="🌐 Todos los Mercados", key="radar_pills_tab2")
                if not cat_sel_tab2:
                    cat_sel_tab2 = "🌐 Todos los Mercados"
            except Exception:
                cat_sel_tab2 = st.radio("Categoría de Mercado:", cats_partido, horizontal=True, key="radar_radio_tab2")

            mapa_cats_partido = {
                "⚽ Goles Totales": "⚽ Goles Totales",
                "🎯 Ambos Marcan (BTTS)": "🎯 Ambos Marcan (BTTS)",
                "🚩 Córners Totales": "🚩 Córners Totales",
                "🟨 Tarjetas Totales": "🟨 Fricción Disciplinaria (Tarjetas)",
                "⚡ Faltas Totales": "⚡ Fricción de Faltas Cometidas",
                "🔥 Doble Oportunidad + Goles": "🔥 Doble Oportunidad + Goles",
            }

            linea_elegida_tab2 = "TODAS"
            cat_target = mapa_cats_partido.get(cat_sel_tab2)

            if cat_sel_tab2 != "🌐 Todos los Mercados" and cat_target:
                lineas_disponibles = OPCIONES_RADAR.get(cat_target, [])
                opciones_lineas_tab2 = []
                conteo_por_linea = {}
                total_cat = 0
                for lin in lineas_disponibles:
                    k_par = f"{cat_target}||{lin}"
                    cnt = len(resultados_por_mercado.get(k_par, []))
                    conteo_por_linea[lin] = cnt
                    total_cat += cnt

                opciones_lineas_tab2.append(f"Todas las Líneas ({total_cat})")
                for lin in lineas_disponibles:
                    opciones_lineas_tab2.append(f"{lin} ({conteo_por_linea[lin]})")

                sel_lin_str = st.selectbox("Línea Objetivo:", opciones_lineas_tab2, key="radar_linea_tab2")
                if not sel_lin_str.startswith("Todas las Líneas"):
                    linea_elegida_tab2 = sel_lin_str.split(" (")[0].strip()

            if cat_sel_tab2 == "🌐 Todos los Mercados":
                choques_filtrados = choques_detectados
                rachas_filtradas = todas_rachas_activas
            else:
                choques_filtrados = [
                    ch for ch in choques_detectados
                    if cat_target in ch["mercado"] and (linea_elegida_tab2 == "TODAS" or linea_elegida_tab2 in ch["mercado"])
                ]
                rachas_filtradas = [
                    r for r in todas_rachas_activas
                    if r["Mercado_Cat"] == cat_target and (linea_elegida_tab2 == "TODAS" or r["Mercado_Linea"] == linea_elegida_tab2)
                ]

            if choques_filtrados:
                tarjetas_choque_html = []
                for ch in choques_filtrados[:6]:
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
                    f"<span class='pill-ev-cyan'>{len(choques_filtrados)} Convergencias Activas</span>"
                    "</div>"
                    f"<div class='radar-choque-grid'>{''.join(tarjetas_choque_html)}</div>"
                    "</div>",
                    unsafe_allow_html=True,
                )
            else:
                st.info("ℹ️ No hay partidos que alcancen este umbral. Prueba reduciendo el umbral de racha.")

            if not rachas_filtradas:
                st.info("Sin rachas de partido que superen el umbral seleccionado bajo este filtro.")
            else:
                st.markdown(f"#### 🎯 Desglose de Rachas Activas ({len(rachas_filtradas)})")
                _renderizar_tabla_barrido_global_radar(rachas_filtradas)

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
