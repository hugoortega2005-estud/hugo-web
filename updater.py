# updater.py
# ─────────────────────────────────────────────────────────────────────────────
# Pipeline Híbrido de Sincronización en Dos Capas (Two-Tier Hybrid Pipeline)
# para Hugo's Hub (100% libre de bloqueos 403 Cloudflare):
#   • Tier 1 (Base Line): Procesamiento por lotes CSV desde Football-Data.co.uk
#                         (temporada 2026/2027, 5 grandes ligas).
#   • Tier 2 (Live Delta Sync): Ingesta en tiempo real mediante la API JSON
#                         pública y abierta de ESPN (HTTP 200 directo, sin WAF)
#                         para las 5 grandes ligas europeas.
#
# Tabla objetivo: 'historico_reciente' en 'macro_base_futbol.db'
# Columnas: Date (DD/MM/YYYY), HomeTeam, AwayTeam, FTHG, FTAG, FTR,
#           Liga_Origen, Temporada_Origen ('2627'), y columna Div de la BD.
# ─────────────────────────────────────────────────────────────────────────────

from datetime import datetime, timedelta
import io
import logging
import os
import re
import sqlite3
import time
import unicodedata
from pathlib import Path

import pandas as pd
import requests

# ── Constantes ────────────────────────────────────────────────────────────────
RUTA_BD = Path(__file__).parent / "macro_base_futbol.db"

TEMPORADA_ACTUAL = "2627"

# Tier 1: Endpoints CSV de Football-Data.co.uk (Temporada 2026/2027)
URLS_FOOTBALL_DATA = {
    "SP1": "https://www.football-data.co.uk/mmz4281/2627/SP1.csv",  # LaLiga
    "E0":  "https://www.football-data.co.uk/mmz4281/2627/E0.csv",   # Premier League
    "I1":  "https://www.football-data.co.uk/mmz4281/2627/I1.csv",   # Serie A
    "D1":  "https://www.football-data.co.uk/mmz4281/2627/D1.csv",   # Bundesliga
    "F1":  "https://www.football-data.co.uk/mmz4281/2627/F1.csv",   # Ligue 1
}

# Tier 2: Endpoints JSON públicos de ESPN (sin Cloudflare, HTTP 200 directo)
URLS_ESPN_SCOREBOARD = {
    "I1":  "https://site.api.espn.com/apis/site/v2/sports/soccer/ita.1/scoreboard",  # Serie A
    "SP1": "https://site.api.espn.com/apis/site/v2/sports/soccer/esp.1/scoreboard",  # LaLiga
    "E0":  "https://site.api.espn.com/apis/site/v2/sports/soccer/eng.1/scoreboard",  # Premier League
    "D1":  "https://site.api.espn.com/apis/site/v2/sports/soccer/ger.1/scoreboard",  # Bundesliga
    "F1":  "https://site.api.espn.com/apis/site/v2/sports/soccer/fra.1/scoreboard",  # Ligue 1
    "uefa.champions":   "https://site.api.espn.com/apis/site/v2/sports/soccer/uefa.champions/scoreboard",    # Champions League
    "uefa.europa":      "https://site.api.espn.com/apis/site/v2/sports/soccer/uefa.europa/scoreboard",       # Europa League
    "uefa.europa.conf": "https://site.api.espn.com/apis/site/v2/sports/soccer/uefa.europa.conf/scoreboard",  # Conference League
}

# Nombre físico de la columna Div en historico_reciente
COL_DIV_BD = "\u00ef\u00bb\u00bfDiv"

TIMEOUT_CSV = 30
TIMEOUT_ESPN = 10

HEADERS_HTTP = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/130.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json,text/csv,text/plain,*/*;q=0.8",
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
}

# Mapa canónico de equipos alineado con 'historico_reciente' en macro_base_futbol.db
CANONICAL_TEAM_MAP = {
    # LaLiga (SP1)
    "espanyol": "Espanol", "rcd espanyol": "Espanol",
    "real sociedad": "Sociedad", "la real": "Sociedad",
    "athletic club": "Ath Bilbao", "athletic bilbao": "Ath Bilbao",
    "atlético de madrid": "Ath Madrid", "atletico madrid": "Ath Madrid",
    "atletico de madrid": "Ath Madrid", "atlético madrid": "Ath Madrid",
    "club atlético de madrid": "Ath Madrid",
    "celta de vigo": "Celta", "celta vigo": "Celta", "rc celta": "Celta", "rc celta de vigo": "Celta",
    "betis": "Betis", "real betis": "Betis", "real betis balompié": "Betis",
    "deportivo la coruña": "La Coruna", "deportivo": "La Coruna",
    "deportivo de la coruña": "La Coruna", "rc deportivo": "La Coruna",
    "dep. a coruña": "La Coruna", "dep. a coruna": "La Coruna", "deportivo a coruna": "La Coruna",
    "sporting de gijón": "Gijon", "sporting gijon": "Gijon",
    "rayo vallecano": "Vallecano",
    "alavés": "Alaves", "deportivo alaves": "Alaves", "deportivo alavés": "Alaves",
    "real valladolid": "Valladolid",
    "cd leganés": "Leganes", "leganés": "Leganes", "leganes": "Leganes",
    "ca osasuna": "Osasuna", "osasuna": "Osasuna",
    "ud las palmas": "Las Palmas", "las palmas": "Las Palmas",
    "rcd mallorca": "Mallorca", "mallorca": "Mallorca",
    "girona fc": "Girona", "girona": "Girona",
    "getafe cf": "Getafe", "getafe": "Getafe",
    "sevilla fc": "Sevilla", "sevilla": "Sevilla",
    "valencia cf": "Valencia", "valencia": "Valencia",
    "villarreal cf": "Villarreal", "villarreal": "Villarreal",
    "fc barcelona": "Barcelona", "barcelona": "Barcelona",
    "real madrid cf": "Real Madrid", "real madrid": "Real Madrid",
    "elche cf": "Elche", "elche": "Elche",
    "levante ud": "Levante", "levante": "Levante",
    "real oviedo": "Oviedo", "oviedo": "Oviedo",
    "racing santander": "Santander", "real racing club": "Santander", "santander": "Santander",
    "racing sant": "Santander",
    "málaga": "Malaga", "malaga": "Malaga", "málaga cf": "Malaga", "malaga cf": "Malaga",
    # Premier League (E0)
    "manchester united": "Man United", "man utd": "Man United", "manchester utd": "Man United",
    "manchester city": "Man City",
    "nottingham forest": "Nott'm Forest", "nottingham": "Nott'm Forest",
    "newcastle united": "Newcastle", "newcastle": "Newcastle",
    "tottenham hotspur": "Tottenham", "spurs": "Tottenham", "tottenham": "Tottenham",
    "west ham united": "West Ham", "west ham": "West Ham",
    "wolverhampton wanderers": "Wolves", "wolverhampton": "Wolves", "wolves": "Wolves",
    "brighton & hove albion": "Brighton", "brighton and hove albion": "Brighton", "brighton": "Brighton",
    "leicester city": "Leicester", "leicester": "Leicester",
    "ipswich town": "Ipswich", "ipswich": "Ipswich",
    "afc bournemouth": "Bournemouth", "bournemouth": "Bournemouth",
    "arsenal fc": "Arsenal", "arsenal": "Arsenal",
    "aston villa": "Aston Villa",
    "brentford fc": "Brentford", "brentford": "Brentford",
    "chelsea fc": "Chelsea", "chelsea": "Chelsea",
    "crystal palace": "Crystal Palace",
    "everton fc": "Everton", "everton": "Everton",
    "fulham fc": "Fulham", "fulham": "Fulham",
    "liverpool fc": "Liverpool", "liverpool": "Liverpool",
    "southampton fc": "Southampton", "southampton": "Southampton",
    "leeds united": "Leeds", "leeds": "Leeds",
    "burnley fc": "Burnley", "burnley": "Burnley",
    "sunderland afc": "Sunderland", "sunderland": "Sunderland",
    "hull city": "Hull", "hull": "Hull",
    "coventry city": "Coventry", "coventry": "Coventry",
    # Serie A (I1)
    "inter milan": "Inter", "internazionale": "Inter", "fc internazionale": "Inter", "inter": "Inter",
    "ac milan": "Milan", "milan": "Milan",
    "hellas verona": "Verona", "verona": "Verona",
    "as roma": "Roma", "roma": "Roma",
    "ss lazio": "Lazio", "lazio": "Lazio",
    "atalanta bc": "Atalanta", "atalanta": "Atalanta",
    "ssc napoli": "Napoli", "napoli": "Napoli",
    "acf fiorentina": "Fiorentina", "fiorentina": "Fiorentina",
    "torino fc": "Torino", "torino": "Torino",
    "como 1907": "Como", "como": "Como",
    "juventus fc": "Juventus", "juventus": "Juventus",
    "bologna fc 1909": "Bologna", "bologna fc": "Bologna", "bologna": "Bologna",
    "udinese calcio": "Udinese", "udinese": "Udinese",
    "genoa cfc": "Genoa", "genoa": "Genoa",
    "cagliari calcio": "Cagliari", "cagliari": "Cagliari",
    "parma calcio 1913": "Parma", "parma": "Parma",
    "us lecce": "Lecce", "lecce": "Lecce",
    "empoli fc": "Empoli", "empoli": "Empoli",
    "ac monza": "Monza", "monza": "Monza",
    "venezia fc": "Venezia", "venezia": "Venezia",
    "us sassuolo": "Sassuolo", "sassuolo": "Sassuolo", "sassuolo calcio": "Sassuolo",
    "us cremonese": "Cremonese", "cremonese": "Cremonese",
    "pisa sc": "Pisa", "ac pisa 1909": "Pisa", "pisa": "Pisa",
    # Bundesliga (D1)
    "bayern munich": "Bayern Munich", "fc bayern münchen": "Bayern Munich",
    "bayern münchen": "Bayern Munich", "fc bayern munich": "Bayern Munich",
    "bayer leverkusen": "Leverkusen", "bayer 04 leverkusen": "Leverkusen",
    "borussia dortmund": "Dortmund", "dortmund": "Dortmund",
    "borussia mönchengladbach": "M'gladbach", "borussia monchengladbach": "M'gladbach",
    "borussia m'gladbach": "M'gladbach", "gladbach": "M'gladbach", "m'gladbach": "M'gladbach",
    "eintracht frankfurt": "Ein Frankfurt", "frankfurt": "Ein Frankfurt",
    "rb leipzig": "RB Leipzig", "rasenballsport leipzig": "RB Leipzig",
    "vfb stuttgart": "Stuttgart", "stuttgart": "Stuttgart",
    "sc freiburg": "Freiburg", "freiburg": "Freiburg",
    "vfl wolfsburg": "Wolfsburg", "wolfsburg": "Wolfsburg",
    "1. fsv mainz 05": "Mainz", "mainz 05": "Mainz", "fsv mainz 05": "Mainz", "mainz": "Mainz",
    "tsg 1899 hoffenheim": "Hoffenheim", "tsg hoffenheim": "Hoffenheim", "hoffenheim": "Hoffenheim",
    "fc augsburg": "Augsburg", "augsburg": "Augsburg",
    "1. fc heidenheim": "Heidenheim", "fc heidenheim": "Heidenheim",
    "1. fc heidenheim 1846": "Heidenheim", "heidenheim": "Heidenheim",
    "sv werder bremen": "Werder Bremen", "werder bremen": "Werder Bremen",
    "fc st. pauli": "St Pauli", "st. pauli": "St Pauli", "st pauli": "St Pauli",
    "vfl bochum 1848": "Bochum", "vfl bochum": "Bochum", "bochum": "Bochum",
    "1. fc union berlin": "Union Berlin", "union berlin": "Union Berlin",
    "holstein kiel": "Holstein Kiel",
    "1. fc köln": "FC Koln", "fc köln": "FC Koln", "1. fc koln": "FC Koln", "fc koln": "FC Koln",
    "fc cologne": "FC Koln", "1. fc cologne": "FC Koln", "köln": "FC Koln", "koln": "FC Koln",
    "hamburger sv": "Hamburg", "hamburg": "Hamburg",
    "sv elversberg": "Elversberg", "elversberg": "Elversberg",
    "sc paderborn 07": "Paderborn", "sc paderborn": "Paderborn", "paderborn 07": "Paderborn", "paderborn": "Paderborn",
    "fc schalke 04": "Schalke 04", "schalke 04": "Schalke 04", "schalke": "Schalke 04",
    # Ligue 1 (F1)
    "paris saint-germain": "Paris SG", "paris sg": "Paris SG", "psg": "Paris SG",
    "olympique de marseille": "Marseille", "olympique marseille": "Marseille", "marseille": "Marseille",
    "olympique lyonnais": "Lyon", "lyon": "Lyon",
    "as monaco": "Monaco", "monaco": "Monaco",
    "losc lille": "Lille", "lille osc": "Lille", "lille": "Lille",
    "ogc nice": "Nice", "nice": "Nice",
    "rc lens": "Lens", "racing club de lens": "Lens", "lens": "Lens",
    "stade rennais": "Rennes", "stade rennais fc": "Rennes", "rennes": "Rennes",
    "stade brestois 29": "Brest", "stade brestois": "Brest", "brest": "Brest",
    "rc strasbourg alsace": "Strasbourg", "rc strasbourg": "Strasbourg", "strasbourg": "Strasbourg",
    "fc nantes": "Nantes", "nantes": "Nantes",
    "toulouse fc": "Toulouse", "toulouse": "Toulouse",
    "montpellier hsc": "Montpellier", "montpellier": "Montpellier",
    "stade de reims": "Reims", "reims": "Reims",
    "aj auxerre": "Auxerre", "auxerre": "Auxerre",
    "angers sco": "Angers", "angers": "Angers",
    "as saint-étienne": "St Etienne", "saint-etienne": "St Etienne",
    "le havre ac": "Le Havre", "le havre": "Le Havre",
    "fc lorient": "Lorient", "lorient": "Lorient",
    "fc metz": "Metz", "metz": "Metz",
    "paris fc": "Paris FC",
    # Clubes europeos (UEFA)
    "bodø/glimt": "Bodo/Glimt", "bodo/glimt": "Bodo/Glimt", "fk bodø/glimt": "Bodo/Glimt",
    "shakhtar donetsk": "Shakhtar Donetsk", "shakhtar d.": "Shakhtar Donetsk", "fc shakhtar donetsk": "Shakhtar Donetsk",
    "fenerbahçe": "Fenerbahce", "fenerbahce": "Fenerbahce", "fenerbahçe sk": "Fenerbahce",
    "slavia prague": "Slavia Prague", "sk slavia prague": "Slavia Prague", "slavia praha": "Slavia Prague",
    "slovan bratislava": "Slovan Bratislava", "šk slovan bratislava": "Slovan Bratislava",
    "sporting cp": "Sporting CP", "sporting lisbon": "Sporting CP",
    "psv eindhoven": "PSV", "psv": "PSV",
    "fc porto": "Porto", "porto": "Porto",
    "club brugge": "Club Brugge", "club brugge kv": "Club Brugge",
    "feyenoord rotterdam": "Feyenoord", "feyenoord": "Feyenoord",
    "lask linz": "LASK", "lask": "LASK",
    "viking fk": "Viking", "viking": "Viking",
}

# ── Logger ────────────────────────────────────────────────────────────────────
log = logging.getLogger(__name__)
if not log.handlers:
    _h = logging.StreamHandler()
    _h.setFormatter(logging.Formatter("%(asctime)s [updater] %(levelname)s: %(message)s"))
    log.addHandler(_h)
log.setLevel(logging.INFO)


def _norm_txt(texto: str) -> str:
    t = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode("ascii")
    return " ".join(t.strip().lower().split())


_CANONICAL_NORM = {_norm_txt(k): v for k, v in CANONICAL_TEAM_MAP.items()}


def resolver_equipo_canonico(nombre: str, mapa_bd: dict[str, str] | None = None) -> str:
    """
    Resuelve el nombre de un equipo a su nombre canónico en `historico_reciente`
    utilizando `CANONICAL_TEAM_MAP` y el índice de equipos de la BD.
    """
    if not nombre:
        return ""
    limpio = str(nombre).strip()
    clave_low = limpio.lower()
    if clave_low in CANONICAL_TEAM_MAP:
        canon = CANONICAL_TEAM_MAP[clave_low]
    else:
        clave_norm = _norm_txt(limpio)
        if clave_norm in _CANONICAL_NORM:
            canon = _CANONICAL_NORM[clave_norm]
        else:
            sin_afijos = re.sub(
                r"\b(fc|cf|ac|sc|bc|cfc|ssc|afc|ud|cd|rcd|rc|ca|sv|vfl|vfb|tsg|fsv|1\.\s*fc|1\.\s*fsv|1907|1909|1913|1899|1848|1846|04|05|29)\b",
                "",
                clave_norm,
            )
            sin_afijos = " ".join(sin_afijos.split())
            canon = _CANONICAL_NORM.get(sin_afijos, limpio)

    if mapa_bd:
        k_norm = _norm_txt(canon)
        if k_norm in mapa_bd:
            return mapa_bd[k_norm]
    return canon


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 1 – UTILIDADES DE BASE DE DATOS Y NORMALIZACIÓN
# ══════════════════════════════════════════════════════════════════════════════

def _garantizar_indices(conn: sqlite3.Connection) -> None:
    """Garantiza el índice UNIQUE en `historico_reciente` y depura alias no canónicos."""
    cur = conn.cursor()
    cur.execute(
        f'CREATE UNIQUE INDEX IF NOT EXISTS idx_partidos_pk '
        f'ON historico_reciente ("{COL_DIV_BD}", Date, HomeTeam, AwayTeam)'
    )
    # Limpiar posibles filas duplicadas con nombres sin normalizar de pruebas previas
    cur.execute(
        """
        DELETE FROM historico_reciente
        WHERE HomeTeam IN ('Racing Santander', 'Hull City', 'Coventry City', 'FC Cologne', 'SC Paderborn 07')
           OR AwayTeam IN ('SV Elversberg')
        """
    )
    conn.commit()


def _obtener_equipos_conocidos(conn: sqlite3.Connection) -> dict[str, str]:
    """Devuelve {clave_normalizada: nombre_canonico_bd} de `historico_reciente`."""
    mapa = {}
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT DISTINCT HomeTeam FROM historico_reciente WHERE HomeTeam IS NOT NULL "
            "UNION "
            "SELECT DISTINCT AwayTeam FROM historico_reciente WHERE AwayTeam IS NOT NULL"
        )
        for (eq,) in cur.fetchall():
            if eq:
                mapa[_norm_txt(eq)] = str(eq).strip()
    except sqlite3.Error:
        pass
    return mapa


def _normalizar_fecha_ddmmyyyy(valor_fecha) -> str:
    """Convierte cualquier fecha ISO o CSV al formato 'DD/MM/YYYY'."""
    if pd.isna(valor_fecha):
        return ""
    s = str(valor_fecha).strip()
    if not s:
        return ""
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s[:10], fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return s


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 2 – TIER 1: FOOTBALL-DATA.CO.UK CSV BATCH INGESTION
# ══════════════════════════════════════════════════════════════════════════════

def _descargar_csv_liga(liga: str, url: str) -> pd.DataFrame | None:
    """Descarga y parsea el CSV de una liga desde Football-Data.co.uk."""
    log.info(f"[Tier 1 - CSV] Descargando {liga} ({TEMPORADA_ACTUAL}): {url}")
    try:
        resp = requests.get(url, headers=HEADERS_HTTP, timeout=TIMEOUT_CSV)
        log.info(f"[HTTP] GET {url} -> Status: {resp.status_code}")
        resp.raise_for_status()

        df = pd.read_csv(io.StringIO(resp.content.decode("latin-1"))).copy()
        df = df.assign(Temporada_Origen=TEMPORADA_ACTUAL, Liga_Origen=liga)
        return df
    except Exception as exc:
        log.warning(f"  [Tier 1 - CSV] No disponible {liga} ({url}): {exc}")
        return None


def _normalizar_df_csv(
    df: pd.DataFrame,
    liga_default: str,
    mapa_equipos_bd: dict[str, str],
) -> pd.DataFrame:
    """Normaliza las columnas del CSV de Football-Data y resuelve nombres canónicos."""
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()
    df.columns = [
        c.replace("\ufeff", "").replace("\u00ef\u00bb\u00bf", "").strip()
        if c not in ("Temporada_Origen", "Liga_Origen") else c
        for c in df.columns
    ]

    if "Div" not in df.columns and COL_DIV_BD not in df.columns:
        df["Div"] = liga_default
    if "Div" in df.columns and COL_DIV_BD not in df.columns:
        df = df.rename(columns={"Div": COL_DIV_BD})

    cols_req = ["Date", "HomeTeam", "AwayTeam", "FTHG", "FTAG"]
    for c in cols_req:
        if c not in df.columns:
            return pd.DataFrame()

    df["FTHG"] = pd.to_numeric(df["FTHG"], errors="coerce")
    df["FTAG"] = pd.to_numeric(df["FTAG"], errors="coerce")
    df = df.dropna(subset=cols_req).copy()

    df["Date"] = df["Date"].apply(_normalizar_fecha_ddmmyyyy)
    df["HomeTeam"] = df["HomeTeam"].apply(lambda x: resolver_equipo_canonico(x, mapa_equipos_bd))
    df["AwayTeam"] = df["AwayTeam"].apply(lambda x: resolver_equipo_canonico(x, mapa_equipos_bd))

    df = df[
        (df["Date"] != "")
        & (df["HomeTeam"] != "")
        & (df["AwayTeam"] != "")
    ].copy()

    def _calc_ftr(r):
        val = r.get("FTR")
        if pd.notna(val) and str(val).strip() in ("H", "D", "A"):
            return str(val).strip()
        gl = int(r["FTHG"])
        gv = int(r["FTAG"])
        return "H" if gl > gv else ("A" if gv > gl else "D")

    df["FTR"] = df.apply(_calc_ftr, axis=1)
    df["Liga_Origen"] = liga_default
    df["Temporada_Origen"] = TEMPORADA_ACTUAL
    return df.reset_index(drop=True)


def _ingestar_df_csv(
    conn: sqlite3.Connection,
    df: pd.DataFrame,
) -> tuple[int, int, list[str]]:
    """Inserta o actualiza en `historico_reciente` los registros procedentes del CSV."""
    if df is None or df.empty:
        return 0, 0, []

    cur = conn.cursor()
    cur.execute("PRAGMA table_info(historico_reciente)")
    cols_bd = [row[1] for row in cur.fetchall()]

    cols_comunes = [c for c in df.columns if c in cols_bd]
    df_filtrado = df[cols_comunes].copy()

    placeholders = ", ".join(["?" for _ in cols_comunes])
    cols_quoted  = ", ".join([f'"{c}"' for c in cols_comunes])
    sql_insert = (
        f"INSERT OR IGNORE INTO historico_reciente ({cols_quoted}) "
        f"VALUES ({placeholders})"
    )

    sql_check = """
        SELECT rowid, FTHG, FTAG
        FROM historico_reciente
        WHERE HomeTeam = ? AND AwayTeam = ? AND Date = ?
        LIMIT 1
    """

    sql_update = """
        UPDATE historico_reciente
        SET FTHG = ?, FTAG = ?, FTR = ?
        WHERE rowid = ?
    """

    insertados = 0
    actualizados = 0
    nuevos_detalle = []

    for idx_row, row in df_filtrado.iterrows():
        raw_row = df.loc[idx_row]
        date_val = str(raw_row["Date"]).strip()
        home_val = str(raw_row["HomeTeam"]).strip()
        away_val = str(raw_row["AwayTeam"]).strip()
        liga_val = str(raw_row.get("Liga_Origen", "")).strip()
        goles_l = int(raw_row["FTHG"])
        goles_v = int(raw_row["FTAG"])
        ftr_val = str(raw_row["FTR"]).strip()

        cur.execute(sql_check, (home_val, away_val, date_val))
        existente = cur.fetchone()

        if existente is not None:
            rowid_ex, db_fthg, db_ftag = existente
            if db_fthg is None or db_ftag is None:
                cur.execute(sql_update, (goles_l, goles_v, ftr_val, rowid_ex))
                if cur.rowcount > 0:
                    actualizados += 1
        else:
            valores_ins = tuple(
                int(raw_row["FTHG"]) if col_n == "FTHG" and pd.notna(v)
                else (
                    int(raw_row["FTAG"]) if col_n == "FTAG" and pd.notna(v)
                    else (None if (isinstance(v, float) and pd.isna(v)) else v)
                )
                for col_n, v in zip(cols_comunes, row)
            )
            cur.execute(sql_insert, valores_ins)
            if cur.rowcount == 1:
                insertados += 1
                det = f"[Tier 1 CSV] {liga_val} | {date_val}: {home_val} {goles_l} - {goles_v} {away_val}"
                nuevos_detalle.append(det)
                print(det)

    conn.commit()
    return insertados, actualizados, nuevos_detalle


def ejecutar_tier1_csv(
    conn: sqlite3.Connection,
    mapa_equipos_bd: dict[str, str],
) -> tuple[int, int, list[str], list[str]]:
    """Ejecuta el Tier 1 descargando los CSVs de Football-Data.co.uk."""
    t1_nuevos = 0
    t1_act = 0
    t1_detalles = []
    errores = []
    descargas_ok = 0

    for liga, url in URLS_FOOTBALL_DATA.items():
        df_raw = _descargar_csv_liga(liga, url)
        if df_raw is None:
            continue
        descargas_ok += 1
        df_clean = _normalizar_df_csv(df_raw, liga_default=liga, mapa_equipos_bd=mapa_equipos_bd)
        ins, act, det = _ingestar_df_csv(conn, df_clean)
        t1_nuevos += ins
        t1_act += act
        t1_detalles.extend(det)
        log.info(f"  [Tier 1 - {liga}] {ins} nuevos, {act} actualizados.")

    if descargas_ok == 0:
        errores.append("Tier 1: No se pudo descargar ningún CSV desde Football-Data.co.uk.")

    return t1_nuevos, t1_act, t1_detalles, errores


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 3 – TIER 2: API JSON PÚBLICA DE ESPN (SIN 403, TIEMPO REAL)
# ══════════════════════════════════════════════════════════════════════════════

# Mapa de código ESPN / liga -> nombre canónico de competición para torneos europeos
_NOMBRE_COMPETICION_PARTIDOS = {
    "uefa.champions": "Champions League",
    "uefa.europa": "Europa League",
    "uefa.europa.conf": "Conference League",
    "Champions League": "Champions League",
    "Europa League": "Europa League",
    "Conference League": "Conference League",
}


def _extraer_eventos_espn_liga(
    liga: str,
    url: str,
    mapa_equipos_bd: dict[str, str],
    dias_ventana: int = 5,
) -> list[dict]:
    """
    Descarga el JSON de ESPN Scoreboard con `requests.get(url, timeout=10)`.
    Para ligas domésticas consulta los últimos `dias_ventana` días; para torneos UEFA
    barre además las jornadas europeas (martes, miércoles y jueves desde el inicio de
    la fase de liga en septiembre de 2026) para consolidar todos los partidos disputados.
    """
    partidos_extraidos = []
    vistos = set()

    hoy = datetime.now()
    es_uefa = str(liga).startswith("uefa.") or liga in _NOMBRE_COMPETICION_PARTIDOS
    liga_canon = _NOMBRE_COMPETICION_PARTIDOS.get(liga, liga)

    fechas_str = [
        (hoy - timedelta(days=delta)).strftime("%Y%m%d")
        for delta in range(dias_ventana)
    ]
    if es_uefa:
        inicio_uefa = datetime(hoy.year if hoy.month >= 8 else hoy.year - 1, 9, 1)
        cursor_d = inicio_uefa
        while cursor_d <= hoy:
            if cursor_d.weekday() in (1, 2, 3):  # Martes, Miércoles, Jueves
                f_s = cursor_d.strftime("%Y%m%d")
                if f_s not in fechas_str:
                    fechas_str.append(f_s)
            cursor_d += timedelta(days=1)

    consultas = [url] + [f"{url}?dates={f_s}" for f_s in fechas_str]

    for target_url in consultas:
        try:
            resp = requests.get(target_url, timeout=TIMEOUT_ESPN)
            log.info(f"[HTTP] GET {target_url} -> Status: {resp.status_code}")
            resp.raise_for_status()
            data = resp.json()
        except Exception as exc:
            log.warning(f"  [Tier 2 - ESPN] Error consultando {target_url}: {exc}")
            continue

        eventos = data.get("events") or []
        for event in eventos:
            status_type = (event.get("status") or {}).get("type") or {}
            if status_type.get("completed") is not True:
                continue

            competitions = event.get("competitions") or []
            if not competitions:
                continue
            comp = competitions[0]
            competitors = comp.get("competitors") or []
            if len(competitors) < 2:
                continue

            home_comp = next((c for c in competitors if c.get("homeAway") == "home"), None)
            away_comp = next((c for c in competitors if c.get("homeAway") == "away"), None)
            if not home_comp or not away_comp:
                continue

            home_raw = (
                (home_comp.get("team") or {}).get("displayName")
                or (home_comp.get("team") or {}).get("name")
                or ""
            )
            away_raw = (
                (away_comp.get("team") or {}).get("displayName")
                or (away_comp.get("team") or {}).get("name")
                or ""
            )

            try:
                goles_local = int(float(home_comp.get("score")))
                goles_visitante = int(float(away_comp.get("score")))
            except (TypeError, ValueError):
                continue

            raw_date = comp.get("date") or event.get("date") or ""
            fecha_ddmmyyyy = _normalizar_fecha_ddmmyyyy(raw_date[:10])
            if not fecha_ddmmyyyy:
                continue

            equipo_local = resolver_equipo_canonico(home_raw, mapa_equipos_bd)
            equipo_visitante = resolver_equipo_canonico(away_raw, mapa_equipos_bd)
            if not equipo_local or not equipo_visitante:
                continue

            ftr = (
                "H" if goles_local > goles_visitante
                else ("A" if goles_visitante > goles_local else "D")
            )

            clave = (liga_canon, fecha_ddmmyyyy, equipo_local, equipo_visitante)
            if clave in vistos:
                continue
            vistos.add(clave)

            partidos_extraidos.append({
                "Date": fecha_ddmmyyyy,
                "HomeTeam": equipo_local,
                "AwayTeam": equipo_visitante,
                "FTHG": goles_local,
                "FTAG": goles_visitante,
                "FTR": ftr,
                "Liga_Origen": liga_canon,
                "Temporada_Origen": TEMPORADA_ACTUAL,
            })

    return partidos_extraidos


def ejecutar_tier2_espn(
    conn: sqlite3.Connection,
    mapa_equipos_bd: dict[str, str],
) -> tuple[int, int, list[str], list[str]]:
    """
    Tier 2 (Live Delta Sync): Recorre los endpoints de ESPN Scoreboard (5 grandes ligas
    + Champions League, Europa League y Conference League), comprueba existencia en
    `historico_reciente` y en `partidos`, e inserta/actualiza registrando el nombre
    canónico de la competición.
    """
    log.info("[Tier 2 - ESPN API] Consultando resultados en tiempo real (5 grandes ligas + UEFA)...")
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(historico_reciente)")
    cols_bd = {row[1] for row in cur.fetchall()}
    tiene_col_div_bom = COL_DIV_BD in cols_bd

    tablas_existentes = {
        r[0] for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
    }
    tiene_tabla_partidos = "partidos" in tablas_existentes

    sql_check = """
        SELECT rowid, FTHG, FTAG
        FROM historico_reciente
        WHERE HomeTeam = ? AND AwayTeam = ? AND Date = ?
        LIMIT 1
    """

    if tiene_col_div_bom:
        sql_insert = f"""
            INSERT INTO historico_reciente (
                "{COL_DIV_BD}", Date, HomeTeam, AwayTeam,
                FTHG, FTAG, FTR, Liga_Origen, Temporada_Origen
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
    else:
        sql_insert = """
            INSERT INTO historico_reciente (
                Date, HomeTeam, AwayTeam,
                FTHG, FTAG, FTR, Liga_Origen, Temporada_Origen
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """

    sql_update = """
        UPDATE historico_reciente
        SET FTHG             = ?,
            FTAG             = ?,
            FTR              = ?,
            Liga_Origen      = COALESCE(Liga_Origen, ?),
            Temporada_Origen = COALESCE(Temporada_Origen, ?)
        WHERE rowid = ?
    """

    sql_upsert_tabla_partidos = """
        INSERT INTO partidos (
            fecha, liga, equipo_local, equipo_visitante,
            goles_local, goles_visitante, resultado, temporada, fuente
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(fecha, liga, equipo_local, equipo_visitante) DO UPDATE SET
            goles_local     = excluded.goles_local,
            goles_visitante = excluded.goles_visitante,
            resultado       = excluded.resultado,
            temporada       = COALESCE(partidos.temporada, excluded.temporada),
            fuente          = COALESCE(partidos.fuente, excluded.fuente)
    """

    t2_insertados = 0
    t2_actualizados = 0
    t2_detalles = []
    errores = []
    ligas_ok = 0

    for liga, url in URLS_ESPN_SCOREBOARD.items():
        liga_canon = _NOMBRE_COMPETICION_PARTIDOS.get(liga, liga)
        partidos = _extraer_eventos_espn_liga(liga, url, mapa_equipos_bd)
        if partidos:
            ligas_ok += 1

        ins_liga = 0
        act_liga = 0
        ins_tabla_partidos = 0
        for p in partidos:
            date_val = p["Date"]
            home_val = p["HomeTeam"]
            away_val = p["AwayTeam"]
            fthg_val = p["FTHG"]
            ftag_val = p["FTAG"]
            ftr_val = p["FTR"]
            liga_val = p["Liga_Origen"]
            temp_val = p["Temporada_Origen"]

            try:
                if tiene_tabla_partidos:
                    antes_cambios = conn.total_changes
                    cur.execute(
                        sql_upsert_tabla_partidos,
                        (
                            date_val, liga_val, home_val, away_val,
                            fthg_val, ftag_val, ftr_val, temp_val, "espn_api",
                        ),
                    )
                    if conn.total_changes > antes_cambios:
                        ins_tabla_partidos += 1

                cur.execute(sql_check, (home_val, away_val, date_val))
                existente = cur.fetchone()

                if existente is None:
                    if tiene_col_div_bom:
                        cur.execute(
                            sql_insert,
                            (
                                liga_val, date_val, home_val, away_val,
                                fthg_val, ftag_val, ftr_val, liga_val, temp_val,
                            ),
                        )
                    else:
                        cur.execute(
                            sql_insert,
                            (
                                date_val, home_val, away_val,
                                fthg_val, ftag_val, ftr_val, liga_val, temp_val,
                            ),
                        )
                    if cur.rowcount == 1:
                        ins_liga += 1
                        t2_insertados += 1
                        msg = (
                            f"[Tier 2 ESPN - Partido Nuevo Añadido] "
                            f"{liga_val} ({date_val}): {home_val} {fthg_val} - {ftag_val} {away_val} "
                            f"[FTR: {ftr_val}]"
                        )
                        t2_detalles.append(msg)
                        print(msg)
                        log.info(msg)
                else:
                    rowid_ex, db_fthg, db_ftag = existente
                    if db_fthg is None or db_ftag is None:
                        cur.execute(
                            sql_update,
                            (fthg_val, ftag_val, ftr_val, liga_val, temp_val, rowid_ex),
                        )
                        if cur.rowcount > 0:
                            act_liga += 1
                            t2_actualizados += 1
                            msg_u = (
                                f"[Tier 2 ESPN - Partido Actualizado] "
                                f"{liga_val} ({date_val}): {home_val} {fthg_val} - {ftag_val} {away_val}"
                            )
                            print(msg_u)
                            log.info(msg_u)
            except sqlite3.Error as exc:
                log.debug(f"  Error SQLite en ({date_val}, {home_val} vs {away_val}): {exc}")

        log.info(
            f"  [Tier 2 - {liga_canon}] {len(partidos)} partidos finalizados evaluados -> "
            f"{ins_liga} insertados en historico_reciente, {act_liga} actualizados "
            f"(tabla 'partidos': {ins_tabla_partidos} sincronizados)."
        )

    conn.commit()
    return t2_insertados, t2_actualizados, t2_detalles, errores


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 3B – ESTADÍSTICAS DE JUGADORES (MIGRACIÓN + INGESTA ESPN ROSTER)
# ══════════════════════════════════════════════════════════════════════════════

# Esquema FINAL de `estadisticas_jugadores` (nombres intuitivos en español).
# Se conservan sin cambios los nombres ya legibles (Pases_*, Entradas*, Intercepciones,
# Faltas_Recibidas) porque el Scouting del dashboard los consume.
_COLUMNAS_JUGADORES_BASE = [
    ("Jugador", "TEXT"), ("Equipo", "TEXT"), ("posicion", "TEXT"),
    ("minutos", "REAL DEFAULT 0"), ("partidos_jugados", "INTEGER DEFAULT 0"),
    ("titularidades", "INTEGER DEFAULT 0"), ("goles", "REAL DEFAULT 0"),
    ("asistencias", "REAL DEFAULT 0"), ("xg", "REAL DEFAULT 0.0"),
    ("tarjetas_amarillas", "REAL DEFAULT 0"), ("tarjetas_rojas", "REAL DEFAULT 0"),
    ("tiros_totales", "INTEGER"), ("tiros_a_puerta", "INTEGER"),
    ("faltas_cometidas", "REAL"), ("Faltas_Recibidas", "REAL"),
    ("fueras_juego", "INTEGER DEFAULT 0"), ("entradas_tackles", "INTEGER DEFAULT 0"),
    ("Pases_Intentados", "REAL"), ("Pases_Completados", "REAL"), ("Pases_Clave", "REAL"),
    ("Entradas", "REAL"), ("Entradas_Ganadas", "REAL"), ("Intercepciones", "REAL"),
]
# Compatibilidad: ya no hay columnas "nuevas" aparte del esquema base.
_COLUMNAS_JUGADORES_NUEVAS: list = []

_LIGAS_DOMESTICAS = ("LaLiga", "Premier League", "Serie A", "Bundesliga", "Ligue 1")
_COMPETICIONES_UEFA = ("Champions League", "Europa League", "Conference League")
_CODIGO_A_COMPETICION = {
    "SP1": "LaLiga", "E0": "Premier League", "I1": "Serie A",
    "D1": "Bundesliga", "F1": "Ligue 1",
    "uefa.champions": "Champions League",
    "uefa.europa": "Europa League",
    "uefa.europa.conf": "Conference League",
}
_COMPETICION_SIN_ASIGNAR = "Sin Asignar"

# Renombrado de columnas legadas -> nombres intuitivos.
_RENOMBRES_JUGADORES = {
    "Min": "minutos", "Gls": "goles", "Ast": "asistencias", "xG": "xg",
    "CrdY": "tarjetas_amarillas", "CrdR": "tarjetas_rojas",
    "Tiros": "tiros_totales", "Shots": "tiros_totales",
    "Tiros_Puerta": "tiros_a_puerta", "Shots_On": "tiros_a_puerta",
    "Faltas_Cometidas": "faltas_cometidas", "Faltas": "faltas_cometidas", "Fouls": "faltas_cometidas",
    "Tackles": "entradas_tackles", "Offsides": "fueras_juego",
}
# Columnas crípticas residuales que se eliminan.
_COLUMNAS_JUGADORES_OBSOLETAS = ["90s", "G+A", "G-PK", "PKatt", "PK", "Pos", "posicion_tactica"]


def _migrar_esquema_estadisticas_jugadores(conn: sqlite3.Connection) -> None:
    """
    Migración estructural idempotente y sin pérdida de datos de `estadisticas_jugadores`:
      1. Crea la tabla con el esquema final si no existe.
      2. Fusiona Pos / posicion_tactica en `posicion` (POR, DFC, MC, DC...).
      3. Renombra columnas legadas (Min -> minutos, Gls -> goles, ...).
      4. Elimina columnas crípticas (90s, G+A, G-PK, PKatt, Pos, posicion_tactica...).
      5. Añade las columnas del esquema final que falten y garantiza PK (Jugador, Equipo).
    Debe invocarse dentro de un `with sqlite3.connect(...)` (rollback automático).
    """
    cur = conn.cursor()
    info = cur.execute("PRAGMA table_info(estadisticas_jugadores)").fetchall()
    if not info:
        defs = ", ".join(f'"{n}" {t}' for n, t in _COLUMNAS_JUGADORES_BASE)
        cur.execute(
            f'CREATE TABLE estadisticas_jugadores ({defs}, "competicion" TEXT NOT NULL, '
            f'PRIMARY KEY ("Jugador", "Equipo", "competicion"))'
        )
        log.info("[Jugadores] Tabla estadisticas_jugadores creada (esquema final, PK Jugador+Equipo+competicion).")
        return

    def _cols() -> dict:
        return {str(f[1]).lower(): str(f[1]) for f in cur.execute("PRAGMA table_info(estadisticas_jugadores)")}

    cols = _cols()

    # 2. posicion = posicion existente > posicion_tactica (ESPN) > mapeo de Pos (FBref)
    if "posicion" not in cols:
        cur.execute('ALTER TABLE estadisticas_jugadores ADD COLUMN "posicion" TEXT')
        cols = _cols()
    tact = f'NULLIF(NULLIF("{cols["posicion_tactica"]}", \'\'), \'S/P\')' if "posicion_tactica" in cols else "NULL"
    if "pos" in cols:
        p = f'"{cols["pos"]}"'
        raw = (f"CASE WHEN {p} LIKE 'GK%' THEN 'POR' WHEN {p} LIKE 'DF%' THEN 'DFC' "
               f"WHEN {p} LIKE 'MF%' THEN 'MC' WHEN {p} LIKE 'FW%' THEN 'DC' ELSE NULLIF({p}, '') END")
    else:
        raw = "NULL"
    cur.execute(
        f'UPDATE estadisticas_jugadores SET "posicion" = COALESCE({tact}, {raw}) '
        f'WHERE "posicion" IS NULL OR "posicion" IN (\'\', \'S/P\')'
    )

    # 3. Renombrados (si el destino ya existe, la columna legada simplemente se descarta)
    for viejo, nuevo in _RENOMBRES_JUGADORES.items():
        if viejo.lower() not in cols:
            continue
        real = cols[viejo.lower()]
        if viejo.lower() == nuevo.lower() or nuevo.lower() not in cols:
            if real != nuevo:
                cur.execute(f'ALTER TABLE estadisticas_jugadores RENAME COLUMN "{real}" TO "{nuevo}"')
                log.info(f"[Jugadores] Columna renombrada: {real} -> {nuevo}")
        else:
            cur.execute(f'UPDATE estadisticas_jugadores SET "{cols[nuevo.lower()]}" = COALESCE("{cols[nuevo.lower()]}", "{real}")')
            cur.execute(f'ALTER TABLE estadisticas_jugadores DROP COLUMN "{real}"')
        cols = _cols()

    # 4. Eliminación de columnas obsoletas
    for obs in _COLUMNAS_JUGADORES_OBSOLETAS:
        if obs.lower() in cols:
            cur.execute(f'ALTER TABLE estadisticas_jugadores DROP COLUMN "{cols[obs.lower()]}"')
            log.info(f"[Jugadores] Columna obsoleta eliminada: {obs}")
            cols = _cols()

    # 5. Columnas del esquema final que falten
    for nombre, tipo in _COLUMNAS_JUGADORES_BASE:
        if nombre.lower() not in cols:
            cur.execute(f'ALTER TABLE estadisticas_jugadores ADD COLUMN "{nombre}" {tipo}')
            cols = _cols()
            log.info(f"[Jugadores] Columna añadida: {nombre}")

    _asegurar_pk_competicion(conn)


def _resolver_ruta_fbref(archivo: str):
    """Ruta existente del CSV de FBref (tolera alias de nombre) o None si no existe."""
    base = Path(__file__).parent
    alias = {"fbref_seriea.csv": "fbref_seriaA.csv"}.get(archivo)
    for r in (archivo, str(base / archivo), alias, str(base / alias) if alias else None):
        if r and os.path.exists(r):
            return r
    return None


def _mapa_club_competicion(conn: sqlite3.Connection) -> dict:
    """{club normalizado: liga doméstica} a partir de los CSVs de FBref y de historico_reciente."""
    mapa: dict = {}
    for archivo, liga in ARCHIVOS_FBREF:
        if liga not in _LIGAS_DOMESTICAS:
            continue
        ruta = _resolver_ruta_fbref(archivo)
        if not ruta:
            continue
        try:
            squads = pd.read_csv(ruta, header=1, usecols=["Squad"])["Squad"].dropna().astype(str)
        except Exception as exc:
            log.warning(f"[Migración] No se pudo leer {archivo}: {exc}")
            continue
        for sq in squads.unique():
            limpio = re.sub(r"^[a-z]{2,3}\s+", "", sq).strip()
            mapa[normalizar_nombre(resolver_equipo_canonico(limpio, None))] = liga
    try:
        cols = [r[1] for r in conn.execute("PRAGMA table_info(historico_reciente)")]
        div = next((c for c in cols if str(c).endswith("Div")), None)
        if div:
            for d, h in conn.execute(f'SELECT DISTINCT "{div}", HomeTeam FROM historico_reciente'):
                comp = _CODIGO_A_COMPETICION.get(str(d).strip())
                if comp and h:
                    mapa.setdefault(normalizar_nombre(h), comp)
    except sqlite3.Error as exc:
        log.warning(f"[Migración] historico_reciente no disponible para mapear ligas: {exc}")
    return mapa


def _registrar_comp_de(conn: sqlite3.Connection) -> None:
    """Registra la función SQL comp_de(equipo) -> liga doméstica (con resolución canónica del club)."""
    mapa = _mapa_club_competicion(conn)

    def _comp(eq):
        return (mapa.get(normalizar_nombre(eq))
                or mapa.get(normalizar_nombre(resolver_equipo_canonico(str(eq or ""), None)))
                or _COMPETICION_SIN_ASIGNAR)

    conn.create_function("comp_de", 1, _comp)


def _reasignar_competicion_pendiente(conn: sqlite3.Connection) -> None:
    """Reasigna filas 'Sin Asignar' cuyo club ya se puede resolver a su liga doméstica."""
    cur = conn.cursor()
    if not cur.execute("SELECT 1 FROM estadisticas_jugadores WHERE competicion = ? LIMIT 1",
                       (_COMPETICION_SIN_ASIGNAR,)).fetchone():
        return
    _registrar_comp_de(conn)
    cur.execute(
        "UPDATE OR IGNORE estadisticas_jugadores SET competicion = comp_de(Equipo) WHERE competicion = ?",
        (_COMPETICION_SIN_ASIGNAR,),
    )
    log.info(f"[Jugadores] Filas pendientes reasignadas a su liga: {cur.rowcount}")


def _asegurar_pk_competicion(conn: sqlite3.Connection) -> None:
    """
    Garantiza la columna `competicion` (TEXT NOT NULL) y la clave primaria
    (Jugador, Equipo, competicion). Reconstruye la tabla copiando TODAS las filas y asigna
    la liga doméstica según el club. Verifica que no se pierda ninguna fila.
    """
    cur = conn.cursor()
    info = cur.execute("PRAGMA table_info(estadisticas_jugadores)").fetchall()
    nombres = {str(f[1]).lower() for f in info}
    pk_cols = {str(f[1]).lower() for f in info if f[5]}
    if "competicion" in nombres and pk_cols == {"jugador", "equipo", "competicion"}:
        _reasignar_competicion_pendiente(conn)
        armonizar_clubes_estadisticas_jugadores(conn)
        return

    _registrar_comp_de(conn)
    cols_t = [(str(f[1]), str(f[2]) or "TEXT") for f in info if str(f[1]).lower() != "competicion"]
    defs = ", ".join(f'"{n}" {t}' for n, t in cols_t)
    lista = ", ".join(f'"{n}"' for n, _ in cols_t)
    expr = ('COALESCE(NULLIF("competicion", \'\'), comp_de("Equipo"))'
            if "competicion" in nombres else 'comp_de("Equipo")')
    n_antes = cur.execute(
        'SELECT COUNT(*) FROM estadisticas_jugadores WHERE "Jugador" IS NOT NULL AND "Equipo" IS NOT NULL'
    ).fetchone()[0]
    cur.execute("DROP TABLE IF EXISTS estadisticas_jugadores_nueva")
    cur.execute(
        f'CREATE TABLE estadisticas_jugadores_nueva ({defs}, "competicion" TEXT NOT NULL, '
        f'PRIMARY KEY ("Jugador", "Equipo", "competicion"))'
    )
    cur.execute(
        f'INSERT INTO estadisticas_jugadores_nueva ({lista}, "competicion") '
        f'SELECT {lista}, {expr} FROM estadisticas_jugadores '
        f'WHERE "Jugador" IS NOT NULL AND "Equipo" IS NOT NULL'
    )
    n_despues = cur.execute("SELECT COUNT(*) FROM estadisticas_jugadores_nueva").fetchone()[0]
    if n_despues != n_antes:
        raise sqlite3.IntegrityError(f"Migración abortada: {n_antes} filas origen vs {n_despues} destino.")
    cur.execute("DROP TABLE estadisticas_jugadores")
    cur.execute("ALTER TABLE estadisticas_jugadores_nueva RENAME TO estadisticas_jugadores")
    sin = cur.execute(
        "SELECT COUNT(*) FROM estadisticas_jugadores WHERE competicion = ?", (_COMPETICION_SIN_ASIGNAR,)
    ).fetchone()[0]
    log.info(f"[Jugadores] Esquema migrado: columna competicion + PK (Jugador, Equipo, competicion); "
             f"{n_despues} filas conservadas ({sin} sin liga asignada).")


def armonizar_clubes_estadisticas_jugadores(conn: sqlite3.Connection) -> dict:
    """
    Aplica `resolver_equipo_canonico` a todas las filas de `estadisticas_jugadores`
    (y tablas de partidos) para unificar grafías duplicadas (ej. 'Real Betis' -> 'Betis',
    'Atlético Madrid' -> 'Ath Madrid', 'Manchester Utd' -> 'Man United').
    Cuando el mismo futbolista figura dos veces en la misma competición bajo grafías
    distintas del mismo club (una fila de ESPN con tiros/faltas y otra de FBref con
    minutos reales), fusiona ambas filas conservando el máximo de minutos/titularidades
    de FBref y los tiros/faltas de ESPN.
    """
    cur = conn.cursor()
    mapa_bd = _obtener_equipos_conocidos(conn)
    filas = cur.execute(
        "SELECT rowid, Jugador, Equipo, competicion FROM estadisticas_jugadores"
    ).fetchall()
    if not filas:
        return {"renombrados": 0, "fusionados": 0}

    # Paso 1: normalizar el nombre de Equipo de cada fila a su club canónico
    grupos: dict = {}
    cambios_club = 0
    for rid, jug, eq, comp in filas:
        eq_canon = resolver_equipo_canonico(str(eq or ""), mapa_bd) or str(eq or "").strip()
        if eq_canon != eq:
            cambios_club += 1
        grupos.setdefault((eq_canon, comp), []).append((rid, jug, eq))

    fusionados = 0
    renombrados = 0

    for (eq_canon, comp), items in grupos.items():
        # Si todas las filas del grupo ya tienen eq_canon y no hay nombres duplicados, saltar
        if all(eq == eq_canon for _, _, eq in items):
            norms = [normalizar_nombre(j) for _, j, _ in items]
            if len(norms) == len(set(norms)):
                continue

        # Cargar detalle completo de las filas de este (club_canon, competicion)
        rids = [r for r, _, _ in items]
        placeholders = ",".join("?" for _ in rids)
        detalles = cur.execute(
            f"""
            SELECT rowid, Jugador, Equipo, posicion, minutos, partidos_jugados, titularidades,
                   goles, asistencias, xg, tarjetas_amarillas, tarjetas_rojas,
                   tiros_totales, tiros_a_puerta, faltas_cometidas, Faltas_Recibidas,
                   fueras_juego, entradas_tackles, Pases_Intentados, Pases_Completados,
                   Pases_Clave, Entradas, Entradas_Ganadas, Intercepciones
            FROM estadisticas_jugadores
            WHERE rowid IN ({placeholders})
            """,
            rids,
        ).fetchall()

        # Separar filas que ya estaban en eq_canon de las que vienen de un alias no canónico,
        # o agrupar por coincidencia difusa de jugador dentro del mismo (eq_canon, comp).
        # Ordenar priorizando filas con tiros_totales > 0 o minutos > 0.
        por_rid = {d[0]: d for d in detalles}
        libres: dict = {}  # rid_canon -> norm_nombre

        # Primero registrar como anclas las filas que ya tienen el nombre canónico del club
        # y, a igualdad, las que tienen datos ESPN (tiros/faltas no nulos).
        ordenadas = sorted(
            detalles,
            key=lambda d: (
                1 if d[2] == eq_canon else 0,
                1 if (d[12] is not None or d[14] is not None) else 0,
                float(d[4] or 0.0),
            ),
            reverse=True,
        )

        for d in ordenadas:
            rid = d[0]
            jug = d[1]
            eq_actual = d[2]
            norm_j = normalizar_nombre(jug)
            match_rid = _buscar_coincidencia(norm_j, libres)
            if match_rid is None:
                # No existe aún otro registro de este jugador en (eq_canon, comp)
                if eq_actual != eq_canon:
                    cur.execute(
                        "UPDATE OR IGNORE estadisticas_jugadores SET Equipo = ? WHERE rowid = ?",
                        (eq_canon, rid),
                    )
                    renombrados += cur.rowcount
                libres[rid] = norm_j
            else:
                # Fusionar `d` (secundaria) dentro de `base` (match_rid) y borrar `rid`
                base = por_rid[match_rid]
                # Elegir la fila con minutos reales de FBref (la que no tiene tiros ESPN o tiene más precisión)
                # Si una tiene tiros_totales IS NULL (vino de FBref) y la otra tiene tiros_totales IS NOT NULL (ESPN),
                # tomamos minutos/PJ/titularidades/goles/asistencias/xg/tarjetas de FBref y tiros/faltas de ESPN.
                base_es_espn = base[12] is not None
                sec_es_fbref = d[12] is None
                if base_es_espn and sec_es_fbref:
                    fb, es = d, base
                elif (not base_es_espn) and (d[12] is not None):
                    fb, es = base, d
                else:
                    # Ambas de la misma fuente: conservar el máximo de cada métrica
                    fb = d if float(d[4] or 0) >= float(base[4] or 0) else base
                    es = base if fb is d else d

                pos_f = fb[3] if (fb[3] and fb[3] != "S/P") else (es[3] or "S/P")
                min_f = float(fb[4] if fb[4] is not None else (es[4] or 0.0))
                # En ligas domésticas y UEFA, los partidos_jugados ('MP') y titularidades ('Starts')
                # de FBref son la fuente canónica exacta (incluso si MP=0).
                pj_f = int(fb[5] if fb[5] is not None else (es[5] or 0))
                tit_f = int(fb[6] if fb[6] is not None else (es[6] or 0))
                gol_f = float(fb[7] if fb[7] is not None else (es[7] or 0.0))
                ast_f = float(fb[8] if fb[8] is not None else (es[8] or 0.0))
                xg_f = max(float(fb[9] or 0.0), float(es[9] or 0.0))
                crd_y = float(fb[10] if fb[10] is not None else (es[10] or 0.0))
                crd_r = float(fb[11] if fb[11] is not None else (es[11] or 0.0))

                tiros_t = es[12] if es[12] is not None else fb[12]
                tiros_p = es[13] if es[13] is not None else fb[13]
                faltas_c = es[14] if es[14] is not None else fb[14]
                faltas_r = es[15] if es[15] is not None else fb[15]
                fueras_j = es[16] if (es[16] or 0) > 0 else (fb[16] or 0)
                tackles = es[17] if (es[17] or 0) > 0 else (fb[17] or 0)
                p_int = es[18] if es[18] is not None else fb[18]
                p_cmp = es[19] if es[19] is not None else fb[19]
                p_clv = es[20] if es[20] is not None else fb[20]
                ent = es[21] if es[21] is not None else fb[21]
                ent_g = es[22] if es[22] is not None else fb[22]
                inter = es[23] if es[23] is not None else fb[23]

                cur.execute("DELETE FROM estadisticas_jugadores WHERE rowid = ?", (rid,))
                cur.execute(
                    """
                    UPDATE estadisticas_jugadores
                    SET Equipo=?, posicion=?, minutos=?, partidos_jugados=?, titularidades=?,
                        goles=?, asistencias=?, xg=?, tarjetas_amarillas=?, tarjetas_rojas=?,
                        tiros_totales=?, tiros_a_puerta=?, faltas_cometidas=?, Faltas_Recibidas=?,
                        fueras_juego=?, entradas_tackles=?, Pases_Intentados=?, Pases_Completados=?,
                        Pases_Clave=?, Entradas=?, Entradas_Ganadas=?, Intercepciones=?
                    WHERE rowid=?
                    """,
                    (
                        eq_canon, pos_f, min_f, pj_f, tit_f,
                        gol_f, ast_f, xg_f, crd_y, crd_r,
                        tiros_t, tiros_p, faltas_c, faltas_r,
                        fueras_j, tackles, p_int, p_cmp,
                        p_clv, ent, ent_g, inter,
                        match_rid,
                    ),
                )
                fusionados += 1

    if cambios_club or fusionados:
        log.info(
            f"[Jugadores] Armonización de clubes completada: "
            f"{renombrados} filas renombradas y {fusionados} duplicados fusionados."
        )
    return {"renombrados": renombrados, "fusionados": fusionados}


_MAPA_POS_ESPN = {
    "G": "POR", "GK": "POR", "GOALKEEPER": "POR",
    "D": "DFC", "CD": "DFC", "CD-L": "DFC", "CD-R": "DFC", "CB": "DFC", "DEFENDER": "DFC",
    "LB": "LI", "LWB": "LI", "RB": "LD", "RWB": "LD", "SW": "DFC",
    "DM": "MCD", "CDM": "MCD", "CM": "MC", "CM-L": "MC", "CM-R": "MC", "M": "MC",
    "MIDFIELDER": "MC", "AM": "MCO", "CAM": "MCO",
    "LM": "EI", "LW": "EI", "RM": "ED", "RW": "ED",
    "F": "DC", "FW": "DC", "CF": "DC", "CF-L": "DC", "CF-R": "DC", "ST": "DC", "FORWARD": "DC",
}


def _normalizar_posicion_tactica(pos_raw: str) -> str:
    """Normaliza una posición ESPN (abreviatura o nombre) al código táctico del Scouting."""
    p = str(pos_raw or "").strip().upper()
    if not p:
        return "S/P"
    if p in _MAPA_POS_ESPN:
        return _MAPA_POS_ESPN[p]
    for token in re.split(r"[\s/]+", p):
        if token in _MAPA_POS_ESPN:
            return _MAPA_POS_ESPN[token]
    return "S/P"


def _stats_roster_plano(atleta: dict) -> dict:
    """Aplana athlete['statistics'] (splits → categories → stats) a {name: value}."""
    plano: dict = {}
    try:
        stats = atleta.get("statistics") or {}
        splits = stats.get("splits") if isinstance(stats, dict) else None
        for cat in (splits or {}).get("categories", []) or []:
            for st_ in cat.get("stats", []) or []:
                nombre = st_.get("name")
                if nombre is not None:
                    plano[nombre] = float(st_.get("value") or 0.0)
    except Exception:
        pass
    return plano


_SQL_UPSERT_JUGADOR = """
INSERT INTO estadisticas_jugadores
    (Jugador, Equipo, posicion, minutos, partidos_jugados, titularidades,
     goles, asistencias, tiros_totales, tiros_a_puerta, faltas_cometidas, Faltas_Recibidas,
     tarjetas_amarillas, tarjetas_rojas, fueras_juego, entradas_tackles, Pases_Completados,
     competicion)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
ON CONFLICT(Jugador, Equipo, competicion) DO UPDATE SET
    posicion=excluded.posicion, minutos=excluded.minutos,
    partidos_jugados=excluded.partidos_jugados, titularidades=excluded.titularidades,
    goles=excluded.goles, asistencias=excluded.asistencias,
    tiros_totales=excluded.tiros_totales, tiros_a_puerta=excluded.tiros_a_puerta,
    faltas_cometidas=excluded.faltas_cometidas, Faltas_Recibidas=excluded.Faltas_Recibidas,
    tarjetas_amarillas=excluded.tarjetas_amarillas, tarjetas_rojas=excluded.tarjetas_rojas,
    fueras_juego=excluded.fueras_juego, entradas_tackles=excluded.entradas_tackles,
    Pases_Completados=excluded.Pases_Completados
"""


_SQL_UPDATE_ESPN_UEFA = """
UPDATE estadisticas_jugadores
SET tiros_totales      = ?,
    tiros_a_puerta     = ?,
    faltas_cometidas   = ?,
    Faltas_Recibidas   = ?,
    fueras_juego       = ?,
    posicion           = CASE WHEN posicion IS NULL OR posicion IN ('', 'S/P') THEN ? ELSE posicion END,
    minutos            = CASE WHEN minutos IS NULL OR minutos <= 0 THEN ? ELSE minutos END,
    partidos_jugados   = CASE WHEN partidos_jugados IS NULL OR partidos_jugados <= 0 THEN ? ELSE partidos_jugados END,
    titularidades      = CASE WHEN titularidades IS NULL OR titularidades <= 0 THEN ? ELSE titularidades END
WHERE rowid = ?
"""


def sincronizar_estadisticas_jugadores(conn: sqlite3.Connection) -> int:
    """
    Sincroniza `estadisticas_jugadores` desde el roster público de ESPN por equipo.
    - Para las 5 grandes ligas domésticas: purga la plantilla previa del club en esa liga
      e inserta/actualiza con `ON CONFLICT(Jugador, Equipo, competicion) DO UPDATE`
      (el Paso 3 posterior de FBref sobreescribe minutos y titularidades reales).
    - Para torneos europeos (Champions League, Europa League, Conference League):
      NO purga las filas existentes de FBref ni sobreescribe `minutos`, `partidos_jugados`
      o `titularidades` cuando ya vienen de FBref; hace coincidencia difusa por jugador
      dentro del club y actualiza quirúrgicamente `tiros_totales`, `tiros_a_puerta`,
      `faltas_cometidas`, `Faltas_Recibidas` y `fueras_juego` (o inserta el jugador si
      solo tiene minutos/apariciones en ESPN).
    """
    total_jugadores_actualizados = 0
    try:
        _migrar_esquema_estadisticas_jugadores(conn)
        mapa_bd = _obtener_equipos_conocidos(conn)
    except Exception as exc:
        log.error(f"[Jugadores] Error preparando esquema: {exc}")
        return 0

    cursor = conn.cursor()
    for liga, url_sb in URLS_ESPN_SCOREBOARD.items():
        comp_nombre = _CODIGO_A_COMPETICION.get(liga, liga)
        es_uefa = comp_nombre in _COMPETICIONES_UEFA
        base = url_sb.rsplit("/scoreboard", 1)[0]
        try:
            r_t = requests.get(f"{base}/teams?limit=100", timeout=TIMEOUT_ESPN)
            r_t.raise_for_status()
            equipos = r_t.json()["sports"][0]["leagues"][0]["teams"]
        except Exception as exc:
            log.warning(f"  [Jugadores - {comp_nombre}] No se pudo listar equipos: {exc}")
            continue

        tuplas_liga: list = []
        purgas: list = []
        for item in equipos:
            try:
                eq = item.get("team", {})
                eq_id, eq_nom = eq.get("id"), eq.get("displayName") or eq.get("name")
                if not eq_id or not eq_nom:
                    continue
                equipo_normalizado = resolver_equipo_canonico(eq_nom, mapa_bd)
                r = requests.get(f"{base}/teams/{eq_id}/roster", timeout=TIMEOUT_ESPN)
                r.raise_for_status()
                atletas = r.json().get("athletes", []) or []
                planos = []
                for a in atletas:
                    if isinstance(a, dict) and isinstance(a.get("items"), list):
                        planos.extend(a["items"])
                    else:
                        planos.append(a)

                tuplas_club = []
                for a in planos:
                    nombre = str(a.get("displayName") or a.get("fullName") or "").strip()
                    if not nombre:
                        continue
                    pos_obj = a.get("position") or {}
                    pos_raw = str(pos_obj.get("abbreviation") or pos_obj.get("displayName") or "").strip()
                    s = _stats_roster_plano(a)
                    pj = int(s.get("appearances", 0))
                    subs = int(s.get("subIns", 0))
                    tit = max(pj - subs, 0)
                    minutos = float(tit * 80 + subs * 25)
                    goles = int(s.get("totalGoals", 0))
                    asist = int(s.get("goalAssists", 0))
                    tuplas_club.append((
                        nombre, equipo_normalizado, _normalizar_posicion_tactica(pos_raw),
                        minutos, pj, tit, goles, asist,
                        int(s.get("totalShots", 0)), int(s.get("shotsOnTarget", 0)),
                        int(s.get("foulsCommitted", 0)), int(s.get("foulsSuffered", 0)),
                        int(s.get("yellowCards", 0)), int(s.get("redCards", 0)),
                        int(s.get("offsides", 0)), 0, 0,
                        comp_nombre,
                    ))
                if not tuplas_club:
                    log.warning(f"  [Jugadores - {comp_nombre}] Roster vacío para {eq_nom}; se conserva el previo.")
                    continue
                purgas.append(equipo_normalizado)
                tuplas_liga.extend(tuplas_club)
            except Exception as exc:
                log.warning(f"  [Jugadores - {comp_nombre}] Fallo en club {item.get('team', {}).get('displayName')}: {exc}")
                continue

        if not tuplas_liga:
            continue

        try:
            unicas = {(t[0], t[1]): t for t in tuplas_liga}
            if not es_uefa:
                for equipo_normalizado in purgas:
                    cursor.execute(
                        "DELETE FROM estadisticas_jugadores WHERE equipo = ? AND competicion = ?",
                        (equipo_normalizado, comp_nombre),
                    )
                cursor.executemany(_SQL_UPSERT_JUGADOR, list(unicas.values()))
                total_jugadores_actualizados += len(unicas)
                log.info(f"  [Jugadores - {comp_nombre}] {len(unicas)} jugadores sincronizados ({len(purgas)} clubes).")
            else:
                # Torneos UEFA: asociar tiros y faltas de ESPN sin sobreescribir minutos/titularidades de FBref
                existentes_uefa: dict = {}
                for rid, jug_db, eq_db in cursor.execute(
                    "SELECT rowid, Jugador, Equipo FROM estadisticas_jugadores WHERE competicion = ?",
                    (comp_nombre,),
                ):
                    existentes_uefa.setdefault(normalizar_nombre(eq_db), {})[rid] = normalizar_nombre(jug_db)

                usados_uefa: set = set()
                updates_uefa = []
                inserts_uefa = []
                for t in unicas.values():
                    nom_j, eq_j, pos_j, min_j, pj_j, tit_j, gol_j, ast_j = t[:8]
                    tiros_t, tiros_p, fls_c, fls_r, crd_y, crd_r, offs_j = t[8:15]
                    cands = {
                        k: v for k, v in existentes_uefa.get(normalizar_nombre(eq_j), {}).items()
                        if k not in usados_uefa
                    }
                    rid_m = _buscar_coincidencia(normalizar_nombre(nom_j), cands)
                    if rid_m is not None:
                        usados_uefa.add(rid_m)
                        updates_uefa.append((
                            tiros_t, tiros_p, fls_c, fls_r, offs_j,
                            pos_j, min_j, pj_j, tit_j, rid_m,
                        ))
                    elif pj_j > 0 or tiros_t > 0 or fls_c > 0 or gol_j > 0 or ast_j > 0:
                        inserts_uefa.append(t)

                if updates_uefa:
                    cursor.executemany(_SQL_UPDATE_ESPN_UEFA, updates_uefa)
                if inserts_uefa:
                    cursor.executemany(_SQL_UPSERT_JUGADOR, inserts_uefa)
                n_uefa = len(updates_uefa) + len(inserts_uefa)
                total_jugadores_actualizados += n_uefa
                log.info(
                    f"  [Jugadores - {comp_nombre}] {n_uefa} jugadores actualizados con tiros/faltas ESPN "
                    f"({len(updates_uefa)} fusionados sobre FBref + {len(inserts_uefa)} nuevos)."
                )
        except sqlite3.Error as exc:
            log.error(f"  [Jugadores - {comp_nombre}] Error de escritura SQLite: {exc}")
    conn.commit()
    return total_jugadores_actualizados


_MAPA_POS_FBREF = {"GK": "POR", "DF": "DFC", "MF": "MC", "FW": "DC"}

# Partículas y sufijos ignorados al buscar el "apellido principal".
_PARTICULAS_NOMBRE = {"jr", "junior", "sr", "de", "da", "do", "dos", "das", "del", "la", "el",
                      "van", "von", "di", "le", "y"}


def normalizar_nombre(texto) -> str:
    """Minúsculas, sin acentos/tildes (NFKD) y sin signos de puntuación."""
    t = unicodedata.normalize("NFKD", str(texto or ""))
    t = "".join(ch for ch in t if not unicodedata.combining(ch)).lower()
    t = t.replace("-", " ")
    t = re.sub(r"[^\w\s]", "", t)
    return " ".join(t.split())


def _tokens_nombre(norm: str) -> list:
    toks = [t for t in norm.split() if t not in _PARTICULAS_NOMBRE]
    return toks or norm.split()


def _token_principal(norm: str) -> str:
    """Token más largo (a igualdad, el último: suele ser el apellido)."""
    toks = _tokens_nombre(norm)
    return max(reversed(toks), key=len) if toks else ""


def _buscar_coincidencia(norm: str, candidatos: dict):
    """
    candidatos: {rowid: nombre_normalizado} (solo filas libres del MISMO equipo).
    Devuelve rowid o None. Solo acepta coincidencias unívocas.
    """
    if not candidatos:
        return None
    exactos = [rid for rid, n in candidatos.items() if n == norm]
    if len(exactos) == 1:
        return exactos[0]
    if len(exactos) > 1:
        return None

    tok_p = _token_principal(norm)
    toks = set(_tokens_nombre(norm))
    if not tok_p:
        return None

    def _desambiguar(rids):
        if len(rids) == 1:
            return rids[0]
        # Varios candidatos con mismo apellido: exigir >=2 tokens compartidos o misma inicial
        mejores = [r for r in rids if len(toks & set(_tokens_nombre(candidatos[r]))) >= 2]
        if len(mejores) == 1:
            return mejores[0]
        ini = norm[:1]
        mejores = [r for r in rids if candidatos[r][:1] == ini and ini]
        return mejores[0] if len(mejores) == 1 else None

    # Nivel 2: mismo apellido principal
    nivel2 = [r for r, n in candidatos.items()
              if tok_p in _tokens_nombre(n) or _token_principal(n) in toks]
    if nivel2:
        return _desambiguar(nivel2)
    # Nivel 3: prefijo ("vini" -> "vinicius"), mínimo 4 caracteres
    nivel3 = []
    for r, n in candidatos.items():
        tp = _token_principal(n)
        corto, largo = sorted((tok_p, tp), key=len)
        if len(corto) >= 4 and largo.startswith(corto):
            nivel3.append(r)
    return _desambiguar(nivel3) if nivel3 else None


def _num_fbref(serie: pd.Series, tipo=int):
    """Convierte una columna FBref a numérico (quita comas de miles, NaN -> 0)."""
    limpia = serie.astype(str).str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(limpia, errors="coerce").fillna(0).astype(tipo)


def _fusionar_csv_fbref(conn: sqlite3.Connection, ruta_csv, liga: str = "", confirmar: bool = True) -> dict:
    """
    Fusiona el CSV nativo de FBref (cabecera doble) en `estadisticas_jugadores` mediante
    fuzzy matching POR EQUIPO:
      * Coincidencia (nombre exacto normalizado o apellido principal unívoco) ->
        UPDATE quirúrgico solo de minutos, partidos_jugados, titularidades, goles,
        asistencias, xg, tarjetas (y posicion si faltaba). No se tocan tiros, faltas ni pases.
      * Sin coincidencia -> INSERT del registro limpio.
    Con confirmar=True hace commit/rollback propio; con False deja la transacción al llamador.
    Devuelve {"total", "fusionados", "insertados"}.
    """
    ruta = Path(ruta_csv)
    if not ruta.is_absolute() and not ruta.exists():
        ruta = Path(__file__).parent / ruta_csv
    df = pd.read_csv(ruta, header=1)
    df = df[df["Player"].notna() & (df["Player"] != "Player")].copy()
    if df.empty:
        log.warning(f"[FBref CSV {liga}] Sin filas válidas; no se modifica la BD.")
        return {"total": 0, "fusionados": 0, "insertados": 0}

    mapa_bd = _obtener_equipos_conocidos(conn)
    df["Squad"] = (
        df["Squad"].astype(str).str.replace(r"^[a-z]{2,3}\s+", "", regex=True).str.strip()
        .map(lambda s: resolver_equipo_canonico(s, mapa_bd))
    )
    df["pos_tact"] = df["Pos"].fillna("").astype(str).map(
        lambda p: _MAPA_POS_FBREF.get(p.split(",")[0].strip().upper(), "S/P")
    )
    df["v_min"] = _num_fbref(df["Min"])
    df["v_mp"] = _num_fbref(df["MP"])
    df["v_st"] = _num_fbref(df["Starts"])
    df["v_gls"] = _num_fbref(df["Gls"])
    df["v_ast"] = _num_fbref(df["Ast"])
    df["v_y"] = _num_fbref(df["CrdY"])
    df["v_r"] = _num_fbref(df["CrdR"])
    df["v_xg"] = _num_fbref(df["xG"], float) if "xG" in df.columns else 0.0
    df["Player"] = df["Player"].astype(str).str.strip()
    df = df.sort_values("v_min", ascending=False).drop_duplicates(["Player", "Squad"])

    sql_upd = (
        "UPDATE estadisticas_jugadores SET minutos=?, partidos_jugados=?, titularidades=?, "
        "goles=?, asistencias=?, xg=?, tarjetas_amarillas=?, tarjetas_rojas=?, "
        "posicion=CASE WHEN posicion IS NULL OR posicion IN ('', 'S/P') THEN ? ELSE posicion END "
        "WHERE rowid=?"
    )
    sql_ins = """
    INSERT INTO estadisticas_jugadores
        (Jugador, Equipo, posicion, minutos, partidos_jugados, titularidades,
         goles, asistencias, xg, tarjetas_amarillas, tarjetas_rojas, competicion)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(Jugador, Equipo, competicion) DO UPDATE SET
        minutos=excluded.minutos, partidos_jugados=excluded.partidos_jugados,
        titularidades=excluded.titularidades, goles=excluded.goles,
        asistencias=excluded.asistencias, xg=excluded.xg,
        tarjetas_amarillas=excluded.tarjetas_amarillas, tarjetas_rojas=excluded.tarjetas_rojas
    """
    competicion = liga or "LaLiga"

    try:
        _migrar_esquema_estadisticas_jugadores(conn)
        cur = conn.cursor()
        # Candidatos por equipo normalizado: {equipo_norm: {rowid: nombre_norm}}
        por_equipo: dict = {}
        # Misma competición: candidatos de UPDATE. Nombres globales: solo para canonizar en UEFA.
        globales: dict = {}
        for rid, jug, eq, comp in cur.execute(
                "SELECT rowid, Jugador, Equipo, competicion FROM estadisticas_jugadores"):
            if comp == competicion:
                por_equipo.setdefault(normalizar_nombre(eq), {})[rid] = normalizar_nombre(jug)
            globales.setdefault(normalizar_nombre(eq), {}).setdefault(normalizar_nombre(jug), jug)

        registros = [
            dict(
                jugador=r.Player, equipo=r.Squad, norm=normalizar_nombre(r.Player),
                pos=r.pos_tact,
                vals=(float(r.v_min), int(r.v_mp), int(r.v_st), float(r.v_gls), float(r.v_ast),
                      float(r.v_xg), float(r.v_y), float(r.v_r)),
            )
            for r in df.itertuples(index=False)
        ]
        if competicion in _COMPETICIONES_UEFA:
            # Reutiliza el nombre ya existente del jugador en su club (ESPN/liga) para que la
            # vista global agregue correctamente liga + torneo europeo.
            for reg in registros:
                glob = globales.get(normalizar_nombre(reg["equipo"]), {})
                nombres = {i: n for i, n in enumerate(glob)}
                reales = list(glob.values())
                k = _buscar_coincidencia(reg["norm"], nombres)
                if k is not None:
                    reg["jugador"] = reales[k]
                    reg["norm"] = normalizar_nombre(reales[k])
        usados: set = set()
        actualizaciones, inserciones = [], []

        # Pasada 1: coincidencias exactas; pasada 2: difusas sobre los restantes.
        pendientes = []
        for reg in registros:
            cands = {k: v for k, v in por_equipo.get(normalizar_nombre(reg["equipo"]), {}).items() if k not in usados}
            exactos = [k for k, v in cands.items() if v == reg["norm"]]
            if len(exactos) == 1:
                usados.add(exactos[0])
                actualizaciones.append((*reg["vals"], reg["pos"], exactos[0]))
            else:
                pendientes.append(reg)
        for reg in pendientes:
            cands = {k: v for k, v in por_equipo.get(normalizar_nombre(reg["equipo"]), {}).items() if k not in usados}
            rid = _buscar_coincidencia(reg["norm"], cands)
            if rid is not None:
                usados.add(rid)
                actualizaciones.append((*reg["vals"], reg["pos"], rid))
            else:
                inserciones.append((reg["jugador"], reg["equipo"], reg["pos"], *reg["vals"], competicion))

        cur.executemany(sql_upd, actualizaciones)
        cur.executemany(sql_ins, inserciones)
        if confirmar:
            conn.commit()
    except sqlite3.Error as exc:
        if confirmar:
            conn.rollback()
        log.error(f"[FBref CSV {liga}] Error SQLite: {exc}")
        raise
    log.info(f"[FBref CSV {liga}] {len(registros)} consolidados: "
             f"{len(actualizaciones)} fusionados (fuzzy) + {len(inserciones)} insertados.")
    return {"total": len(registros), "fusionados": len(actualizaciones), "insertados": len(inserciones)}


def sincronizar_desde_fbref_csv(conn: sqlite3.Connection, ruta_csv="fbref_jugadores.csv") -> int:
    """Compatibilidad: fusiona un único CSV de FBref y devuelve nº de jugadores consolidados."""
    return _fusionar_csv_fbref(conn, ruta_csv, "LaLiga")["total"]


ARCHIVOS_FBREF = [
    ('fbref_jugadores.csv', 'LaLiga'),
    ('fbref_premier.csv', 'Premier League'),
    ('fbref_seriea.csv', 'Serie A'),
    ('fbref_bundesliga.csv', 'Bundesliga'),
    ('fbref_ligue1.csv', 'Ligue 1'),
    ('fbref_champions.csv', 'Champions League'),
    ('fbref_europaleague.csv', 'Europa League'),
    ('fbref_conference.csv', 'Conference League'),
]


def sincronizar_todas_las_ligas_fbref(conn: sqlite3.Connection) -> dict:
    """
    Fusiona los 5 CSVs de FBref (Top 5 ligas) en `estadisticas_jugadores` con fuzzy matching
    por equipo (UPDATE quirúrgico si coincide, INSERT si no). Los archivos ausentes se omiten.
    Un único commit al final (rollback total si algún archivo falla).
    Devuelve {"por_liga": {liga: {...}}, "total_consolidados": int, "total_tabla": int}.
    """
    base = Path(__file__).parent
    por_liga: dict = {}
    try:
        _migrar_esquema_estadisticas_jugadores(conn)
        for archivo, liga in ARCHIVOS_FBREF:
            ruta = _resolver_ruta_fbref(archivo)
            if not ruta:
                log.info(f"[FBref] {archivo} no existe; se ignora {liga}.")
                continue
            por_liga[liga] = _fusionar_csv_fbref(conn, ruta, liga, confirmar=False)
        armonizar_clubes_estadisticas_jugadores(conn)
        conn.commit()
    except Exception as exc:
        conn.rollback()
        log.error(f"[FBref] Sincronización multi-liga abortada con rollback: {exc}")
        raise
    total = sum(v["total"] for v in por_liga.values())
    total_tabla = conn.execute("SELECT COUNT(*) FROM estadisticas_jugadores").fetchone()[0]
    print("\n[FBref] Resumen de consolidación por liga:")
    for liga, v in por_liga.items():
        print(f"  • {liga:<15} {v['total']:>4} jugadores ({v['fusionados']} fusionados + {v['insertados']} insertados)")
    print(f"  Total consolidados: {total} | Total en 'estadisticas_jugadores': {total_tabla}")
    return {"por_liga": por_liga, "total_consolidados": total, "total_tabla": total_tabla}


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 4 – PUNTO DE ENTRADA PÚBLICO
# ══════════════════════════════════════════════════════════════════════════════

def sincronizar_jornada() -> dict:
    """
    Sincroniza `historico_reciente` en `macro_base_futbol.db` mediante:
      1. Tier 1: CSVs de Football-Data.co.uk (2026/2027).
      2. Tier 2: API JSON abierta de ESPN (sin Cloudflare 403).
    """
    t_inicio = time.perf_counter()

    resultado = {
        "exito": False,
        "partidos_nuevos": 0,
        "partidos_actualizados": 0,
        "tier1_insertados": 0,
        "tier1_actualizados": 0,
        "tier2_insertados": 0,
        "tier2_actualizados": 0,
        "partidos_añadidos_detalle": [],
        "jugadores_act": 0,
        "jugadores_total": 0,
        "columnas_jugadores": [],
        "tiempo_ejecucion": 0.0,
        "errores": [],
        "mensaje": "",
    }

    if not RUTA_BD.exists():
        msg = f"Base de datos no encontrada en: {RUTA_BD}"
        resultado["mensaje"] = msg
        resultado["errores"].append(msg)
        log.error(msg)
        return resultado

    log.info(f"=== Iniciando sincronizar_jornada() [Tier 1 CSV + Tier 2 ESPN JSON] → {RUTA_BD.name} ===")

    try:
        with sqlite3.connect(RUTA_BD) as conn:
            _garantizar_indices(conn)
            mapa_equipos_bd = _obtener_equipos_conocidos(conn)

            # ── TIER 1: Base Line (Football-Data.co.uk CSVs) ──────────────────
            t1_ins, t1_act, t1_det, err_t1 = ejecutar_tier1_csv(conn, mapa_equipos_bd)
            resultado["tier1_insertados"] = t1_ins
            resultado["tier1_actualizados"] = t1_act
            resultado["partidos_añadidos_detalle"].extend(t1_det)

            if t1_ins > 0:
                mapa_equipos_bd = _obtener_equipos_conocidos(conn)

            # ── TIER 2: Live Delta Sync (ESPN Open JSON API, sin 403) ─────────
            t2_ins, t2_act, t2_det, err_t2 = ejecutar_tier2_espn(conn, mapa_equipos_bd)
            resultado["tier2_insertados"] = t2_ins
            resultado["tier2_actualizados"] = t2_act
            resultado["partidos_añadidos_detalle"].extend(t2_det)

            resultado["partidos_nuevos"] = t1_ins + t2_ins
            resultado["partidos_actualizados"] = t1_act + t2_act

            # ── PASO 2: faltas/tiros desde ESPN (aislado: un fallo no aborta la sync) ──
            try:
                total_jugadores = sincronizar_estadisticas_jugadores(conn)
            except Exception as exc_j:
                total_jugadores = 0
                log.error(f"[Jugadores] Sincronización ESPN fallida: {exc_j}")

            # ── PASO 3: fusión definitiva de minutos reales desde los 5 CSVs de FBref ──
            try:
                total_csv = sincronizar_todas_las_ligas_fbref(conn)["total_consolidados"]
                if total_csv:
                    total_jugadores = total_csv
            except Exception as exc_c:
                log.error(f"[Jugadores] Ingesta FBref CSV omitida: {exc_c}")
            conn.commit()  # commit atómico final
            resultado["jugadores_total"] = conn.execute(
                "SELECT COUNT(*) FROM estadisticas_jugadores").fetchone()[0]
            resultado["columnas_jugadores"] = [
                f[1] for f in conn.execute("PRAGMA table_info(estadisticas_jugadores)")]
            print(f"[Jugadores] Futbolistas consolidados en BD: {resultado['jugadores_total']}")
            print(f"[Jugadores] Columnas finales: {', '.join(resultado['columnas_jugadores'])}")
            resultado["jugadores_act"] = total_jugadores

            if err_t1 and err_t2:
                resultado["errores"].extend(err_t1)
                resultado["errores"].extend(err_t2)

    except sqlite3.Error as exc:
        msg = f"Error crítico de base de datos: {exc}"
        resultado["mensaje"] = msg
        resultado["errores"].append(msg)
        log.error(msg)
        return resultado

    duracion = round(time.perf_counter() - t_inicio, 2)
    resultado["tiempo_ejecucion"] = duracion

    n_ins = resultado["partidos_nuevos"]
    m_act = resultado["partidos_actualizados"]

    if not resultado["errores"]:
        resultado["exito"] = True
        if n_ins > 0 or m_act > 0:
            resultado["mensaje"] = (
                f"Sincronización completada: {n_ins} partidos nuevos "
                f"(Tier 1 CSV: {resultado['tier1_insertados']}, "
                f"Tier 2 ESPN: {resultado['tier2_insertados']}) y "
                f"{m_act} actualizados."
            )
        else:
            resultado["mensaje"] = (
                "Base de datos al día (Tier 1 CSV + Tier 2 ESPN). "
                "No hay partidos nuevos pendientes."
            )
    else:
        resultado["exito"] = False
        resultado["mensaje"] = "Sincronización fallida. Revisa los errores."

    lineas_detalle = (
        "\n".join(f"  ✓ {d}" for d in resultado["partidos_añadidos_detalle"])
        if resultado["partidos_añadidos_detalle"]
        else "  • (Ningún partido nuevo insertado en esta ejecución)"
    )

    resumen_stdout = (
        "\n============================================================\n"
        "[RESUMEN DE EJECUCIÓN - PIPELINE HÍBRIDO (CSV + ESPN API)]\n"
        f"• Tier 1 (Football-Data CSV) -> Insertados: {resultado['tier1_insertados']} | "
        f"Actualizados: {resultado['tier1_actualizados']}\n"
        f"• Tier 2 (ESPN Live API)     -> Insertados: {resultado['tier2_insertados']} | "
        f"Actualizados: {resultado['tier2_actualizados']}\n"
        f"• Total Global               -> Insertados: {n_ins} | Actualizados: {m_act}\n"
        f"• Tiempo de ejecución        -> {duracion:.2f} s\n"
        "------------------------------------------------------------\n"
        "Confirmación de partidos añadidos:\n"
        f"{lineas_detalle}\n"
        "============================================================"
    )
    print(resumen_stdout)
    log.info(resumen_stdout)
    return resultado


if __name__ == "__main__":
    import json
    res = sincronizar_jornada()
    with sqlite3.connect(RUTA_BD) as _c:
        _filas_27 = _c.execute(
            "SELECT Liga_Origen, Date, HomeTeam, FTHG, FTAG, AwayTeam, FTR "
            "FROM historico_reciente WHERE Date = '27/10/2026'"
        ).fetchall()
        print("\n[VERIFICACIÓN EN BD - Partidos del 27/10/2026 en 'historico_reciente']:")
        for _f in _filas_27:
            print(f"  • {_f[0]} | {_f[1]}: {_f[2]} {_f[3]} - {_f[4]} {_f[5]} (FTR={_f[6]})")
    print(json.dumps(res, indent=2, ensure_ascii=False))
