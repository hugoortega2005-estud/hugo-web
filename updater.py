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
    # Premier League (E0)
    "manchester united": "Man United", "man utd": "Man United",
    "manchester city": "Man City",
    "nottingham forest": "Nott'm Forest",
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
    "borussia m'gladbach": "M'gladbach",
    "eintracht frankfurt": "Ein Frankfurt",
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
    "fc cologne": "FC Koln", "1. fc cologne": "FC Koln",
    "hamburger sv": "Hamburg", "hamburg": "Hamburg",
    "sv elversberg": "Elversberg", "elversberg": "Elversberg",
    "sc paderborn 07": "Paderborn", "sc paderborn": "Paderborn", "paderborn": "Paderborn",
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

def _extraer_eventos_espn_liga(
    liga: str,
    url: str,
    mapa_equipos_bd: dict[str, str],
    dias_ventana: int = 5,
) -> list[dict]:
    """
    Descarga el JSON de ESPN Scoreboard con `requests.get(url, timeout=10)`
    para los últimos `dias_ventana` días (`?dates=YYYYMMDD`) y extrae los
    partidos finalizados (`event['status']['type']['completed'] == True`).
    """
    partidos_extraidos = []
    vistos = set()

    hoy = datetime.now()
    consultas = [
        url if delta == 0 else f"{url}?dates={(hoy - timedelta(days=delta)).strftime('%Y%m%d')}"
        for delta in range(dias_ventana)
    ]

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

            clave = (liga, fecha_ddmmyyyy, equipo_local, equipo_visitante)
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
                "Liga_Origen": liga,
                "Temporada_Origen": TEMPORADA_ACTUAL,
            })

    return partidos_extraidos


def ejecutar_tier2_espn(
    conn: sqlite3.Connection,
    mapa_equipos_bd: dict[str, str],
) -> tuple[int, int, list[str], list[str]]:
    """
    Tier 2 (Live Delta Sync): Recorre los 5 endpoints de ESPN Scoreboard,
    comprueba existencia en `historico_reciente` por `(HomeTeam, AwayTeam, Date)`,
    ejecuta `INSERT` si no existe o `UPDATE` si existe sin marcador, e imprime
    confirmación de cada partido nuevo añadido.
    """
    log.info("[Tier 2 - ESPN API] Consultando resultados en tiempo real de las 5 grandes ligas...")
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(historico_reciente)")
    cols_bd = {row[1] for row in cur.fetchall()}
    tiene_col_div_bom = COL_DIV_BD in cols_bd

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

    t2_insertados = 0
    t2_actualizados = 0
    t2_detalles = []
    errores = []
    ligas_ok = 0

    for liga, url in URLS_ESPN_SCOREBOARD.items():
        partidos = _extraer_eventos_espn_liga(liga, url, mapa_equipos_bd)
        if partidos:
            ligas_ok += 1

        ins_liga = 0
        act_liga = 0
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
            f"  [Tier 2 - {liga}] {len(partidos)} partidos finalizados evaluados -> "
            f"{ins_liga} insertados, {act_liga} actualizados."
        )

    conn.commit()
    return t2_insertados, t2_actualizados, t2_detalles, errores


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
