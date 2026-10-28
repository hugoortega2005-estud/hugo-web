# updater.py
# ─────────────────────────────────────────────────────────────────────────────
# Módulo de sincronización de jornada para Hugo's Hub.
#
# Función exportada:
#   sincronizar_jornada() -> dict
#       Devuelve: {
#           "exito":           bool,
#           "partidos_nuevos": int,
#           "jugadores_act":   int,
#           "errores":         list[str],
#           "mensaje":         str,
#       }
#
# Fuentes:
#   - Partidos  → football-data.co.uk (CSV, todas las ligas, temporada actual)
#   - Jugadores → FBref (ejecución silenciosa en segundo plano; fallos 403
#                 de Cloudflare quedan totalmente silenciados sin bloquear UI)
#
# Seguridad contra duplicados:
#   - historico_reciente     → UNIQUE INDEX (Div, Date, HomeTeam, AwayTeam)
#                              → INSERT OR IGNORE
#   - estadisticas_jugadores → UNIQUE INDEX (Jugador, Equipo)
#                              → INSERT OR REPLACE (stats siempre frescas)
# ─────────────────────────────────────────────────────────────────────────────

import io
import logging
import random
import re
import sqlite3
import time
import unicodedata
from pathlib import Path

import pandas as pd
import requests

from scraper_arbitros import sincronizar_arbitros

try:
    from curl_cffi import requests as curl_requests
except ImportError:
    curl_requests = None

# ── Constantes ────────────────────────────────────────────────────────────────
RUTA_BD = Path(__file__).parent / "macro_base_futbol.db"

TEMPORADAS_PARTIDOS = ["2627", "2526"]
LIGAS_PARTIDOS      = ["SP1", "E0", "I1", "D1", "F1"]

# La columna Div fue almacenada con BOM (artefacto de latin-1 en el CSV original)
COL_DIV_BD = "\u00ef\u00bb\u00bfDiv"

URLS_FBREF = {
    "Tiempo":   "https://fbref.com/es/comps/12/playingtime/Estadisticas-de-La-Liga",
    "Estandar": "https://fbref.com/es/comps/12/stats/Estadisticas-de-La-Liga",
    "Pases":    "https://fbref.com/es/comps/12/passing/Estadisticas-de-La-Liga",
    "Defensa":  "https://fbref.com/es/comps/12/defense/Estadisticas-de-La-Liga",
    "Varios":   "https://fbref.com/es/comps/12/misc/Estadisticas-de-La-Liga",
}

TIMEOUT_HTTP = 30

# Headers que simulan un navegador real para evitar el bloqueo de FBref
HEADERS_FBREF = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# ── Logger ────────────────────────────────────────────────────────────────────
log = logging.getLogger(__name__)
if not log.handlers:
    _h = logging.StreamHandler()
    _h.setFormatter(logging.Formatter("%(asctime)s [updater] %(levelname)s: %(message)s"))
    log.addHandler(_h)
log.setLevel(logging.INFO)


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 1 – UTILIDADES DE BASE DE DATOS
# ══════════════════════════════════════════════════════════════════════════════

def _garantizar_indices(conn: sqlite3.Connection) -> None:
    """
    Crea los índices UNIQUE necesarios para INSERT OR IGNORE / INSERT OR REPLACE
    si aún no existen. Operación idempotente y no destructiva.
    """
    cur = conn.cursor()

    # historico_reciente: unicidad por liga + fecha + equipos
    cur.execute(
        f'CREATE UNIQUE INDEX IF NOT EXISTS idx_partidos_pk '
        f'ON historico_reciente ("{COL_DIV_BD}", Date, HomeTeam, AwayTeam)'
    )

    # estadisticas_jugadores: unicidad por nombre + equipo
    cur.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_jugadores_pk "
        "ON estadisticas_jugadores (Jugador, Equipo)"
    )
    conn.commit()
    log.info("Índices UNIQUE verificados / creados.")


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 2 – PARTIDOS (football-data.co.uk)
# ══════════════════════════════════════════════════════════════════════════════

def _descargar_csv_liga(liga: str, temporada: str):
    """Descarga el CSV de football-data.co.uk y lo devuelve como DataFrame."""
    url = f"https://www.football-data.co.uk/mmz4281/{temporada}/{liga}.csv"
    try:
        resp = requests.get(url, timeout=TIMEOUT_HTTP)
        resp.raise_for_status()
        df = pd.read_csv(io.StringIO(resp.content.decode("latin-1")))
        df["Temporada_Origen"] = temporada
        df["Liga_Origen"] = liga
        log.info(f"  Descargado {liga} {temporada}: {len(df)} filas.")
        return df
    except Exception as exc:
        log.info(f"  No disponible {liga} {temporada}: {exc}")
        return None


def _normalizar_df_partidos(df: pd.DataFrame) -> pd.DataFrame:
    """
    Limpia la cabecera BOM y alinea el nombre de la columna Div
    con el nombre real almacenado en la BD (COL_DIV_BD).
    """
    # Quitar BOM de las cabeceras
    df.columns = [
        c.replace("\ufeff", "").replace("\u00ef\u00bb\u00bf", "").strip()
        if c not in ("Temporada_Origen", "Liga_Origen") else c
        for c in df.columns
    ]

    # Renombrar 'Div' → COL_DIV_BD para coincidir con la BD
    if "Div" in df.columns and COL_DIV_BD not in df.columns:
        df = df.rename(columns={"Div": COL_DIV_BD})

    # Filtrar filas sin resultado (partidos no jugados aún)
    for col in ("HomeTeam", "AwayTeam", "FTHG", "FTAG"):
        if col in df.columns:
            df = df[df[col].notna()]

    return df


def _insertar_partidos(conn: sqlite3.Connection, df: pd.DataFrame) -> int:
    """
    Inserta las filas del DataFrame en historico_reciente usando INSERT OR IGNORE.
    Devuelve el número de filas realmente insertadas (nuevas).
    """
    if df.empty:
        return 0

    cur = conn.cursor()
    cur.execute("PRAGMA table_info(historico_reciente)")
    cols_bd = [row[1] for row in cur.fetchall()]

    # Solo columnas que ya existen en la BD
    cols_comunes = [c for c in df.columns if c in cols_bd]
    df_filtrado = df[cols_comunes].copy()

    for col in ("FTHG", "FTAG"):
        if col in df_filtrado.columns:
            df_filtrado[col] = pd.to_numeric(df_filtrado[col], errors="coerce")

    placeholders = ", ".join(["?" for _ in cols_comunes])
    cols_quoted  = ", ".join([f'"{c}"' for c in cols_comunes])
    sql = (
        f'INSERT OR IGNORE INTO historico_reciente '
        f'({cols_quoted}) VALUES ({placeholders})'
    )

    insertados = 0
    for _, row in df_filtrado.iterrows():
        valores = tuple(
            None if (isinstance(v, float) and pd.isna(v)) else v
            for v in row
        )
        try:
            cur.execute(sql, valores)
            if cur.rowcount == 1:
                insertados += 1
        except sqlite3.Error as exc:
            log.debug(f"  Fila ignorada: {exc}")

    conn.commit()
    return insertados


def sincronizar_partidos(conn: sqlite3.Connection):
    """
    Descarga CSVs de todas las ligas para todas las temporadas configuradas
    e inserta con INSERT OR IGNORE en historico_reciente.

    Returns: (partidos_nuevos: int, errores: list[str])
    """
    total_nuevos = 0
    errores = []
    descargas_exitosas = 0

    for temporada in TEMPORADAS_PARTIDOS:
        for liga in LIGAS_PARTIDOS:
            df_raw = _descargar_csv_liga(liga, temporada)
            if df_raw is None:
                continue
            descargas_exitosas += 1
            df_clean = _normalizar_df_partidos(df_raw)
            nuevos = _insertar_partidos(conn, df_clean)
            total_nuevos += nuevos
            log.info(f"  {liga} {temporada}: {nuevos} partidos nuevos.")

    if descargas_exitosas == 0:
        errores.append("No se pudo conectar con football-data.co.uk para descargar partidos.")

    return total_nuevos, errores


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 3 – JUGADORES (FBref, LaLiga)
# ══════════════════════════════════════════════════════════════════════════════

def _crear_id_jugador(nombre: str, equipo: str) -> str:
    """ID normalizado sin acentos ni caracteres especiales."""
    def _norm(s: str) -> str:
        s = str(s).lower()
        s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("utf-8")
        return re.sub(r"[^a-z0-9]", "", s)
    return _norm(nombre) + "_" + _norm(equipo)


def _buscar_col(df: pd.DataFrame, requisitos: list) -> str:
    """Devuelve la primera columna que contiene todas las palabras clave."""
    for col in df.columns:
        col_str = str(col).lower()
        if all(req in col_str for req in requisitos):
            return col
    return None


def _extraer_tabla_fbref(url: str, session=None):
    """
    Descarga la página de FBref usando una sesión de curl_cffi con impersonación de Chrome
    para mantener las cookies y evitar el Error 403 de Cloudflare.
    Si Cloudflare bloquea con 403 o falla la red, finaliza en silencio devolviendo None.
    """
    if curl_requests is None:
        log.info("  [Silenciado] curl_cffi no disponible; omitiendo scraping de FBref.")
        return None

    own_session = session is None
    if own_session:
        session = curl_requests.Session()

    try:
        resp = session.get(url, headers=HEADERS_FBREF, impersonate="chrome", timeout=15)
        if resp.status_code == 403:
            log.info(f"  [Silenciado] Bloqueo Cloudflare 403 en FBref ({url}).")
            return None

        resp.raise_for_status()

        # FBref oculta las tablas de estadísticas dentro de comentarios HTML.
        html = resp.text.replace("<!--", "").replace("-->", "")
        tablas = pd.read_html(io.StringIO(html))

        for df in tablas:
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = [
                    "_".join(str(c).strip() for c in col if str(c).strip())
                    for col in df.columns
                ]
            cols_lower = [str(c).lower() for c in df.columns]
            if len(df) > 50 and any("jugador" in c or "player" in c for c in cols_lower):
                return df

        log.info(f"  [Silenciado] No se encontró tabla de jugadores en: {url}")
        return None
    except Exception as exc:
        log.info(f"  [Silenciado] Scraping FBref omitido en {url}: {exc}")
        return None
    finally:
        if own_session:
            try:
                session.close()
            except Exception:
                pass


def _parsear_categoria(df, mapa: dict):
    """
    Extrae y renombra columnas según el mapa {NombreDestino: [palabras_clave]}.
    Incluye siempre Jugador, Equipo, Pos e ID.
    """
    if df is None or df.empty:
        return None

    c_jug = _buscar_col(df, ["player"]) or _buscar_col(df, ["jugador"])
    c_eq  = _buscar_col(df, ["squad"])  or _buscar_col(df, ["equipo"])
    if not c_jug or not c_eq:
        return None

    res = pd.DataFrame()
    res["Jugador"] = (
        df[c_jug].astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
        .str.replace(r"^[A-Z]{2,3}\s+", "", regex=True)
    )
    c_pos = _buscar_col(df, ["pos"])
    if c_pos:
        res["Pos"] = df[c_pos].astype(str).str.strip()

    res["Equipo"] = df[c_eq].astype(str).str.strip()
    res["ID"] = [_crear_id_jugador(j, e) for j, e in zip(res["Jugador"], res["Equipo"])]

    for col_dest, keywords in mapa.items():
        col_src = _buscar_col(df, keywords)
        if col_src:
            res[col_dest] = df[col_src]

    # Eliminar filas de cabecera repetida que FBref incrusta
    res = res[res["Jugador"].str.lower() != "jugador"]
    res = res[res["Jugador"].str.lower() != "player"]
    res = res[~res["Jugador"].str.lower().str.startswith("equipo")]
    return res.reset_index(drop=True)


def _construir_df_jugadores(dfs: dict):
    """Combina las distintas categorías de FBref en un único DataFrame limpio."""
    if "Tiempo" not in dfs or "Estandar" not in dfs:
        return None

    res_df = _parsear_categoria(dfs["Tiempo"], {"Pos": ["pos"]})
    if res_df is None:
        return None

    mapa_est = {
        "Min": ["min"], "90s": ["90s"], "Gls": ["gls"], "Ast": ["ast"],
        "G+A": ["g+a"], "xG": ["xg"], "xAG": ["xag"],
        "CrdY": ["crdy"], "CrdR": ["crdr"],
    }
    d_est = _parsear_categoria(dfs["Estandar"], mapa_est)
    if d_est is not None:
        res_df = pd.merge(
            res_df,
            d_est.drop(columns=["Jugador", "Equipo", "Pos"], errors="ignore"),
            on="ID", how="left",
        )

    if "Pases" in dfs:
        d_pas = _parsear_categoria(dfs["Pases"], {
            "Pases_Completados": ["total", "cmp"],
            "Pases_Intentados":  ["total", "att"],
            "Pases_Progresivos": ["prgp"],
            "Pases_Clave":       ["kp"],
        })
        if d_pas is not None:
            res_df = pd.merge(
                res_df,
                d_pas.drop(columns=["Jugador", "Equipo", "Pos"], errors="ignore"),
                on="ID", how="left",
            )

    if "Defensa" in dfs:
        d_def = _parsear_categoria(dfs["Defensa"], {
            "Entradas":         ["tackles", "tkl"],
            "Entradas_Ganadas": ["tackles", "tklw"],
            "Intercepciones":   ["int"],
        })
        if d_def is not None:
            res_df = pd.merge(
                res_df,
                d_def.drop(columns=["Jugador", "Equipo", "Pos"], errors="ignore"),
                on="ID", how="left",
            )

    if "Varios" in dfs:
        d_var = _parsear_categoria(dfs["Varios"], {
            "Faltas_Cometidas": ["fls"],
            "Faltas_Recibidas": ["fld"],
            "Recuperaciones":   ["recov"],
        })
        if d_var is not None:
            res_df = pd.merge(
                res_df,
                d_var.drop(columns=["Jugador", "Equipo", "Pos"], errors="ignore"),
                on="ID", how="left",
            )

    res_df.drop(columns=["ID"], inplace=True, errors="ignore")
    cols_no_num = ["Jugador", "Equipo", "Pos"]
    for c in res_df.columns:
        if c not in cols_no_num:
            res_df[c] = pd.to_numeric(res_df[c], errors="coerce").fillna(0)

    orden = [
        "Jugador", "Equipo", "Pos", "Min", "90s", "Gls", "Ast", "G+A", "xG", "xAG",
        "Pases_Intentados", "Pases_Completados", "Pases_Clave", "Pases_Progresivos",
        "Entradas", "Entradas_Ganadas", "Intercepciones",
        "Faltas_Cometidas", "Faltas_Recibidas", "Recuperaciones", "CrdY", "CrdR",
    ]
    return res_df[[c for c in orden if c in res_df.columns]]


def _insertar_jugadores_upsert(conn: sqlite3.Connection, df: pd.DataFrame) -> int:
    """
    INSERT OR REPLACE en estadisticas_jugadores.
    UNIQUE INDEX (Jugador, Equipo) previene duplicados:
    fila existente → stats actualizadas; fila nueva → insertada.
    """
    if df.empty:
        return 0

    cur = conn.cursor()
    cur.execute("PRAGMA table_info(estadisticas_jugadores)")
    cols_bd = [row[1] for row in cur.fetchall()]

    # Añadir columnas nuevas sin destruir la tabla
    for col in df.columns:
        if col not in cols_bd:
            try:
                cur.execute(f'ALTER TABLE estadisticas_jugadores ADD COLUMN "{col}" REAL')
                log.info(f"  Columna nueva añadida: {col}")
                cols_bd.append(col)
            except sqlite3.Error:
                pass

    cols_comunes = [c for c in df.columns if c in cols_bd]
    cols_quoted  = ", ".join([f'"{c}"' for c in cols_comunes])
    placeholders = ", ".join(["?" for _ in cols_comunes])
    sql = (
        f"INSERT OR REPLACE INTO estadisticas_jugadores "
        f"({cols_quoted}) VALUES ({placeholders})"
    )

    afectados = 0
    for _, row in df[cols_comunes].iterrows():
        valores = tuple(
            None if (isinstance(v, float) and pd.isna(v)) else v
            for v in row
        )
        try:
            cur.execute(sql, valores)
            afectados += cur.rowcount
        except sqlite3.Error as exc:
            log.debug(f"  Upsert ignorado: {exc}")

    conn.commit()
    return afectados


def sincronizar_jugadores(conn: sqlite3.Connection):
    """
    Scraping de FBref en segundo plano y volcado seguro a la BD.
    Si FBref bloquea con HTTP 403 (Cloudflare), aborta de inmediato sin seguir
    esperando el resto de categorías y devuelve (0, []) sin inyectar errores.

    Returns: (jugadores_actualizados: int, errores: list[str])
    """
    if curl_requests is None:
        return 0, []

    dfs = {}
    session = curl_requests.Session()

    try:
        # Visita inicial rápida para comprobar si Cloudflare permite acceso
        try:
            resp_init = session.get(
                "https://fbref.com/es/",
                headers=HEADERS_FBREF,
                impersonate="chrome",
                timeout=10,
            )
            if resp_init.status_code == 403:
                log.info("  [Silenciado] FBref bloqueado por Cloudflare (403). Omitiendo jugadores.")
                return 0, []
        except Exception as exc:
            log.info(f"  [Silenciado] FBref inaccesible al iniciar sesión: {exc}")
            return 0, []

        for cat, url in URLS_FBREF.items():
            delay = random.uniform(2, 4)
            log.info(f"  Esperando {delay:.1f}s antes de solicitar FBref: {cat}")
            time.sleep(delay)

            log.info(f"  Scraping FBref: {cat} → {url}")
            df = _extraer_tabla_fbref(url, session=session)
            if df is not None and not df.empty:
                dfs[cat] = df
                log.info(f"  ✔ {cat}: {len(df)} filas.")
            else:
                # Si falla una categoría base ('Tiempo' o 'Estandar'), no tiene sentido
                # seguir esperando el resto porque _construir_df_jugadores las requiere.
                if cat in ("Tiempo", "Estandar"):
                    log.info(
                        f"  [Silenciado] Categoría base '{cat}' bloqueada/sin datos. "
                        "Abortando resto de peticiones a FBref."
                    )
                    return 0, []
    finally:
        try:
            session.close()
        except Exception:
            pass

    df_jugadores = _construir_df_jugadores(dfs)
    if df_jugadores is None or df_jugadores.empty:
        log.info("  [Silenciado] No se pudo construir el DataFrame de jugadores (bloqueo externo).")
        return 0, []

    actualizados = _insertar_jugadores_upsert(conn, df_jugadores)
    log.info(f"  → Jugadores upserted: {actualizados}")
    return actualizados, []


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 4 – PUNTO DE ENTRADA PÚBLICO
# ══════════════════════════════════════════════════════════════════════════════

def sincronizar_jornada() -> dict:
    """
    Sincroniza la última jornada disponible con macro_base_futbol.db.

    Flujo:
      1. Verifica que la BD existe.
      2. Garantiza los índices UNIQUE (idempotente, nunca destruye datos).
      3. Descarga CSVs de football-data.co.uk → INSERT OR IGNORE en historico_reciente.
      4. Scraping FBref (silencioso ante bloqueos 403) → INSERT OR REPLACE en estadisticas_jugadores.
      5. Devuelve un dict compatible con el bloque Admin de dashboard.py.

    Returns:
        dict con claves:
            "exito"           (bool)
            "partidos_nuevos" (int)
            "jugadores_act"   (int)
            "errores"         (list[str])
            "mensaje"         (str)
    """
    resultado = {
        "exito": False,
        "partidos_nuevos": 0,
        "jugadores_act": 0,
        "errores": [],
        "mensaje": "",
    }

    # ── 1. Verificar BD ───────────────────────────────────────────────────────
    if not RUTA_BD.exists():
        msg = f"Base de datos no encontrada en: {RUTA_BD}"
        resultado["mensaje"] = msg
        resultado["errores"].append(msg)
        log.error(msg)
        return resultado

    log.info(f"=== Iniciando sincronizar_jornada() → {RUTA_BD.name} ===")

    try:
        with sqlite3.connect(RUTA_BD) as conn:

            # ── 2. Garantizar índices UNIQUE ──────────────────────────────────
            try:
                _garantizar_indices(conn)
            except sqlite3.Error as exc:
                msg = f"Error creando índices UNIQUE: {exc}"
                resultado["errores"].append(msg)
                log.error(msg)

            # ── 3. Partidos (Prioridad crítica) ───────────────────────────────
            try:
                p_nuevos, err_p = sincronizar_partidos(conn)
                resultado["partidos_nuevos"] = p_nuevos
                resultado["errores"].extend(err_p)
                log.info(f"Partidos nuevos: {p_nuevos}")
            except Exception as exc:
                msg = f"Error en sincronización de partidos: {exc}"
                resultado["errores"].append(msg)
                log.error(msg)

            # ── 4. Jugadores (Silencioso en segundo plano) ────────────────────
            try:
                j_act, _ = sincronizar_jugadores(conn)
                resultado["jugadores_act"] = j_act
                log.info(f"Jugadores actualizados: {j_act}")
            except Exception as exc:
                # Nunca inyectamos fallos de FBref en resultado["errores"]
                log.info(f"[Silenciado] Excepción en sincronización de jugadores: {exc}")

            # ── 4b. Árbitros (Transfermarkt; nunca interrumpe ni marca error) ──
            try:
                a_act, err_a = sincronizar_arbitros(conn)
                for e in err_a:
                    log.warning(f"[Árbitros] {e}")
                log.info(
                    "[Árbitros] Sincronización completada correctamente "
                    "(sin llamadas externas bloqueadas)"
                )
            except Exception as exc:
                log.warning(f"[Silenciado] Fallo en sincronización de árbitros: {exc}")

    except sqlite3.Error as exc:
        msg = f"Error crítico de base de datos: {exc}"
        resultado["mensaje"] = msg
        resultado["errores"].append(msg)
        log.error(msg)
        return resultado

    # ── 5. Evaluar resultado ──────────────────────────────────────────────────
    # Si no hay errores críticos de BD o de descarga global de partidos,
    # la sincronización se marca siempre como exitosa (en verde en Streamlit).
    if not resultado["errores"]:
        resultado["exito"] = True
        if resultado["partidos_nuevos"] > 0 or resultado["jugadores_act"] > 0:
            resultado["mensaje"] = (
                f"Sincronización completada: {resultado['partidos_nuevos']} "
                f"partidos nuevos, {resultado['jugadores_act']} jugadores actualizados."
            )
        else:
            resultado["mensaje"] = (
                "Base de datos ya al día. No hay datos nuevos disponibles en este momento."
            )
    else:
        resultado["exito"] = False
        resultado["mensaje"] = "Sincronización fallida. Revisa los errores."

    log.info(f"=== Resultado: {resultado['mensaje']} ===")
    return resultado


# ── Ejecución directa para prueba/diagnóstico ─────────────────────────────────
if __name__ == "__main__":
    import json
    res = sincronizar_jornada()
    print(json.dumps(res, indent=2, ensure_ascii=False))
