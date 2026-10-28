# scraper_arbitros.py
# ─────────────────────────────────────────────────────────────────────────────
# Módulo de extracción de colegiados oficiales y estadísticas disciplinarias
# para las 5 grandes ligas europeas desde Transfermarkt:
#   - LaLiga (España)           → SP1
#   - Premier League (Inglaterra) → E0
#   - Serie A (Italia)          → I1
#   - Bundesliga (Alemania)     → D1
#   - Ligue 1 (Francia)         → F1
#
# Almacena/actualiza los registros en la tabla `arbitros_stats` de
# `macro_base_futbol.db` manteniendo el esquema exacto que consume dashboard.py.
# ─────────────────────────────────────────────────────────────────────────────

import logging
import sqlite3
import time
import unicodedata
from pathlib import Path

import requests
from bs4 import BeautifulSoup

try:
    from curl_cffi import requests as curl_requests
except ImportError:
    curl_requests = None

RUTA_BD = Path(__file__).parent / "macro_base_futbol.db"
TIMEOUT_HTTP = 20

HEADERS_BROWSER = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;q=0.9,"
        "image/avif,image/webp,image/apng,*/*;q=0.8"
    ),
    "Accept-Encoding": "gzip, deflate, br",
    "Referer": "https://www.transfermarkt.es/",
    "Sec-Ch-Ua": '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"Windows"',
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
}


def _nueva_sesion():
    """Sesión con huella TLS de Chrome (curl_cffi) o requests como alternativa."""
    if curl_requests is not None:
        return curl_requests.Session()
    return requests.Session()


def _get_html(session, url: str):
    """GET con cabeceras de navegador. Devuelve el HTML solo si HTTP == 200; si no, None."""
    kwargs = {"headers": HEADERS_BROWSER, "timeout": TIMEOUT_HTTP}
    if curl_requests is not None and isinstance(session, curl_requests.Session):
        kwargs["impersonate"] = "chrome"
    resp = session.get(url, **kwargs)
    if resp.status_code != 200:
        log.warning(f"  HTTP {resp.status_code} en {url}; se conservan datos previos.")
        return None
    return resp.text

COMPETICIONES_ARBITROS = [
    {
        "div": "SP1",
        "liga": "LaLiga (España)",
        "comite": "CTA / RFEF (España)",
        "url": "https://www.transfermarkt.es/laliga/schiedsrichter/wettbewerb/ES1",
    },
    {
        "div": "E0",
        "liga": "Premier League (Inglaterra)",
        "comite": "PGMOL (Inglaterra)",
        "url": "https://www.transfermarkt.es/premier-league/schiedsrichter/wettbewerb/GB1",
    },
    {
        "div": "I1",
        "liga": "Serie A (Italia)",
        "comite": "AIA (Italia)",
        "url": "https://www.transfermarkt.es/serie-a/schiedsrichter/wettbewerb/IT1",
    },
    {
        "div": "D1",
        "liga": "Bundesliga (Alemania)",
        "comite": "DFB (Alemania)",
        "url": "https://www.transfermarkt.es/bundesliga/schiedsrichter/wettbewerb/L1",
    },
    {
        "div": "F1",
        "liga": "Ligue 1 (Francia)",
        "comite": "FFF / DTA (Francia)",
        "url": "https://www.transfermarkt.es/ligue-1/schiedsrichter/wettbewerb/FR1",
    },
]

log = logging.getLogger(__name__)
if not log.handlers:
    _h = logging.StreamHandler()
    _h.setFormatter(logging.Formatter("%(asctime)s [scraper_arbitros] %(levelname)s: %(message)s"))
    log.addHandler(_h)
log.setLevel(logging.INFO)


def _parse_int(texto: str) -> int:
    """Convierte texto de celda Transfermarkt ('-', '', '12') a entero seguro."""
    if not texto:
        return 0
    limpio = texto.strip().replace(".", "").replace(",", "")
    if limpio in ("-", ""):
        return 0
    try:
        return int(limpio)
    except ValueError:
        return 0


def clasificar_perfil_disciplinario(media_amarillas: float) -> str:
    """
    Clasifica automáticamente el perfil disciplinario del colegiado:
      - > 5.0       → "Alto"
      - 4.0 a 5.0   → "Medio"
      - < 4.0       → "Permisivo"
    """
    if media_amarillas > 5.0:
        return "Alto"
    elif media_amarillas >= 4.0:
        return "Medio"
    else:
        return "Permisivo"


def estimar_faltas_por_partido(media_amarillas: float, div_code: str) -> float:
    """
    Calcula una estimación coherente de faltas por partido cuando la web externa
    no desglosa faltas por colegiado, calibrada según la media de cada liga.
    """
    base_liga = {
        "SP1": 18.5,
        "E0": 15.8,
        "I1": 18.2,
        "D1": 16.5,
        "F1": 17.4,
    }.get(div_code, 17.5)
    return round(base_liga + (media_amarillas * 1.45), 2)


def extraer_arbitros_liga(comp: dict, session: requests.Session = None) -> list[dict]:
    """
    Descarga y parsea la tabla de árbitros de una competición desde Transfermarkt.
    Devuelve una lista de diccionarios listos para insertar en `arbitros_stats`.
    """
    own_session = session is None
    if own_session:
        session = _nueva_sesion()

    url = comp["url"]
    div_code = comp["div"]
    liga_nombre = comp["liga"]
    comite = comp["comite"]
    resultados = []

    try:
        html = _get_html(session, url)
        if html is None:
            return []
        soup = BeautifulSoup(html, "html.parser")

        tabla = soup.find("table", class_="items")
        if not tabla or not tabla.find("tbody"):
            log.warning(f"No se encontró tabla 'items' para {liga_nombre} ({url})")
            return []

        filas = tabla.find("tbody").find_all("tr", recursive=False)
        for tr in filas:
            celdas = tr.find_all("td", recursive=False)
            if len(celdas) < 6:
                continue

            # Nombre del árbitro dentro de <td class="hauptlink"><a>
            celda_nombre = tr.find("td", class_="hauptlink")
            tag_a = celda_nombre.find("a") if celda_nombre else celdas[0].find("a")
            nombre = tag_a.get_text(strip=True) if tag_a else celdas[0].get_text(strip=True)
            if not nombre:
                continue

            partidos = _parse_int(celdas[2].get_text(strip=True))
            if partidos <= 0:
                continue

            total_amarillas = _parse_int(celdas[3].get_text(strip=True))
            total_segundas_amarillas = _parse_int(celdas[4].get_text(strip=True))
            total_rojas_directas = _parse_int(celdas[5].get_text(strip=True))
            total_rojas = total_segundas_amarillas + total_rojas_directas

            media_amarillas = round(total_amarillas / partidos, 2)
            media_rojas = round(total_rojas / partidos, 2)
            perfil = clasificar_perfil_disciplinario(media_amarillas)
            media_faltas = estimar_faltas_por_partido(media_amarillas, div_code)

            resultados.append(
                {
                    "Árbitro": nombre,
                    "Div": div_code,
                    "Liga": liga_nombre,
                    "Comité": comite,
                    "Partidos": partidos,
                    "Perfil Disciplinario": perfil,
                    "Media Amarillas Est.": media_amarillas,
                    "Media Rojas": media_rojas,
                    "Media Faltas": media_faltas,
                }
            )

        log.info(f"  ✔ {liga_nombre} ({div_code}): {len(resultados)} árbitros extraídos.")
        return resultados

    except Exception as exc:
        log.warning(f"  Error extrayendo árbitros de {liga_nombre}: {exc}")
        return []
    finally:
        if own_session:
            session.close()


def _garantizar_esquema_arbitros(conn: sqlite3.Connection) -> None:
    """Crea la tabla `arbitros_stats` si no existe con el esquema esperado por dashboard.py."""
    cur = conn.cursor()
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
    conn.commit()


def _sincronizar_arbitros_web(conn: sqlite3.Connection = None) -> tuple[int, list[str]]:
    """
    [OPCIONAL] Extrae los árbitros desde Transfermarkt (puede devolver 403 por
    Cloudflare). Ya no se usa por defecto; ver `sincronizar_arbitros(usar_web=True)`.

    Returns:
        (total_arbitros_actualizados: int, errores: list[str])
    """
    cerrar_conn = False
    if conn is None:
        conn = sqlite3.connect(RUTA_BD)
        cerrar_conn = True

    errores = []
    total_actualizados = 0

    try:
        _garantizar_esquema_arbitros(conn)
        cur = conn.cursor()

        session = _nueva_sesion()
        todos_registros = []

        for idx, comp in enumerate(COMPETICIONES_ARBITROS):
            if idx > 0:
                time.sleep(1.0)
            registros_liga = extraer_arbitros_liga(comp, session=session)
            if registros_liga:
                todos_registros.extend(registros_liga)
            else:
                errores.append(f"Sin datos de árbitros para {comp['liga']}")

        session.close()

        if todos_registros:
            # Reemplazar los registros de las ligas sincronizadas exitosamente
            divs_exitosas = list({r["Div"] for r in todos_registros})
            placeholders_div = ", ".join(["?" for _ in divs_exitosas])
            cur.execute(
                f'DELETE FROM arbitros_stats WHERE "Div" IN ({placeholders_div})',
                divs_exitosas,
            )

            sql_upsert = """
                INSERT OR REPLACE INTO arbitros_stats (
                    "Árbitro", "Div", "Liga", "Comité", "Partidos",
                    "Perfil Disciplinario", "Media Amarillas Est.", "Media Rojas", "Media Faltas"
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            for reg in todos_registros:
                cur.execute(
                    sql_upsert,
                    (
                        reg["Árbitro"],
                        reg["Div"],
                        reg["Liga"],
                        reg["Comité"],
                        reg["Partidos"],
                        reg["Perfil Disciplinario"],
                        reg["Media Amarillas Est."],
                        reg["Media Rojas"],
                        reg["Media Faltas"],
                    ),
                )
                total_actualizados += 1

            conn.commit()
            log.info(f"Total árbitros guardados en arbitros_stats: {total_actualizados}")

        return total_actualizados, errores

    finally:
        if cerrar_conn:
            conn.close()


def _normalizar_nombre(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode()
    return " ".join(s.lower().replace(".", " ").split())


def _clave_nombre(nombre: str) -> tuple[str, str]:
    """('a madley') y ('andrew madley') -> ('a', 'madley')."""
    partes = _normalizar_nombre(nombre).split()
    if len(partes) < 2:
        return ("", partes[0] if partes else "")
    return partes[0][0], " ".join(partes[1:])


def _sincronizar_arbitros_local(conn: sqlite3.Connection) -> tuple[int, list[str]]:
    """
    Recalcula con SQL real las estadísticas de los árbitros de las ligas cuya tabla
    `historico_reciente` incluye la columna `Referee` (p. ej. Premier League).
    Las ligas sin esa columna no se tocan.
    """
    errores: list[str] = []
    total = 0

    cols = [r[1] for r in conn.execute("PRAGMA table_info(historico_reciente)")]
    if "Referee" not in cols or not cols:
        return 0, errores
    col_div = cols[0]  # columna Div (almacenada con BOM)

    cur = conn.cursor()
    for comp in COMPETICIONES_ARBITROS:
        div = comp["div"]
        filas = cur.execute(
            f'''
            SELECT Referee,
                   COUNT(*),
                   SUM(COALESCE(HY,0) + COALESCE(AY,0)),
                   SUM(COALESCE(HR,0) + COALESCE(AR,0)),
                   SUM(COALESCE(HF,0) + COALESCE(AF,0))
            FROM historico_reciente
            WHERE "{col_div}" = ? AND Referee IS NOT NULL AND TRIM(Referee) <> ''
              AND HY IS NOT NULL AND AY IS NOT NULL
            GROUP BY Referee
            ''',
            (div,),
        ).fetchall()
        if not filas:
            continue  # liga sin datos de Referee: se conservan los registros previos

        existentes: dict[tuple[str, str], list[str]] = {}
        for (nom,) in cur.execute('SELECT "Árbitro" FROM arbitros_stats WHERE "Div" = ?', (div,)):
            existentes.setdefault(_clave_nombre(nom), []).append(nom)

        # Agrupar variantes del mismo árbitro (misma clave inicial+apellido)
        agregados: dict[tuple[str, str], list] = {}
        nombres_csv: dict[tuple[str, str], str] = {}
        for ref, n, amar, roj, faltas in filas:
            k = _clave_nombre(ref)
            acc = agregados.setdefault(k, [0, 0, 0, 0])
            acc[0] += n
            acc[1] += amar or 0
            acc[2] += roj or 0
            acc[3] += faltas or 0
            nombres_csv.setdefault(k, ref.strip())

        for k, (n, amar, roj, faltas) in agregados.items():
            coincid = existentes.get(k, [])
            if len(coincid) > 1:
                continue  # ambiguo: no se arriesga a sobrescribir
            if not coincid and n < 3:
                continue  # árbitro nuevo con muestra irrelevante
            media_am = round(amar / n, 2)
            media_roj = round(roj / n, 2)
            media_fal = round(faltas / n, 2) if faltas else estimar_faltas_por_partido(media_am, div)
            perfil = clasificar_perfil_disciplinario(media_am)
            if coincid:
                cur.execute(
                    'UPDATE arbitros_stats SET "Partidos"=?, "Perfil Disciplinario"=?, '
                    '"Media Amarillas Est."=?, "Media Rojas"=?, "Media Faltas"=? '
                    'WHERE "Árbitro"=? AND "Div"=?',
                    (n, perfil, media_am, media_roj, media_fal, coincid[0], div),
                )
            else:
                cur.execute(
                    'INSERT OR REPLACE INTO arbitros_stats ("Árbitro","Div","Liga","Comité","Partidos",'
                    '"Perfil Disciplinario","Media Amarillas Est.","Media Rojas","Media Faltas") '
                    'VALUES (?,?,?,?,?,?,?,?,?)',
                    (nombres_csv[k], div, comp["liga"], comp["comite"], n, perfil,
                     media_am, media_roj, media_fal),
                )
            total += 1
        log.info(f"  {comp['liga']} ({div}): árbitros recalculados desde la BD local.")

    conn.commit()
    return total, errores


def sincronizar_arbitros(
    conn: sqlite3.Connection = None, usar_web: bool = False
) -> tuple[int, list[str]]:
    """
    Actualiza `arbitros_stats` sin llamadas externas: recalcula con SQL, desde
    `historico_reciente`, las ligas que traen la columna `Referee`. El resto de
    ligas conserva intactos sus registros. La vía web (Transfermarkt) es opcional
    (`usar_web=True`) y sus fallos quedan silenciados.

    Returns:
        (total_arbitros_actualizados: int, errores: list[str])
    """
    cerrar_conn = False
    if conn is None:
        conn = sqlite3.connect(RUTA_BD)
        cerrar_conn = True

    try:
        _garantizar_esquema_arbitros(conn)
        total, errores = _sincronizar_arbitros_local(conn)
        if usar_web:
            try:
                t_web, _ = _sincronizar_arbitros_web(conn)
                total += t_web
            except Exception as exc:
                log.info(f"[Silenciado] Vía web de árbitros omitida: {exc}")
        return total, errores
    finally:
        if cerrar_conn:
            conn.close()


if __name__ == "__main__":
    total, errs = sincronizar_arbitros()
    print(f"\nResumen sincronización árbitros: {total} colegiados actualizados.")
    if errs:
        print("Avisos:", errs)

    with sqlite3.connect(RUTA_BD) as _conn:
        resumen = _conn.execute(
            'SELECT "Liga", "Div", COUNT(*), MIN("Partidos"), MAX("Partidos"), '
            'ROUND(AVG("Media Amarillas Est."), 2) '
            'FROM arbitros_stats GROUP BY "Liga", "Div" ORDER BY "Liga"'
        ).fetchall()
        print("\nDistribución en BD (arbitros_stats):")
        for fila in resumen:
            print(
                f"  - {fila[0]} ({fila[1]}): {fila[2]} árbitros | "
                f"Partidos [{fila[3]}..{fila[4]}] | Media Amarillas: {fila[5]}"
            )
