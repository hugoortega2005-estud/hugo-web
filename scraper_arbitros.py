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
import re
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
        "comite": "LFP / FFF (Francia)",
        "url": "https://www.transfermarkt.es/ligue-1/schiedsrichter/wettbewerb/FR1",
    },
    {
        "div": "UCL",
        "liga": "Champions League (UEFA)",
        "comite": "Comité Arbitral UEFA",
        "url": "https://www.transfermarkt.es/uefa-champions-league/schiedsrichter/wettbewerb/CL/saison_id/2026",
    },
    {
        "div": "UEL",
        "liga": "Europa League (UEFA)",
        "comite": "Comité Arbitral UEFA",
        "url": "https://www.transfermarkt.es/europa-league/schiedsrichter/wettbewerb/EL/saison_id/2026",
    },
    {
        "div": "UECL",
        "liga": "Conference League (UEFA)",
        "comite": "Comité Arbitral UEFA",
        "url": "https://www.transfermarkt.es/uefa-conference-league/schiedsrichter/wettbewerb/UCOL/saison_id/2026",
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


def extraer_arbitros_liga(*args, **kwargs) -> list[dict]:
    """
    Descarga y parsea la tabla de árbitros de una competición desde Transfermarkt.
    Soporta firmas polimórficas:
      - extraer_arbitros_liga(session, comp_info)
      - extraer_arbitros_liga(comp_info, session=session, html_content=html_content)
    Devuelve una lista de diccionarios con el esquema canónico y relacional de arbitros_stats.
    """
    session = kwargs.get("session")
    comp_info = kwargs.get("comp_info") or kwargs.get("comp")
    html_content = kwargs.get("html_content")

    for arg in args:
        if isinstance(arg, dict) and comp_info is None:
            comp_info = arg
        elif (
            (hasattr(arg, "get") or hasattr(arg, "request"))
            and session is None
            and not isinstance(arg, dict)
        ):
            session = arg
        elif isinstance(arg, str) and html_content is None:
            html_content = arg

    if comp_info is None:
        comp_info = {}

    own_session = session is None and html_content is None
    if own_session:
        session = _nueva_sesion()

    div = comp_info.get("div", "")
    liga_nombre = comp_info.get("liga", "")
    comite = comp_info.get("comite", "Comité Arbitral")
    url = comp_info.get("url", "")

    try:
        html = html_content if html_content is not None else _get_html(session, url)
        if not html:
            # Reintento defensivo sin sufijo de temporada si venía con él
            if "/saison_id/" in url:
                url_alt = url.split("/saison_id/")[0]
                html = _get_html(session, url_alt)

        if not html:
            log.warning(f"[{div}] No se pudo obtener respuesta HTTP válida de {url}")
            return []

        soup = BeautifulSoup(html, "html.parser")

        # Selector en cascada: 1) table.items, 2) .responsive-table table, 3) tabla con enlaces a perfiles
        tabla = soup.select_one("table.items") or soup.select_one("div.responsive-table table")

        if not tabla:
            # Búsqueda heurística entre todas las tablas del documento
            for t in soup.find_all("table"):
                if t.find("a", href=lambda href: href and "/profil/schiedsrichter/" in href):
                    tabla = t
                    break

        # Reintento defensivo adicional para torneos UEFA/copas: endpoint 'pokalwettbewerb' si venía con 'wettbewerb'
        if not tabla and "/wettbewerb/" in url:
            url_pokal = url.replace("/wettbewerb/", "/pokalwettbewerb/")
            html_p = _get_html(session, url_pokal)
            if not html_p and "/saison_id/" in url_pokal:
                url_p_alt = url_pokal.split("/saison_id/")[0]
                html_p = _get_html(session, url_p_alt)
            if html_p:
                soup_p = BeautifulSoup(html_p, "html.parser")
                tabla = soup_p.select_one("table.items") or soup_p.select_one("div.responsive-table table")
                if not tabla:
                    for t in soup_p.find_all("table"):
                        if t.find("a", href=lambda href: href and "/profil/schiedsrichter/" in href):
                            tabla = t
                            break

        if not tabla:
            log.warning(f"[{div}] Estructura de tabla no localizada en {url}.")
            return []

        # Mapeo dinámico de columnas por cabecera si thead está presente
        idx_col_partidos = -1
        idx_col_amarillas = -1
        idx_col_segundas = -1
        idx_col_rojas = -1
        thead = tabla.find("thead")
        if thead:
            for i, th in enumerate(thead.find_all("th")):
                th_str = str(th).lower()
                if "einsaetze" in th_str or "alineaciones" in th_str or "spiele" in th_str:
                    idx_col_partidos = i
                elif "gelbrote" in th_str or "gelb-rote" in th_str or "segundas" in th_str:
                    idx_col_segundas = i
                elif "gelbe" in th_str or "amarillas" in th_str:
                    idx_col_amarillas = i
                elif "rote" in th_str or "rojas" in th_str:
                    idx_col_rojas = i

        tbody = tabla.find("tbody") or tabla
        filas = tbody.find_all("tr", recursive=False)
        if not filas:
            filas = tbody.find_all("tr")

        resultados = []
        for fila in filas:
            # Localizar celda/enlace del colegiado
            link_ref = fila.find(
                "a",
                href=lambda href: href
                and ("/profil/schiedsrichter/" in href or "/schiedsrichter/" in href),
            )
            if not link_ref:
                continue

            nombre_ref = link_ref.get_text(strip=True)
            if not nombre_ref or len(nombre_ref) < 3:
                continue

            # Extraer celdas numéricas de la fila
            tds = fila.find_all("td", recursive=False)
            if not tds:
                tds = fila.find_all("td")

            textos_td = [td.get_text(strip=True).replace("-", "0") for td in tds]

            # Buscar valores numéricos de partidos y tarjetas
            numeros = []
            for txt in textos_td:
                limpio = re.sub(r"[^\d]", "", txt)
                if limpio.isdigit():
                    numeros.append(int(limpio))

            # En la estructura Transfermarkt estándar:
            # columnas habituales tras el nombre: [Edad/Nacionalidad] -> Partidos -> Amarillas -> Dobles Amarillas -> Rojas -> Penaltis
            partidos = 0
            amarillas = 0
            rojas = 0
            faltas = 0.0  # Transfermarkt no audita faltas de equipo; se inicializa a 0.0 para compatibilidad con la BD

            if idx_col_partidos != -1 and len(tds) > max(idx_col_partidos, idx_col_amarillas):
                def _num_td(idx):
                    if 0 <= idx < len(tds):
                        clean = re.sub(r"[^\d]", "", tds[idx].get_text(strip=True).replace("-", "0"))
                        return int(clean) if clean.isdigit() else 0
                    return 0

                partidos = _num_td(idx_col_partidos)
                amarillas = _num_td(idx_col_amarillas)
                rojas = _num_td(idx_col_segundas) + _num_td(idx_col_rojas)
            else:
                if len(numeros) >= 1:
                    partidos = numeros[0]
                if len(numeros) >= 2:
                    amarillas = numeros[1]
                if len(numeros) >= 4:
                    # Suma de segundas amarillas + rojas directas
                    rojas = numeros[2] + numeros[3]
                elif len(numeros) >= 3:
                    rojas = numeros[2]

            if partidos > 0:
                media_amarillas = round(amarillas / partidos, 2)
                media_rojas = round(rojas / partidos, 2)

                perfil = "Medio"
                if media_amarillas >= 5.0:
                    perfil = "Alto"
                elif media_amarillas <= 3.8:
                    perfil = "Permisivo"

                resultados.append(
                    {
                        "referee": nombre_ref,
                        "comp": div,
                        "partidos": partidos,
                        "media_amarillas": media_amarillas,
                        "media_rojas": media_rojas,
                        "media_faltas": faltas,
                        "perfil": perfil,
                        # Claves para compatibilidad con el esquema relacional de SQLite y dashboard.py
                        "Árbitro": nombre_ref,
                        "Div": div,
                        "Liga": liga_nombre,
                        "Comité": comite,
                        "Partidos": partidos,
                        "Perfil Disciplinario": perfil,
                        "Media Amarillas Est.": media_amarillas,
                        "Media Rojas": media_rojas,
                        "Media Faltas": faltas,
                    }
                )

        log.info(f"[{div}] {len(resultados)} colegiados extraídos satisfactoriamente.")
        return resultados

    except Exception as exc:
        log.warning(f"[{div}] Error extrayendo árbitros de {liga_nombre}: {exc}")
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

        for idx, comp_info in enumerate(COMPETICIONES_ARBITROS):
            if idx > 0:
                time.sleep(1.0)
            div = comp_info.get("div")
            liga_nombre = comp_info.get("liga")
            url_tm = comp_info.get("url")

            print(f"\n[+] Procesando: {liga_nombre} ({div}) -> {url_tm}")
            html_content = _get_html(session, url_tm)

            if html_content is None:
                print(f"    [!] Fallo HTTP/Bloqueo en {div}. Intentando recálculo fallback desde BD local...")
                errores.append(f"Sin datos de árbitros para {liga_nombre}")
            else:
                print(f"    [OK] HTML recibido ({len(html_content)} bytes). Parseando tabla Transfermarkt...")
                registros_liga = extraer_arbitros_liga(
                    comp_info, session=session, html_content=html_content
                )
                if registros_liga:
                    todos_registros.extend(registros_liga)
                else:
                    errores.append(f"Sin datos de árbitros para {liga_nombre}")

        session.close()

        if todos_registros:
            # Reemplazar los registros de las ligas sincronizadas exitosamente
            divs_exitosas = list({r.get("Div") or r.get("comp") for r in todos_registros})
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
                        reg.get("Árbitro") or reg.get("referee"),
                        reg.get("Div") or reg.get("comp"),
                        reg.get("Liga") or reg.get("comp"),
                        reg.get("Comité") or "Comité Arbitral",
                        reg.get("Partidos") if reg.get("Partidos") is not None else reg.get("partidos", 0),
                        reg.get("Perfil Disciplinario") or reg.get("perfil", "Medio"),
                        reg.get("Media Amarillas Est.")
                        if reg.get("Media Amarillas Est.") is not None
                        else reg.get("media_amarillas", 0.0),
                        reg.get("Media Rojas")
                        if reg.get("Media Rojas") is not None
                        else reg.get("media_rojas", 0.0),
                        reg.get("Media Faltas")
                        if reg.get("Media Faltas") is not None
                        else reg.get("media_faltas", 0.0),
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
    conn: sqlite3.Connection = None, usar_web: bool = True
) -> tuple[int, list[str]]:
    """
    Actualiza `arbitros_stats` ejecutando la extracción web (Transfermarkt) para
    todas las competiciones de COMPETICIONES_ARBITROS (incluyendo UCL, UEL, UECL)
    y aplicando recálculo fallback con SQL desde `historico_reciente` para las
    ligas con columna `Referee`.

    Returns:
        (total_arbitros_actualizados: int, errores: list[str])
    """
    cerrar_conn = False
    if conn is None:
        conn = sqlite3.connect(RUTA_BD)
        cerrar_conn = True

    try:
        _garantizar_esquema_arbitros(conn)
        total = 0
        errores: list[str] = []
        if usar_web:
            try:
                t_web, errs_web = _sincronizar_arbitros_web(conn)
                total += t_web
                errores.extend(errs_web)
            except Exception as exc:
                log.info(f"[Silenciado] Vía web de árbitros omitida: {exc}")
        t_local, errs_local = _sincronizar_arbitros_local(conn)
        total += t_local
        errores.extend(errs_local)
        return total, errores
    finally:
        if cerrar_conn:
            conn.close()


if __name__ == "__main__":
    total, errs = sincronizar_arbitros(usar_web=True)
    print(f"\nResumen sincronización árbitros: {total} colegiados actualizados.")
    if errs:
        print("Avisos:", errs)

    with sqlite3.connect(RUTA_BD) as _conn:
        resumen = _conn.execute(
            'SELECT "Liga", "Div", COUNT(*), MIN("Partidos"), MAX("Partidos"), '
            'ROUND(AVG("Media Amarillas Est."), 2) '
            'FROM arbitros_stats '
            "WHERE \"Div\" IN ('SP1', 'E0', 'I1', 'D1', 'F1', 'UCL', 'UEL', 'UECL') "
            'GROUP BY "Liga", "Div" ORDER BY "Liga"'
        ).fetchall()
        print("\nDistribución en BD (arbitros_stats):")
        divs_presentes = set()
        for fila in resumen:
            divs_presentes.add(fila[1])
            print(
                f"  - {fila[0]} ({fila[1]}): {fila[2]} árbitros | "
                f"Partidos [{fila[3]}..{fila[4]}] | Media Amarillas: {fila[5]}"
            )
        for div_uefa in ("UCL", "UEL", "UECL"):
            if div_uefa not in divs_presentes:
                print(f"  - [{div_uefa}]: 0 árbitros en arbitros_stats (sin datos extraídos)")
