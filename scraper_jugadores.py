import sqlite3
import sys
import pandas as pd
import logging
import time
import unicodedata
import re
from pathlib import Path
try:
    from playwright.sync_api import sync_playwright
except ImportError:  # Permite importar el módulo (p. ej. desde updater) sin Playwright instalado
    sync_playwright = None
from io import StringIO

# Configuración del log
logging.basicConfig(level=logging.INFO, format='%(asctime)s - [%(levelname)s] - %(message)s')

# URLs de FBref para extraer todas las métricas (Añadida la de Tiempo de Juego como Lista Maestra)
URLS_CATEGORIAS = {
    "Tiempo": "https://fbref.com/es/comps/12/playingtime/Estadisticas-de-La-Liga",
    "Estandar": "https://fbref.com/es/comps/12/stats/Estadisticas-de-La-Liga",
    "Pases": "https://fbref.com/es/comps/12/passing/Estadisticas-de-La-Liga",
    "Defensa": "https://fbref.com/es/comps/12/defense/Estadisticas-de-La-Liga",
    "Varios": "https://fbref.com/es/comps/12/misc/Estadisticas-de-La-Liga"
}

class ScraperJugadoresDefinitivo:
    def __init__(self):
        self.ruta_bd = Path(__file__).parent / "macro_base_futbol.db"
        self.ruta_sesion = Path(__file__).parent / "sesion_chrome_fbref"

    def _crear_id(self, serie_jug, serie_eq):
        """Crea un ID único uniendo nombre y equipo sin caracteres especiales."""
        def norm(s):
            s = str(s).lower()
            s = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('utf-8')
            return re.sub(r'[^a-z0-9]', '', s)
        return serie_jug.apply(norm) + "_" + serie_eq.apply(norm)

    def _extraer_df(self, page, url):
        """Navega a la URL y extrae la tabla principal de estadísticas."""
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            page.evaluate("window.scrollBy(0, 1000)")
            page.wait_for_timeout(3500)
            
            # Limpiar comentarios HTML que ocultan tablas en FBref
            html = page.content().replace('<!--', '').replace('-->', '')
            tablas = pd.read_html(StringIO(html))
            
            for df in tablas:
                # Aplanar MultiIndex si existe
                if isinstance(df.columns, pd.MultiIndex):
                    df.columns = ['_'.join(str(c).strip() for c in col if str(c).strip()) for col in df.columns]
                
                # Validar que es la tabla de jugadores
                cols_lower = [str(c).lower() for c in df.columns]
                if len(df) > 50 and any('jugador' in c or 'player' in c for c in cols_lower):
                    return df
            return None
        except Exception as e:
            logging.error(f"Error procesando {url}: {e}")
            return None

    def _buscar_col(self, df, requisitos):
        """Busca una columna que contenga todas las palabras clave requeridas."""
        for col in df.columns:
            col_str = str(col).lower()
            if all(req in col_str for req in requisitos):
                return col
        return None

    def _parsear(self, df, mapa):
        """Filtra y renombra las columnas de la tabla según el mapa proporcionado."""
        if df is None or df.empty:
            return None
            
        c_jug = self._buscar_col(df, ['player']) or self._buscar_col(df, ['jugador'])
        c_eq = self._buscar_col(df, ['squad']) or self._buscar_col(df, ['equipo'])
        
        if not c_jug or not c_eq:
            return None

        res = pd.DataFrame()
        res['Jugador'] = df[c_jug].astype(str).str.replace('\xa0', ' ', regex=False).str.strip()
        
        # Limpieza crucial: Eliminar prefijos de país que mete FBref (ej. "es Atlético Madrid" -> "Atlético Madrid")
        res['Equipo'] = df[c_eq].astype(str).str.replace('\xa0', ' ', regex=False).str.strip()
        res['Equipo'] = res['Equipo'].str.replace(r'^[a-z]{2}\s', '', regex=True)
        
        # Eliminar filas de cabecera repetidas
        res = res[~res['Jugador'].isin(['Jugador', 'Player', 'nan', '', 'Jugador_Jugador'])]

        res['ID'] = self._crear_id(res['Jugador'], res['Equipo'])

        for alias, requisitos in mapa.items():
            c_match = self._buscar_col(df, requisitos)
            if c_match is not None:
                res[alias] = df[c_match]

        return res.drop_duplicates(subset=['ID'])

    def extraer(self):
        dfs = {}
        with sync_playwright() as p:
            ctx = p.chromium.launch_persistent_context(
                user_data_dir=str(self.ruta_sesion),
                headless=False,
                channel="chrome",
                args=["--disable-blink-features=AutomationControlled"]
            )
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            
            for cat, url in URLS_CATEGORIAS.items():
                logging.info(f"Descargando datos de {cat}...")
                df = self._extraer_df(page, url)
                if df is not None:
                    dfs[cat] = df
                    logging.info(f"✔ {cat} extraída: {len(df)} filas.")
                time.sleep(3)
            ctx.close()

        if "Tiempo" not in dfs or "Estandar" not in dfs:
            logging.error("Fallo crítico: No se obtuvieron las tablas maestras.")
            return

        # 1. TABLA MAESTRA (Tiempo de juego) -> Garantiza a los de 0 minutos
        mapa_tiempo = {'Pos': ['pos']}
        res_df = self._parsear(dfs["Tiempo"], mapa_tiempo)

        # 2. ESTÁNDAR
        mapa_est = {
            'Min': ['min'], '90s': ['90s'], 'Gls': ['gls'], 'Ast': ['ast'], 'G+A': ['g+a'],
            'xG': ['xg'], 'xAG': ['xag'], 'CrdY': ['crdy'], 'CrdR': ['crdr']
        }
        d_est = self._parsear(dfs["Estandar"], mapa_est)
        if d_est is not None:
            res_df = pd.merge(res_df, d_est.drop(columns=['Jugador', 'Equipo', 'Pos'], errors='ignore'), on='ID', how='left')

        # 3. PASES
        if "Pases" in dfs:
            mapa_pas = {
                'Pases_Completados': ['total', 'cmp'], 'Pases_Intentados': ['total', 'att'],
                'Pases_Progresivos': ['prgp'], 'Pases_Clave': ['kp']
            }
            d_pas = self._parsear(dfs["Pases"], mapa_pas)
            if d_pas is not None:
                res_df = pd.merge(res_df, d_pas.drop(columns=['Jugador', 'Equipo', 'Pos'], errors='ignore'), on='ID', how='left')

        # 4. DEFENSA
        if "Defensa" in dfs:
            mapa_def = {
                'Entradas': ['tackles', 'tkl'], 'Entradas_Ganadas': ['tackles', 'tklw'],
                'Intercepciones': ['int']
            }
            d_def = self._parsear(dfs["Defensa"], mapa_def)
            if d_def is not None:
                res_df = pd.merge(res_df, d_def.drop(columns=['Jugador', 'Equipo', 'Pos'], errors='ignore'), on='ID', how='left')

        # 5. VARIOS
        if "Varios" in dfs:
            mapa_var = {
                'Faltas_Cometidas': ['fls'], 'Faltas_Recibidas': ['fld'], 'Recuperaciones': ['recov']
            }
            d_var = self._parsear(dfs["Varios"], mapa_var)
            if d_var is not None:
                res_df = pd.merge(res_df, d_var.drop(columns=['Jugador', 'Equipo', 'Pos'], errors='ignore'), on='ID', how='left')

        # LIMPIEZA FINAL Y VOLCADO
        res_df.drop(columns=['ID'], inplace=True)
        cols_no_num = ['Jugador', 'Equipo', 'Pos']
        
        # Rellenar con 0 a los que no han jugado (suplentes/inéditos)
        for c in res_df.columns:
            if c not in cols_no_num:
                res_df[c] = pd.to_numeric(res_df[c], errors='coerce').fillna(0)

        # Ordenar columnas para el dashboard
        orden = [
            'Jugador', 'Equipo', 'Pos', 'Min', '90s', 'Gls', 'Ast', 'G+A', 'xG', 'xAG',
            'Pases_Intentados', 'Pases_Completados', 'Pases_Clave', 'Pases_Progresivos',
            'Entradas', 'Entradas_Ganadas', 'Intercepciones',
            'Faltas_Cometidas', 'Faltas_Recibidas', 'Recuperaciones', 'CrdY', 'CrdR'
        ]
        res_df = res_df[[c for c in orden if c in res_df.columns]]

        # Guardar en base de datos
        with sqlite3.connect(self.ruta_bd) as conn:
            conn.execute("DROP TABLE IF EXISTS estadisticas_jugadores")
            res_df.to_sql('estadisticas_jugadores', conn, if_exists='replace', index=False)
            
        logging.info(f"\n¡ÉXITO TOTAL! {len(res_df)} jugadores (plantillas completas) guardados con todas sus métricas.")

# ══════════════════════════════════════════════════════════════════════════════
# FBREF – INGESTA ROBUSTA (headless, recursos bloqueados, hilo aislado, UPSERT)
# ══════════════════════════════════════════════════════════════════════════════
import threading

URL_FBREF_STANDARD_LALIGA = "https://fbref.com/es/comps/12/stats/Estadisticas-de-La-Liga"
UA_CHROME_WINDOWS = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)
_JS_EXTRAER_TABLA = """
() => {
  let t = document.querySelector('table#stats_standard');
  if (!t) {  // La tabla puede venir envuelta en un comentario HTML
    const w = document.createTreeWalker(document, NodeFilter.SHOW_COMMENT);
    let n;
    while ((n = w.nextNode())) {
      if (n.nodeValue.includes('id="stats_standard"')) {
        const d = document.createElement('div');
        d.innerHTML = n.nodeValue;
        t = d.querySelector('table#stats_standard');
        break;
      }
    }
  }
  if (!t) return null;
  // Se apunta al <tbody>: se ignoran las cabeceras dobles (over_header) y las repetidas (thead)
  return Array.from(t.querySelectorAll('tbody tr'))
    .filter(tr => !tr.classList.contains('thead') && !tr.classList.contains('over_header'))
    .map(tr => {
      const o = {};
      tr.querySelectorAll('th,td').forEach(c => {
        const k = c.getAttribute('data-stat');
        if (k) o[k] = c.textContent.trim();
      });
      return o;
    });
}
"""


def _entero(valor, defecto: int = 0) -> int:
    try:
        return int(float(str(valor).replace(",", "").strip()))
    except (ValueError, TypeError):
        return defecto


def _flotante(valor, defecto: float = 0.0) -> float:
    try:
        return float(str(valor).replace(",", "").strip())
    except (ValueError, TypeError):
        return defecto


RUTA_SESION_FBREF = Path(__file__).parent / "sesion_chrome_fbref"
ESPERA_VERIFICACION_S = 60  # tiempo máximo para que la verificación de seguridad la complete una persona


def _hay_desafio(page) -> bool:
    titulo = (page.title() or "").lower()
    return any(k in titulo for k in ("just a moment", "attention required", "un momento", "verificaci"))


def _extraer_filas_fbref(url: str, resultado: dict) -> None:
    """
    Se ejecuta en un hilo propio (sin event loop previo). Deja las filas en `resultado`.
    Usa Chrome real con perfil persistente (cookies reutilizables) y NO bloquea recursos.
    Si FBref muestra la verificación de seguridad, la ventana visible permite que la complete
    una persona; no se aplican técnicas de enmascaramiento ni de resolución automática.
    """
    try:
        if sync_playwright is None:
            raise RuntimeError("Playwright no está instalado (pip install playwright).")
        with sync_playwright() as p:
            ctx = p.chromium.launch_persistent_context(
                user_data_dir=str(RUTA_SESION_FBREF),
                channel="chrome",
                headless=False,
            )
            try:
                page = ctx.pages[0] if ctx.pages else ctx.new_page()
                page.goto(url, wait_until="domcontentloaded", timeout=45000)
                if _hay_desafio(page):
                    logging.warning(
                        f"FBref muestra la verificación de seguridad; esperando hasta {ESPERA_VERIFICACION_S}s "
                        "a que se complete en la ventana visible."
                    )
                    fin = time.time() + ESPERA_VERIFICACION_S
                    while time.time() < fin and _hay_desafio(page):
                        page.wait_for_timeout(1000)
                try:
                    page.wait_for_selector("table#stats_standard", timeout=12000, state="attached")
                except Exception:
                    pass  # la tabla puede venir en un comentario HTML: se intenta igualmente
                filas = page.evaluate(_JS_EXTRAER_TABLA)
                if not filas:
                    logging.error("No se encontró la tabla stats_standard en FBref (¿verificación sin completar?).")
                    page.screenshot(path="error_fbref_stealth.png")
                    resultado["filas"] = []
                    return
                resultado["filas"] = filas
            finally:
                ctx.close()
    except Exception as exc:
        logging.error(f"Error en extracción FBref: {exc}")
        resultado["filas"] = []


def extraer_fbref_standard(url: str = URL_FBREF_STANDARD_LALIGA) -> list[dict]:
    """Wrapper síncrono: ejecuta Playwright en un hilo aislado (evita 'event loop is already running')."""
    resultado: dict = {}
    hilo = threading.Thread(target=_extraer_filas_fbref, args=(url, resultado), daemon=True)
    hilo.start()
    hilo.join(timeout=ESPERA_VERIFICACION_S + 90)
    return resultado.get("filas", []) or []


_SQL_UPSERT_FBREF = """
INSERT INTO estadisticas_jugadores
    (Jugador, Equipo, Pos, posicion_tactica, Min, "90s", partidos_jugados, titularidades,
     Gls, Ast, "G+A", xG, CrdY, CrdR, Tiros, Tiros_Puerta, Faltas_Cometidas, Faltas_Recibidas, fueras_juego)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
ON CONFLICT(Jugador, Equipo) DO UPDATE SET
    Pos=excluded.Pos, posicion_tactica=excluded.posicion_tactica, Min=excluded.Min,
    "90s"=excluded."90s", partidos_jugados=excluded.partidos_jugados,
    titularidades=excluded.titularidades, Gls=excluded.Gls, Ast=excluded.Ast,
    "G+A"=excluded."G+A", xG=excluded.xG, CrdY=excluded.CrdY, CrdR=excluded.CrdR,
    Tiros=excluded.Tiros, Tiros_Puerta=excluded.Tiros_Puerta,
    Faltas_Cometidas=excluded.Faltas_Cometidas, Faltas_Recibidas=excluded.Faltas_Recibidas,
    fueras_juego=excluded.fueras_juego
"""


def _clave_nombre(nombre: str) -> str:
    s = unicodedata.normalize("NFKD", str(nombre or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]", "", s.lower())


def ejecutar_ingesta_fbref(conn: sqlite3.Connection) -> int:
    """
    Ingesta FBref (LaLiga, estadísticas estándar) → estadisticas_jugadores.
    Devuelve el nº de jugadores escritos. Debe llamarse dentro de `with sqlite3.connect(...)`.
    La tabla estándar de FBref no incluye Sh/SoT/Fls: se conservan los valores previos
    (p. ej. del roster ESPN) de cada jugador y, si no existen, quedan a 0.
    """
    from updater import (
        _migrar_esquema_estadisticas_jugadores,
        _normalizar_posicion_tactica,
        _obtener_equipos_conocidos,
        resolver_equipo_canonico,
    )

    _migrar_esquema_estadisticas_jugadores(conn)
    filas = extraer_fbref_standard()
    if not filas:
        logging.warning("FBref no devolvió filas; no se purga ni se modifica la tabla.")
        return 0

    mapa_bd = _obtener_equipos_conocidos(conn)
    por_equipo: dict[str, list[tuple]] = {}
    for f in filas:
        jugador = str(f.get("player", "")).replace("\xa0", " ").strip()
        if not jugador or jugador.lower() in ("jugador", "player"):
            continue
        equipo_raw = re.sub(r"^[a-z]{2,3}\s+", "", str(f.get("team", "")).replace("\xa0", " ").strip())
        if not equipo_raw:
            continue
        equipo = resolver_equipo_canonico(equipo_raw, mapa_bd)
        primer_rol = re.split(r"[,/\s]+", str(f.get("position", "")).strip())[0].upper()
        pos_tact = {"GK": "POR", "DF": "DFC", "MF": "MC", "FW": "DC"}.get(primer_rol) \
            or _normalizar_posicion_tactica(primer_rol)
        minutos = _entero(f.get("minutes"))
        goles, asist = _entero(f.get("goals")), _entero(f.get("assists"))
        por_equipo.setdefault(equipo, []).append((
            jugador, equipo, primer_rol or "S/P", pos_tact,
            minutos, _flotante(f.get("minutes_90s"), round(minutos / 90.0, 1)),
            _entero(f.get("games")), _entero(f.get("games_starts")),
            goles, asist, goles + asist, _flotante(f.get("xg")),
            _entero(f.get("cards_yellow")), _entero(f.get("cards_red")),
        ))

    cursor = conn.cursor()
    total_jugadores = 0
    lote: dict[tuple, tuple] = {}
    for equipo, tuplas in por_equipo.items():
        try:
            previos = {
                _clave_nombre(r[0]): r[1:]
                for r in cursor.execute(
                    "SELECT Jugador, Tiros, Tiros_Puerta, Faltas_Cometidas, Faltas_Recibidas, fueras_juego "
                    "FROM estadisticas_jugadores WHERE equipo = ?", (equipo,)
                ).fetchall()
            }
            # PURGA DE TRASPASOS: instantánea real de la plantilla del club
            cursor.execute("DELETE FROM estadisticas_jugadores WHERE equipo = ?", (equipo,))
            for t in tuplas:
                extra = previos.get(_clave_nombre(t[0]), (0, 0, 0, 0, 0))
                extra = tuple(0 if v is None else v for v in extra)
                lote[(t[0], t[1])] = t + extra
        except sqlite3.Error as exc:
            logging.error(f"Error purgando/leyendo {equipo}: {exc}")
    if lote:
        cursor.executemany(_SQL_UPSERT_FBREF, list(lote.values()))
        total_jugadores = len(lote)
    conn.commit()
    logging.info(f"FBref: {total_jugadores} jugadores sincronizados en {len(por_equipo)} equipos.")
    return total_jugadores


if __name__ == "__main__":
    if "--fbref-stealth" in sys.argv:
        with sqlite3.connect(Path(__file__).parent / "macro_base_futbol.db") as _conn:
            print("total_jugadores:", ejecutar_ingesta_fbref(_conn))
    else:
        ScraperJugadoresDefinitivo().extraer()