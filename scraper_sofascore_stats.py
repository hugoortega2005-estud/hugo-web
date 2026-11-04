"""
scraper_sofascore_stats.py — Extractor de estadísticas de jugadores por torneo (SofaScore JSON).

Consume los endpoints JSON del backend de SofaScore (sin BeautifulSoup) y construye un
DataFrame estandarizado. NO escribe en ninguna base de datos SQLite.

Política de acceso: las peticiones se identifican de forma transparente y, si el servidor
responde 403/429 (bloqueo anti-bot o rate limit), el script se detiene y lo informa. NO se
implementan técnicas de evasión (cloudscraper / undetected_chromedriver / suplantación de
navegador), ya que eludirían una medida de protección del servicio. Revisa los términos de
uso de SofaScore antes de usar estos datos; para datos con licencia usa una API oficial.
"""
from __future__ import annotations

import sys
import time

import pandas as pd
import requests

BASE_URL = "https://api.sofascore.com/api/v1"
TIMEOUT = 15
PAUSA_ENTRE_PAGINAS = 1.0  # cortesía con el servidor

HEADERS = {
    "User-Agent": "BigData_Futbol-research-scraper/1.0 (+personal analytics)",
    "Accept": "application/json",
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
}

_MAPA_METRICAS = {
    "minutesPlayed": "Minutos Jugados",
    "appearances": "Partidos Jugados",
    "matchesStarted": "Titularidades",
    "goals": "Goles",
    "assists": "Asistencias",
    "totalShots": "Tiros Totales",
    "shotsOnTarget": "Tiros a Puerta",
    "fouls": "Faltas Cometidas",
    "tackles": "Entradas (Tackles)",
    "yellowCards": "Tarjetas Amarillas",
    "redCards": "Tarjetas Rojas",
}
CAMPOS = list(_MAPA_METRICAS)
COLUMNAS_SALIDA = ["Jugador", "Equipo", "Posición"] + list(_MAPA_METRICAS.values())


class AccesoBloqueado(RuntimeError):
    """El servidor rechazó la petición (403/429). No se intenta eludir."""


def _get_json(session: requests.Session, url: str, params: dict | None = None) -> dict:
    resp = session.get(url, params=params, timeout=TIMEOUT)
    if resp.status_code in (403, 429):
        raise AccesoBloqueado(
            f"HTTP {resp.status_code} en {url}. El servidor bloquea el acceso automatizado; "
            "no se aplican técnicas de evasión. Usa una fuente/API con licencia."
        )
    resp.raise_for_status()
    return resp.json()


def obtener_season_id(session: requests.Session, tournament_id: int) -> int:
    data = _get_json(session, f"{BASE_URL}/unique-tournament/{tournament_id}/seasons")
    seasons = data.get("seasons") or []
    if not seasons:
        raise ValueError(f"Sin temporadas para el torneo {tournament_id}.")
    return int(seasons[0]["id"])  # la primera es la temporada más reciente


def extraer_estadisticas_torneo(tournament_id: int = 8, season_id: int | None = None,
                                limite_paginas: int = 40) -> pd.DataFrame:
    """Devuelve un DataFrame con las columnas de COLUMNAS_SALIDA para el torneo indicado."""
    filas: list[dict] = []
    with requests.Session() as session:
        session.headers.update(HEADERS)
        if season_id is None:
            season_id = obtener_season_id(session, tournament_id)
        url = f"{BASE_URL}/unique-tournament/{tournament_id}/season/{season_id}/statistics"

        for pagina in range(limite_paginas):
            params = {
                "limit": 100, "offset": pagina * 100, "order": "-minutesPlayed",
                "accumulation": "total", "fields": ",".join(CAMPOS),
            }
            data = _get_json(session, url, params)
            resultados = data.get("results") or []
            for r in resultados:
                jugador = r.get("player") or {}
                equipo = r.get("team") or {}
                fila = {
                    "Jugador": jugador.get("name"),
                    "Equipo": equipo.get("name"),
                    "Posición": jugador.get("position"),
                }
                for campo, col in _MAPA_METRICAS.items():
                    fila[col] = r.get(campo, 0) or 0
                filas.append(fila)
            if not resultados or pagina + 1 >= int(data.get("pages", 1) or 1):
                break
            time.sleep(PAUSA_ENTRE_PAGINAS)

    df = pd.DataFrame(filas, columns=COLUMNAS_SALIDA)
    cols_num = COLUMNAS_SALIDA[3:]
    df[cols_num] = df[cols_num].apply(pd.to_numeric, errors="coerce").fillna(0).round().astype(int)
    return df


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    try:
        df_test = extraer_estadisticas_torneo(tournament_id=8)  # LaLiga
    except AccesoBloqueado as exc:
        print(f"[BLOQUEADO] {exc}")
        sys.exit(2)
    except Exception as exc:
        print(f"[ERROR] {type(exc).__name__}: {exc}")
        sys.exit(1)

    print(f"Jugadores extraídos: {len(df_test)}")
    print(df_test.head(10).to_string(index=False))
    df_test.to_csv("test_sofascore.csv", index=False, encoding="utf-8-sig")
    print("Exportado a test_sofascore.csv")

