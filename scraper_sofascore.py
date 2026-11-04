# scraper_sofascore.py
# ─────────────────────────────────────────────────────────────────────────────
# Módulo de extracción en tiempo real de las 5 grandes ligas europeas
# utilizando los endpoints públicos de ESPN Scoreboard JSON API (HTTP 200
# directo, sin bloqueo Cloudflare WAF 403).
# ─────────────────────────────────────────────────────────────────────────────

from datetime import datetime, timedelta
import logging
import re
import sqlite3
import unicodedata
from pathlib import Path

import pandas as pd
import requests

from updater import (
    CANONICAL_TEAM_MAP,
    TEMPORADA_ACTUAL,
    TIMEOUT_ESPN,
    URLS_ESPN_SCOREBOARD,
    _extraer_eventos_espn_liga,
    resolver_equipo_canonico,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - [%(levelname)s] - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


class GestorBaseDatosFutbol:
    def __init__(self, nombre_bd="macro_base_futbol.db"):
        self.ruta_bd = Path(__file__).parent / nombre_bd

    def guardar_dataframe(self, df, nombre_tabla):
        try:
            with sqlite3.connect(self.ruta_bd) as conexion:
                df.to_sql(nombre_tabla, conexion, if_exists="append", index=False)
            logging.info(f"Guardados {len(df)} eventos en '{nombre_tabla}'.")
        except Exception as error:
            logging.error(f"Error al guardar en BD: {error}")


class ApiSofascore:
    """
    Cliente compatible hacia atrás que consume los endpoints abiertos de ESPN
    para las 5 grandes ligas europeas evitando por completo los bloqueos 403.
    """

    def extraer_ultimos_dias(self, dias: int = 7, temporada_actual: str = TEMPORADA_ACTUAL) -> pd.DataFrame:
        registros = []
        for liga, url in URLS_ESPN_SCOREBOARD.items():
            partidos = _extraer_eventos_espn_liga(
                liga=liga,
                url=url,
                mapa_equipos_bd={},
                dias_ventana=dias,
            )
            registros.extend(partidos)
        if not registros:
            return pd.DataFrame()
        return pd.DataFrame(registros)

    def extraer_partidos_por_fecha(self, fecha_str: str) -> pd.DataFrame | None:
        df = self.extraer_ultimos_dias(dias=7)
        if df.empty:
            return None
        return df


def extraer_ultimos_dias(dias: int = 7, temporada_actual: str = TEMPORADA_ACTUAL) -> pd.DataFrame:
    return ApiSofascore().extraer_ultimos_dias(dias=dias, temporada_actual=temporada_actual)


if __name__ == "__main__":
    df_eventos = extraer_ultimos_dias(dias=7)
    if not df_eventos.empty:
        print(f"\n--- Partidos Finalizados Extraídos (ESPN API, sin 403): {len(df_eventos)} ---")
        print(df_eventos[["Liga_Origen", "Date", "HomeTeam", "AwayTeam", "FTHG", "FTAG", "FTR"]].to_string(index=False))
    else:
        print("\nNo se encontraron partidos finalizados.")
