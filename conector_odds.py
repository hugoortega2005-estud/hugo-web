import requests
import sqlite3
import pandas as pd
import numpy as np
import math
from datetime import datetime
from pathlib import Path
import streamlit as st

API_KEY = "0aba9582adcfd166b4a9ab037d1400cf"
RUTA_BD = Path(__file__).parent / "macro_base_futbol.db"

CASAS_TOP = [
    "Winamax", "Winamax (FR)", "1xBet", "William Hill", 
    "Marathon Bet", "Pinnacle", "Betclic (FR)", "Betfair", "Bwin", "Codere"
]

MAPA_EQUIPOS_API = {
    "Rayo Vallecano": "Vallecano",
    "Athletic Bilbao": "Ath Bilbao",
    "Atletico Madrid": "Ath Madrid",
    "Atlético Madrid": "Ath Madrid",
    "Celta Vigo": "Celta",
    "Espanyol": "Espanol",
    "Real Betis": "Betis",
    "Real Sociedad": "Sociedad",
    "Deportivo Alaves": "Alaves",
    "Alavés": "Alaves",
    "CA Osasuna": "Osasuna",
    "Girona FC": "Girona",
    "RCD Mallorca": "Mallorca",
    "Getafe CF": "Getafe",
    "Sevilla FC": "Sevilla",
    "Valencia CF": "Valencia",
    "Villarreal CF": "Villarreal",
    "Manchester United": "Man United",
    "Manchester City": "Man City",
    "Nottingham Forest": "Nott'm Forest",
    "Wolverhampton Wanderers": "Wolves",
    "Newcastle United": "Newcastle",
    "Tottenham Hotspur": "Tottenham",
    "Brighton and Hove Albion": "Brighton",
    "West Ham United": "West Ham",
    "Leicester City": "Leicester",
    "Ipswich Town": "Ipswich",
    "Inter Milan": "Inter",
    "AC Milan": "Milan",
    "AS Roma": "Roma",
    "Atalanta BC": "Atalanta",
    "Hellas Verona": "Verona",
    "Parma Calcio 1913": "Parma",
    "Bayer Leverkusen": "Leverkusen",
    "Borussia Dortmund": "Dortmund",
    "Borussia Monchengladbach": "M'gladbach",
    "Eintracht Frankfurt": "Ein Frankfurt",
    "VfB Stuttgart": "Stuttgart",
    "SC Freiburg": "Freiburg",
    "TSG Hoffenheim": "Hoffenheim",
    "1. FC Heidenheim": "Heidenheim",
    "FC St. Pauli": "St Pauli",
    "Mainz 05": "Mainz",
    "VfL Wolfsburg": "Wolfsburg",
    "FC Augsburg": "Augsburg",
}

ULTIMO_ESTADO_API = {
    "ok": True,
    "status_code": 200,
    "mensaje": "OK",
}


class _RespuestaFallbackAPI:
    """Objeto de respuesta seguro para fallos de red o ausencia de clave en entornos cloud."""
    def __init__(self, status_code=503, message="Servicio no disponible temporalmente"):
        self.status_code = status_code
        self.text = message
        self._message = message

    def json(self):
        return {"message": self._message}


def _obtener_api_key(api_key_override=None):
    if api_key_override:
        return str(api_key_override).strip()
    try:
        if hasattr(st, "secrets") and "ODDS_API_KEY" in st.secrets and st.secrets["ODDS_API_KEY"]:
            return str(st.secrets["ODDS_API_KEY"]).strip()
    except Exception:
        pass
    return API_KEY


def _pedir_cuotas(sport_key, markets="h2h,totals", api_key=None):
    key_efectiva = _obtener_api_key(api_key)
    if not key_efectiva:
        return _RespuestaFallbackAPI(status_code=401, message="Falta clave ODDS_API_KEY")
    url = f"https://api.the-odds-api.com/v4/sports/{sport_key}/odds/?apiKey={key_efectiva}&regions=eu&markets={markets}"
    try:
        return requests.get(url, timeout=10)
    except Exception as exc:
        return _RespuestaFallbackAPI(status_code=503, message=f"Error de conexión a la API: {exc}")


def calcular_poisson(l, x):
    if l <= 0:
        return 1.0 if x == 0 else 0.0
    return (math.exp(-l) * (l**x)) / math.factorial(x)


def _media_con_decay(df_sub, col, usar_time_decay=False, factor_decay=0.01, default=1.25):
    try:
        if df_sub is None or df_sub.empty or col not in df_sub.columns:
            return float(default)
        serie = pd.to_numeric(df_sub[col], errors="coerce").dropna()
        if serie.empty:
            return float(default)
        if not usar_time_decay or "Date" not in df_sub.columns:
            return float(serie.mean())
        fechas = pd.to_datetime(df_sub.loc[serie.index, "Date"], errors="coerce", dayfirst=True)
        if fechas.notna().sum() < 2:
            return float(serie.mean())
        f_max = fechas.max()
        dias = (f_max - fechas).dt.days.fillna(0).clip(lower=0).to_numpy(dtype=float)
        pesos = np.exp(-float(factor_decay) * dias)
        if pesos.sum() <= 0:
            return float(serie.mean())
        return float(np.average(serie.to_numpy(dtype=float), weights=pesos))
    except Exception:
        return float(default)


def obtener_probabilidades_bd(local_api, visitante_api, usar_time_decay=False, factor_decay=0.01, **kwargs):
    local = MAPA_EQUIPOS_API.get(local_api, local_api)
    vis = MAPA_EQUIPOS_API.get(visitante_api, visitante_api)
    decay_eff = float(kwargs.get("xi", factor_decay) if factor_decay is None else factor_decay)

    if not RUTA_BD.exists():
        return None

    try:
        with sqlite3.connect(RUTA_BD) as conn:
            df = pd.read_sql("SELECT * FROM historico_reciente", conn)
    except Exception:
        return None

    if df is None or df.empty:
        return None

    p_local = df[df['HomeTeam'] == local]
    if p_local.empty:
        p_local = df[df['HomeTeam'].astype(str).str.lower() == str(local).lower()]
    p_vis = df[df['AwayTeam'] == vis]
    if p_vis.empty:
        p_vis = df[df['AwayTeam'].astype(str).str.lower() == str(vis).lower()]

    gf_l = _media_con_decay(p_local, 'FTHG', usar_time_decay=usar_time_decay, factor_decay=decay_eff, default=1.5)
    gc_l = _media_con_decay(p_local, 'FTAG', usar_time_decay=usar_time_decay, factor_decay=decay_eff, default=1.0)
    gf_v = _media_con_decay(p_vis, 'FTAG', usar_time_decay=usar_time_decay, factor_decay=decay_eff, default=1.0)
    gc_v = _media_con_decay(p_vis, 'FTHG', usar_time_decay=usar_time_decay, factor_decay=decay_eff, default=1.5)

    xg_l = max(0.15, (gf_l + gc_v) / 2)
    xg_v = max(0.15, (gf_v + gc_l) / 2)

    prob_1 = prob_X = prob_2 = 0
    prob_o05 = prob_o15 = prob_o25 = prob_o35 = prob_o45 = 0
    prob_btts_si = 0
    total = 0

    for h in range(7):
        for a in range(7):
            pr = calcular_poisson(xg_l, h) * calcular_poisson(xg_v, a)
            total += pr
            if h > a: prob_1 += pr
            elif h == a: prob_X += pr
            else: prob_2 += pr

            suma = h + a
            if suma > 0.5: prob_o05 += pr
            if suma > 1.5: prob_o15 += pr
            if suma > 2.5: prob_o25 += pr
            if suma > 3.5: prob_o35 += pr
            if suma > 4.5: prob_o45 += pr

            if h > 0 and a > 0: prob_btts_si += pr

    norm = total if total > 0 else 1.0

    return {
        "local_norm": local, "vis_norm": vis,
        "prob_1": prob_1 / norm, "prob_X": prob_X / norm, "prob_2": prob_2 / norm,
        "prob_o05": prob_o05 / norm, "prob_u05": max(1 - (prob_o05 / norm), 0),
        "prob_o15": prob_o15 / norm, "prob_u15": max(1 - (prob_o15 / norm), 0),
        "prob_o25": prob_o25 / norm, "prob_u25": max(1 - (prob_o25 / norm), 0),
        "prob_o35": prob_o35 / norm, "prob_u35": max(1 - (prob_o35 / norm), 0),
        "prob_o45": prob_o45 / norm, "prob_u45": max(1 - (prob_o45 / norm), 0),
        "prob_btts_si": prob_btts_si / norm, "prob_btts_no": max(1 - (prob_btts_si / norm), 0)
    }

def añadir_ev(oportunidades, fecha, datos_m, mercado, tipo, casa, cuota, prob_mod, cuota_min, cuota_max):
    try:
        cuota = float(cuota or 0.0)
        prob_mod = float(prob_mod or 0.0)
    except Exception:
        return
    if cuota < cuota_min or cuota > cuota_max or prob_mod <= 0: return
    ev = (prob_mod * cuota) - 1.0
    if ev > 0.01:
        b = cuota - 1.0
        kelly = (prob_mod * b - (1 - prob_mod)) / b if b > 0 else 0
        stake = round(100 * (kelly * 0.5), 2)
        oportunidades.append({
            "Fecha / Hora": fecha,
            "Partido": f"{datos_m['local_norm']} vs {datos_m['vis_norm']}",
            "Mercado": mercado,
            "Tipo_Mercado": tipo,
            "Casa": casa,
            "Cuota": cuota,
            "Prob_Modelo": f"{(prob_mod*100):.1f}%",
            "Cuota_Justa": round(1/prob_mod, 2),
            "EV_Num": round(ev * 100, 2),
            "EV": f"+{(ev*100):.1f}%",
            "Stake_Recomendado": f"{stake} €"
        })

def escaneo_apuestas_ev(
    sport_key="soccer_spain_la_liga",
    cuota_minima=1.12,
    cuota_maxima=3.50,
    usar_time_decay=False,
    factor_decay=0.01,
    **kwargs,
):
    global ULTIMO_ESTADO_API
    decay_eff = float(kwargs.get("xi", factor_decay) if factor_decay is not None else 0.01)
    api_key_eff = _obtener_api_key(kwargs.get("api_key"))

    try:
        res = _pedir_cuotas(sport_key, "h2h,totals", api_key_eff)

        if getattr(res, "status_code", 500) != 200:
            res = _pedir_cuotas(sport_key, "h2h,totals", api_key_eff)
            if getattr(res, "status_code", 500) != 200:
                try:
                    err_payload = res.json() if hasattr(res, "json") else {}
                    error_critico = err_payload.get("message", getattr(res, "text", "")) if isinstance(err_payload, dict) else str(getattr(res, "text", ""))
                except Exception:
                    error_critico = getattr(res, "text", "Límite de API alcanzado o servicio no disponible")
                ULTIMO_ESTADO_API = {
                    "ok": False,
                    "status_code": getattr(res, "status_code", 503),
                    "mensaje": error_critico or "Mercados no disponibles temporalmente o límite de API alcanzado",
                }
                return []

        partidos = res.json()
        if not isinstance(partidos, list):
            ULTIMO_ESTADO_API = {
                "ok": False,
                "status_code": 422,
                "mensaje": "Formato de respuesta inesperado desde el proveedor de cuotas",
            }
            return []

        ULTIMO_ESTADO_API = {
            "ok": True,
            "status_code": 200,
            "mensaje": "OK",
        }
    except Exception as exc_req:
        ULTIMO_ESTADO_API = {
            "ok": False,
            "status_code": 503,
            "mensaje": f"Mercados no disponibles temporalmente o límite de API alcanzado ({exc_req})",
        }
        return []

    oportunidades = []

    for p in partidos:
        try:
            if not isinstance(p, dict):
                continue
            loc_api = p.get('home_team', '')
            vis_api = p.get('away_team', '')
            if not loc_api or not vis_api:
                continue

            raw_date = p.get('commence_time', '')
            try:
                dt = datetime.fromisoformat(str(raw_date).replace('Z', '+00:00'))
                fecha_fmt = dt.strftime('%d/%m/%Y %H:%M')
            except Exception:
                fecha_fmt = str(raw_date)[:10]

            datos_m = obtener_probabilidades_bd(
                loc_api,
                vis_api,
                usar_time_decay=usar_time_decay,
                factor_decay=decay_eff,
                **kwargs,
            )
            if not datos_m:
                continue

            for bm in p.get('bookmakers', []):
                if not isinstance(bm, dict):
                    continue
                casa = str(bm.get('title', ''))
                if not any(top.lower() in casa.lower() for top in CASAS_TOP):
                    continue

                for market in bm.get('markets', []):
                    if not isinstance(market, dict):
                        continue
                    m_key = market.get('key')

                    if m_key == 'h2h':
                        for outcome in market.get('outcomes', []):
                            opc = outcome.get('name')
                            cuota = outcome.get('price')
                            if opc == loc_api: prob, nom = datos_m['prob_1'], "Victoria Local (1)"
                            elif opc == "Draw": prob, nom = datos_m['prob_X'], "Empate (X)"
                            else: prob, nom = datos_m['prob_2'], "Victoria Visitante (2)"
                            añadir_ev(oportunidades, fecha_fmt, datos_m, nom, "1X2", casa, cuota, prob, cuota_minima, cuota_maxima)

                    elif m_key == 'totals':
                        for outcome in market.get('outcomes', []):
                            name = outcome.get('name')
                            point = outcome.get('point')
                            cuota = outcome.get('price')

                            prob = 0; nom = ""
                            if name == 'Over':
                                if point == 0.5: prob, nom = datos_m['prob_o05'], "⚽ Más de 0.5 Goles"
                                elif point == 1.5: prob, nom = datos_m['prob_o15'], "⚽ Más de 1.5 Goles"
                                elif point == 2.5: prob, nom = datos_m['prob_o25'], "🔥 Más de 2.5 Goles"
                                elif point == 3.5: prob, nom = datos_m['prob_o35'], "🚀 Más de 3.5 Goles"
                                elif point == 4.5: prob, nom = datos_m['prob_o45'], "🚀 Más de 4.5 Goles"
                            elif name == 'Under':
                                if point == 0.5: prob, nom = datos_m['prob_u05'], "Menos de 0.5 Goles"
                                elif point == 1.5: prob, nom = datos_m['prob_u15'], "Menos de 1.5 Goles"
                                elif point == 2.5: prob, nom = datos_m['prob_u25'], "Menos de 2.5 Goles"
                                elif point == 3.5: prob, nom = datos_m['prob_u35'], "Menos de 3.5 Goles"
                                elif point == 4.5: prob, nom = datos_m['prob_u45'], "Menos de 4.5 Goles"

                            if prob > 0:
                                añadir_ev(oportunidades, fecha_fmt, datos_m, nom, "Goles", casa, cuota, prob, cuota_minima, cuota_maxima)

                    elif m_key == 'btts':
                        for outcome in market.get('outcomes', []):
                            name = outcome.get('name')
                            cuota = outcome.get('price')

                            if name == 'Yes': prob, nom = datos_m['prob_btts_si'], "🎯 Ambos Marcan (Sí)"
                            elif name == 'No': prob, nom = datos_m['prob_btts_no'], "🛑 Ambos Marcan (No)"
                            else: continue

                            añadir_ev(oportunidades, fecha_fmt, datos_m, nom, "BTTS", casa, cuota, prob, cuota_minima, cuota_maxima)
        except Exception:
            continue

    return oportunidades
