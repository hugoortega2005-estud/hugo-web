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
    "Atletico Madrid": "Atlético Madrid",
    "Celta Vigo": "Celta",
    "Espanyol": "Espanol",
    "Real Betis": "Betis",
    "Real Sociedad": "Sociedad"
}

def calcular_poisson(l, x):
    return (math.exp(-l) * (l**x)) / math.factorial(x)

def obtener_probabilidades_bd(local_api, visitante_api):
    local = MAPA_EQUIPOS_API.get(local_api, local_api)
    vis = MAPA_EQUIPOS_API.get(visitante_api, visitante_api)
    
    if not RUTA_BD.exists():
        return None
        
    with sqlite3.connect(RUTA_BD) as conn:
        try:
            df = pd.read_sql("SELECT * FROM historico_reciente", conn)
        except Exception:
            return None
            
    p_local = df[df['HomeTeam'] == local]
    p_vis = df[df['AwayTeam'] == vis]
    
    gf_l = p_local['FTHG'].mean() if not p_local.empty else 1.5
    gc_l = p_local['FTAG'].mean() if not p_local.empty else 1.0
    gf_v = p_vis['FTAG'].mean() if not p_vis.empty else 1.0
    gc_v = p_vis['FTHG'].mean() if not p_vis.empty else 1.5
    
    xg_l = (gf_l + gc_v) / 2
    xg_v = (gf_v + gc_l) / 2
    
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

def escaneo_apuestas_ev(sport_key="soccer_spain_la_liga", cuota_minima=1.12, cuota_maxima=3.50):
    url = f"https://api.the-odds-api.com/v4/sports/{sport_key}/odds/?apiKey={API_KEY}&regions=eu&markets=h2h,totals"
    res = requests.get(url)
    
    # Sistema de auto-recuperación si la API rechaza el mercado múltiple o nos quedamos sin cuota
    if res.status_code != 200:
        error_msg = res.json().get('message', res.text) if res.text else "Error desconocido"
        st.warning(f"⚠️ La API ha devuelto un error (Código {res.status_code}): {error_msg}. Activando modo de emergencia sin mercado BTTS...")
        
        url_segura = f"https://api.the-odds-api.com/v4/sports/{sport_key}/odds/?apiKey={API_KEY}&regions=eu&markets=h2h,totals"
        res = requests.get(url_segura)
        
        if res.status_code != 200:
            error_critico = res.json().get('message', res.text) if res.text else "Límite agotado"
            st.error(f"❌ Error crítico de API (Código {res.status_code}): {error_critico}. Es probable que hayas agotado tu límite de peticiones.")
            return []
            
    partidos = res.json()
    oportunidades = []
    
    for p in partidos:
        loc_api = p['home_team']
        vis_api = p['away_team']
        
        raw_date = p.get('commence_time', '')
        try:
            dt = datetime.fromisoformat(raw_date.replace('Z', '+00:00'))
            fecha_fmt = dt.strftime('%d/%m/%Y %H:%M')
        except: 
            fecha_fmt = raw_date[:10]
        
        datos_m = obtener_probabilidades_bd(loc_api, vis_api)
        if not datos_m: continue
            
        for bm in p['bookmakers']:
            casa = bm['title']
            if not any(top.lower() in casa.lower() for top in CASAS_TOP): continue

            for market in bm.get('markets', []):
                m_key = market.get('key')
                
                if m_key == 'h2h':
                    for outcome in market['outcomes']:
                        opc = outcome['name']
                        cuota = outcome['price']
                        if opc == loc_api: prob, nom = datos_m['prob_1'], "Victoria Local (1)"
                        elif opc == "Draw": prob, nom = datos_m['prob_X'], "Empate (X)"
                        else: prob, nom = datos_m['prob_2'], "Victoria Visitante (2)"
                        añadir_ev(oportunidades, fecha_fmt, datos_m, nom, "1X2", casa, cuota, prob, cuota_minima, cuota_maxima)

                elif m_key == 'totals':
                    for outcome in market['outcomes']:
                        name = outcome.get('name')
                        point = outcome.get('point')
                        cuota = outcome['price']
                        
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
                        
                        if prob > 0: añadir_ev(oportunidades, fecha_fmt, datos_m, nom, "Goles", casa, cuota, prob, cuota_minima, cuota_maxima)

                elif m_key == 'btts':
                    for outcome in market['outcomes']:
                        name = outcome.get('name')
                        cuota = outcome['price']
                        
                        if name == 'Yes': prob, nom = datos_m['prob_btts_si'], "🎯 Ambos Marcan (Sí)"
                        elif name == 'No': prob, nom = datos_m['prob_btts_no'], "🛑 Ambos Marcan (No)"
                        else: continue
                            
                        añadir_ev(oportunidades, fecha_fmt, datos_m, nom, "BTTS", casa, cuota, prob, cuota_minima, cuota_maxima)

    return oportunidades