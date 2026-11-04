import sqlite3
import pandas as pd
import logging
from pathlib import Path

# Configuración de logs
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - [%(levelname)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

class GestorBigDataFutbol:
    def __init__(self, nombre_bd="macro_base_futbol.db"):
        self.ruta_bd = Path(__file__).parent / nombre_bd
        logging.info(f"Iniciando motor Big Data en: {self.ruta_bd}")

    def vaciar_tabla_si_existe(self, nombre_tabla):
        """Limpia la tabla para evitar duplicados en cada ejecución."""
        with sqlite3.connect(self.ruta_bd) as conexion:
            conexion.cursor().execute(f"DROP TABLE IF EXISTS {nombre_tabla}")

    def guardar_datos_masivos(self, df, nombre_tabla):
        try:
            with sqlite3.connect(self.ruta_bd) as conexion:
                df.to_sql(nombre_tabla, conexion, if_exists='append', index=False)
            logging.info(f"Bloque integrado: {len(df)} partidos insertados.")
        except Exception as error:
            logging.error(f"Error en la ingesta de datos: {error}")

# ==========================================
# EJECUCIÓN PRINCIPAL DEL SCRIPT
# ==========================================
if __name__ == "__main__":
    motor_bd = GestorBigDataFutbol()
    tabla_destino = 'historico_reciente'
    
    motor_bd.vaciar_tabla_si_existe(tabla_destino)

    # Las 5 grandes ligas: España, Inglaterra, Italia, Alemania, Francia
    ligas = ['SP1', 'E0', 'I1', 'D1', 'F1']
    
    # Foco estricto: Temporada actual (26/27) y máximo dos años atrás (25/26, 24/25)
    temporadas = ['2627', '2526', '2425'] 
    
    lista_dataframes = []

    for temp in temporadas:
        for liga in ligas:
            url = f"https://www.football-data.co.uk/mmz4281/{temp}/{liga}.csv"
            logging.info(f"Ingestando datos: {liga} - Temporada {temp}")
            try:
                # Se descargan absolutamente todas las columnas (más de 100 variables)
                df_temp = pd.read_csv(url, encoding='latin1')
                
                # Añadimos variables de contexto para identificar los datos fácilmente
                df_temp['Temporada_Origen'] = temp
                df_temp['Liga_Origen'] = liga
                
                lista_dataframes.append(df_temp)
            except Exception as e:
                # Es normal que salte un warning si la web aún no ha subido el CSV de una liga concreta esta semana
                logging.warning(f"Archivo no disponible para {liga} ({temp}): {e}")

    # Consolidación final
    if lista_dataframes:
        # Concatenamos todo en un único macro-conjunto de datos
        df_macro = pd.concat(lista_dataframes, ignore_index=True)
        
        # Guardamos en la base de datos local
        motor_bd.guardar_datos_masivos(df_macro, tabla_destino)
        
        logging.info("\n==============================================")
        logging.info("   REPORTE DEL DATA LAKE (DATOS RECIENTES)")
        logging.info("==============================================")
        logging.info(f"-> Volumen total de partidos: {len(df_macro)}")
        logging.info(f"-> Variables (columnas) por partido: {len(df_macro.columns)}")
        logging.info("==============================================")