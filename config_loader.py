import configparser
import os

def cargar_config(config_file="config.ini"):
    """
    Carga el archivo de configuración y devuelve los valores en forma de diccionario.
    Si el archivo no existe o hay errores de lectura, lanza una excepción.
    """
    if not os.path.exists(config_file):
        raise FileNotFoundError(f"[ERROR] No se encontró el archivo de configuración: {config_file}")

    config = configparser.ConfigParser()
    config.read(config_file, encoding="utf-8")

    try:
        return {
            "PING_INTERVAL": int(config.get("GENERAL", "PING_INTERVAL", fallback="30")),

            "AD_SERVER": config.get("ACTIVE_DIRECTORY", "AD_SERVER"),
            "AD_USER": config.get("ACTIVE_DIRECTORY", "AD_USER"),
            "AD_PASSWORD": config.get("ACTIVE_DIRECTORY", "AD_PASSWORD"),
            "AD_SEARCH_BASE": config.get("ACTIVE_DIRECTORY", "AD_SEARCH_BASE"),

            "DB_DRIVER": config.get("DATABASE", "DB_DRIVER"),
            "DB_SERVER": config.get("DATABASE", "DB_SERVER"),
            "DB_NAME": config.get("DATABASE", "DB_NAME"),
            "DB_TRUSTED": config.get("DATABASE", "DB_TRUSTED", fallback="yes"),
            "DB_USER": config.get("DATABASE", "DB_USER", fallback=None),
            "DB_PASSWORD": config.get("DATABASE", "DB_PASSWORD", fallback=None)
        }
    except Exception as e:
        raise ValueError(f"[ERROR] No se pudieron leer los valores del archivo de configuración: {e}")
