# db_conexion.py
import pyodbc
import configparser
import time

CONFIG_FILE = "config.ini"
config = configparser.ConfigParser()
config.read(CONFIG_FILE, encoding="utf-8")

DB_DRIVER = config.get("DATABASE", "DB_DRIVER")
DB_SERVER = config.get("DATABASE", "DB_SERVER")
DB_NAME = config.get("DATABASE", "DB_NAME")
DB_TRUSTED = config.get("DATABASE", "DB_TRUSTED")

def conectar_sql(reintentos=3, espera=5):
    for intento in range(1, reintentos + 1):
        try:
            if DB_TRUSTED.lower() == "yes":
                conn_str = (
                    f"DRIVER={DB_DRIVER};"
                    f"SERVER={DB_SERVER};"
                    f"DATABASE={DB_NAME};"
                    "Trusted_Connection=yes;"
                )
            else:
                DB_USER = config.get("DATABASE", "DB_USER")
                DB_PASSWORD = config.get("DATABASE", "DB_PASSWORD")
                conn_str = (
                    f"DRIVER={DB_DRIVER};"
                    f"SERVER={DB_SERVER};"
                    f"DATABASE={DB_NAME};"
                    f"UID={DB_USER};"
                    f"PWD={DB_PASSWORD};"
                )

            conn = pyodbc.connect(conn_str, timeout=5)
            print("[OK] Conectado a SQL Server correctamente.")
            return conn
        except pyodbc.Error as e:
            print(f"[ERROR] Intento {intento}: No se pudo conectar a SQL Server: {e}")
            if intento < reintentos:
                print(f"  Reintentando en {espera} segundos...")
                time.sleep(espera)
            else:
                print("[FATAL] No se pudo establecer conexión con SQL Server.")
                return None
