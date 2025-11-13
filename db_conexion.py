'''
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
'''


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
DB_USER = config.get("DATABASE", "DB_USER") if DB_TRUSTED.lower() != "yes" else None
DB_PASSWORD = config.get("DATABASE", "DB_PASSWORD") if DB_TRUSTED.lower() != "yes" else None

def conectar_sql():
    """
    Intenta conectarse a SQL Server indefinidamente hasta que tenga éxito.
    Devuelve la conexión activa.
    """
    while True:
        try:
            if DB_TRUSTED.lower() == "yes":
                conn_str = (
                    f"DRIVER={DB_DRIVER};"
                    f"SERVER={DB_SERVER};"
                    f"DATABASE={DB_NAME};"
                    "Trusted_Connection=yes;"
                )
            else:
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
            print(f"[ERROR] No se pudo conectar a SQL Server: {e}")
            print("  Reintentando en 5 segundos...")
            time.sleep(5)

def ejecutar_sql(conn, query, params=(), reintentos=3, espera=5):
    """
    Ejecuta un query SQL con reconexión automática en caso de fallo.
    """
    for intento in range(1, reintentos + 1):
        try:
            cursor = conn.cursor()
            cursor.execute(query, params)
            conn.commit()
            return True
        except pyodbc.OperationalError as e:
            print(f"[ERROR] Fallo de conexión: {e}")
            print(f"  Reconectando y reintentando ({intento}/{reintentos})...")
            conn = conectar_sql()
            time.sleep(espera)
        except pyodbc.Error as e:
            print(f"[ERROR] SQL Error: {e}")
            return False
    print("[FATAL] No se pudo ejecutar la consulta tras varios intentos.")
    return False

