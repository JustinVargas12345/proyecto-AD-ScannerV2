#import pyodbc
from ldap3 import Server, Connection, ALL
from ping3 import ping
import socket
import threading
import time
import datetime
import os 




'''
SQL_SERVER = r"DC\SQLEXPRESS"           # Nombre del servidor o instancia SQL
DB_NAME = "DbAlgoritmo"                 # Base de datos ya creada
TABLE_NAME = "EquiposAD"                # Tabla que usaremos
DOMINIO = "lab.local"                   # Dominio del Active Directory
USUARIO = "administrador@lab.local"     # Usuario con permisos de lectura
PASSWORD = "Laboratorio1"               # Contraseña del usuario
PING_INTERVAL = 30                      # Segundos entre cada ciclo de ping

# =====================================================
# FUNCIÓN: CONECTAR A SQL SERVER
# =====================================================
def conectar_sql():
    try:
        conexion = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={SQL_SERVER};"
            f"DATABASE={DB_NAME};"
            f"Trusted_Connection=yes;"
        )
        return conexion
    except Exception as e:
        print(f"[ERROR] No se pudo conectar a SQL Server: {e}")
        return None

# =====================================================
# FUNCIÓN: CREAR TABLA SI NO EXISTE
# =====================================================
def crear_tabla_si_no_existe(conexion):
    cursor = conexion.cursor()
    cursor.execute(f"""
    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='{TABLE_NAME}' AND xtype='U')
    CREATE TABLE {TABLE_NAME} (
        id INT IDENTITY(1,1) PRIMARY KEY,
        nombre NVARCHAR(255),
        ip NVARCHAR(50),
        sistema_operativo NVARCHAR(255),
        descripcion NVARCHAR(500),
        estado_registro NVARCHAR(100),
        estado_ping NVARCHAR(20),
        ultima_actualizacion DATETIME
    )
    """)
    conexion.commit()

# =====================================================
# FUNCIÓN: OBTENER EQUIPOS DESDE ACTIVE DIRECTORY
# =====================================================
def obtener_equipos_ad():
    equipos = []
    try:
        server = Server(DOMINIO, get_info=ALL)
        conn = Connection(server, user=USUARIO, password=PASSWORD, auto_bind=True)

        base_dn = ','.join([f"DC={p}" for p in DOMINIO.split('.')])
        conn.search(
            base_dn,
            '(objectClass=computer)',
            attributes=['name', 'dNSHostName', 'operatingSystem', 'description']
        )

        for entry in conn.entries: 
            nombre = str(entry.name)
            sistema = str(entry.operatingSystem) if 'operatingSystem' in entry else 'Desconocido'
            descripcion = str(entry.description) if 'description' in entry else 'N/A'

            try:
                ip = socket.gethostbyname(nombre)
                estado_registro = "Registrado"
            except:
                ip = "N/A"
                estado_registro = "No Resuelto"

            equipos.append({
                "nombre": nombre,
                "ip": ip,
                "sistema": sistema,
                "descripcion": descripcion,
                "estado_registro": estado_registro
            })

    except Exception as e:
        print("Error al conectar con AD:", e)
    return equipos

# =====================================================
# FUNCIÓN: GUARDAR DATOS EN SQL
# =====================================================
def guardar_en_sql(conexion, equipos):
    cursor = conexion.cursor()
    cursor.execute(f"TRUNCATE TABLE {TABLE_NAME}")
    conexion.commit()

    for equipo in equipos:
        cursor.execute(f"""
        INSERT INTO {TABLE_NAME} 
        (nombre, ip, sistema_operativo, descripcion, estado_registro, estado_ping, ultima_actualizacion)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            equipo["nombre"],
            equipo["ip"],
            equipo["sistema"],
            equipo["descripcion"],
            equipo["estado_registro"],
            "Desconocido",
            datetime.datetime.now()
        ))
    conexion.commit()

# =====================================================
# FUNCIÓN: HACER PING Y ACTUALIZAR SQL
# =====================================================
def actualizar_ping(conexion):
    cursor = conexion.cursor()
    cursor.execute(f"SELECT id, ip, nombre FROM {TABLE_NAME}")
    equipos = cursor.fetchall()

    for equipo in equipos:
        id_equipo, ip, nombre = equipo
        if ip == "N/A":
            estado = "No Resuelto"
        else:
            respuesta = os.system(f"ping -n 1 {ip} >nul 2>&1")
            estado = "Activo" if respuesta == 0 else "Inactivo"

        timestamp = datetime.datetime.now()

        cursor.execute(f"""
        UPDATE {TABLE_NAME}
        SET estado_ping = ?, ultima_actualizacion = ?
        WHERE id = ?
        """, (estado, timestamp, id_equipo))
        conexion.commit()

        print(f"[{timestamp.strftime('%H:%M:%S')}] Ping a {nombre} ({ip}) → {estado}")

# =====================================================
# PROGRAMA PRINCIPAL
# =====================================================
if __name__ == "__main__":
    print("\n======================================")
    print("   Escaneo de Servidores en AD (v2)  ")
    print("======================================")

    conexion = conectar_sql()
    if not conexion:
        print("[FALLO] No se pudo establecer conexión con SQL Server.")
        exit()

    crear_tabla_si_no_existe(conexion)

    print("\n[INFO] Consultando equipos en Active Directory...")
    equipos = obtener_equipos_ad()

    if equipos:
        print(f"[OK] {len(equipos)} equipos encontrados. Guardando en base de datos...")
        guardar_en_sql(conexion, equipos)
    else:
        print("[WARN] No se encontraron equipos en AD.")

    print(f"\n[INFO] Iniciando monitoreo de red (cada {PING_INTERVAL} segundos)...\n")

    while True:
        actualizar_ping(conexion)
        time.sleep(PING_INTERVAL)
'''
import pyodbc
import configparser
import os
import time
import platform
import subprocess
from ldap3 import Server, Connection, ALL

# ==========================================================
# 1️⃣ LECTURA DEL ARCHIVO CONFIG.INI
# ==========================================================
CONFIG_FILE = "config.ini"
config = configparser.ConfigParser()
config.read(CONFIG_FILE, encoding="utf-8")

DB_DRIVER = config.get("DATABASE", "DB_DRIVER")
DB_SERVER = config.get("DATABASE", "DB_SERVER")
DB_NAME = config.get("DATABASE", "DB_NAME")
DB_TRUSTED = config.get("DATABASE", "DB_TRUSTED")

AD_SERVER = config.get("ACTIVE_DIRECTORY", "AD_SERVER")
AD_USER = config.get("ACTIVE_DIRECTORY", "AD_USER")
AD_PASSWORD = config.get("ACTIVE_DIRECTORY", "AD_PASSWORD")
AD_SEARCH_BASE = config.get("ACTIVE_DIRECTORY", "AD_SEARCH_BASE")

# ==========================================================
# 2️⃣ CONEXIÓN A SQL SERVER
# ==========================================================
def conectar_sql():
    try:
        if DB_TRUSTED.lower() == "yes":
            conn_str = (
                f"DRIVER={DB_DRIVER};"
                f"SERVER={DB_SERVER};"
                f"DATABASE={DB_NAME};"
                f"Trusted_Connection=yes;"
            )
        else:
            conn_str = (
                f"DRIVER={DB_DRIVER};"
                f"SERVER={DB_SERVER};"
                f"DATABASE={DB_NAME};"
                f"UID={config.get('DATABASE', 'DB_USER')};"
                f"PWD={config.get('DATABASE', 'DB_PASSWORD')};"
            )
        conn = pyodbc.connect(conn_str)
        print("[OK] Conectado a SQL Server correctamente.")
        return conn
    except Exception as e:
        print("[ERROR] No se pudo conectar a SQL Server:", e)
        return None

# ==========================================================
# 3️⃣ CONEXIÓN AL ACTIVE DIRECTORY
# ==========================================================
def obtener_equipos_ad():
    equipos = []
    try:
        server = Server(AD_SERVER, get_info=ALL)
        conn = Connection(server, user=AD_USER, password=AD_PASSWORD, auto_bind=True)
        conn.search(
            AD_SEARCH_BASE,
            "(objectClass=computer)",
            attributes=["name", "operatingSystem", "description"]
        )

        for entry in conn.entries:
            nombre = str(entry.name)
            so = str(entry.operatingSystem) if hasattr(entry, 'operatingSystem') else "N/A"
            desc = str(entry.description) if hasattr(entry, 'description') else "N/A"
            equipos.append({"nombre": nombre, "so": so, "descripcion": desc})

        print(f"[OK] Equipos obtenidos desde AD: {len(equipos)} encontrados.")
    except Exception as e:
        print("[ERROR] Al leer equipos desde AD:", e)

    return equipos

# ==========================================================
# 4️⃣ HACER PING A CADA EQUIPO
# ==========================================================
def hacer_ping(host):
    try:
        param = "-n" if platform.system().lower() == "windows" else "-c"
        result = subprocess.run(["ping", param, "1", host], capture_output=True)
        return "Activo" if result.returncode == 0 else "Inactivo"
    except:
        return "Error"

# ==========================================================
# 5️⃣ CREAR TABLA SI NO EXISTE
# ==========================================================
def crear_tabla(conn):
    try:
        cursor = conn.cursor()
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='EquiposAD' AND xtype='U')
            CREATE TABLE EquiposAD (
                Nombre NVARCHAR(255) PRIMARY KEY,
                SO NVARCHAR(255),
                Descripcion NVARCHAR(255),
                PingStatus NVARCHAR(50),
                UltimaActualizacion DATETIME DEFAULT GETDATE()
            )
        """)
        conn.commit()
        print("[OK] Tabla 'EquiposAD' verificada o creada.")
    except Exception as e:
        print("[ERROR] Al crear/verificar la tabla:", e)

# ==========================================================
# 6️⃣ INSERTAR O ACTUALIZAR DATOS
# ==========================================================
def insertar_o_actualizar(conn, equipos):
    cursor = conn.cursor()
    for eq in equipos:
        ping = hacer_ping(eq["nombre"])
        cursor.execute("""
            MERGE EquiposAD AS target
            USING (SELECT ? AS Nombre, ? AS SO, ? AS Descripcion, ? AS PingStatus) AS src
            ON target.Nombre = src.Nombre
            WHEN MATCHED THEN
                UPDATE SET target.SO = src.SO,
                           target.Descripcion = src.Descripcion,
                           target.PingStatus = src.PingStatus,
                           target.UltimaActualizacion = GETDATE()
            WHEN NOT MATCHED THEN
                INSERT (Nombre, SO, Descripcion, PingStatus)
                VALUES (src.Nombre, src.SO, src.Descripcion, src.PingStatus);
        """, (eq["nombre"], eq["so"], eq["descripcion"], ping))
        conn.commit()
        print(f"[PING] {eq['nombre']} → {ping}")

# ==========================================================
# 7️⃣ PROCESO PRINCIPAL
# ==========================================================
def main():
    conn = conectar_sql()
    if not conn:
        return

    crear_tabla(conn)
    equipos = obtener_equipos_ad()
    if not equipos:
        print("[WARN] No se encontraron equipos en AD.")
        return

    while True:
        insertar_o_actualizar(conn, equipos)
        print("[INFO] Actualización completada. Esperando 30 segundos...\n")
        time.sleep(30)

if __name__ == "__main__":
    main()
