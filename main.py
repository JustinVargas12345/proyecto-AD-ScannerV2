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
                ip = socket.gethostbyname(nombre)##
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

#################
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
        ping = hacer_ping(eq["nombre"])#######
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
'''

####
'''
import pyodbc
import configparser
import os
import time
import platform
import subprocess
from ldap3 import Server, Connection, ALL
import socket   # 🔹 agregado para obtener la IP

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

            # 🔹 Obtener IP del equipo
            try:
                ip = socket.gethostbyname(nombre)
            except:
                ip = "No resuelve"

            equipos.append({"nombre": nombre, "so": so, "descripcion": desc, "ip": ip})

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
                IP NVARCHAR(50),               -- 🔹 Campo nuevo
                PingStatus NVARCHAR(50),
                UltimaActualizacion DATETIME DEFAULT GETDATE()
            )
            ELSE
            BEGIN
                IF NOT EXISTS (
                    SELECT * FROM sys.columns 
                    WHERE Name = N'IP' AND Object_ID = Object_ID(N'EquiposAD')
                )
                ALTER TABLE EquiposAD ADD IP NVARCHAR(50);  -- 🔹 Agrega columna si no existe
            END
        """)
        conn.commit()
        print("[OK] Tabla 'EquiposAD' verificada o creada (con campo IP).")
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
            USING (SELECT ? AS Nombre, ? AS SO, ? AS Descripcion, ? AS IP, ? AS PingStatus) AS src
            ON target.Nombre = src.Nombre
            WHEN MATCHED THEN
                UPDATE SET target.SO = src.SO,
                           target.Descripcion = src.Descripcion,
                           target.IP = src.IP,                      -- 🔹 Se actualiza IP
                           target.PingStatus = src.PingStatus,
                           target.UltimaActualizacion = GETDATE()
            WHEN NOT MATCHED THEN
                INSERT (Nombre, SO, Descripcion, IP, PingStatus)
                VALUES (src.Nombre, src.SO, src.Descripcion, src.IP, src.PingStatus);
        """, (eq["nombre"], eq["so"], eq["descripcion"], eq["ip"], ping))
        conn.commit()
        print(f"[PING] {eq['nombre']} ({eq['ip']}) → {ping}")

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
    '''
    
    #####################
    
'''
import pyodbc
import configparser
import os
import time
import platform
import subprocess
from ldap3 import Server, Connection, ALL
import socket   # 🔹 Para obtener la IP

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
            attributes=[
                "name",
                "dNSHostName",
                "operatingSystem",
                "operatingSystemVersion",
                "description",
                "whenCreated",
                "lastLogonTimestamp",
                "managedBy",
                "location",
                "userAccountControl"
            ]
        )

        for entry in conn.entries:
            nombre = str(entry.name)
            so = str(entry.operatingSystem) if hasattr(entry, 'operatingSystem') else "N/A"
            desc = str(entry.description) if hasattr(entry, 'description') else "N/A"
            nombre_dns = str(entry.dNSHostName) if hasattr(entry, 'dNSHostName') else "N/A"
            version_so = str(entry.operatingSystemVersion) if hasattr(entry, 'operatingSystemVersion') else "N/A"
            creado_el = str(entry.whenCreated) if hasattr(entry, 'whenCreated') else "N/A"
            ultimo_logon = str(entry.lastLogonTimestamp) if hasattr(entry, 'lastLogonTimestamp') else "N/A"
            responsable = str(entry.managedBy) if hasattr(entry, 'managedBy') else "N/A"
            ubicacion = str(entry.location) if hasattr(entry, 'location') else "N/A"
            estado_cuenta = str(entry.userAccountControl) if hasattr(entry, 'userAccountControl') else "N/A"

            # 🔹 Obtener IP del equipo
            try:
                ip = socket.gethostbyname(nombre)
            except:
                ip = "No resuelve"

            equipos.append({
                "nombre": nombre,
                "so": so,
                "descripcion": desc,
                "ip": ip,
                "nombredns": nombre_dns,
                "versionso": version_so,
                "creadoel": creado_el,
                "ultimologon": ultimo_logon,
                "responsable": responsable,
                "ubicacion": ubicacion,
                "estadocuenta": estado_cuenta
            })

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
                IP NVARCHAR(50),
                NombreDNS NVARCHAR(255),
                VersionSO NVARCHAR(255),
                CreadoEl NVARCHAR(100),
                UltimoLogon NVARCHAR(100),
                Responsable NVARCHAR(255),
                Ubicacion NVARCHAR(255),
                EstadoCuenta NVARCHAR(50),
                PingStatus NVARCHAR(50),
                UltimaActualizacion DATETIME DEFAULT GETDATE()
            )
        """)
        conn.commit()
        print("[OK] Tabla 'EquiposAD' verificada o creada con nuevos campos.")
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
            USING (SELECT ? AS Nombre, ? AS SO, ? AS Descripcion, ? AS IP, ? AS NombreDNS,
                          ? AS VersionSO, ? AS CreadoEl, ? AS UltimoLogon, ? AS Responsable,
                          ? AS Ubicacion, ? AS EstadoCuenta, ? AS PingStatus) AS src
            ON target.Nombre = src.Nombre
            WHEN MATCHED THEN
                UPDATE SET target.SO = src.SO,
                           target.Descripcion = src.Descripcion,
                           target.IP = src.IP,
                           target.NombreDNS = src.NombreDNS,
                           target.VersionSO = src.VersionSO,
                           target.CreadoEl = src.CreadoEl,
                           target.UltimoLogon = src.UltimoLogon,
                           target.Responsable = src.Responsable,
                           target.Ubicacion = src.Ubicacion,
                           target.EstadoCuenta = src.EstadoCuenta,
                           target.PingStatus = src.PingStatus,
                           target.UltimaActualizacion = GETDATE()
            WHEN NOT MATCHED THEN
                INSERT (Nombre, SO, Descripcion, IP, NombreDNS, VersionSO, CreadoEl, UltimoLogon,
                        Responsable, Ubicacion, EstadoCuenta, PingStatus)
                VALUES (src.Nombre, src.SO, src.Descripcion, src.IP, src.NombreDNS, src.VersionSO,
                        src.CreadoEl, src.UltimoLogon, src.Responsable, src.Ubicacion, src.EstadoCuenta, src.PingStatus);
        """, (
            eq["nombre"], eq["so"], eq["descripcion"], eq["ip"], eq["nombredns"],
            eq["versionso"], eq["creadoel"], eq["ultimologon"], eq["responsable"],
            eq["ubicacion"], eq["estadocuenta"], ping
        ))
        conn.commit()
        print(f"[PING] {eq['nombre']} ({eq['ip']}) → {ping}")

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
'''

'''
import pyodbc
import configparser
import os
import time
import platform
import subprocess
from ldap3 import Server, Connection, ALL
import socket   # 🔹 Para obtener la IP

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
            attributes=[
                "name",
                "dNSHostName",
                "operatingSystem",
                "operatingSystemVersion",
                "description",
                "whenCreated",
                "lastLogonTimestamp",
                "managedBy",
                "location",
                "userAccountControl"
            ]
        )

        for entry in conn.entries:
            nombre = str(entry.name)
            so = str(entry.operatingSystem) if hasattr(entry, 'operatingSystem') else "N/A"
            desc = str(entry.description) if hasattr(entry, 'description') else "N/A"
            nombre_dns = str(entry.dNSHostName) if hasattr(entry, 'dNSHostName') else "N/A"
            version_so = str(entry.operatingSystemVersion) if hasattr(entry, 'operatingSystemVersion') else "N/A"
            creado_el = str(entry.whenCreated) if hasattr(entry, 'whenCreated') else "N/A"
            ultimo_logon = str(entry.lastLogonTimestamp) if hasattr(entry, 'lastLogonTimestamp') else "N/A"
            responsable = str(entry.managedBy) if hasattr(entry, 'managedBy') else "N/A"
            ubicacion = str(entry.location) if hasattr(entry, 'location') else "N/A"
            estado_cuenta = str(entry.userAccountControl) if hasattr(entry, 'userAccountControl') else "N/A"

            # 🔹 Obtener IP del equipo
            try:
                ip = socket.gethostbyname(nombre)
            except:
                ip = "No resuelve"

            equipos.append({
                "nombre": nombre,
                "so": so,
                "descripcion": desc,
                "ip": ip,
                "nombredns": nombre_dns,
                "versionso": version_so,
                "creadoel": creado_el,
                "ultimologon": ultimo_logon,
                "responsable": responsable,
                "ubicacion": ubicacion,
                "estadocuenta": estado_cuenta
            })

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
# 5️⃣ CREAR TABLA SI NO EXISTE (con TiempoPing)
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
                IP NVARCHAR(50),
                NombreDNS NVARCHAR(255),
                VersionSO NVARCHAR(255),
                CreadoEl NVARCHAR(100),
                UltimoLogon NVARCHAR(100),
                Responsable NVARCHAR(255),
                Ubicacion NVARCHAR(255),
                EstadoCuenta NVARCHAR(50),
                PingStatus NVARCHAR(50),
                TiempoPing NVARCHAR(20),
                UltimaActualizacion DATETIME DEFAULT GETDATE()
            )
        """)
        conn.commit()
        print("[OK] Tabla 'EquiposAD' verificada o creada con nuevos campos.")
    except Exception as e:
        print("[ERROR] Al crear/verificar la tabla:", e)

# ==========================================================
# Diccionario global para llevar el conteo de ping
# ==========================================================
estado_ping = {}  # {"NombreEquipo": {"estado": "Activo", "contador": 3}}

# ==========================================================
# 6️⃣ INSERTAR O ACTUALIZAR DATOS (con contador y tiempo)
# ==========================================================
def insertar_o_actualizar(conn, equipos):
    cursor = conn.cursor()
    for eq in equipos:
        ping = hacer_ping(eq["nombre"])

        # 🔹 Actualizar contador de estado
        if eq["nombre"] in estado_ping:
            if estado_ping[eq["nombre"]]["estado"] == ping:
                estado_ping[eq["nombre"]]["contador"] += 1
            else:
                estado_ping[eq["nombre"]]["estado"] = ping
                estado_ping[eq["nombre"]]["contador"] = 1
        else:
            estado_ping[eq["nombre"]] = {"estado": ping, "contador": 1}

        # 🔹 Calcular tiempo total en hh:mm:ss
        tiempo_total_segundos = estado_ping[eq["nombre"]]["contador"] * 30
        horas = tiempo_total_segundos // 3600
        minutos = (tiempo_total_segundos % 3600) // 60
        segundos = tiempo_total_segundos % 60
        tiempo_formateado = f"{horas:02}:{minutos:02}:{segundos:02}"

        # 🔹 Insertar o actualizar en SQL
        cursor.execute("""
            MERGE EquiposAD AS target
            USING (SELECT ? AS Nombre, ? AS SO, ? AS Descripcion, ? AS IP, ? AS NombreDNS,
                          ? AS VersionSO, ? AS CreadoEl, ? AS UltimoLogon, ? AS Responsable,
                          ? AS Ubicacion, ? AS EstadoCuenta, ? AS PingStatus, ? AS TiempoPing) AS src
            ON target.Nombre = src.Nombre
            WHEN MATCHED THEN
                UPDATE SET target.SO = src.SO,
                           target.Descripcion = src.Descripcion,
                           target.IP = src.IP,
                           target.NombreDNS = src.NombreDNS,
                           target.VersionSO = src.VersionSO,
                           target.CreadoEl = src.CreadoEl,
                           target.UltimoLogon = src.UltimoLogon,
                           target.Responsable = src.Responsable,
                           target.Ubicacion = src.Ubicacion,
                           target.EstadoCuenta = src.EstadoCuenta,
                           target.PingStatus = src.PingStatus,
                           target.TiempoPing = src.TiempoPing,
                           target.UltimaActualizacion = GETDATE()
            WHEN NOT MATCHED THEN
                INSERT (Nombre, SO, Descripcion, IP, NombreDNS, VersionSO, CreadoEl, UltimoLogon,
                        Responsable, Ubicacion, EstadoCuenta, PingStatus, TiempoPing)
                VALUES (src.Nombre, src.SO, src.Descripcion, src.IP, src.NombreDNS, src.VersionSO,
                        src.CreadoEl, src.UltimoLogon, src.Responsable, src.Ubicacion, src.EstadoCuenta, src.PingStatus, src.TiempoPing);
        """, (
            eq["nombre"], eq["so"], eq["descripcion"], eq["ip"], eq["nombredns"],
            eq["versionso"], eq["creadoel"], eq["ultimologon"], eq["responsable"],
            eq["ubicacion"], eq["estadocuenta"], ping, tiempo_formateado
        ))
        conn.commit()
        print(f"[PING] {eq['nombre']} ({eq['ip']}) → {ping} ({tiempo_formateado})")

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
'''
######################
'''
#import logging
#from tkinter import messagebox, Tk
import pyodbc
import configparser
import os
import time
import platform
import subprocess
from ldap3 import Server, Connection, ALL
import socket
from datetime import datetime  # 🔹 Para registrar la fecha de inactividad

# ==========================================================
# 📋 CONFIGURACIÓN DE LOGS
# ==========================================================
'''
#if not os.path.exists("logs"):
    #os.makedirs("logs")

#logging.basicConfig(
    #filename="logs/estado_app.log",
    #level=logging.INFO,
    #format="%(asctime)s - %(levelname)s - %(message)s",
    #encoding="utf-8"
#)
'''
################################################



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
            attributes=[
                "name",
                "dNSHostName",
                "operatingSystem",
                "operatingSystemVersion",
                "description",
                "whenCreated",
                "lastLogonTimestamp",
                "managedBy",
                "location",
                "userAccountControl"
            ]
        )

        for entry in conn.entries:
            nombre = str(entry.name)
            so = str(entry.operatingSystem) if hasattr(entry, 'operatingSystem') else "N/A"
            desc = str(entry.description) if hasattr(entry, 'description') else "N/A"
            nombre_dns = str(entry.dNSHostName) if hasattr(entry, 'dNSHostName') else "N/A"
            version_so = str(entry.operatingSystemVersion) if hasattr(entry, 'operatingSystemVersion') else "N/A"
            creado_el = str(entry.whenCreated) if hasattr(entry, 'whenCreated') else "N/A"
            ultimo_logon = str(entry.lastLogonTimestamp) if hasattr(entry, 'lastLogonTimestamp') else "N/A"
            responsable = str(entry.managedBy) if hasattr(entry, 'managedBy') else "N/A"
            ubicacion = str(entry.location) if hasattr(entry, 'location') else "N/A"
            estado_cuenta = str(entry.userAccountControl) if hasattr(entry, 'userAccountControl') else "N/A"

            try:
                ip = socket.gethostbyname(nombre)
            except:
                ip = "No resuelve"

            equipos.append({
                "nombre": nombre,
                "so": so,
                "descripcion": desc,
                "ip": ip,
                "nombredns": nombre_dns,
                "versionso": version_so,
                "creadoel": creado_el,
                "ultimologon": ultimo_logon,
                "responsable": responsable,
                "ubicacion": ubicacion,
                "estadocuenta": estado_cuenta
            })

        print(f"[OK] Equipos obtenidos desde AD: {len(equipos)} encontrados.")
    except Exception as e:
        print("[ERROR] Al leer equipos desde AD:", e)

    return equipos

'''
#def mostrar_mensaje(titulo, mensaje):
    #try:
       # ventana = Tk()
       # ventana.withdraw()  # No mostrar ventana principal
       # messagebox.showinfo(titulo, mensaje)
       # ventana.destroy()
    #except:
        #pass  # En caso de que tkinter no funcione (modo servidor)
'''

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
# 5️⃣ CREAR TABLA SI NO EXISTE (agregando InactivoDesde)
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
                IP NVARCHAR(50),
                NombreDNS NVARCHAR(255),
                VersionSO NVARCHAR(255),
                CreadoEl NVARCHAR(100),
                UltimoLogon NVARCHAR(100),
                Responsable NVARCHAR(255),
                Ubicacion NVARCHAR(255),
                EstadoCuenta NVARCHAR(50),
                PingStatus NVARCHAR(50),
                TiempoPing NVARCHAR(20),
                InactivoDesde DATETIME NULL,
                UltimaActualizacion DATETIME DEFAULT GETDATE()
            )
        """)
        conn.commit()
        print("[OK] Tabla 'EquiposAD' verificada o creada con nuevos campos.")
    except Exception as e:
        print("[ERROR] Al crear/verificar la tabla:", e)

# ==========================================================
# Diccionario global para llevar el conteo y la fecha
# ==========================================================
estado_ping = {}  # {"Equipo": {"estado": "Activo", "contador": 3, "inactivo_desde": fecha}}

# ==========================================================
# 6️⃣ INSERTAR O ACTUALIZAR DATOS
# ==========================================================
def insertar_o_actualizar(conn, equipos):
    cursor = conn.cursor()
    for eq in equipos:
        ping = hacer_ping(eq["nombre"])

        # 🔹 Actualizar estado y contador
        if eq["nombre"] in estado_ping:
            anterior = estado_ping[eq["nombre"]]["estado"]
            if anterior == ping:
                estado_ping[eq["nombre"]]["contador"] += 1
            else:
                estado_ping[eq["nombre"]]["estado"] = ping
                estado_ping[eq["nombre"]]["contador"] = 1

                # 🔹 Si cambió a Inactivo, registrar fecha actual
                if ping == "Inactivo":
                    estado_ping[eq["nombre"]]["inactivo_desde"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                else:
                    estado_ping[eq["nombre"]]["inactivo_desde"] = None
        else:
            estado_ping[eq["nombre"]] = {
                "estado": ping,
                "contador": 1,
                "inactivo_desde": datetime.now().strftime("%Y-%m-%d %H:%M:%S") if ping == "Inactivo" else None
            }

        # 🔹 Calcular tiempo total (en base a ciclos de 30 seg)
        tiempo_total_segundos = estado_ping[eq["nombre"]]["contador"] * 30
        horas = tiempo_total_segundos // 3600
        minutos = (tiempo_total_segundos % 3600) // 60
        segundos = tiempo_total_segundos % 60
        tiempo_formateado = f"{horas:02}:{minutos:02}:{segundos:02}"

        inactivo_desde = estado_ping[eq["nombre"]]["inactivo_desde"]

        # 🔹 Guardar en SQL
        cursor.execute("""
            MERGE EquiposAD AS target
            USING (SELECT ? AS Nombre, ? AS SO, ? AS Descripcion, ? AS IP, ? AS NombreDNS,
                          ? AS VersionSO, ? AS CreadoEl, ? AS UltimoLogon, ? AS Responsable,
                          ? AS Ubicacion, ? AS EstadoCuenta, ? AS PingStatus, ? AS TiempoPing, ? AS InactivoDesde) AS src
            ON target.Nombre = src.Nombre
            WHEN MATCHED THEN
                UPDATE SET target.SO = src.SO,
                           target.Descripcion = src.Descripcion,
                           target.IP = src.IP,
                           target.NombreDNS = src.NombreDNS,
                           target.VersionSO = src.VersionSO,
                           target.CreadoEl = src.CreadoEl,
                           target.UltimoLogon = src.UltimoLogon,
                           target.Responsable = src.Responsable,
                           target.Ubicacion = src.Ubicacion,
                           target.EstadoCuenta = src.EstadoCuenta,
                           target.PingStatus = src.PingStatus,
                           target.TiempoPing = src.TiempoPing,
                           target.InactivoDesde = src.InactivoDesde,
                           target.UltimaActualizacion = GETDATE()
            WHEN NOT MATCHED THEN
                INSERT (Nombre, SO, Descripcion, IP, NombreDNS, VersionSO, CreadoEl, UltimoLogon,
                        Responsable, Ubicacion, EstadoCuenta, PingStatus, TiempoPing, InactivoDesde)
                VALUES (src.Nombre, src.SO, src.Descripcion, src.IP, src.NombreDNS, src.VersionSO,
                        src.CreadoEl, src.UltimoLogon, src.Responsable, src.Ubicacion,
                        src.EstadoCuenta, src.PingStatus, src.TiempoPing, src.InactivoDesde);
        """, (
            eq["nombre"], eq["so"], eq["descripcion"], eq["ip"], eq["nombredns"],
            eq["versionso"], eq["creadoel"], eq["ultimologon"], eq["responsable"],
            eq["ubicacion"], eq["estadocuenta"], ping, tiempo_formateado, inactivo_desde
        ))
        conn.commit()

        texto_fecha = f" | Inactivo desde: {inactivo_desde}" if inactivo_desde else ""
        print(f"[PING] {eq['nombre']} ({eq['ip']}) → {ping} ({tiempo_formateado}){texto_fecha}")



'''
#def check_conexiones():
    #errores = []

    # 🔹 Verificar conexión SQL
    #try:
       # conn = conectar_sql()
        #if conn:
            #logging.info("[OK] Conexión a SQL Server exitosa.")
            #conn.close()
        #else:
            #errores.append("No se pudo conectar a SQL Server.")
            #logging.error("[ERROR] No se pudo conectar a SQL Server.")
    #except Exception as e:
        #errores.append(f"Error SQL: {e}")
        #logging.error(f"[ERROR] SQL: {e}")

    # 🔹 Verificar conexión AD
    #try:
        #server = Server(AD_SERVER, get_info=ALL)
        #conn_ad = Connection(server, user=AD_USER, password=AD_PASSWORD, auto_bind=True)
        #conn_ad.search(AD_SEARCH_BASE, "(objectClass=computer)", attributes=["name"])
        #logging.info("[OK] Conexión al Active Directory exitosa.")
        #conn_ad.unbind()
    #except Exception as e:
        #errores.append(f"Error AD: {e}")
        #logging.error(f"[ERROR] Active Directory: {e}")

    # 🔹 Resultado final
    #if errores:
       # mostrar_mensaje("Error de conexión", "\n".join(errores))
        #return False
   # else:
        #mostrar_mensaje("Conexiones exitosas", "Conexión a SQL Server y Active Directory establecida correctamente.")
        #return True



# ==========================================================
# 7️⃣ PROCESO PRINCIPAL
# ==========================================================
'''
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
'''


'''
def main():
    if not check_conexiones():
        logging.error("❌ Falló la verificación de conexiones. Cerrando aplicación.")
        return

    conn = conectar_sql()
    if not conn:
        return

    crear_tabla(conn)
    equipos = obtener_equipos_ad()
    if not equipos:
        logging.warning("[WARN] No se encontraron equipos en AD.")
        mostrar_mensaje("Advertencia", "No se encontraron equipos en el Active Directory.")
        return

    while True:
        insertar_o_actualizar(conn, equipos)
        logging.info("[INFO] Actualización completada. Esperando 30 segundos...")
        time.sleep(30)
'''
####################################

import configparser
import time
import platform
import subprocess
from ldap3 import Server, Connection, ALL
import socket
from datetime import datetime
from db_conexion import conectar_sql
# from db_table import crear_tabla  # Solo si quieres crearla manualmente

# ======================
# CONFIGURACIÓN
# ======================
CONFIG_FILE = "config.ini"
config = configparser.ConfigParser()
config.read(CONFIG_FILE, encoding="utf-8")
PING_INTERVAL = int(config.get("GENERAL", "PING_INTERVAL", fallback="30"))

AD_SERVER = config.get("ACTIVE_DIRECTORY", "AD_SERVER")
AD_USER = config.get("ACTIVE_DIRECTORY", "AD_USER")
AD_PASSWORD = config.get("ACTIVE_DIRECTORY", "AD_PASSWORD")
AD_SEARCH_BASE = config.get("ACTIVE_DIRECTORY", "AD_SEARCH_BASE")

# ======================
# ESTADO GLOBAL DE PING
# ======================
estado_ping = {}

# ======================
# FUNCIONES
# ======================
def obtener_equipos_ad():
    equipos = []
    try:
        server = Server(AD_SERVER, get_info=ALL)
        conn = Connection(server, user=AD_USER, password=AD_PASSWORD, auto_bind=True)
        conn.search(
            AD_SEARCH_BASE,
            "(objectClass=computer)",
            attributes=[
                "name","dNSHostName","operatingSystem","operatingSystemVersion",
                "description","whenCreated","lastLogonTimestamp","managedBy",
                "location","userAccountControl"
            ]
        )
        for entry in conn.entries:
            nombre = str(entry.name)
            so = str(entry.operatingSystem) if hasattr(entry, 'operatingSystem') else "N/A"
            desc = str(entry.description) if hasattr(entry, 'description') else "N/A"
            nombre_dns = str(entry.dNSHostName) if hasattr(entry, 'dNSHostName') else "N/A"
            version_so = str(entry.operatingSystemVersion) if hasattr(entry, 'operatingSystemVersion') else "N/A"
            creado_el = str(entry.whenCreated) if hasattr(entry, 'whenCreated') else "N/A"
            ultimo_logon = str(entry.lastLogonTimestamp) if hasattr(entry, 'lastLogonTimestamp') else "N/A"
            responsable = str(entry.managedBy) if hasattr(entry, 'managedBy') else "N/A"
            ubicacion = str(entry.location) if hasattr(entry, 'location') else "N/A"
            estado_cuenta = str(entry.userAccountControl) if hasattr(entry, 'userAccountControl') else "N/A"

            try:
                ip = socket.gethostbyname(nombre)
            except socket.gaierror:
                ip = "No resuelve"

            equipos.append({
                "nombre": nombre,"so": so,"descripcion": desc,"ip": ip,
                "nombredns": nombre_dns,"versionso": version_so,"creadoel": creado_el,
                "ultimologon": ultimo_logon,"responsable": responsable,
                "ubicacion": ubicacion,"estadocuenta": estado_cuenta
            })
        print(f"[OK] Equipos obtenidos desde AD: {len(equipos)} encontrados.")
    except Exception as e:
        print("[ERROR] Excepción al leer AD:", e)
    return equipos

def hacer_ping(host):
    try:
        param = "-n" if platform.system().lower() == "windows" else "-c"
        result = subprocess.run(["ping", param, "1", host], capture_output=True, timeout=3)
        return "Activo" if result.returncode == 0 else "Inactivo"
    except subprocess.TimeoutExpired:
        return "Timeout"
    except Exception:
        return "Error"

def insertar_o_actualizar(conn, equipos, equipos_ad_actuales):
    cursor = conn.cursor()
    for eq in equipos:
        ping = hacer_ping(eq["nombre"])
        estado_ad = "Dentro de AD" if eq["nombre"] in equipos_ad_actuales else "Removido de AD"

        if eq["nombre"] in estado_ping:
            anterior = estado_ping[eq["nombre"]]["estado"]
            if anterior == ping:
                estado_ping[eq["nombre"]]["contador"] += 1
            else:
                estado_ping[eq["nombre"]]["estado"] = ping
                estado_ping[eq["nombre"]]["contador"] = 1
                if ping == "Inactivo":
                    estado_ping[eq["nombre"]]["inactivo_desde"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                else:
                    estado_ping[eq["nombre"]]["inactivo_desde"] = None
        else:
            estado_ping[eq["nombre"]] = {
                "estado": ping,
                "contador": 1,
                "inactivo_desde": datetime.now().strftime("%Y-%m-%d %H:%M:%S") if ping == "Inactivo" else None
            }

        tiempo_total_segundos = estado_ping[eq["nombre"]]["contador"] * PING_INTERVAL
        horas = tiempo_total_segundos // 3600
        minutos = (tiempo_total_segundos % 3600) // 60
        segundos = tiempo_total_segundos % 60
        tiempo_formateado = f"{horas:02}:{minutos:02}:{segundos:02}"

        inactivo_desde = estado_ping[eq["nombre"]]["inactivo_desde"]

        try:
            cursor.execute("""
                MERGE EquiposAD AS target
                USING (SELECT ? AS Nombre, ? AS SO, ? AS Descripcion, ? AS IP, ? AS NombreDNS,
                              ? AS VersionSO, ? AS CreadoEl, ? AS UltimoLogon, ? AS Responsable,
                              ? AS Ubicacion, ? AS EstadoCuenta, ? AS PingStatus, ? AS TiempoPing,
                              ? AS InactivoDesde, ? AS EstadoAD) AS src
                ON target.Nombre = src.Nombre
                WHEN MATCHED THEN
                    UPDATE SET target.SO = src.SO,
                               target.Descripcion = src.Descripcion,
                               target.IP = src.IP,
                               target.NombreDNS = src.NombreDNS,
                               target.VersionSO = src.VersionSO,
                               target.CreadoEl = src.CreadoEl,
                               target.UltimoLogon = src.UltimoLogon,
                               target.Responsable = src.Responsable,
                               target.Ubicacion = src.Ubicacion,
                               target.EstadoCuenta = src.EstadoCuenta,
                               target.PingStatus = src.PingStatus,
                               target.TiempoPing = src.TiempoPing,
                               target.InactivoDesde = src.InactivoDesde,
                               target.EstadoAD = src.EstadoAD,
                               target.UltimaActualizacion = GETDATE()
                WHEN NOT MATCHED THEN
                    INSERT (Nombre, SO, Descripcion, IP, NombreDNS, VersionSO, CreadoEl,
                            UltimoLogon, Responsable, Ubicacion, EstadoCuenta, PingStatus,
                            TiempoPing, InactivoDesde, EstadoAD)
                    VALUES (src.Nombre, src.SO, src.Descripcion, src.IP, src.NombreDNS,
                            src.VersionSO, src.CreadoEl, src.UltimoLogon, src.Responsable,
                            src.Ubicacion, src.EstadoCuenta, src.PingStatus, src.TiempoPing,
                            src.InactivoDesde, src.EstadoAD);
            """, (
                eq["nombre"], eq["so"], eq["descripcion"], eq["ip"], eq["nombredns"],
                eq["versionso"], eq["creadoel"], eq["ultimologon"], eq["responsable"],
                eq["ubicacion"], eq["estadocuenta"], ping, tiempo_formateado, inactivo_desde, estado_ad
            ))
            conn.commit()
        except Exception as e:
            print(f"[ERROR] SQL al actualizar {eq['nombre']}: {e}")

        texto_fecha = f" | Inactivo desde: {inactivo_desde}" if inactivo_desde else ""
        print(f"[PING] {eq['nombre']} ({eq['ip']}) → {ping} | {estado_ad} ({tiempo_formateado}){texto_fecha}")

# ======================
# BUCLE PRINCIPAL
# ======================
def main():
    conn = conectar_sql()
    if not conn:
        return

    # Si quieres crear la tabla manualmente alguna vez:
    # crear_tabla(conn)

    while True:
        try:
            equipos = obtener_equipos_ad()
            if not equipos:
                print("[WARN] No se encontraron equipos en AD.")
                time.sleep(PING_INTERVAL)
                continue

            equipos_ad_actuales = [eq["nombre"] for eq in equipos]
            insertar_o_actualizar(conn, equipos, equipos_ad_actuales)
            print(f"[INFO] Actualización completada. Esperando {PING_INTERVAL} segundos...\n")
            time.sleep(PING_INTERVAL)
        except KeyboardInterrupt:
            print("\n[INFO] Script detenido manualmente.")
            break
        except Exception as e:
            print("[ERROR] Ocurrió un error inesperado en el bucle principal:", e)
            time.sleep(PING_INTERVAL)

if __name__ == "__main__":
    main()
