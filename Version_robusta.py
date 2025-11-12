
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
# 5️⃣ CREAR TABLA SI NO EXISTE (con EstadoAD)
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
                EstadoAD NVARCHAR(50) DEFAULT 'Dentro de AD',  -- 🔹 Nueva columna
                UltimaActualizacion DATETIME DEFAULT GETDATE()
            )
        """)
        conn.commit()
        print("[OK] Tabla 'EquiposAD' verificada o creada con columna EstadoAD.")
    except Exception as e:
        print("[ERROR] Al crear/verificar la tabla:", e)

# ==========================================================
# Diccionario global para llevar el conteo y la fecha
# ==========================================================
estado_ping = {}

# ==========================================================
# 6️⃣ INSERTAR O ACTUALIZAR DATOS
# ==========================================================
def insertar_o_actualizar(conn, equipos, equipos_ad_actuales):
    cursor = conn.cursor()
    for eq in equipos:
        ping = hacer_ping(eq["nombre"])

        # 🔹 EstadoAD: verificar si el equipo sigue en el AD actual
        estado_ad = "Dentro de AD" if eq["nombre"] in equipos_ad_actuales else "Removido de AD"

        # 🔹 Actualizar estado y contador
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

        tiempo_total_segundos = estado_ping[eq["nombre"]]["contador"] * 30
        horas = tiempo_total_segundos // 3600
        minutos = (tiempo_total_segundos % 3600) // 60
        segundos = tiempo_total_segundos % 60
        tiempo_formateado = f"{horas:02}:{minutos:02}:{segundos:02}"

        inactivo_desde = estado_ping[eq["nombre"]]["inactivo_desde"]

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

        texto_fecha = f" | Inactivo desde: {inactivo_desde}" if inactivo_desde else ""
        print(f"[PING] {eq['nombre']} ({eq['ip']}) → {ping} | {estado_ad} ({tiempo_formateado}){texto_fecha}")

# ==========================================================
# 7️⃣ PROCESO PRINCIPAL
# ==========================================================
def main():
    conn = conectar_sql()
    if not conn:
        return

    crear_tabla(conn)

    while True:
        equipos = obtener_equipos_ad()
        if not equipos:
            print("[WARN] No se encontraron equipos en AD.")
            time.sleep(30)
            continue

        equipos_ad_actuales = [eq["nombre"] for eq in equipos]  # 🔹 Lista actual de nombres en AD
        insertar_o_actualizar(conn, equipos, equipos_ad_actuales)
        print("[INFO] Actualización completada. Esperando 30 segundos...\n")
        time.sleep(30)

if __name__ == "__main__":
    main()
