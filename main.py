'''
#import pyodbc
from ldap3 import Server, Connection, ALL
from ping3 import ping
import socket
import threading
import time
import datetime
import os 
import time
import platform
import subprocess
from ldap3 import Server, Connection, ALL
import socket
from datetime import datetime
from db_conexion import conectar_sql
from db_table import crear_tabla  # Solo si quieres crearla manualmente
import utils
from config_loader import cargar_config  # ✅ Nuevo import

# ======================
# CONFIGURACIÓN
# ======================
config = cargar_config()  # ✅ Se carga desde config_loader
PING_INTERVAL = config["PING_INTERVAL"]
AD_SERVER = config["AD_SERVER"]
AD_USER = config["AD_USER"]
AD_PASSWORD = config["AD_PASSWORD"]
AD_SEARCH_BASE = config["AD_SEARCH_BASE"]

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
    crear_tabla(conn)

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
'''

###########################################################################
'''
import time
from datetime import datetime
from db_conexion import conectar_sql#, crear_tabla
from db_table import crear_tabla
from config_loader import cargar_config
from ad_utils import obtener_equipos_ad, insertar_o_actualizar  # ✅ ahora importadas

# ======================
# CONFIGURACIÓN
# ======================
config = cargar_config()  # ✅ Se carga desde config_loader
PING_INTERVAL = config["PING_INTERVAL"]

# ======================
# BUCLE PRINCIPAL
# ======================
def main():
    conn = conectar_sql()
    if not conn:
        return

    # Si deseas crear la tabla manualmente alguna vez:
    crear_tabla(conn)

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
'''
####################################################################################


# main.py
import time
from db_conexion import conectar_sql
from db_table import crear_tabla
from ad_utils import obtener_equipos_ad, insertar_o_actualizar
from config_loader import cargar_config

config = cargar_config()
PING_INTERVAL = config["PING_INTERVAL"]

def main():
    conn = conectar_sql()
    if not conn:
        return  # Si no hay conexión, termina

    crear_tabla(conn)

    while True:
        try:
            equipos = obtener_equipos_ad()
            equipos_ad_actuales = [eq["nombre"] for eq in equipos]  # lista de nombres actuales
            insertar_o_actualizar(conn, equipos, equipos_ad_actuales)
        except Exception as e:
            print(f"[ERROR] Ocurrió un error inesperado en el bucle principal: {e}")
        except KeyboardInterrupt:
            print("\n[INFO] Script detenido manualmente.")
            break
        time.sleep(PING_INTERVAL)




if __name__ == "__main__":
        main()

    


