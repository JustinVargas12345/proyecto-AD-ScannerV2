import socket
import subprocess
import platform
import time
from datetime import datetime
from ldap3 import Server, Connection, ALL
from config_loader import cargar_config  # ✅ para obtener la configuración AD

# ======================
# CONFIGURACIÓN
# ======================
config = cargar_config()
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
    """
    Lee los equipos desde Active Directory y devuelve una lista de diccionarios.
    """
    equipos = []
    try:
        server = Server(AD_SERVER, get_info=ALL)
        conn = Connection(server, user=AD_USER, password=AD_PASSWORD, auto_bind=True)
        conn.search(
            AD_SEARCH_BASE,
            "(objectClass=computer)",
            attributes=[
                "name", "dNSHostName", "operatingSystem", "operatingSystemVersion",
                "description", "whenCreated", "lastLogonTimestamp", "managedBy",
                "location", "userAccountControl"
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
        print("[ERROR] Excepción al leer AD:", e)
    return equipos


def hacer_ping(host):
    """
    Realiza ping a un host y devuelve su estado.
    """
    try:
        param = "-n" if platform.system().lower() == "windows" else "-c"
        result = subprocess.run(["ping", param, "1", host], capture_output=True, timeout=6)
        return "Activo" if result.returncode == 0 else "Inactivo"
    except subprocess.TimeoutExpired:
        return "Timeout"
    except Exception:
        return "Error"


def insertar_o_actualizar(conn, equipos, equipos_ad_actuales):
    """
    Inserta o actualiza los equipos obtenidos desde AD en la base de datos SQL.
    También mantiene el seguimiento del estado de ping y del tiempo inactivo.
    """
    cursor = conn.cursor()
    for eq in equipos:
        ping = hacer_ping(eq["nombre"])
        estado_ad = "Dentro de AD" if eq["nombre"] in equipos_ad_actuales else "Removido de AD"

        # --- Seguimiento de ping ---
        if eq["nombre"] in estado_ping:
            anterior = estado_ping[eq["nombre"]]["estado"]
            if anterior == ping:
                estado_ping[eq["nombre"]]["contador"] += 1
            else:
                estado_ping[eq["nombre"]]["estado"] = ping
                estado_ping[eq["nombre"]]["contador"] = 1

            # Actualiza inactividad considerando Timeout/Error como inactivo
            if ping in ("Inactivo", "Timeout", "Error"):
                if not estado_ping[eq["nombre"]].get("inactivo_desde"):
                    estado_ping[eq["nombre"]]["inactivo_desde"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            else:
                estado_ping[eq["nombre"]]["inactivo_desde"] = None
        else:
            estado_ping[eq["nombre"]] = {
                "estado": ping,
                "contador": 1,
                "inactivo_desde": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                if ping in ("Inactivo", "Timeout", "Error") else None
            }

        # --- Tiempo acumulado activo/inactivo ---
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
