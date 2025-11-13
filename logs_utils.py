import os
from datetime import datetime

LOG_FILE = "ad_scanner.log"

def escribir_log(mensaje, tipo="INFO"):
    """
    Escribe un mensaje en el log con timestamp y tipo de mensaje.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linea = f"[{timestamp}] [{tipo}] {mensaje}\n"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(linea)
