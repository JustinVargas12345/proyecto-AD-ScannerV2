'''
from db_conexion import conectar_sql

def crear_tabla(conn):
    """
    Crea la tabla EquiposAD si no existe.
    """
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
                EstadoAD NVARCHAR(50) DEFAULT 'Dentro de AD',
                UltimaActualizacion DATETIME DEFAULT GETDATE()
            )
        """)
        conn.commit()
        print("[OK] Tabla 'EquiposAD' verificada o creada.")
    except Exception as e:
        print("[ERROR] Al crear/verificar la tabla:", e)
'''

# db_table.py
from db_conexion import conectar_sql, ejecutar_sql

def crear_tabla(conn):
    """
    Crea la tabla EquiposAD si no existe, con reconexión automática.
    """
    query = """
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
            EstadoAD NVARCHAR(50) DEFAULT 'Dentro de AD',
            UltimaActualizacion DATETIME DEFAULT GETDATE()
        )
    """
    if ejecutar_sql(conn, query):
        print("[OK] Tabla 'EquiposAD' verificada o creada.")
    else:
        print("[ERROR] No se pudo crear/verificar la tabla.")
