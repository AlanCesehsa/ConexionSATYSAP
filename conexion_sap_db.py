"""
========================================================================================
CONEXIÓN DIRECTA A BASE DE DATOS SAP BUSINESS ONE (SQL SERVER)
========================================================================================
Servidor: 192.168.1.250
Base de datos: REJ
Usuario: TESCH
Tablas/Procedimientos por Sociedad:
- CESEHSA -> ReporteSAPvsSAT
- CSH     -> ReporteSAPvsSATCSH
- YUKME   -> ReporteSAPvsSATYuk
========================================================================================
"""

import pyodbc
import pandas as pd
import warnings
from typing import Optional, Dict

# Silenciar advertencia informativa de pandas al usar conexión directa pyodbc
warnings.filterwarnings('ignore', category=UserWarning)

# Configuración del servidor SQL de SAP
SAP_DB_CONFIG = {
    "server": "192.168.1.250",
    "database": "REJ",
    "username": "TESCH",
    "password": "B1Admin"
}

TABLAS_POR_SOCIEDAD = {
    "CESEHSA": "ReporteSAPvsSAT",
    "CSH": "ReporteSAPvsSATCSH",
    "YUKME": "ReporteSAPvsSATYuk"
}

SAP_B1_DBS = {
    "CESEHSA": "CESEHSA_PRODUCTS",
    "CSH": "CSH_PRODUCTS",
    "YUKME": "YUKMEB"
}


def probar_conexion_sap() -> bool:
    """Verifica si el servidor de SAP 192.168.1.250 es alcanzable."""
    try:
        conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={SAP_DB_CONFIG['server']};"
            f"DATABASE={SAP_DB_CONFIG['database']};"
            f"UID={SAP_DB_CONFIG['username']};"
            f"PWD={SAP_DB_CONFIG['password']};"
            f"TrustServerCertificate=yes;"
            f"Connection Timeout=3"
        )
        conn = pyodbc.connect(conn_str)
        conn.close()
        return True
    except Exception:
        try:
            # Reintento con driver estándar de SQL Server
            conn_str = (
                f"DRIVER={{SQL Server}};"
                f"SERVER={SAP_DB_CONFIG['server']};"
                f"DATABASE={SAP_DB_CONFIG['database']};"
                f"UID={SAP_DB_CONFIG['username']};"
                f"PWD={SAP_DB_CONFIG['password']};"
                f"Connection Timeout=3"
            )
            conn = pyodbc.connect(conn_str)
            conn.close()
            return True
        except Exception:
            return False


def obtener_facturas_sap_db(sociedad: str, anio: int, mes_ini: int, mes_fin: Optional[int] = None) -> pd.DataFrame:
    """
    Obtiene las facturas emitidas de SAP directamente de SQL Server
    filtradas por Sociedad, Año y rango de Meses (desde mes_ini hasta mes_fin).
    """
    if mes_fin is None:
        mes_fin = mes_ini

    soc = sociedad.upper().strip()
    if soc not in TABLAS_POR_SOCIEDAD:
        raise ValueError(f"Sociedad '{sociedad}' no reconocida. Opciones: {list(TABLAS_POR_SOCIEDAD.keys())}")

    tabla = TABLAS_POR_SOCIEDAD[soc]

    # Intentar con ODBC Driver 17 o SQL Server
    conn = None
    for drv in ["ODBC Driver 17 for SQL Server", "SQL Server"]:
        try:
            conn_str = (
                f"DRIVER={{{drv}}};"
                f"SERVER={SAP_DB_CONFIG['server']};"
                f"DATABASE={SAP_DB_CONFIG['database']};"
                f"UID={SAP_DB_CONFIG['username']};"
                f"PWD={SAP_DB_CONFIG['password']};"
                f"TrustServerCertificate=yes;"
                f"Connection Timeout=5"
            )
            conn = pyodbc.connect(conn_str)
            break
        except Exception:
            continue

    if not conn:
        raise ConnectionError(
            f"No se pudo conectar al servidor de SAP ({SAP_DB_CONFIG['server']}). "
            f"Verifica estar conectado a la red local de la empresa."
        )

    try:
        # Consulta filtrando por fecha real del documento (FechaDocumento) para incluir
        # todas las facturas y notas de crédito emitidas en el periodo contable
        if mes_ini == mes_fin:
            query = f"SELECT * FROM {tabla} WHERE YEAR(FechaDocumento) = ? AND MONTH(FechaDocumento) = ?"
            params = [anio, mes_ini]
        else:
            query = f"SELECT * FROM {tabla} WHERE YEAR(FechaDocumento) = ? AND MONTH(FechaDocumento) >= ? AND MONTH(FechaDocumento) <= ?"
            params = [anio, mes_ini, mes_fin]

        cursor = conn.cursor()
        cursor.execute(query, params)
        cols = [column[0] for column in cursor.description]
        rows = cursor.fetchall()
        # Convertir a lista de tuplas/valores normales para DataFrame
        data = [tuple(row) for row in rows]
        df = pd.DataFrame(data, columns=cols)
        cursor.close()

        # Enriquecer con datos nativos de SAP Business One (Xml, UUI DRelacionado, Domicilio Fiscal Receptor, No Identificacion)
        if not df.empty and soc in SAP_B1_DBS and "UUID" in df.columns:
            try:
                db_b1 = SAP_B1_DBS[soc]
                uuids_validos = [str(u).strip().upper() for u in df["UUID"].dropna() if str(u).strip() != ""]
                if uuids_validos:
                    conn_b1 = None
                    for drv in ["ODBC Driver 17 for SQL Server", "SQL Server"]:
                        try:
                            c_str = (
                                f"DRIVER={{{drv}}};"
                                f"SERVER={SAP_DB_CONFIG['server']};"
                                f"DATABASE={db_b1};"
                                f"UID={SAP_DB_CONFIG['username']};"
                                f"PWD={SAP_DB_CONFIG['password']};"
                                f"TrustServerCertificate=yes;"
                                f"Connection Timeout=5"
                            )
                            conn_b1 = pyodbc.connect(c_str)
                            break
                        except Exception:
                            continue

                    if conn_b1:
                        cur_b1 = conn_b1.cursor()
                        placeholders = ",".join(["?"] * len(uuids_validos))
                        q_b1 = f"""
                        SELECT 
                            UPPER(CAST(T0.U_UUID AS VARCHAR(100))) AS [UUID_B1],
                            COALESCE(T0.U_Xml, '') AS [Xml],
                            CASE WHEN T0.U_CFDIRel = 'NA' THEN '' ELSE COALESCE(T0.U_CFDIRel, '') END AS [UUI DRelacionado],
                            COALESCE(T12.ZipCodeB, T12.ZipCodeS, '') AS [Domicilio Fiscal Receptor],
                            COALESCE((SELECT TOP 1 CAST(T1.ItemCode AS VARCHAR(100)) FROM INV1 T1 WHERE T1.DocEntry = T0.DocEntry ORDER BY T1.LineNum), '') AS [No Identificacion]
                        FROM OINV T0
                        LEFT JOIN INV12 T12 ON T12.DocEntry = T0.DocEntry
                        WHERE UPPER(CAST(T0.U_UUID AS VARCHAR(100))) IN ({placeholders})
                        """
                        cur_b1.execute(q_b1, uuids_validos)
                        b1_rows = cur_b1.fetchall()
                        b1_cols = [col[0] for col in cur_b1.description]
                        df_b1 = pd.DataFrame([tuple(r) for r in b1_rows], columns=b1_cols)
                        cur_b1.close()
                        conn_b1.close()

                        if not df_b1.empty:
                            df_b1["UUID_CLEAN"] = df_b1["UUID_B1"].astype(str).str.strip().str.upper()
                            df["UUID_CLEAN"] = df["UUID"].astype(str).str.strip().str.upper()
                            df = df.merge(
                                df_b1[["UUID_CLEAN", "Xml", "UUI DRelacionado", "Domicilio Fiscal Receptor", "No Identificacion"]],
                                on="UUID_CLEAN",
                                how="left"
                            ).drop(columns=["UUID_CLEAN"])
            except Exception as e_b1:
                # Silencioso para no interrumpir el flujo si una tabla auxiliar no responde
                pass

        # Asegurar que las columnas enriquecidas siempre existan en el DataFrame
        for col_extra in ["Xml", "UUI DRelacionado", "Domicilio Fiscal Receptor", "No Identificacion"]:
            if col_extra not in df.columns:
                df[col_extra] = ""

        return df
    finally:
        conn.close()

