"""
========================================================================================
MOTOR DE CONCILIACIÓN: SAP VS. SERVICIOS DEL SAT
========================================================================================
Este script realiza la auditoría y conciliación fiscal:
1. Lee comprobantes (desde una carpeta de XMLs o desde un reporte exportado de SAP).
2. Consulta el estatus en tiempo real al Web Service oficial del SAT (multihilo para rapidez).
3. Detecta incongruencias críticas:
   - Cancelados en el SAT pero vigentes en SAP.
   - Faltantes en el SAT (No Encontrados).
   - Vigentes en el SAT pero cancelados en SAP.
4. Exporta un informe ejecutivo en Excel con semáforo de colores.
"""

import os
import glob
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Any, Optional
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

from sat_service import SATValidator
from cfdi_parser import extraer_datos_xml


def evaluar_incongruencia(estatus_sap: str, estado_sat: str, total_sap: Optional[float] = None, total_sat: Optional[float] = None) -> tuple[str, str]:
    """
    Determina el nivel de alerta y el diagnóstico comparando SAP vs. SAT.
    Retorna: (diagnostico, severidad) donde severidad es 'CRITICO', 'ADVERTENCIA', 'OK', 'ERROR'.
    """
    estatus_sap = str(estatus_sap).strip().upper() if estatus_sap else "DESCONOCIDO"
    estado_sat = str(estado_sat).strip().upper() if estado_sat else "DESCONOCIDO"

    # Caso 1: En el SAT figura como Cancelado pero en SAP está Vigente/Abierta
    if estado_sat == "CANCELADO" and estatus_sap in ["VIGENTE", "ABIERTO", "PAGADO", "ACTIVO", "CERRADO"]:
        return "[ALERTA CRITICA] Cancelado en SAT pero Vigente en SAP", "CRITICO"

    # Caso 2: El UUID no existe en los registros del SAT
    if estado_sat in ["NO ENCONTRADO", "NO EXISTE"]:
        return "[ALERTA CRITICA] Comprobante no encontrado en el SAT (No timbrado o UUID erroneo)", "CRITICO"

    # Caso 3: Cancelado en SAP pero ante el SAT sigue Vigente
    if estado_sat == "VIGENTE" and estatus_sap in ["CANCELADO", "ANULADO", "INACTIVO"]:
        return "[ADVERTENCIA] Cancelado en SAP pero sigue Vigente en SAT (Falta cancelar)", "ADVERTENCIA"

    # Caso 4: Error de comunicación
    if "ERROR" in estado_sat or "TIMEOUT" in estado_sat:
        return "[ERROR] No se pudo verificar en el SAT", "ERROR"

    # Caso 5: Diferencia de importes
    if total_sap is not None and total_sat is not None:
        if abs(float(total_sap) - float(total_sat)) > 0.05:
            return f"[ADVERTENCIA] Diferencia de importe (SAP: ${total_sap:,.2f} vs SAT: ${total_sat:,.2f})", "ADVERTENCIA"

    # Si ambos son vigentes
    if estado_sat == "VIGENTE":
        return "[OK] Vigente en ambos sistemas", "OK"

    # Si ambos son cancelados
    if estado_sat == "CANCELADO" and estatus_sap in ["CANCELADO", "ANULADO"]:
        return "[OK] Cancelado en ambos sistemas", "OK"

    return f"[INFO] SAT: {estado_sat} | SAP: {estatus_sap}", "INFO"


def validar_lote_cfdi(items: List[Dict[str, Any]], max_workers: int = 8) -> List[Dict[str, Any]]:
    """
    Consulta en paralelo múltiples comprobantes contra el SAT para máxima velocidad.
    """
    total_docs = len(items)
    print(f"\n[+] Iniciando consulta de {total_docs} documentos en los servidores del SAT (Conexión multihilo)...")

    resultados = []

    def _consultar_uno(doc):
        resp = SATValidator.consultar_estatus(
            rfc_emisor=doc.get("rfc_emisor", ""),
            rfc_receptor=doc.get("rfc_receptor", ""),
            total=doc.get("total", 0.0),
            uuid=doc.get("uuid", ""),
            sello=doc.get("sello")
        )
        
        # Evaluar incongruencia
        diag, severidad = evaluar_incongruencia(
            estatus_sap=doc.get("estatus_sap", "VIGENTE"),
            estado_sat=resp["estado"],
            total_sap=doc.get("total_sap", doc.get("total")),
            total_sat=doc.get("total")
        )

        doc_resultado = {**doc}
        doc_resultado["sat_estado"] = resp["estado"]
        doc_resultado["sat_es_cancelable"] = resp["es_cancelable"]
        doc_resultado["sat_estatus_cancelacion"] = resp["estatus_cancelacion"]
        doc_resultado["sat_validacion_efos"] = resp["validacion_efos"]
        doc_resultado["sat_codigo_respuesta"] = resp["codigo_estatus"]
        doc_resultado["diagnostico"] = diag
        doc_resultado["severidad"] = severidad
        return doc_resultado

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_consultar_uno, doc): doc for doc in items}
        completados = 0
        for future in as_completed(futures):
            completados += 1
            if completados % 10 == 0 or completados == total_docs:
                print(f"    Progreso: {completados}/{total_docs} comprobantes consultados...")
            try:
                res = future.result()
                resultados.append(res)
            except Exception as e:
                print(f"    Error procesando un documento: {e}")

    return resultados


def exportar_reporte_excel(resultados: List[Dict[str, Any]], ruta_salida: str = "Reporte_Conciliacion_SAT_SAP.xlsx"):
    """
    Genera un archivo Excel con formato profesional y semáforo visual de auditoría.
    """
    if not resultados:
        print("[!] No hay datos para exportar.")
        return

    df = pd.DataFrame(resultados)

    # Ordenar columnas deseadas
    columnas_orden = [
        "severidad", "diagnostico", "folio_completo", "tipo_comprobante",
        "fecha", "rfc_receptor", "nombre_receptor", "total", "moneda",
        "sat_estado", "estatus_sap", "uuid", "sat_es_cancelable",
        "sat_validacion_efos", "sat_codigo_respuesta"
    ]
    cols_existentes = [c for c in columnas_orden if c in df.columns]
    otros_cols = [c for c in df.columns if c not in cols_existentes and c not in ["sello", "ruta_completa"]]
    df = df[cols_existentes + otros_cols]

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Conciliacion_SAT_SAP")
        workbook = writer.book
        worksheet = writer.sheets["Conciliacion_SAT_SAP"]

        # Estilos
        fill_critico = PatternFill(start_color="FFD2D2", end_color="FFD2D2", fill_type="solid") # Rojo claro
        font_critico = Font(color="9C0006", bold=True)
        fill_adver = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")   # Amarillo claro
        font_adver = Font(color="9C6500", bold=True)
        fill_ok = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")      # Verde claro
        font_ok = Font(color="155724")

        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid") # Azul corporativo
        header_font = Font(color="FFFFFF", bold=True)

        for col_idx, col in enumerate(worksheet.columns, start=1):
            # Formatear encabezado
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

            # Auto-ancho de columna
            max_len = max(len(str(c.value or '')) for c in col)
            col_letter = cell.column_letter
            worksheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

        # Aplicar colores de semáforo por fila según la severidad
        sev_col_idx = 1 # severidad es la primera columna
        diag_col_idx = 2

        for row_idx in range(2, worksheet.max_row + 1):
            sev_val = worksheet.cell(row=row_idx, column=sev_col_idx).value
            cell_diag = worksheet.cell(row=row_idx, column=diag_col_idx)

            if sev_val == "CRITICO":
                cell_diag.fill = fill_critico
                cell_diag.font = font_critico
            elif sev_val == "ADVERTENCIA":
                cell_diag.fill = fill_adver
                cell_diag.font = font_adver
            elif sev_val == "OK":
                cell_diag.fill = fill_ok
                cell_diag.font = font_ok

    print(f"\n[✓] Reporte generado exitosamente: {os.path.abspath(ruta_salida)}")


def procesar_carpeta_xml(carpeta_xml: str, limite: Optional[int] = None) -> List[Dict[str, Any]]:
    """
    Lee archivos XML directamente de una carpeta local o de red y los prepara para validación.
    """
    patron = os.path.join(carpeta_xml, "*.xml")
    archivos = glob.glob(patron)
    if limite:
        archivos = archivos[:limite]

    print(f"[+] Leyendo {len(archivos)} archivos XML desde {carpeta_xml}...")
    documentos = []
    for f in archivos:
        datos = extraer_datos_xml(f)
        if datos:
            # Por defecto asumimos que si el XML está en creadas de SAP, en SAP está Vigente
            datos["estatus_sap"] = "VIGENTE"
            datos["total_sap"] = datos["total"]
            documentos.append(datos)

    return documentos
