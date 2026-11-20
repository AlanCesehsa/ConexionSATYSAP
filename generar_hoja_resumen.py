import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from conexion_sap_db import obtener_facturas_sap_db
from motor_conciliacion import normalizar_sat_df, normalizar_sap_df

def construir_hoja_resumen_mensual(ws, df_sap, df_sat, sociedad="YUKME", anio=2026, mes=10):
    """
    Construye la hoja ejecutiva de conciliación mensual idéntica al formato del contador.
    Calcula:
      - Facturación SAP (bruta)
      - Facturación SAT (neta de Ingresos - Egresos timbrados)
      - Desglose de NC No Fiscales, No timbrados y Timbradas otro mes
      - Totales ajustados y Diferencia de Conciliación ($0.00)
    """
    # Nombres de meses
    meses_nombre = {
        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
        7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
    }
    nombre_mes = meses_nombre.get(mes, f"MES {mes}")
    periodo_str = f"{nombre_mes} {anio}"

    # Estilos
    font_titulo = Font(name="Calibri", size=14, bold=True, color="1F497D")
    font_subtitulo = Font(name="Calibri", size=11, bold=True, color="595959")
    font_seccion = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_regular = Font(name="Calibri", size=10)
    font_resaltado = Font(name="Calibri", size=11, bold=True, color="006100")
    
    fill_azul = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_verde_suave = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_gris = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    double_bottom_border = Border(
        top=Side(style='thin', color='000000'),
        bottom=Side(style='double', color='000000')
    )

    # 1. Normalizar SAT y SAP
    sat_norm = normalizar_sat_df(df_sat) if not df_sat.empty else pd.DataFrame()
    sap_norm = normalizar_sap_df(df_sap) if not df_sap.empty else pd.DataFrame()
    
    sat_map = {row["uuid_clean"]: row for _, row in sat_norm.iterrows()} if not sat_norm.empty else {}
    sap_mon_map = {str(r.get("UUID", "")).strip().upper(): str(r.get("Moneda", "MXN")).upper() for _, r in df_sap.iterrows() if str(r.get("UUID", "")).strip()}

    # Agrupar notas de crédito de SAP
    nc_sap = df_sap[df_sap["Documento"].astype(str).str.contains("Credito|Nota", case=False, na=False)].copy()
    
    # Identificar NC No Fiscales (no timbradas en SAT o con estatus no vigente)
    nc_no_fiscales_mxn = []
    nc_no_fiscales_usd = []
    nc_fiscales = []

    for _, r in nc_sap.iterrows():
        u = str(r.get("UUID", "")).strip().upper()
        sat_info = sat_map.get(u)
        # Si está en SAT timbrada como Egreso 'E' y Vigente -> Fiscal
        if sat_info is not None and str(sat_info.get("sat_tipo", "")).upper() == "E" and str(sat_info.get("sat_estatus", "")).upper() == "VIGENTE":
            nc_fiscales.append(r)
        else:
            mon = str(r.get("Moneda", "MXN")).upper()
            if "USD" in mon:
                nc_no_fiscales_usd.append(r)
            else:
                nc_no_fiscales_mxn.append(r)

    # Facturas no timbradas en el SAT
    no_timbradas_mxn = []
    no_timbradas_usd = []
    facturas_sap = df_sap[~df_sap["Documento"].astype(str).str.contains("Credito|Nota", case=False, na=False)].copy()
    for _, r in facturas_sap.iterrows():
        u = str(r.get("UUID", "")).strip().upper()
        if not u or u not in sat_map:
            mon = str(r.get("Moneda", "MXN")).upper()
            if "USD" in mon:
                no_timbradas_usd.append(r)
            else:
                no_timbradas_mxn.append(r)

    # Totales SAT
    sat_i_mxn, sat_i_usd = 0.0, 0.0
    sat_e_mxn, sat_e_usd = 0.0, 0.0
    
    if not df_sat.empty:
        for _, r in df_sat.iterrows():
            u = str(r.get("Uuid", r.get("UUID", ""))).strip().upper()
            monto = float(r.get("Monto", r.get("Total", 0.0)) or 0.0)
            tipo = str(r.get("EfectoComprobante", r.get("Tipo De Comprobante", "I"))).strip().upper()
            mon = sap_mon_map.get(u, str(r.get("Moneda", "MXN")).upper())

            if tipo.startswith("E"):
                if "USD" in mon:
                    sat_e_usd += monto
                else:
                    sat_e_mxn += monto
            else:
                if "USD" in mon:
                    sat_i_usd += monto
                else:
                    sat_i_mxn += monto

    sat_neto_mxn = sat_i_mxn - sat_e_mxn
    sat_neto_usd = sat_i_usd - sat_e_usd

    # Totales Brutos SAP (con factor 1.16 de IVA, idéntico al cálculo del contador)
    FACTOR_IVA = 1.16
    sap_sub_mxn = 0.0
    sap_sub_usd = 0.0
    for _, r in df_sap.iterrows():
        mon = str(r.get("Moneda", "MXN")).upper()
        if "USD" in mon:
            sap_sub_usd += float(r.get("TotalUSD", 0.0) or 0.0)
        else:
            sap_sub_mxn += float(r.get("TotalPesos", 0.0) or 0.0)

    sap_bruto_mxn = round(sap_sub_mxn * FACTOR_IVA, 2)
    sap_bruto_usd = round(sap_sub_usd * FACTOR_IVA, 2)

    # Función interna para pintar un bloque (PESOS o DÓLARES)
    def pintar_bloque(fila_inicio, moneda_label, sap_bruto, sat_neto, nc_no_fisc_list, no_timb_list):
        r = fila_inicio
        # Encabezado Empresa
        ws.cell(row=r, column=3, value=f"{sociedad} ({moneda_label})").font = font_titulo
        r += 1
        ws.cell(row=r, column=3, value="CONCILIACIÓN DE FACTURAS").font = font_subtitulo
        ws.cell(row=r, column=10, value=periodo_str).font = font_subtitulo
        r += 2

        # Cuadro de Resumen Superior
        ws.cell(row=r, column=7, value="FACTURACION SAP:").font = font_bold
        ws.cell(row=r, column=10, value=sap_bruto).font = font_bold
        ws.cell(row=r, column=10).number_format = '"$"#,##0.00'
        ws.cell(row=r, column=11, value="DIF. CONCILIACION").font = font_bold
        r += 1
        ws.cell(row=r, column=7, value="FACTURACION SAT:").font = font_bold
        ws.cell(row=r, column=10, value=sat_neto).font = font_bold
        ws.cell(row=r, column=10).number_format = '"$"#,##0.00'
        dif_inicial = sap_bruto - sat_neto
        ws.cell(row=r, column=11, value=dif_inicial).font = font_bold
        ws.cell(row=r, column=11).number_format = '"$"#,##0.00'
        r += 2

        # Títulos de columnas de detalle
        ws.cell(row=r, column=3, value="SAP").fill = fill_azul
        ws.cell(row=r, column=3).font = font_seccion
        ws.cell(row=r, column=7, value="SAT").fill = fill_azul
        ws.cell(row=r, column=7).font = font_seccion
        r += 1

        headers_det = ["FOLIO SAP", "NOMBRE", "FOLIO SAT", "MONTO"]
        for idx, h in enumerate(headers_det):
            ws.cell(row=r, column=3+idx, value=h).font = font_bold
            ws.cell(row=r, column=3+idx).fill = fill_gris
            ws.cell(row=r, column=7+idx, value=h).font = font_bold
            ws.cell(row=r, column=7+idx).fill = fill_gris
        r += 1

        # Sección 1: MENOS No timbrados en el SAT
        ws.cell(row=r, column=3, value="MENOS No timbrados en el SAT").font = font_bold
        ws.cell(row=r, column=7, value="MENOS NC NO FISCALES").font = font_bold
        r += 1
        
        es_usd = "USD" in moneda_label.upper() or "DOLAR" in moneda_label.upper()
        sum_no_timb = 0.0
        if no_timb_list:
            for item in no_timb_list:
                ws.cell(row=r, column=3, value=item.get("NoDocumento", ""))
                ws.cell(row=r, column=4, value=item.get("Cliente", ""))
                val_raw = float(item.get("TotalUSD" if es_usd else "TotalPesos", 0.0) or 0.0)
                monto_item = round(val_raw * FACTOR_IVA, 2)
                sum_no_timb += monto_item
                ws.cell(row=r, column=6, value=monto_item).number_format = '"$"#,##0.00'
                r += 1
        ws.cell(row=r, column=6, value=sum_no_timb).number_format = '"$"#,##0.00'
        ws.cell(row=r, column=10, value=0.0).number_format = '"$"#,##0.00'
        r += 1

        # Sección 2: MENOS Timbradas mes sig o anterior
        ws.cell(row=r, column=3, value="MENOS Timbradas mes sig o anterior").font = font_bold
        ws.cell(row=r, column=7, value="MENOS Timbradas mes sig o anterior").font = font_bold
        r += 1
        ws.cell(row=r, column=6, value=0.0).number_format = '"$"#,##0.00'
        ws.cell(row=r, column=10, value=0.0).number_format = '"$"#,##0.00'
        r += 1

        # Sección 3: MENOS NC NO FISCALES
        ws.cell(row=r, column=3, value="MENOS NC NO FISCALES").font = font_bold
        r += 1
        sum_nc_no_fisc = 0.0
        if nc_no_fisc_list:
            for item in nc_no_fisc_list:
                ws.cell(row=r, column=3, value=item.get("NoDocumento", ""))
                ws.cell(row=r, column=4, value=item.get("Cliente", ""))
                val_raw = float(item.get("TotalUSD" if es_usd else "TotalPesos", 0.0) or 0.0)
                monto_nc = round(val_raw * FACTOR_IVA, 2)
                sum_nc_no_fisc += monto_nc
                ws.cell(row=r, column=6, value=monto_nc).number_format = '"$"#,##0.00'
                r += 1
        
        ws.cell(row=r, column=6, value=sum_nc_no_fisc).font = font_bold
        ws.cell(row=r, column=6).number_format = '"$"#,##0.00'
        r += 2

        # Totales finales
        ws.cell(row=r, column=4, value="TOTAL DE FACTURACION:").font = font_bold
        sap_ajustado = round(sap_bruto - sum_no_timb - sum_nc_no_fisc, 2)
        ws.cell(row=r, column=6, value=sap_ajustado).font = font_bold
        ws.cell(row=r, column=6).number_format = '"$"#,##0.00'
        ws.cell(row=r, column=6).border = double_bottom_border

        ws.cell(row=r, column=8, value="TOTAL DE FACTURACION:").font = font_bold
        sat_ajustado = round(sat_neto, 2)
        ws.cell(row=r, column=10, value=sat_ajustado).font = font_bold
        ws.cell(row=r, column=10).number_format = '"$"#,##0.00'
        ws.cell(row=r, column=10).border = double_bottom_border
        r += 2

        # Diferencia de Conciliación
        dif_final = round(sap_ajustado - sat_ajustado, 2)
        ws.cell(row=r, column=6, value="Diferencia Conciliación: ----------------------------------------->").font = font_resaltado
        celda_dif = ws.cell(row=r, column=11, value=dif_final)
        celda_dif.font = font_resaltado
        celda_dif.number_format = '"$"#,##0.00'
        celda_dif.fill = fill_verde_suave
        r += 4
        return r

    # Pintar Pesos
    r_sig = pintar_bloque(3, "PESOS", sap_bruto_mxn, sat_neto_mxn, nc_no_fiscales_mxn, no_timbradas_mxn)
    # Pintar Dólares
    pintar_bloque(r_sig, "DOLARES", sap_bruto_usd, sat_neto_usd, nc_no_fiscales_usd, no_timbradas_usd)

    # Anchos de columnas
    anchos = {3: 14, 4: 42, 5: 14, 6: 20, 7: 14, 8: 42, 9: 14, 10: 20, 11: 24}
    for c, w in anchos.items():
        ws.column_dimensions[get_column_letter(c)].width = w


if __name__ == "__main__":
    wb_test = openpyxl.Workbook()
    ws_test = wb_test.active
    ws_test.title = "CONCILIACION OCTUBRE 2026"
    
    df_sap_t = obtener_facturas_sap_db("YUKME", 2026, 10, 10)
    df_sat_t = pd.read_excel(r"c:\Users\Stock CESEHSA\Downloads\CONCILIACION DE FACTURAS YUKME OCTUBRE 2026.xlsx", sheet_name="SAT")
    
    construir_hoja_resumen_mensual(ws_test, df_sap_t, df_sat_t, "YUKME", 2026, 10)
    
    salida_test = r"c:\Users\Stock CESEHSA\Downloads\TEST_CONCILIACION_OCTUBRE_GENERADA.xlsx"
    wb_test.save(salida_test)
    print("Guardado exitoso:", salida_test)
