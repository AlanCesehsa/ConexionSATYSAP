"""
========================================================================================
HERRAMIENTA DEDICADA: DESCARGA ANUAL DE COMPROBANTES DEL SAT
========================================================================================
Descarga todos los comprobantes emitidos (Facturas y Notas de Crédito) de un AÑO completo
directamente desde los servidores centrales del SAT con la e.firma de la empresa.

Uso:
  python descargar_sat.py --sociedad YUKME --anio 2024
  python descargar_sat.py --sociedad CESEHSA --anio 2024
  python descargar_sat.py --sociedad CSH --anio 2024

O interactivo (sin argumentos):
  python descargar_sat.py
========================================================================================
"""

import os
import sys
import argparse
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# Asegurar codificación UTF-8 en consola de Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from sat_descarga_masiva import SATDescargaMasiva


def formatear_y_guardar_excel(df: pd.DataFrame, ruta_salida: str, sociedad: str, anio: int):
    """Guarda los metadatos descargados en un Excel estilizado y con resumen por mes."""
    if df.empty:
        print("[!] No hay datos para guardar.")
        return

    # Normalizar columnas
    df = df.copy()
    
    # Renombrar columnas para que sean 100% amigables y legibles
    mapeo_columnas = {
        "Uuid": "UUID",
        "RfcEmisor": "RFC Emisor",
        "NombreEmisor": "Nombre Emisor",
        "RfcReceptor": "RFC Receptor",
        "NombreReceptor": "Nombre Cliente / Receptor",
        "PacCertifico": "PAC que Certificó",
        "FechaEmision": "Fecha Emisión",
        "FechaCertificacionSat": "Fecha Certificación SAT",
        "Monto": "Total",
        "EfectoComprobante": "Tipo (Efecto)",
        "Estatus": "Estatus SAT",
        "FechaCancelacion": "Fecha Cancelación"
    }
    df.rename(columns=mapeo_columnas, inplace=True)

    # Traducir Estatus SAT: 1 -> Vigente, 0 -> Cancelado
    if "Estatus SAT" in df.columns:
        df["Estatus SAT"] = df["Estatus SAT"].replace({"1": "Vigente", "0": "Cancelado", 1: "Vigente", 0: "Cancelado"})

    # Traducir Tipo: I -> Ingreso (Factura), E -> Egreso (Nota de Crédito), P -> Pago
    if "Tipo (Efecto)" in df.columns:
        tipos_map = {"I": "I - Factura (Ingreso)", "E": "E - Nota de Crédito (Egreso)", "P": "P - Pago", "N": "N - Nómina"}
        df["Tipo (Efecto)"] = df["Tipo (Efecto)"].replace(tipos_map)

    # Formatear montos numéricos
    if "Total" in df.columns:
        df["Total"] = pd.to_numeric(df["Total"].astype(str).str.replace("$", "").str.replace(",", ""), errors="coerce").fillna(0.0)

    # Extraer Mes para análisis
    if "Fecha Emisión" in df.columns:
        fechas = pd.to_datetime(df["Fecha Emisión"], errors="coerce")
        df["Mes"] = fechas.dt.month
        df["Mes_Nombre"] = fechas.dt.strftime("%B")

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        # Pestaña 1: Resumen Anual por Mes
        meses_nombres = {
            1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
            5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
            9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
        }
        resumen_mes = []
        for m in range(1, 13):
            sub_m = df[df["Mes"] == m]
            cant_vig = len(sub_m[sub_m["Estatus SAT"] == "Vigente"])
            cant_canc = len(sub_m[sub_m["Estatus SAT"] == "Cancelado"])
            tot_facturado = sub_m[sub_m["Estatus SAT"] == "Vigente"]["Total"].sum()
            resumen_mes.append({
                "Mes": f"{m:02d} - {meses_nombres[m]}",
                "Total Facturas": len(sub_m),
                "Vigentes": cant_vig,
                "Canceladas": cant_canc,
                "Importe Total Vigente (SAT)": tot_facturado
            })

        df_resumen = pd.DataFrame(resumen_mes)
        df_resumen.to_excel(writer, sheet_name="RESUMEN_ANUAL", index=False)

        # Pestaña 2: Listado Completo de Facturas
        cols_orden = [
            "UUID", "Estatus SAT", "Fecha Emisión", "Fecha Cancelación",
            "RFC Receptor", "Nombre Cliente / Receptor", "Total",
            "Tipo (Efecto)", "PAC que Certificó"
        ]
        cols_presentes = [c for c in cols_orden if c in df.columns]
        otros = [c for c in df.columns if c not in cols_presentes and c not in ["Mes", "Mes_Nombre"]]
        df[cols_presentes + otros].to_excel(writer, sheet_name="COMPROBANTES_SAT", index=False)

        # Pestaña 3: Solo Canceladas (Para revisión rápida)
        df_canc = df[df["Estatus SAT"] == "Cancelado"]
        if not df_canc.empty:
            df_canc[cols_presentes + otros].to_excel(writer, sheet_name="SOLO_CANCELADAS", index=False)

        # Estilizar
        wb = writer.book
        fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        font_header = Font(color="FFFFFF", bold=True)
        fill_canc = PatternFill(start_color="FFD2D2", end_color="FFD2D2", fill_type="solid")
        font_canc = Font(color="9C0006", bold=True)

        for sheet in wb.worksheets:
            for col_idx in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=1, column=col_idx)
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = Alignment(horizontal="center", vertical="center")

                # Auto ancho
                col_letter = get_column_letter(col_idx)
                max_len = max(len(str(sheet.cell(row=r, column=col_idx).value or "")) for r in range(1, sheet.max_row + 1))
                sheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)

            # Si es la hoja de comprobantes, colorear los cancelados
            if sheet.title in ["COMPROBANTES_SAT", "SOLO_CANCELADAS"]:
                for r in range(2, sheet.max_row + 1):
                    val_estatus = str(sheet.cell(row=r, column=2).value or "")
                    if val_estatus == "Cancelado":
                        sheet.cell(row=r, column=2).fill = fill_canc
                        sheet.cell(row=r, column=2).font = font_canc

    print(f"\n[OK] Archivo Excel Anual generado exitosamente:")
    print(f"    {os.path.abspath(ruta_salida)}")


def descargar_anio_completo(sociedad: str, anio: int, carpeta_salida: Optional[str] = None, abrir_excel: bool = True):
    """Descarga todos los comprobantes emitidos del SAT para un año completo."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    cert_dir = os.path.join(base_dir, "certificados")

    if not carpeta_salida:
        carpeta_salida = os.path.join(base_dir, "reportes")
    os.makedirs(carpeta_salida, exist_ok=True)

    print("\n" + "="*75)
    print(f"  DESCARGA OFICIAL SAT CON E.FIRMA — AÑO {anio}")
    print(f"  Sociedad: {sociedad}")
    print(f"  Carpeta destino: {carpeta_salida}")
    print("="*75)

    sat = SATDescargaMasiva(sociedad, base_cert_dir=cert_dir)
    sat.autenticar()

    ahora = datetime.datetime.now()
    fecha_ini = datetime.datetime(anio, 1, 1, 0, 0, 0)
    fecha_fin_teorica = datetime.datetime(anio, 12, 31, 23, 59, 59)

    if fecha_ini > ahora:
        print(f"\n[ERROR] El año {anio} está completamente en el futuro.")
        return None

    # El SAT rechaza con error 301 'Fecha final invalida' si la fecha excede el momento actual
    fecha_fin = min(fecha_fin_teorica, ahora - datetime.timedelta(minutes=1))

    print(f"\n[*] Solicitando metadatos al SAT del {fecha_ini.strftime('%Y-%m-%d')} al {fecha_fin.strftime('%Y-%m-%d %H:%M:%S')}...")
    id_sol, res = sat.solicitar_descarga_metadatos(fecha_ini, fecha_fin)

    if not id_sol:
        print(f"\n[ERROR] El SAT no generó ID de solicitud: {res.get('mensaje')}")
        return None

    print(f"[*] Esperando procesamiento del SAT (ID: {id_sol})...")
    df_sat = sat.verificar_y_descargar(id_sol, max_intentos=20, espera_segundos=4)

    if df_sat.empty:
        print("[!] El SAT no devolvió registros para este año.")
        return None

    total_cfdis = len(df_sat)
    total_canc = len(df_sat[df_sat["Estatus"] == "0"])
    total_vig = len(df_sat[df_sat["Estatus"] == "1"])

    print("\n" + "="*75)
    print(f"RESULTADO DE LA DESCARGA SAT — {sociedad} {anio}:")
    print(f"  • Total Comprobantes Oficiales : {total_cfdis}")
    print(f"  • Comprobantes Vigentes        : {total_vig}")
    print(f"  • Comprobantes Cancelados      : {total_canc}")
    print("="*75)

    # Guardar en Excel dentro de la carpeta elegida
    ruta_salida = os.path.join(carpeta_salida, f"SAT_{sociedad}_{anio}_COMPLETO.xlsx")
    formatear_y_guardar_excel(df_sat, ruta_salida, sociedad, anio)

    if abrir_excel:
        try:
            print("[*] Abriendo archivo Excel en pantalla...")
            os.startfile(ruta_salida)
        except Exception:
            pass

    return df_sat


def main():
    parser = argparse.ArgumentParser(description="Descarga Anual de Facturas del SAT")
    parser.add_argument("--sociedad", choices=["CESEHSA", "CSH", "YUKME"], help="Sociedad a descargar")
    parser.add_argument("--anio", type=int, help="Año completo a descargar (ej. 2024)")
    parser.add_argument("--carpeta", help="Carpeta de destino donde guardar el archivo Excel")
    parser.add_argument("--no-abrir", action="store_true", help="No abrir automáticamente el Excel")

    args = parser.parse_args()

    carpeta_salida = args.carpeta

    # Si no se pasan argumentos, preguntar amigablemente
    if args.sociedad and args.anio:
        sociedad = args.sociedad
        anio = args.anio
    else:
        print("\n" + "="*70)
        print("   DESCARGA DIRECTA ANUAL DEL SAT (OFICIAL CON E.FIRMA)")
        print("="*70)
        print("Selecciona la Sociedad:")
        print("  1) YUKME (YUK990113KU9) — [Prueba rápida]")
        print("  2) CESEHSA (CPR100602NJ2)")
        print("  3) CSH (CPR1006024T4)")
        opcion = input("Elige una opción [1, 2, 3] (Por defecto: 1): ").strip()
        if opcion == "2":
            sociedad = "CESEHSA"
        elif opcion == "3":
            sociedad = "CSH"
        else:
            sociedad = "YUKME"

        resp_anio = input("Ingresa el Año que deseas descargar [Por defecto: 2024]: ").strip()
        anio = int(resp_anio) if resp_anio.isdigit() else 2024

        if not carpeta_salida:
            print("\n¿Dónde deseas guardar el archivo descargado?")
            print("  1) En la carpeta 'reportes/' de esta aplicación (Recomendado)")
            print("  2) En el Escritorio")
            print("  3) En la carpeta de Descargas")
            print("  4) Escribir otra ruta manualmente")
            opc_dest = input("Elige una opción [1, 2, 3, 4] (Por defecto: 1): ").strip()
            
            user_profile = os.environ.get("USERPROFILE", "")
            if opc_dest == "2":
                carpeta_salida = os.path.join(user_profile, "Desktop")
            elif opc_dest == "3":
                carpeta_salida = os.path.join(user_profile, "Downloads")
            elif opc_dest == "4":
                carpeta_salida = input("Escribe la ruta completa de la carpeta: ").strip()
            else:
                carpeta_salida = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes")

    descargar_anio_completo(sociedad, anio, carpeta_salida=carpeta_salida, abrir_excel=not args.no_abrir)


if __name__ == "__main__":
    main()

