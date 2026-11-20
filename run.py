"""  """"""
========================================================================================
EJECUTOR PRINCIPAL (python run.py)
========================================================================================
Modo de uso:

1. Modo Interactivo (Te pregunta qué deseas hacer):
   python run.py

2. Modo Conciliación Mensual contra SAP:
   python run.py --sociedad YUKME --anio 2024 --mes 1

3. Modo SOLO DESCARGAR un AÑO COMPLETO del SAT (Sin cruzar con SAP):
   python run.py --solo-descargar --sociedad YUKME --anio 2024
   python run.py --solo-descargar --sociedad CESEHSA --anio 2024
   python run.py --solo-descargar --sociedad CSH --anio 2024
========================================================================================
"""

import os
import sys
import argparse
import datetime
import warnings
import pandas as pd

# Silenciar avisos informativos
warnings.filterwarnings('ignore')

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from sat_descarga_masiva import SATDescargaMasiva
from motor_conciliacion import conciliar_sociedad, exportar_excel_ejecutivo
from descargar_sat import descargar_anio_completo


def main():
    # Si se ejecuta sin argumentos (ej. 'python run.py' o 'python run'), abrir directamente la interfaz gráfica
    if len(sys.argv) == 1 or "--gui" in sys.argv:
        import tkinter as tk
        from app_conciliacion import AppConciliacionFiscal
        root = tk.Tk()
        app = AppConciliacionFiscal(root)
        root.mainloop()
        return

    parser = argparse.ArgumentParser(description="Sistema Fiscal SAT y SAP")
    parser.add_argument("--gui", action="store_true", help="Abrir la interfaz gráfica visual")
    parser.add_argument("--solo-descargar", action="store_true", help="Solo descargar del SAT sin conciliar con SAP")
    parser.add_argument("--sociedad", choices=["CESEHSA", "CSH", "YUKME"], help="Sociedad a procesar")
    parser.add_argument("--anio", type=int, help="Año a procesar (ej. 2024)")
    parser.add_argument("--mes", type=int, help="Mes a auditar (1 a 12)")
    parser.add_argument("--sap", help="Ruta al archivo Excel de SAP")
    parser.add_argument("--carpeta", help="Carpeta donde guardar el archivo Excel resultante")
    parser.add_argument("--no-abrir", action="store_true", help="No abrir automáticamente el Excel resultante")

    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.abspath(__file__))

    print("\n" + "="*70)
    print("   SISTEMA FISCAL: CONEXIÓN SAT & CONCILIACIÓN SAP")
    print("="*70)

    # 1. Determinar modo de operación
    if args.solo_descargar:
        modo = "descargar_anio"
    elif args.sociedad and args.mes:
        modo = "conciliar_mes"
    else:
        print("\n¿Qué deseas hacer hoy?")
        print("  1) SOLO DESCARGAR comprobantes del SAT de un AÑO COMPLETO (Verificar descarga)")
        print("  2) CONCILIACIÓN COMPLETA SAP vs. SAT de un Mes")
        opc_modo = input("Elige una opción [1 o 2] (Por defecto: 1): ").strip()
        modo = "conciliar_mes" if opc_modo == "2" else "descargar_anio"

    # 2. Sociedad
    if args.sociedad:
        sociedad = args.sociedad.upper()
    else:
        print("\nSelecciona la Sociedad:")
        print("  1) YUKME (YUK990113KU9) — [Prueba rápida]")
        print("  2) CESEHSA (CPR100602NJ2)")
        print("  3) CSH (CPR1006024T4)")
        opc_soc = input("Elige [1, 2, 3] (Por defecto: 1): ").strip()
        if opc_soc == "2":
            sociedad = "CESEHSA"
        elif opc_soc == "3":
            sociedad = "CSH"
        else:
            sociedad = "YUKME"

    # 3. Año
    if args.anio:
        anio = args.anio
    else:
        resp_anio = input("Año a procesar [Por defecto: 2024]: ").strip()
        anio = int(resp_anio) if resp_anio.isdigit() else 2024

    # 4. Carpeta de Salida
    carpeta_salida = args.carpeta
    if not carpeta_salida and not (args.sociedad and args.anio and (args.solo_descargar or args.mes)):
        print("\n¿Dónde deseas guardar el archivo generado?")
        print("  1) En la carpeta 'reportes/' de esta aplicación (Recomendado)")
        print("  2) En el Escritorio")
        print("  3) En la carpeta de Descargas")
        print("  4) Escribir ruta manualmente")
        opc_c = input("Elige una opción [1, 2, 3, 4] (Por defecto: 1): ").strip()
        user_prof = os.environ.get("USERPROFILE", "")
        if opc_c == "2":
            carpeta_salida = os.path.join(user_prof, "Desktop")
        elif opc_c == "3":
            carpeta_salida = os.path.join(user_prof, "Downloads")
        elif opc_c == "4":
            carpeta_salida = input("Ruta completa de la carpeta: ").strip()
        else:
            carpeta_salida = os.path.join(base_dir, "reportes")

    if not carpeta_salida:
        carpeta_salida = os.path.join(base_dir, "reportes")
    os.makedirs(carpeta_salida, exist_ok=True)

    # -------------------------------------------------------------
    # CASO A: SOLO DESCARGAR AÑO DEL SAT
    # -------------------------------------------------------------
    if modo == "descargar_anio":
        descargar_anio_completo(sociedad, anio, carpeta_salida=carpeta_salida, abrir_excel=not args.no_abrir)
        return

    # -------------------------------------------------------------
    # CASO B: CONCILIAR MES CONTRA SAP
    # -------------------------------------------------------------
    if args.mes:
        mes = args.mes
    else:
        resp_mes = input("Mes a auditar (1 a 12) [Por defecto: 1]: ").strip()
        mes = int(resp_mes) if resp_mes.isdigit() and 1 <= int(resp_mes) <= 12 else 1

    if args.sap and os.path.exists(args.sap):
        ruta_sap = args.sap
    else:
        candidatos_sap = [
            os.path.join(base_dir, "Reporte_SAPvsSAT.xlsx"),
            r"c:\Users\Stock CESEHSA\Documents\Reporte_SAPvsSAT (2)\Reporte_SAPvsSAT.xlsx",
            r"c:\Users\Stock CESEHSA\Downloads\EJEMPLO DE CONCILIACION.xls"
        ]
        ruta_sap = next((c for c in candidatos_sap if os.path.exists(c)), None)

    if not ruta_sap:
        print("[!] No se encontró un archivo de reporte de SAP.")
        return

    print("\n" + "-"*70)
    print(f"Parámetros de Conciliación:")
    print(f"  • Sociedad        : {sociedad}")
    print(f"  • Periodo         : {anio}-{mes:02d}")
    print(f"  • Base SAP        : {ruta_sap}")
    print(f"  • Carpeta Destino : {carpeta_salida}")
    print("-"*70 + "\n")

    # Descarga SAT mes
    print(f"[Paso 1/3] Descargando comprobantes oficiales del SAT con e.firma...")
    sat = SATDescargaMasiva(sociedad, base_cert_dir=os.path.join(base_dir, "certificados"))
    sat.autenticar()

    ahora = datetime.datetime.now()
    fecha_ini = datetime.datetime(anio, mes, 1, 0, 0, 0)
    ultimo_dia = 31 if mes in [1, 3, 5, 7, 8, 10, 12] else (30 if mes in [4, 6, 9, 11] else (29 if anio % 4 == 0 else 28))
    fecha_fin = datetime.datetime(anio, mes, ultimo_dia, 23, 59, 59)

    if fecha_ini > ahora:
        print(f"[ERROR] El periodo {anio}-{mes:02d} está completamente en el futuro.")
        return

    if fecha_fin > ahora:
        fecha_fin = ahora - datetime.timedelta(minutes=1)

    id_sol, res = sat.solicitar_descarga_metadatos(fecha_ini, fecha_fin)
    if not id_sol:
        print(f"[ERROR] El SAT rechazó la solicitud: {res.get('mensaje')}")
        return

    df_sat = sat.verificar_y_descargar(id_sol, max_intentos=15, espera_segundos=4)

    if df_sat.empty:
        print("[!] El SAT no devolvió comprobantes para este periodo.")
        return

    # Leer SAP
    print(f"\n[Paso 2/3] Leyendo datos de SAP desde: {os.path.basename(ruta_sap)}...")
    xl_sap = pd.ExcelFile(ruta_sap, engine="xlrd" if ruta_sap.lower().endswith(".xls") else "openpyxl")
    hoja_sap = next((h for h in [f"SAP_{sociedad}", "facturas sap", sociedad] if h in xl_sap.sheet_names), xl_sap.sheet_names[0])
    df_sap = xl_sap.parse(hoja_sap)

    # Conciliar
    print(f"\n[Paso 3/3] Cruzando por UUID y detectando alertas...")
    resultados = conciliar_sociedad(df_sap, df_sat, nombre_sociedad=sociedad)
    salida = os.path.join(carpeta_salida, f"CONCILIACION_{sociedad}_{anio}_{mes:02d}_EJECUTIVA.xlsx")
    exportar_excel_ejecutivo(resultados, salida, nombre_sociedad=sociedad)

    print("\n" + "="*70)
    print("RESUMEN DE AUDITORÍA FISCAL (SAT vs. SAP)")
    print("="*70)
    print(resultados["RESUMEN"].to_string(index=False))
    print("="*70)
    print(f"\n[OK] Archivo Excel generado: {salida}")

    if not args.no_abrir:
        try:
            print("[*] Abriendo reporte en Excel...")
            os.startfile(salida)
        except Exception:
            pass


if __name__ == "__main__":
    main()
