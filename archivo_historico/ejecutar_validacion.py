"""
========================================================================================
SCRIPT PRINCIPAL DE EJECUCIÓN
========================================================================================
Uso:
  1. Validar un solo XML de prueba:
     python ejecutar_validacion.py --xml "ruta\\al\\archivo.xml"

  2. Validar un lote de XMLs de la carpeta de red:
     python ejecutar_validacion.py --carpeta "\\\\192.168.1.253\\Users\\Public\\XML\\creadas\\202611\\Facturas" --limite 20

  3. Validar un archivo Excel exportado de SAP:
     python ejecutar_validacion.py --excel "reporte_sap.xlsx"
"""

import sys
import argparse

# Asegurar codificacion UTF-8 para salida en terminal Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from cfdi_parser import extraer_datos_xml
from conciliador import validar_lote_cfdi, exportar_reporte_excel, procesar_carpeta_xml
from sat_service import SATValidator


def main():
    parser = argparse.ArgumentParser(description="Herramienta de Conciliación Fiscal SAP vs. SAT")
    parser.add_argument("--xml", help="Ruta de un archivo XML individual para probar")
    parser.add_argument("--carpeta", help="Ruta de carpeta con XMLs a validar")
    parser.add_argument("--limite", type=int, default=None, help="Límite máximo de XMLs a procesar en esta corrida")
    parser.add_argument("--salida", default="Reporte_Incongruencias_SAT.xlsx", help="Nombre del archivo Excel de salida")
    
    args = parser.parse_args()

    # Si no se pasan argumentos, procesamos por defecto el XML del usuario como prueba
    if not args.xml and not args.carpeta:
        xml_default = r"\\192.168.1.253\Users\Public\XML\creadas\202611\Facturas\CPR100602NJ2-66-57936.xml"
        print(f"[*] No se especificaron parámetros. Ejecutando prueba con el XML por defecto:\n    {xml_default}\n")
        args.xml = xml_default

    documentos = []

    if args.xml:
        datos = extraer_datos_xml(args.xml)
        if not datos:
            print(f"[!] No se pudo leer el archivo XML: {args.xml}")
            return
        datos["estatus_sap"] = "VIGENTE" # Supuesto inicial en SAP
        documentos.append(datos)

    elif args.carpeta:
        documentos = procesar_carpeta_xml(args.carpeta, limite=args.limite)
        if not documentos:
            print(f"[!] No se encontraron XMLs válidos en la carpeta.")
            return

    # Consultar al SAT y conciliar
    resultados = validar_lote_cfdi(documentos)

    # Mostrar resumen en consola
    print("\n" + "="*80)
    print("RESUMEN DE RESULTADOS")
    print("="*80)
    for r in resultados:
        print(f"Folio: {r.get('folio_completo', 'N/A')} | Total: ${r.get('total', 0):,.2f} | UUID: {r.get('uuid')}")
        print(f"  SAT Estado   : {r.get('sat_estado')} ({r.get('sat_codigo_respuesta')})")
        print(f"  Diagnóstico  : {r.get('diagnostico')}")
        print(f"  EFOS (69-B)  : Código {r.get('sat_validacion_efos')} (200=Limpio)")
        print("-" * 80)

    # Exportar a Excel
    exportar_reporte_excel(resultados, ruta_salida=args.salida)


if __name__ == "__main__":
    main()
