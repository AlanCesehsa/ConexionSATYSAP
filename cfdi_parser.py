"""
========================================================================================
EXTRACTOR DE DATOS DE COMPROBANTES FISCALES DIGITALES (CFDI 3.3 / 4.0)
========================================================================================
Lee archivos XML de facturas, notas de crédito y recibos de pago y extrae
los campos clave necesarios para consultar y conciliar contra el SAT y SAP.
"""

import os
import xml.etree.ElementTree as ET
from typing import Dict, Any, Optional


def extraer_datos_xml(ruta_xml: str) -> Optional[Dict[str, Any]]:
    """
    Parsea un archivo XML CFDI y devuelve un diccionario con los datos clave.
    Soporta CFDI 3.3 y CFDI 4.0.
    """
    if not os.path.exists(ruta_xml):
        return None

    try:
        tree = ET.parse(ruta_xml)
        root = tree.getroot()

        # Namespaces comunes de CFDI y TimbreFiscalDigital
        ns = {
            "cfdi": "http://www.sat.gob.mx/cfd/4" if "cfd/4" in root.tag else "http://www.sat.gob.mx/cfd/3",
            "tfd": "http://www.sat.gob.mx/TimbreFiscalDigital"
        }

        # Datos a nivel de Comprobante
        version = root.attrib.get("Version", "")
        serie = root.attrib.get("Serie", "")
        folio = root.attrib.get("Folio", "")
        fecha = root.attrib.get("Fecha", "")
        total = float(root.attrib.get("Total", 0.0))
        moneda = root.attrib.get("Moneda", "MXN")
        tipo_comprobante = root.attrib.get("TipoDeComprobante", "I") # I=Ingreso, E=Egreso, P=Pago
        sello_cfd = root.attrib.get("Sello", "")

        # Emisor
        emisor_node = root.find("cfdi:Emisor", ns)
        rfc_emisor = emisor_node.attrib.get("Rfc", "") if emisor_node is not None else ""
        nombre_emisor = emisor_node.attrib.get("Nombre", "") if emisor_node is not None else ""

        # Receptor
        receptor_node = root.find("cfdi:Receptor", ns)
        rfc_receptor = receptor_node.attrib.get("Rfc", "") if receptor_node is not None else ""
        nombre_receptor = receptor_node.attrib.get("Nombre", "") if receptor_node is not None else ""

        # Complemento TimbreFiscalDigital
        tfd_node = root.find(".//tfd:TimbreFiscalDigital", ns)
        uuid = ""
        fecha_timbrado = ""
        if tfd_node is not None:
            uuid = tfd_node.attrib.get("UUID", "").strip().upper()
            fecha_timbrado = tfd_node.attrib.get("FechaTimbrado", "")

        if not uuid:
            # Archivo XML sin timbrar o dañado
            return None

        return {
            "archivo": os.path.basename(ruta_xml),
            "ruta_completa": ruta_xml,
            "version": version,
            "serie": serie,
            "folio": folio,
            "folio_completo": f"{serie}{folio}".strip() if serie or folio else "",
            "fecha": fecha,
            "fecha_timbrado": fecha_timbrado,
            "tipo_comprobante": tipo_comprobante, # I=Factura, E=Nota de Crédito
            "rfc_emisor": rfc_emisor,
            "nombre_emisor": nombre_emisor,
            "rfc_receptor": rfc_receptor,
            "nombre_receptor": nombre_receptor,
            "moneda": moneda,
            "total": total,
            "uuid": uuid,
            "sello": sello_cfd
        }
    except Exception as e:
        print(f"Error al leer XML {ruta_xml}: {e}")
        return None

