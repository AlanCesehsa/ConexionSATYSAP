"""
========================================================================================
MÓDULO DE CONEXIÓN CON EL SERVICIO WEB PÚBLICO DEL SAT (Sin e.firma)
========================================================================================
Este módulo implementa la consulta oficial de estatus de CFDI directamente
en los servidores del SAT (Servicio de Administración Tributaria).

¿CÓMO FUNCIONA LA CONEXIÓN?
1. Protocolo: SOAP 1.1 sobre HTTPS.
2. Endpoint oficial:
   https://consultaqr.facturaelectronica.sat.gob.mx/ConsultaCFDIService.svc
3. Parámetros requeridos en la cadena (expresión impresa del código QR):
   - re: RFC del Emisor (quien emite la factura).
   - rr: RFC del Receptor (cliente que recibe la factura).
   - tt: Total del comprobante (formato numérico, hasta 6 decimales).
   - id: Folio Fiscal (UUID) en mayúsculas o minúsculas.
   - fe: (Opcional en 4.0) Últimos 8 caracteres del Sello Digital (SelloCFD).

4. Respuesta del SAT:
   - CodigoEstatus: Ejemplo 'S - Comprobante obtenido satisfactoriamente.'
   - Estado: 'Vigente', 'Cancelado' o 'No Encontrado'.
   - EsCancelable: 'Cancelable sin aceptación', 'Cancelable con aceptación', 'No cancelable'.
   - EstatusCancelacion: Si está en proceso, o el motivo.
   - ValidacionEFOS: '200' si no está en listas negras de empresas fantasma (Art. 69-B).
========================================================================================
"""

import requests
import xml.etree.ElementTree as ET
from typing import Dict, Any, Optional


class SATValidator:
    # URL del Web Service oficial del SAT para consultar estatus de CFDI
    SAT_URL = "https://consultaqr.facturaelectronica.sat.gob.mx/ConsultaCFDIService.svc"
    
    # Headers requeridos por el servicio SOAP del SAT
    HEADERS = {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction": "http://tempuri.org/IConsultaCFDIService/Consulta"
    }

    @staticmethod
    def _crear_envelope_soap(expresion: str) -> str:
        """
        Construye el XML Envelope SOAP requerido por el WSDL del SAT.
        Se usa una sección CDATA para evitar problemas con caracteres especiales (&, <, >).
        """
        return f"""<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:tem="http://tempuri.org/">
   <soapenv:Header/>
   <soapenv:Body>
      <tem:Consulta>
         <tem:expresionImpresa><![CDATA[{expresion}]]></tem:expresionImpresa>
      </tem:Consulta>
   </soapenv:Body>
</soapenv:Envelope>"""

    @classmethod
    def consultar_estatus(cls, rfc_emisor: str, rfc_receptor: str, total: float, uuid: str, sello: Optional[str] = None, timeout: int = 10) -> Dict[str, Any]:
        """
        Consulta el estatus de un CFDI en los servidores del SAT.
        
        Retorna un diccionario con:
        {
            "codigo_estatus": str,       # 'S - Comprobante obtenido satisfactoriamente.' o error
            "estado": str,               # 'Vigente', 'Cancelado', 'No Encontrado', 'Error Conexión'
            "es_cancelable": str,        # 'Cancelable con aceptación', etc.
            "estatus_cancelacion": str,  # Proceso de cancelación
            "validacion_efos": str,      # '200' si está limpio
            "error": Optional[str]       # Mensaje de error si la petición falló
        }
        """
        # Formatear el total con 2 a 6 decimales según estándar SAT
        total_str = f"{float(total):.2f}"
        uuid_str = str(uuid).strip().upper()
        rfc_emisor = str(rfc_emisor).strip().upper()
        rfc_receptor = str(rfc_receptor).strip().upper()

        # Armar la cadena de consulta (?re=...&rr=...&tt=...&id=...)
        expresion = f"?re={rfc_emisor}&rr={rfc_receptor}&tt={total_str}&id={uuid_str}"
        
        # Si se tiene el sello CFD, se anexan los últimos 8 caracteres (estándar CFDI 4.0)
        if sello and len(sello) >= 8:
            fe = sello[-8:]
            expresion += f"&fe={fe}"

        soap_envelope = cls._crear_envelope_soap(expresion)

        try:
            response = requests.post(
                cls.SAT_URL,
                data=soap_envelope.encode("utf-8"),
                headers=cls.HEADERS,
                timeout=timeout
            )
            
            if response.status_code != 200:
                return {
                    "codigo_estatus": f"HTTP {response.status_code}",
                    "estado": "Error Conexión",
                    "es_cancelable": "",
                    "estatus_cancelacion": "",
                    "validacion_efos": "",
                    "error": f"El servidor SAT respondió con código {response.status_code}"
                }

            # Parsear la respuesta XML SOAP
            root = ET.fromstring(response.content)
            
            # Los namespaces que devuelve el SAT
            namespaces = {
                "s": "http://schemas.xmlsoap.org/soap/envelope/",
                "tem": "http://tempuri.org/",
                "a": "http://schemas.datacontract.org/2004/07/Sat.Cfdi.Negocio.ConsultaCfdi.Servicio"
            }

            result_node = root.find(".//a:CodigoEstatus", namespaces)
            estado_node = root.find(".//a:Estado", namespaces)
            es_cancelable_node = root.find(".//a:EsCancelable", namespaces)
            estatus_canc_node = root.find(".//a:EstatusCancelacion", namespaces)
            efos_node = root.find(".//a:ValidacionEFOS", namespaces)

            estado = estado_node.text.strip() if estado_node is not None and estado_node.text else "No Encontrado"
            codigo_estatus = result_node.text.strip() if result_node is not None and result_node.text else ""
            es_cancelable = es_cancelable_node.text.strip() if es_cancelable_node is not None and es_cancelable_node.text else ""
            estatus_cancelacion = estatus_canc_node.text.strip() if estatus_canc_node is not None and estatus_canc_node.text else ""
            validacion_efos = efos_node.text.strip() if efos_node is not None and efos_node.text else ""

            # Si el SAT no lo encuentra en su base de datos
            if "N - 601" in codigo_estatus or "No encontrado" in codigo_estatus:
                estado = "No Encontrado"

            return {
                "codigo_estatus": codigo_estatus,
                "estado": estado,
                "es_cancelable": es_cancelable,
                "estatus_cancelacion": estatus_cancelacion,
                "validacion_efos": validacion_efos,
                "error": None
            }

        except requests.exceptions.Timeout:
            return {
                "codigo_estatus": "Timeout",
                "estado": "Error Conexión",
                "es_cancelable": "",
                "estatus_cancelacion": "",
                "validacion_efos": "",
                "error": "Tiempo de espera agotado con el SAT (> 10s)"
            }
        except Exception as e:
            return {
                "codigo_estatus": "Excepcion",
                "estado": "Error Conexión",
                "es_cancelable": "",
                "estatus_cancelacion": "",
                "validacion_efos": "",
                "error": str(e)
            }

