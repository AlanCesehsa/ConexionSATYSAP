"""
========================================================================================
MÓDULO DE DESCARGA MASIVA DEL SAT CON E.FIRMA (FIEL)
========================================================================================
Automatiza la descarga oficial de comprobantes emitidos (Facturas y Notas de Crédito)
directamente desde los servidores centrales del SAT:
1. Autenticación con e.firma (.cer, .key y contraseña).
2. Solicitud de descarga de Metadatos (hasta 200,000 registros por petición).
3. Verificación asíncrona del estado del paquete.
4. Descarga del archivo ZIP y procesamiento a DataFrame de Pandas.
========================================================================================
"""

import os
import sys
import io
import time
import glob
import base64
import zipfile
import datetime
from typing import Optional, Dict, Any, Tuple
import pandas as pd
from cfdiclient import Fiel, Autenticacion, SolicitaDescargaEmitidos, VerificaSolicitudDescarga, DescargaMasiva


class SATDescargaMasiva:
    # RFCs de las 3 sociedades
    RFCS = {
        "CESEHSA": "CPR100602NJ2",
        "CSH": "CPR1006024T4",
        "YUKME": "YUK990113KU9"
    }

    def __init__(self, sociedad: str, base_cert_dir: Optional[str] = None):
        self.sociedad = sociedad.upper()
        if self.sociedad not in self.RFCS:
            raise ValueError(f"Sociedad '{sociedad}' no reconocida. Opciones: {list(self.RFCS.keys())}")
        
        self.rfc = self.RFCS[self.sociedad]

        # Localizar carpeta de certificados (junto al exe o embebida en bundle)
        candidatos = []
        if base_cert_dir:
            candidatos.append(os.path.join(base_cert_dir, self.sociedad))
        if hasattr(sys, '_MEIPASS'):
            candidatos.append(os.path.join(sys._MEIPASS, 'certificados', self.sociedad))
        if getattr(sys, 'frozen', False):
            candidatos.append(os.path.join(os.path.dirname(sys.executable), 'certificados', self.sociedad))
        candidatos.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'certificados', self.sociedad))
        candidatos.append(os.path.join('certificados', self.sociedad))

        self.dir_soc = None
        for cand in candidatos:
            if os.path.exists(cand):
                self.dir_soc = cand
                break

        if not self.dir_soc:
            raise FileNotFoundError(f"Faltan archivos de credenciales en certificados/{self.sociedad}")
        
        # Cargar archivos de credenciales
        cer_files = glob.glob(os.path.join(self.dir_soc, "*.cer"))
        key_files = glob.glob(os.path.join(self.dir_soc, "*.key"))
        pwd_file = os.path.join(self.dir_soc, "contraseña.txt")

        if not cer_files or not key_files or not os.path.exists(pwd_file):
            raise FileNotFoundError(f"Faltan archivos de credenciales en {self.dir_soc}")

        with open(cer_files[0], "rb") as f:
            self.cer_der = f.read()

        with open(key_files[0], "rb") as f:
            self.key_der = f.read()

        with open(pwd_file, "r", encoding="utf-8") as f:
            self.password = f.read().strip()

        # Instanciar Fiel y Autenticación
        self.fiel = Fiel(self.cer_der, self.key_der, self.password)
        self.auth = Autenticacion(self.fiel)
        self.token = None

    def autenticar(self) -> str:
        """Obtiene un token de sesión de los servidores del SAT."""
        print(f"[*] Autenticando ante el SAT para {self.sociedad} ({self.rfc})...")
        self.token = self.auth.obtener_token()
        print(f"[OK] Token SAT obtenido exitosamente.")
        return self.token

    def solicitar_descarga_metadatos(
        self,
        fecha_inicio: datetime.datetime,
        fecha_fin: datetime.datetime,
        tipo_comprobante: Optional[str] = None
    ) -> Tuple[str, Dict[str, Any]]:
        """
        Envía la solicitud de descarga de metadatos al SAT.
        Devuelve (id_solicitud, respuesta_completa).
        """
        if not self.token:
            self.autenticar()

        # El SAT rechaza con error 301 ("Fecha final invalida") si la fecha final está en el futuro
        ahora = datetime.datetime.now()
        if fecha_inicio > ahora:
            print(f"[ERROR SAT] La fecha inicial ({fecha_inicio.strftime('%Y-%m-%d')}) está en el futuro.")
            return None, {"cod_estatus": "301", "mensaje": "Fecha inicial en el futuro"}

        if fecha_fin > ahora:
            # Ajustar al momento actual menos 1 minuto para evitar rechazo del SAT
            fecha_fin = ahora - datetime.timedelta(minutes=1)

        solicitante = SolicitaDescargaEmitidos(self.fiel)
        print(f"[*] Solicitando metadatos emitidos del {fecha_inicio.strftime('%Y-%m-%d')} al {fecha_fin.strftime('%Y-%m-%d')}...")

        res = solicitante.solicitar_descarga(
            token=self.token,
            rfc_solicitante=self.rfc,
            fecha_inicial=fecha_inicio,
            fecha_final=fecha_fin,
            rfc_emisor=self.rfc,
            tipo_solicitud="Metadata",
            tipo_comprobante=tipo_comprobante, # 'I' para ingresos o None para todo
            estado_comprobante=None # Traer Vigentes Y Cancelados
        )

        id_solicitud = res.get("id_solicitud")
        cod_estatus = res.get("cod_estatus")
        mensaje = res.get("mensaje")

        print(f"   Código SAT: {cod_estatus} | Mensaje: {mensaje}")
        print(f"   ID Solicitud asignado por el SAT: {id_solicitud}")

        if not id_solicitud or str(cod_estatus) != "5000":
            print(f"[ERROR SAT] La solicitud no fue procesada: {mensaje} (Código {cod_estatus})")

        return id_solicitud, res

    def verificar_y_descargar(
        self,
        id_solicitud: str,
        max_intentos: int = 15,
        espera_segundos: int = 5
    ) -> pd.DataFrame:
        """
        Consulta periódicamente al SAT hasta que el paquete esté preparado y lo descarga.
        Devuelve un DataFrame de Pandas con la lista de comprobantes del SAT.
        """
        if not id_solicitud:
            print("[ERROR] No se puede verificar una solicitud sin ID válido.")
            return pd.DataFrame()

        if not self.token:
            self.autenticar()

        verificador = VerificaSolicitudDescarga(self.fiel)
        descargador = DescargaMasiva(self.fiel)

        print(f"[*] Esperando a que el SAT procese la solicitud {id_solicitud}...")

        paquetes_ids = []
        for intento in range(1, max_intentos + 1):
            res_verif = verificador.verificar_descarga(self.token, self.rfc, id_solicitud)
            estado_sol = str(res_verif.get("estado_solicitud", ""))
            cod_estado = str(res_verif.get("codigo_estado_solicitud", ""))
            num_cfdis = res_verif.get("numero_cfdis", 0)
            paquetes = res_verif.get("paquetes", [])

            print(f"   Intento {intento}/{max_intentos}: Estado={estado_sol} (Código={cod_estado}), CFDIs={num_cfdis}, Paquetes={len(paquetes)}")

            # En el SAT: EstadoSolicitud = '3' significa 'Terminada'
            if estado_sol == "3" or cod_estado == "3" or estado_sol.lower() == "terminada":
                paquetes_ids = paquetes
                print(f"[OK] ¡El SAT ha terminado de procesar la solicitud con {num_cfdis} comprobantes!")
                break
            
            # Si el SAT reporta error, rechazo o vencimiento (4, 5, 6), salir de inmediato
            if estado_sol in ["4", "5", "6"]:
                print(f"[ERROR SAT] La solicitud fue terminada con error/rechazo por el SAT (Estado {estado_sol}).")
                break

            time.sleep(espera_segundos)

        if not paquetes_ids:
            print("[!] El paquete aún está en proceso por el SAT o no generó comprobantes.")
            return pd.DataFrame()

        # Descargar cada paquete ZIP
        registros_totales = []
        for id_paq in paquetes_ids:
            print(f"[*] Descargando paquete: {id_paq}...")
            res_descarga = descargador.descargar_paquete(self.token, self.rfc, id_paq)
            paquete_b64 = res_descarga.get("paquete_b64")

            if not paquete_b64:
                continue

            # Decodificar Base64 y abrir ZIP
            zip_bytes = base64.b64decode(paquete_b64)
            with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
                for nombre_archivo in z.namelist():
                    print(f"   -> Archivo de Metadatos encontrado: {nombre_archivo}")
                    contenido_txt = z.read(nombre_archivo).decode("utf-8", errors="replace")
                    lineas = contenido_txt.strip().splitlines()
                    
                    if len(lineas) > 1:
                        encabezados = [c.strip() for c in lineas[0].split("~")]
                        for lin in lineas[1:]:
                            valores = [v.strip() for v in lin.split("~")]
                            if len(valores) == len(encabezados):
                                reg = dict(zip(encabezados, valores))
                                # Convertir estatus técnico numérico a texto legible
                                if "EstadoComprobante" in reg:
                                    if str(reg["EstadoComprobante"]).strip() == "1":
                                        reg["EstadoComprobante"] = "Vigente"
                                        reg["Estatus"] = "Vigente"
                                    elif str(reg["EstadoComprobante"]).strip() == "0":
                                        reg["EstadoComprobante"] = "Cancelado"
                                        reg["Estatus"] = "Cancelado"
                                registros_totales.append(reg)

        df = pd.DataFrame(registros_totales)
        print(f"[OK] Descarga completada exitosamente: {len(df)} registros obtenidos del SAT.")
        return df


def probar_descarga_yukme(año: int = 2024, mes: int = 1):
    """Función de prueba para descargar YUKME de un periodo específico."""
    yukme = SATDescargaMasiva("YUKME")
    
    fecha_ini = datetime.datetime(año, mes, 1, 0, 0, 0)
    # Último día del mes aproximado
    if mes in [1, 3, 5, 7, 8, 10, 12]:
        ultimo_dia = 31
    elif mes in [4, 6, 9, 11]:
        ultimo_dia = 30
    else:
        ultimo_dia = 29 if año % 4 == 0 else 28
    
    fecha_fin = datetime.datetime(año, mes, ultimo_dia, 23, 59, 59)

    # 1. Autenticar y solicitar
    id_solicitud, _ = yukme.solicitar_descarga_metadatos(fecha_ini, fecha_fin)

    # 2. Esperar y descargar
    df_sat = yukme.verificar_y_descargar(id_solicitud)
    
    if not df_sat.empty:
        archivo_salida = f"SAT_YUKME_{año}_{mes:02d}.xlsx"
        df_sat.to_excel(archivo_salida, index=False)
        print(f"[OK] Archivo guardado: {os.path.abspath(archivo_salida)}")
        print(df_sat.head())
    
    return df_sat


if __name__ == "__main__":
    probar_descarga_yukme(2024, 1)
