# Herramienta de Conciliación y Validación Fiscal (SAP vs. SAT)

Herramienta en Python diseñada para auditar y conciliar comprobantes fiscales emitidos (**Facturas** y **Notas de Crédito**), comparando los registros de **SAP** contra el **SAT** (Servicio de Administración Tributaria).

---

## 1. ¿Cómo se conecta al SAT actualmente? (Sin e.firma)

Esta herramienta utiliza el **Web Service oficial de Consulta de Estatus de CFDI del SAT**:

* **URL del Servicio:** `https://consultaqr.facturaelectronica.sat.gob.mx/ConsultaCFDIService.svc`
* **Protocolo:** SOAP 1.1 sobre HTTPS.
* **Autenticación requerida:** **Ninguna.** Es un servicio público y gratuito diseñado por el SAT para validar la vigencia de cualquier factura en tiempo real.
* **Parámetros requeridos (Expresión QR):**
  * `re`: RFC del Emisor (`CPR100602NJ2`)
  * `rr`: RFC del Receptor (Cliente)
  * `tt`: Total de la factura (con formato numérico a 2 o 6 decimales)
  * `id`: Folio Fiscal o UUID
  * `fe`: Últimos 8 caracteres del Sello Digital (SelloCFD)

### ¿Qué responde el SAT?
```xml
<ConsultaResult>
    <CodigoEstatus>S - Comprobante obtenido satisfactoriamente.</CodigoEstatus>
    <Estado>Vigente</Estado>               <!-- Vigente | Cancelado | No Encontrado -->
    <EsCancelable>Cancelable sin aceptación</EsCancelable>
    <EstatusCancelacion/>
    <ValidacionEFOS>200</ValidacionEFOS>  <!-- 200 = Limpio de listas negras (Art. 69-B) -->
</ConsultaResult>
```

---

## 2. Incongruencias y Alertas Detectadas

El motor cruza el estatus registrado en **SAP** contra la respuesta en tiempo real del **SAT**:

1. **`[ALERTA CRÍTICA] Cancelado en SAT pero Vigente en SAP`**:
   * La factura fue cancelada ante el SAT, pero en SAP sigue abierta o activa.
   * **Riesgo:** Estás reconociendo un ingreso o una cuenta por cobrar que fiscalmente ya no existe.
2. **`[ALERTA CRÍTICA] Comprobante no encontrado en el SAT`**:
   * El UUID registrado en SAP no existe en la base de datos del SAT.
   * **Riesgo:** La factura nunca se timbró o el UUID se corrompió.
3. **`[ADVERTENCIA] Cancelado en SAP pero sigue Vigente en SAT`**:
   * En SAP se canceló internamente el documento, pero nunca se envió la solicitud de cancelación al SAT (o fue rechazada).
   * **Riesgo:** Fiscalmente sigue viva y el SAT esperará el pago de impuestos por esa factura.
4. **`[ADVERTENCIA] Diferencia de importe`**:
   * El total en SAP difiere del total timbrado en el SAT por más de $0.05.

---

## 3. Modo de Uso

### A) Validar un solo archivo XML de prueba:
```bash
python ejecutar_validacion.py --xml "\\192.168.1.253\Users\Public\XML\creadas\202611\Facturas\CPR100602NJ2-66-57936.xml"
```

### B) Validar un lote de XMLs de la red (Multihilo / Ultra rápido):
```bash
python ejecutar_validacion.py --carpeta "\\192.168.1.253\Users\Public\XML\creadas\202611\Facturas" --limite 50
```

### C) Salida generada:
Se genera automáticamente un archivo Excel: `Reporte_Incongruencias_SAT.xlsx` con:
* Semáforo de colores: Rojo (Crítico), Amarillo (Advertencia), Verde (OK).
* Columnas completas: Folio, Fecha, RFC Receptor, Cliente, Total, UUID, Estatus SAT, Estatus SAP y Validación EFOS (Listas negras).

---

## 4. ¿Cómo se implementaría con la e.firma (FIEL)?

Cuando cuenten con la e.firma (`.cer`, `.key` y contraseña privada), se puede habilitar la **Descarga Masiva del SAT**, lo cual permite bajar hasta **200,000 registros por solicitud**.

### Flujo Técnico del Web Service de Descarga Masiva:

```
[Tu Aplicación]                                                  [Servidores SAT]
       │                                                                │
       │  1. Autenticación (SOAP Envelope firmado con .key y .cer)       │
       ├───────────────────────────────────────────────────────────────>│
       │<───────────────────────────────────────────────────────────────┤
       │  Devuelve Token de Autorización (Bearer Token temporal)        │
       │                                                                │
       │  2. Solicita Descarga (Filtros: Fecha Inicio, Fin, Emitidos)   │
       ├───────────────────────────────────────────────────────────────>│
       │<───────────────────────────────────────────────────────────────┤
       │  Devuelve IdSolicitud (Ej: "5a8e23...")                        │
       │                                                                │
       │  3. Verifica Solicitud (Polling después de unos minutos)       │
       ├───────────────────────────────────────────────────────────────>│
       │<───────────────────────────────────────────────────────────────┤
       │  Devuelve EstadoSolicitud="Terminada" + Lista de IdPaquetes    │
       │                                                                │
       │  4. Descarga Masiva (Descarga paquete ZIP codificado Base64)   │
       ├───────────────────────────────────────────────────────────────>│
       │<───────────────────────────────────────────────────────────────┤
       │  Devuelve ZIP con archivo de METADATOS (o XMLs completos)      │
```

### El archivo de Metadatos del SAT:
El SAT entrega un archivo de texto delimitado por `~` con todas las facturas emitidas por la empresa:
```
UUID~RfcEmisor~NomEmisor~RfcReceptor~NomReceptor~FechaEmision~FechaCertificacion~PacCertifico~Total~EfectoComprobante~Estatus~FechaCancelacion
27EB3C1D-6F1A...~CPR100602NJ2~...~CPR1006024T4~...~2026-11-04~2026-11-04~...~257.46~I~1~
```
* `Estatus = 1`: Vigente
* `Estatus = 0`: Cancelado (e incluye `FechaCancelacion`)

### Ventaja de la e.firma:
Con la e.firma obtienes la **verdad absoluta del SAT** para todo un mes en un solo archivo, permitiendo saber al instante:
1. Todo lo que el SAT tiene registrado como cancelado.
2. Comprobantes emitidos en el portal del SAT que nadie registró en SAP.

