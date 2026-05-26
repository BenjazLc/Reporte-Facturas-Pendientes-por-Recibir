# Reporte-Facturas-Pendientes-por-Recibir

## 1) Objetivo del flujo
Este proyecto automatiza la construcción del **reporte de facturas pendientes por recibir/facturar** usando SAP GUI Scripting + macros VBA en Excel. El flujo integra información de **tres transacciones SAP** y genera:

- Hoja base consolidada: **Saldos por Facturar**.
- Hoja ejecutiva: **Resumen**.
- Hojas de segmentación: **Nacionales** y **Extranjeros**.

La automatización está diseñada para convivir con una orquestación externa (PAD), por eso varias macros terminan en un punto de “entrega de control” justo antes o durante exportaciones SAP.

---

## 2) Arquitectura funcional (fin a fin)

1. **Extracción SAP #1 – MB51** (`M0_MB51.OTIF_02`):
   - Abre MB51.
   - Carga variante `BLAPA3`.
   - Ajusta rango de fechas desde `01/01/año actual` hasta hoy.
   - Ejecuta y deja lista la pantalla ALV para exportar (`Shift+F4`).
   - Resultado esperado exportado: archivo tipo `ENTREGA_MERCANCIAS_TEMP*`.

2. **Extracción SAP #2 – MB5S** (`M1_MB5S.DESCARGAR_MB5S`):
   - Toma OC desde `ENTREGA_MERCANCIAS_TEMP*` (columna C).
   - Elimina duplicados en memoria.
   - Abre MB5S, aplica variante (fila 3 de lista de variantes), pega OC por selección múltiple y ejecuta.
   - Activa vista de detalle y deja listo export (`Shift+F4`).
   - Resultado esperado exportado: archivo tipo `FACTURAS_TEMP*`.

3. **Extracción SAP #3 – ME2N** (`M2_ME2N.Descargar_ME2N`):
   - Toma OC desde `FACTURAS_TEMP*` (columna A).
   - Elimina duplicados en memoria.
   - Abre ME2N, pega OC por selección múltiple, ejecuta y dispara export (`Ctrl+Shift+F7`).
   - Resultado esperado exportado: archivo tipo `OC_TEMP*`.

4. **Consolidación operativa** (`M3_CREAR_REPORTE.Armar_Reporte_Facturas`):
   - Carga `FACTURAS_TEMP*`, `OC_TEMP*`, `ENTREGA_MERCANCIAS_TEMP*`.
   - Cruza por ID compuesto (**Documento + Posición**, según cada fuente).
   - Construye libro nuevo con hoja **Saldos por Facturar**.
   - Calcula saldo por facturar, datos de guía/ingreso, fecha orden, comprador y conteo de días.
   - Aplica formato, orden, sombreado por OC y guarda `REPORTE_FACTURAS_dd-mm-yyyy.xlsx`.

5. **Resumen ejecutivo** (`M4_CREAR_TABLA.M4_Resumen_Facturas`):
   - Sobre el reporte consolidado abierto, crea/reemplaza hoja **Resumen**.
   - Calcula KPIs globales, por proveedor, por comprador, por moneda y por antigüedad.
   - Convierte montos con tipo de cambio fijo interno `3.45`.

6. **Segmentación final** (`M5_FINAL.M5_Detalle_Nacionales_Extranjeros`):
   - Crea/reemplaza hojas **Nacionales** y **Extranjeros**.
   - Separa por prefijo de código proveedor (`1` nacional, `2` extranjero).
   - Muestra ranking de proveedores y saldos por comprador.

---

## 3) Dependencias técnicas que TI debe validar

### 3.1 SAP GUI y scripting
- SAP GUI instalado en el equipo de ejecución.
- SAP GUI Scripting habilitado en cliente y servidor.
- Sesión SAP activa en `Children(0).Children(0)` (asume primera conexión/sesión).
- Acceso/autorización a transacciones: `MB51`, `MB5S`, `ME2N`.
- Variante existente en MB51: `BLAPA3`.

### 3.2 Microsoft Excel / VBA
- Excel de escritorio con soporte VBA.
- Permitir ejecución de macros firmadas o en ubicación de confianza.
- Referencias COM disponibles:
  - `SAPGUI` automation.
  - `htmlfile` (MSHTML) para uso de portapapeles.
- Permisos de lectura/escritura en carpeta de trabajo.

### 3.3 Rutas y archivos
Ruta hardcodeada principal:

`C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\FACTURACION\`

Ruta adicional declarada en MB51:

`C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\OTIF\`

Archivos que el flujo espera encontrar/generar:
- `ENTREGA_MERCANCIAS_TEMP*`
- `FACTURAS_TEMP*`
- `OC_TEMP*`
- `REPORTE_FACTURAS_dd-mm-yyyy.xlsx`

> Recomendación TI: parametrizar rutas en una hoja de configuración o archivo `.ini/.json` para evitar dependencia de usuario local (`blapa`) y sincronización OneDrive.

---

## 4) Lógica de negocio clave que se aplica en macros

## 4.1 Claves de cruce
- Se construye un ID de cruce con **Documento compras + Posición** para unir datasets.
- Diccionario `dictOC` trae desde `OC_TEMP`:
  - Nombre proveedor
  - Descripción
  - Fecha orden
  - Comprador
- Diccionario `dictEM` trae desde `ENTREGA_MERCANCIAS_TEMP`:
  - Guía de remisión
  - Fecha de ingreso

## 4.2 Cálculos relevantes en “Saldos por Facturar”
- Saldo por facturar en moneda proveedor.
- Fechas para análisis operativo (ingreso / orden).
- Conteo de días (`Conteo Días`) para envejecimiento.

## 4.3 Criterio de Nacional vs Extranjero
- Se usa prefijo del código de proveedor:
  - `1...` => Nacional.
  - `2...` => Extranjero.

## 4.4 Conversión monetaria
- Tipo de cambio fijo embebido en código: `tc = 3.45`.
- Para extranjeros se estandariza a USD con reglas específicas por hoja/campo.
- Para nacionales se usa moneda de la línea y se convierte según corresponda.

> Recomendación TI: mover `tc` a parámetro editable para gobernanza financiera (evitar hardcode).

---

## 5) Orden de ejecución recomendado (operación)

1. Abrir SAP y validar sesión activa.
2. Ejecutar `OTIF_02` (MB51) y completar exportación a `ENTREGA_MERCANCIAS_TEMP*`.
3. Ejecutar `DESCARGAR_MB5S` y completar exportación a `FACTURAS_TEMP*`.
4. Ejecutar `Descargar_ME2N` y completar exportación a `OC_TEMP*`.
5. Ejecutar `Armar_Reporte_Facturas`.
6. Abrir/validar `REPORTE_FACTURAS_dd-mm-yyyy.xlsx`.
7. Ejecutar `M4_Resumen_Facturas`.
8. Ejecutar `M5_Detalle_Nacionales_Extranjeros`.
9. Guardar versión final del libro.

---

## 6) Controles de calidad mínimos (checklist TI)

- Existe exactamente 1 sesión SAP objetivo y está activa.
- Se generaron los 3 temporales (`ENTREGA_MERCANCIAS_TEMP`, `FACTURAS_TEMP`, `OC_TEMP`).
- Las columnas esperadas por macro existen y conservan posiciones.
- El reporte final contiene hojas:
  - `Saldos por Facturar`
  - `Resumen`
  - `Nacionales`
  - `Extranjeros`
- Validación de conteos:
  - Filas base en `Saldos por Facturar` vs fuente `FACTURAS_TEMP`.
- Validación de montos:
  - Revisar muestra de saldos contra SAP.
- Validación de fecha:
  - Nombre de archivo final coincide con la fecha de ejecución.

---

## 7) Riesgos operativos identificados

1. **Dependencia fuerte de UI SAP** (IDs de controles y teclas rápidas). Cambios en SAP GUI pueden romper la macro.
2. **Uso de `SendKeys`** (`Shift+F4`, `Ctrl+Shift+F7`): sensible al foco de ventana.
3. **Ruta local hardcodeada**: no portable entre usuarios/servidores.
4. **Suposición de sesión SAP fija** (`Children(0)`): falla si hay múltiples conexiones.
5. **Dependencia de estructura de columnas** en archivos exportados.
6. **Tipo de cambio fijo** en código sin control de vigencia.

---

## 8) Recomendaciones para industrializar en TI

- Parametrizar:
  - Rutas de entrada/salida.
  - Variante SAP.
  - Tipo de cambio.
  - Fechas desde/hasta.
- Implementar bitácora técnica (log por paso con timestamp y error).
- Validar precondiciones antes de cada macro (archivo existe, columnas mínimas, SAP disponible).
- Reemplazar `SendKeys` por comandos SAP de exportación más deterministas cuando sea posible.
- Controlar versión de layout SAP/ALV y cambios funcionales de transacciones.
- Empaquetar en ejecución controlada (p.ej. PAD/Task Scheduler + usuario de servicio).

---

## 9) Mapa de módulos y responsabilidad

- `M0_MB51.bas` → extracción de entregas/mercancías base desde MB51.
- `M1_MB5S.bas` → extracción de facturas contra OC desde MB5S.
- `M2_ME2N.bas` → extracción de datos de OC desde ME2N.
- `M3_CREAR_REPORTE.bas` → consolidación y creación de hoja operacional.
- `M4_CREAR_TABLA.bas` → dashboard/KPIs ejecutivos en hoja Resumen.
- `M5_FINAL.bas` → vistas detalladas Nacionales vs Extranjeros.

---

## 10) Qué debe recibir formalmente el área de TI

1. Este documento (flujo funcional + técnico).
2. Los `.bas` versionados.
3. Ejemplo de archivos fuente exportados SAP (anonimizados).
4. Ejemplo de `REPORTE_FACTURAS_dd-mm-yyyy.xlsx` final.
5. Matriz de permisos SAP requeridos.
6. Política de actualización de tipo de cambio y calendario de ejecución.

Con estos elementos, TI puede operar, soportar y evolucionar la automatización con menor riesgo.
