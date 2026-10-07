# Aperturar.md — Robot GD / AGPE (Air-e SAS ESP)

## Propósito

Robot Playwright (Python) que automatiza la apertura de solicitudes CREG 174 (Generación Distribuida y Autogeneración) desde el portal web de Air-e.

Por cada SOL el robot:
1. Consulta el número en el portal y abre el popup del formulario
2. Extrae los campos necesarios
3. Determina el tipo de solicitud desde el Unificado (`GD > 100` / `AGPE > 100` / otros)
4. Crea la carpeta del proyecto en DOC. PROYECTOS con la estructura correcta
5. Escribe el nombre en la columna "NOMBRE DEL PROYECTO" del Unificado local (solo si está vacía)
6. Ejecuta validaciones automáticas (NIC / Transformador / Potencia) y emite alertas en consola
7. Descarga todos los anexos a `1. Completitud / 1. Documentacion Inicial`

> El **paso 10 (Revisión Documento)** está desactivado — se hace manualmente.

---

## Estructura del proyecto

```
Automatizacion_Apertura_solicitudes/
├── .env                     ← credenciales del portal (excluido en .gitignore, no subir a git)
├── main.py                  ← script principal
├── renombrar_adjuntos.py    ← utilidad: renombra archivos con prefijo antiguo al nombre correcto
├── Aperturar.md             ← este archivo
├── solicitudes/
│   └── solicitudes.xlsx     ← columna A: "SOL" + números a procesar (desde fila 2)
├── unificado/
│   └── <archivo>.xlsx       ← Unificado colocado manualmente (tiene prioridad)
└── venv/                    ← entorno Python
```

---

## Cómo ejecutar

```
.\venv\Scripts\python.exe main.py
```

Desde el terminal de VS Code. El navegador se abre automáticamente y se cierra al terminar — **no cerrar ventanas del navegador** hasta ver `Cerrando navegador...` en consola.

---

## Rutas clave en OneDrive

| Recurso | Ruta |
|---------|------|
| Base OneDrive | `C:\Users\nelson.velasquez\OneDrive - Air-e SAS ESP (1)` |
| DOC. PROYECTOS | `…\RES 174 2021\DOC. PROYECTOS\<Regional>\` |
| Unificado OneDrive (sync) | `…\RES 174 2021\ARCHIVO DE CONTROL DIARIO\NUEVO UNIFICADO TERRITORIAL 19-03-2025.xlsx` |
| Unificado local (prioridad) | `.\unificado\<archivo>.xlsx` |

Carpetas regionales disponibles: `Atlántico Norte`, `Atlántico Sur`, `Guajira`, `Magdalena`.

---

## Variables de entorno (.env)

```
PORTAL_URL="https://servicios.air-e.com/creg174/form/Login.aspx"
PORTAL_USER="<usuario>"
PORTAL_PASSWORD="<contraseña>"
SHAREPOINT_UNIFICADO_URL=""    ← dejar vacío (sin acceso a cuenta autogeneracion)
```

---

## Nombre de carpeta por tipo de solicitud

El tipo se lee de la columna **TIPO SOLICITUD** del Unificado.

| Tipo | Formato de carpeta | Ejemplo |
|------|--------------------|---------|
| `GD > 100` | `<SOL> - GD <NOMBRE>` | `29106 - GD ESCORPIÓN` |
| `AGPE > 100` | `<SOL> - <NOMBRE CLIENTE>` | `29091 - LA FONTERA ALFREDO STECKERL HIERROS Y ACEROS` |
| Duplicado previo | se agrega sufijo numérico | `29107 - GD ESCORPIÓN 1` |

### Origen del nombre por tipo

**GD > 100** — prioridad:
1. Observaciones del formulario → patrón `GD - <nombre>` (solo si es nombre propio)
2. Historial del formulario → patrón `GD <nombre>`
3. Corregimiento → Vereda → Ciudad (información del inmueble)

**AGPE > 100** — siempre del campo **"Nombre del cliente"** del formulario (nunca del Unificado).
El nombre se coloca sin prefijo. Los espacios/saltos de línea del HTML se colapsan automáticamente.

### Subcarpeta de documentos

Dentro de cada carpeta de proyecto se crea:
```
<SOL> - [GD] <NOMBRE>/
└── 1. Completitud/
    └── 1. Documentacion Inicial/
        ├── 33817266_1DIAGRAMA_UNIFILAR...PDF
        └── ...
```
Los archivos se nombran exactamente como aparecen en la columna "Archivo Adjunto" del portal (nombre original del portal, sin prefijos).

---

## Validaciones automáticas por tipo

Las validaciones solo emiten alertas en consola — **no bloquean** el procesamiento.

### GD > 100
| ID | Regla |
|----|-------|
| V1 | NIC no aplica (GD) — no se valida |
| V2 | Código transformador del formulario debe coincidir con el Unificado |
| V3 | Potencia AC (kW) debe ser **igual** a Potencia entregada a la red (kW) |

### AGPE > 100
| ID | Regla |
|----|-------|
| V1 | NIC debe ser **0** |
| V2 | Código transformador del formulario debe coincidir con el Unificado |
| V3 | Potencia AC (kW) debe ser **mayor o igual** a Potencia entregada a la red (kW) |

### Otros (GD < 100, AGPE < 100, AGGE, sin tipo)
Alerta genérica de revisión manual — validaciones no automáticas.

---

## Actualización automática del Unificado

El robot escribe el nombre generado en la columna **NOMBRE DEL PROYECTO** del Unificado OneDrive (`_RUTA_UNIFICADO_SYNC`) solo si la celda está vacía. OneDrive sincroniza el cambio automáticamente.

- GD: escribe `GD <NOMBRE>`
- AGPE: escribe solo `<NOMBRE>`

---

## Selectores CSS del portal

| Elemento | Selector |
|----------|----------|
| Login usuario | `#txtUsuario` |
| Login password | `#txtPassword` |
| Botón login | `#buttonEnviar` |
| Campo SOL | `#TxtIdSolicitud` |
| Botón Consultar | `#BtnConsultarSol` |
| Filas tabla solicitudes | `#gridSolicitudes tbody tr` |
| Tabla de anexos | `#ListaAnexos tbody tr` |
| Formulario principal | `#formSolicitud` |
| Nombre cliente | `#TxtNombreCli` |
| NIC / N° cuenta | `#TxtNCuentaCli` |
| Código transformador | `#TxtCodTrafoPro` |
| Ciudad proyecto | `#DdlCiudadPro` |
| Corregimiento | `#DdlCorregimientoPro` |
| Vereda | `#DdlVeredaPro` |
| Observación | `#TxtObservacion` |
| Potencia AC | `#TxtPotenciaTotAC` |
| Potencia entregada | `#TxtPotenciaEntregada` |

---

## Flujo "Revisión Documento" (DESACTIVADO — manual)

La función `revision_documento()` está implementada pero no se llama.
Cuando se reactive: botón `REVISIÓN DOCUMENTO` → modal → campo fecha formato `d-mm-yyyy` (sin cero inicial) → CONTINUAR → CERRAR.

---

## Utilidad renombrar_adjuntos.py

Renombra archivos descargados con prefijo antiguo (`142676_Tipo_33817297_...`) al nombre correcto (`33817297_...`).

```
# Renombrar una sola SOL:
.\venv\Scripts\python.exe renombrar_adjuntos.py 28978

# Renombrar todas las SOLs en DOC. PROYECTOS:
.\venv\Scripts\python.exe renombrar_adjuntos.py
```

---

## Historial de cambios

| Fecha | Cambio |
|-------|--------|
| 2026-06-22 | Limpieza del proyecto: eliminados scripts one-time, screenshots de error, log y .browser_profile. AGPE > 100 ya no lleva prefijo "GD" en el nombre de carpeta. Espacios extra del HTML del portal colapsados en nombre_agpe(). |
| 2026-06-18 | Script funcional. +30 solicitudes procesadas. Paso 10 (Revisión Documento) desactivado. |
| 2026-06-10 | Eliminado Power Automate / JSON bitácora. Unificado se lee desde carpeta local `unificado/`. |
| 2026-05-20 | `slow_mo` 150→0. `revision_documento` reescrito con dialog handler JS + press_sequentially fallback. Fecha formato d-mm-yyyy confirmado. |
