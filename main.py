"""
Robot GD — Air-e SAS ESP
Automatización formularios CREG 174

Flujo por solicitud:
  1. Consultar SOL en el portal
  2. Abrir formulario CREG 174 (popup)
  3. Extraer campos necesarios
  4. Crear carpeta en DOC. PROYECTOS (regional correcta según Unificado)
  5. Validar: NIC en 0 / Transformador / Potencia AC
  6. Descargar anexos/PDFs a la carpeta del proyecto
  7. Clic REVISIÓN DOCUMENTO → escribir fecha → CONTINUAR → cerrar
"""

import os
import re
import sys
import unicodedata
import urllib.parse
import warnings
from datetime import date, datetime
from pathlib import Path

# Forzar UTF-8 en la consola Windows para soportar emojis y tildes
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


# ══════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════

load_dotenv()

PORTAL_URL      = os.getenv("PORTAL_URL")
PORTAL_USER     = os.getenv("PORTAL_USER")
PORTAL_PASSWORD = os.getenv("PORTAL_PASSWORD")

if not PORTAL_URL or not PORTAL_USER or not PORTAL_PASSWORD:
    raise ValueError("Faltan variables en .env (PORTAL_URL, PORTAL_USER, PORTAL_PASSWORD)")

# URL del Unificado en SharePoint (opcional — descarga automática si está configurado)
SHAREPOINT_UNIFICADO_URL = os.getenv("SHAREPOINT_UNIFICADO_URL", "").strip()


# Ruta DOC. PROYECTOS en OneDrive
_BASE = r"C:\Users\nelson.velasquez\OneDrive - Air-e SAS ESP (1)"
RUTA_DOC_PROYECTOS = _BASE + r"\RES 174 2021\DOC. PROYECTOS"

# Unificado: ruta OneDrive sincronizada (siempre actualizada por el cliente OneDrive)
# Se puede sobreescribir colocando un .xlsx en la carpeta unificado/
_RUTA_UNIFICADO_SYNC = (
    _BASE
    + r"\RES 174 2021\ARCHIVO DE CONTROL DIARIO"
    + r"\NUEVO UNIFICADO TERRITORIAL 19-03-2025.xlsx"
)
_CARPETA_UNIFICADO   = Path(__file__).parent / "unificado"

# Solicitudes: se leen desde solicitudes/solicitudes.xlsx
_CARPETA_SOLICITUDES = Path(__file__).parent / "solicitudes"
_ARCHIVO_SOLICITUDES = _CARPETA_SOLICITUDES / "solicitudes.xlsx"

# Perfil de browser persistente para SharePoint — fuera de OneDrive (sin espacios en ruta)
_RUTA_PERFIL_BROWSER = str(Path.home() / ".robot_gd_profile")

# Mapa territorio → carpeta regional en DOC. PROYECTOS
TERRITORIO_A_REGIONAL = {
    "atlantico norte": "Atlántico Norte",
    "atlantico sur":   "Atlántico Sur",
    "guajira":         "Guajira",
    "la guajira":      "Guajira",
    "magdalena":       "Magdalena",
}


# ── Selectores del portal ──────────────────────────────────────
SEL_USUARIO       = "#txtUsuario"
SEL_PASSWORD      = "#txtPassword"
SEL_LOGIN         = "#buttonEnviar"
SEL_CAMPO_SOL     = "#TxtIdSolicitud"
SEL_BTN_CONSULTAR = "#BtnConsultarSol"
SEL_FILAS_TABLA   = "#gridSolicitudes tbody tr"
SEL_FILAS_ANEXOS  = "#ListaAnexos tbody tr"


# ══════════════════════════════════════════════════════════════
# UTILIDADES
# ══════════════════════════════════════════════════════════════

def norm(txt: str) -> str:
    """Minúsculas sin tildes para comparaciones robustas."""
    if not txt:
        return ""
    txt = str(txt).strip().lower()
    return unicodedata.normalize("NFD", txt).encode("ascii", "ignore").decode("utf-8")


def limpio(valor) -> str:
    """str limpio; descarta nan/nat/none de pandas."""
    s = str(valor).strip() if valor is not None else ""
    return "" if s.lower() in ("nan", "nat", "none", "<na>") else s


def col(df: pd.DataFrame, *nombres) -> str:
    """Nombre real de columna buscando sin tildes/mayúsculas."""
    for n in nombres:
        nn = norm(n)
        for c in df.columns:
            if norm(c) == nn:
                return c
    return ""


def limpiar_fs(nombre: str, reemplazo: str = "") -> str:
    """Elimina caracteres inválidos para nombres de carpeta en Windows."""
    return re.sub(r'[\\/*?:"<>|]', reemplazo, str(nombre)).strip(" .") or "SIN_NOMBRE"


def a_float(valor) -> float | None:
    try:
        return float(str(valor).replace(",", ".").strip())
    except (ValueError, TypeError):
        return None


def log(msg: str):
    """Print con timestamp HH:MM:SS."""
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


# ══════════════════════════════════════════════════════════════
# DESCARGA UNIFICADO DESDE SHAREPOINT
# ══════════════════════════════════════════════════════════════

def _convertir_a_download_url(url: str) -> str:
    """
    Convierte la URL de vista de SharePoint (la del navegador) a URL de descarga directa.
    Soporta el formato moderno: /x:/r/personal/.../Doc.aspx?sourcedoc={GUID}
    """
    if "download.aspx" in url:
        return url
    parsed = urllib.parse.urlparse(url)
    params = urllib.parse.parse_qs(parsed.query)
    source_doc = params.get("sourcedoc", [""])[0]
    if not source_doc:
        return url
    unique_id   = source_doc.strip("{}")
    path        = parsed.path.replace("/x:/r/", "/")          # quitar el /x:/r/ del formato moderno
    base_path   = path.split("/_layouts/")[0]                 # hasta /personal/usuario
    return (
        f"{parsed.scheme}://{parsed.netloc}{base_path}"
        f"/_layouts/15/download.aspx?UniqueId={unique_id}"
    )


def descargar_unificado_sharepoint(pw) -> str:
    """
    Descarga el Unificado desde SharePoint usando un perfil de browser persistente.

    - Primera ejecución: abre el navegador, el usuario inicia sesión en Microsoft
      y la sesión queda guardada en .browser_profile/.
    - Ejecuciones siguientes: descarga automática sin intervención.

    Retorna la ruta local del archivo descargado, o "" si falla.
    """
    if not SHAREPOINT_UNIFICADO_URL:
        return ""

    url_descarga = _convertir_a_download_url(SHAREPOINT_UNIFICADO_URL)
    _CARPETA_UNIFICADO.mkdir(exist_ok=True)

    log("Descargando Unificado desde SharePoint...")
    log(f"  URL descarga: {url_descarga}")
    try:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=_RUTA_PERFIL_BROWSER,
            headless=False,
            accept_downloads=True,
            no_viewport=True,
        )
        page = ctx.new_page()
        try:
            with page.expect_download(timeout=30000) as dl_info:
                page.goto(url_descarga, wait_until="domcontentloaded")
            descarga = dl_info.value
            nombre   = descarga.suggested_filename or "unificado.xlsx"
            # Reemplazar versión anterior
            for f in _CARPETA_UNIFICADO.glob("*.xlsx"):
                f.unlink(missing_ok=True)
            ruta = _CARPETA_UNIFICADO / nombre
            descarga.save_as(str(ruta))
            log(f"  ✔ Unificado actualizado: {nombre}")
            ctx.close()
            return str(ruta)
        except PlaywrightTimeoutError:
            log("  [SHAREPOINT] Primera ejecucion o sesion expirada.")
            log("  [SHAREPOINT] Un navegador se abrio. Inicia sesion en Microsoft.")
            log("  [SHAREPOINT] Tienes 90 segundos. Despues el script continua.")
            log("  [SHAREPOINT] La sesion queda guardada en .browser_profile/")
            try:
                page.wait_for_timeout(90000)
            except Exception:
                pass
            ctx.close()
            return ""
    except Exception as e:
        log(f"  ⚠ Error descargando Unificado desde SharePoint: {e}")
        return ""


# ══════════════════════════════════════════════════════════════
# CARGAR UNIFICADO
# ══════════════════════════════════════════════════════════════

def buscar_unificado() -> str:
    """
    Prioridad:
      1. Primer .xlsx en carpeta local unificado/ (override manual)
      2. Ruta OneDrive sincronizada automáticamente por el cliente OneDrive
    """
    _CARPETA_UNIFICADO.mkdir(exist_ok=True)
    archivos = sorted(_CARPETA_UNIFICADO.glob("*.xlsx"))
    if archivos:
        log(f"  Unificado local: {archivos[0].name}")
        return str(archivos[0])
    if os.path.isfile(_RUTA_UNIFICADO_SYNC):
        log("  Unificado OneDrive sincronizado")
        return _RUTA_UNIFICADO_SYNC
    log("⚠ Unificado no encontrado. Opciones:")
    log(f"  1. Coloca el xlsx en la carpeta 'unificado/'")
    log(f"  2. Verifica que OneDrive esté sincronizado: {_RUTA_UNIFICADO_SYNC}")
    return ""


# ══════════════════════════════════════════════════════════════
# CARGAR SOLICITUDES
# ══════════════════════════════════════════════════════════════

def cargar_solicitudes() -> list[str]:
    """
    Lee la columna SOL desde solicitudes/solicitudes.xlsx.
    Si el archivo no existe, crea la plantilla y avisa al usuario.
    """
    import openpyxl

    _CARPETA_SOLICITUDES.mkdir(exist_ok=True)

    if not _ARCHIVO_SOLICITUDES.exists():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Solicitudes"
        ws["A1"] = "SOL"
        ws.column_dimensions["A"].width = 15
        wb.save(str(_ARCHIVO_SOLICITUDES))
        log(f"⚠ Plantilla creada: {_ARCHIVO_SOLICITUDES}")
        log("  Agrega los números SOL en la columna A (desde la fila 2) y ejecuta de nuevo.")
        return []

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(str(_ARCHIVO_SOLICITUDES), dtype=str)

    c_sol = col(df, "SOL")
    if not c_sol:
        log(f"⚠ Columna 'SOL' no encontrada en {_ARCHIVO_SOLICITUDES.name}")
        return []

    sols = df[c_sol].dropna().astype(str).str.strip().tolist()
    sols = [s for s in sols if s and s.lower() not in ("nan", "")]

    # Eliminar duplicados manteniendo orden
    seen: set = set()
    unique = []
    for s in sols:
        if s not in seen:
            seen.add(s)
            unique.append(s)

    log(f"  {len(unique)} solicitud(es) en {_ARCHIVO_SOLICITUDES.name}")
    return unique


# ══════════════════════════════════════════════════════════════
# TIPO DE SOLICITUD
# ══════════════════════════════════════════════════════════════

def tipo_solicitud(sol: str, df: pd.DataFrame) -> str:
    """Lee la columna TIPO SOLICITUD del Unificado. Ej: 'GD > 100', 'AGPE > 100'."""
    C_SOL  = col(df, "SOL")
    C_TIPO = col(df, "TIPO SOLICITUD", "TIPO DE SOLICITUD", "TIPO")
    if not C_SOL or not C_TIPO or df.empty:
        return ""
    m = df[df[C_SOL].astype(str).str.strip() == sol]
    if m.empty:
        return ""
    return limpio(m.iloc[0][C_TIPO]).upper().strip()


# ══════════════════════════════════════════════════════════════
# NOMBRE DEL PROYECTO
# ══════════════════════════════════════════════════════════════

_PATRON_GD = re.compile(
    r'\bGD[\s\-]+([A-ZÁÉÍÓÚÑa-záéíóúñ][A-ZÁÉÍÓÚÑa-záéíóúñ\s]{2,})',
    re.IGNORECASE,
)

# Palabras que indican texto genérico, no un nombre de proyecto
_PALABRAS_GENERICAS = {
    "tiene", "una", "de", "el", "la", "los", "las", "con", "por", "para",
    "que", "y", "o", "en", "un", "es", "son", "este", "esta", "hay", "como",
    "se", "del", "al", "no", "si", "su", "sus", "cual", "cuyo",
    "menor", "mayor", "igual", "hasta", "desde", "entre", "sobre",
}

# Sufijos que identifican razones sociales (no se usan como nombre de carpeta GD)
_SUFIJOS_EMPRESA = {"SAS", "SA", "ESP", "LTDA", "SRL", "EU", "BV", "INC", "LLC", "CORP"}

def _es_empresa(nombre: str) -> bool:
    """True si el nombre es una razón social (termina en sufijo corporativo)."""
    if not nombre or not nombre.strip():
        return False
    ultima = nombre.strip().upper().replace(".", "").split()[-1].rstrip(".,;:")
    return ultima in _SUFIJOS_EMPRESA

def _es_nombre_propio(nombre: str) -> bool:
    """Descarta texto genérico de observaciones; solo acepta nombres propios."""
    if not nombre or len(nombre.strip()) < 3:
        return False
    # Debe empezar con mayúscula
    if not nombre.strip()[0].isupper():
        return False
    palabras = nombre.strip().split()
    primera  = palabras[0].lower().rstrip(".,;:")
    if primera not in _PALABRAS_GENERICAS:
        return True
    # Artículo/preposición válido si va seguido de nombre propio en mayúscula (mín 2 letras)
    # ej: "LA BERTINA", "EL CERRO" → aceptar | "MENOR A" → rechazar ("A" es 1 letra)
    return len(palabras) > 1 and len(palabras[1]) > 1 and palabras[1][0].isupper()


# Conectores válidos dentro de nombres de lugar (no cortan el nombre)
_CONECTORES_LUGAR = {"de", "del", "la", "las", "los", "el", "en", "y", "i", "e"}

def _recortar_nombre_proyecto(nombre: str) -> str:
    """
    Recorta el nombre extraído del regex en el primer verbo o texto descriptivo.
    Conserva conectores de lugar ('de', 'del', 'la', etc.) que son parte del topónimo.
    Ej: 'Escorpión requiere conectarse…' → 'Escorpión'
        'San Juan del Cesar'             → 'San Juan del Cesar'
        'La Pantera'                     → 'La Pantera'
    """
    palabras = nombre.strip().split()
    resultado = []
    for p in palabras:
        p_limpia = p.lower().rstrip(".,;:")
        # Palabra en minúscula que no es conector de lugar → inicio de frase → parar
        if p[0].islower() and p_limpia not in _CONECTORES_LUGAR:
            break
        resultado.append(p)
    return " ".join(resultado).strip()


def _nombre_desde_historial(popup) -> str:
    """Busca patrón 'GD <nombre>' en las celdas del historial del formulario."""
    for sel in [
        "#GridHistorial tbody tr td",
        "#gridHistorial tbody tr td",
        "[id*='Historial'] td",
        "[id*='historial'] td",
    ]:
        try:
            celdas = popup.locator(sel)
            for i in range(min(celdas.count(), 60)):
                texto = celdas.nth(i).text_content() or ""
                for m in _PATRON_GD.finditer(texto):
                    if _es_nombre_propio(m.group(1)):
                        recortado = _recortar_nombre_proyecto(m.group(1))
                        if recortado:
                            log(f"  · Nombre en historial: {recortado}")
                            return recortado
        except Exception:
            continue
    return ""


def nombre_gd(popup, datos: dict, fila_tabla: dict) -> str:
    """
    Nombre para GD > 100:
      1. Observaciones → patrón 'GD - <nombre>'
      2. Historial → patrón 'GD <nombre>'
      3. Nombre del cliente si empieza por 'Parque' (parque solar/eólico)
      4. Corregimiento → Vereda  (más específicos que ciudad)
      5. Nombre del cliente (cualquier valor) — antes de usar ciudad genérica
      6. Ciudad (último recurso)
    """
    obs        = datos.get("observacion_detalle", "")
    nombre_cli = " ".join(datos.get("nombre_cliente", "").split())

    # 1. Observaciones → patrón 'GD - <nombre>' (itera todos los matches, no solo el primero)
    if obs:
        for m in _PATRON_GD.finditer(obs):
            if _es_nombre_propio(m.group(1)):
                recortado = _recortar_nombre_proyecto(m.group(1))
                if recortado:
                    log(f"  · Nombre de Observaciones: {recortado}")
                    return recortado

    # 2. Historial
    nombre_hist = _nombre_desde_historial(popup)
    if nombre_hist:
        return nombre_hist

    # 3. Nombre del cliente si es un parque o tecnoparque (solar, eólico, etc.)
    _cli_up = nombre_cli.strip().upper()
    if nombre_cli and (_cli_up.startswith("PARQUE") or _cli_up.startswith("TECNOPARQUE")):
        log(f"  · Nombre del cliente (Parque/Tecnoparque GD): {nombre_cli}")
        return nombre_cli.title()

    # 4. Corregimiento o Vereda — solo si son más específicos que la ciudad
    #    (se descartan si coinciden con la ciudad o la contienen)
    correg = datos.get("corregimiento_proyecto", "") or fila_tabla.get("corregimiento", "")
    vereda = datos.get("vereda_proyecto",        "") or fila_tabla.get("vereda",        "")
    ciudad = datos.get("ciudad_proyecto",        "") or fila_tabla.get("ciudad",        "")

    ciudad_norm = (ciudad or "").strip().upper()
    for fuente in [correg, vereda]:
        v = (fuente or "").strip().upper()
        if v and v != ciudad_norm and ciudad_norm not in v:
            log(f"  · Nombre de inmueble: {fuente.strip()}")
            return fuente.strip()

    # 5. Nombre del cliente — solo si no es razón social (SAS, S.A.S, E.S.P, etc.)
    if nombre_cli and not _es_empresa(nombre_cli):
        log(f"  · Nombre del cliente (GD): {nombre_cli}")
        return nombre_cli

    # 6. Ciudad como último recurso absoluto
    if ciudad and ciudad.strip():
        log(f"  · Nombre de inmueble (ciudad): {ciudad.strip()}")
        return ciudad.strip()

    return "SIN_NOMBRE"


def nombre_agpe(datos: dict) -> str:
    """
    Nombre para AGPE: siempre del campo Nombre del cliente del formulario.
    Colapsa espacios/saltos de línea internos que puede traer el HTML del portal.
    """
    nombre_cli = " ".join(datos.get("nombre_cliente", "").split())
    if nombre_cli:
        log(f"  · Nombre del cliente (AGPE): {nombre_cli}")
        return nombre_cli
    return "SIN_NOMBRE"


# ══════════════════════════════════════════════════════════════
# CARPETA DEL PROYECTO
# ══════════════════════════════════════════════════════════════


def nombre_carpeta(sol: str, nombre: str, tipo: str, batch_nombres: dict) -> str:
    """
    El sufijo numérico se calcula por posición dentro del batch actual (solicitudes.xlsx):
      - Primero con ese nombre base → sin sufijo:  '29224 - GD BARRANCAS'
      - Segundo                     → sufijo ' 1': '29225 - GD BARRANCAS 1'
    batch_nombres es un dict {nombre_base: count} que se acumula en procesar_solicitudes.
    """
    n           = limpiar_fs(nombre)
    prefijo     = "" if tipo.startswith("AGPE") else "GD "
    nombre_base = f"{prefijo}{n}"
    count       = batch_nombres.get(nombre_base, 0)
    sufijo      = f" {count}" if count else ""
    batch_nombres[nombre_base] = count + 1
    return f"{sol} - {nombre_base}{sufijo}"


def regional_de_unificado(sol: str, df: pd.DataFrame) -> str:
    C_SOL  = col(df, "SOL")
    C_TERR = col(df, "TERRITORIO")
    if not C_SOL:
        return ""
    m = df[df[C_SOL].astype(str).str.strip() == sol]
    return limpio(m.iloc[0][C_TERR]) if not m.empty and C_TERR else ""


def carpeta_regional(territorio: str) -> str | None:
    if not os.path.isdir(RUTA_DOC_PROYECTOS):
        log(f"  ⚠ DOC. PROYECTOS no encontrada: {RUTA_DOC_PROYECTOS}")
        return None
    tn = norm(territorio)
    if not tn:
        log("  ⚠ Territorio vacío")
        return None
    # 1. Mapa explícito
    nr = TERRITORIO_A_REGIONAL.get(tn)
    if nr:
        ruta = os.path.join(RUTA_DOC_PROYECTOS, nr)
        if os.path.isdir(ruta):
            log(f"  [REGIONAL] {nr}")
            return ruta
    # 2. Coincidencia normalizada
    try:
        entradas = os.listdir(RUTA_DOC_PROYECTOS)
    except OSError:
        return None
    for e in entradas:
        ruta_e = os.path.join(RUTA_DOC_PROYECTOS, e)
        if os.path.isdir(ruta_e) and tn in norm(e):
            log(f"  [REGIONAL] {e} (normalizada)")
            return ruta_e
    # 3. Primera palabra
    for e in entradas:
        ruta_e = os.path.join(RUTA_DOC_PROYECTOS, e)
        if os.path.isdir(ruta_e) and tn.split()[0] in norm(e):
            log(f"  [REGIONAL] {e} (parcial)")
            return ruta_e
    log(f"  ⚠ Regional no encontrada para '{territorio}' | Disponibles: {entradas}")
    return None


_SUBCARPETA_DOCS = Path("1. Completitud") / "1. Documentacion Inicial"


def crear_carpeta_proyecto(
    sol: str, nombre: str, tipo: str, territorio: str, batch_nombres: dict
) -> str | None:
    """Crea la carpeta en DOC. PROYECTOS / <Regional> / <SOL> - [GD] <NOMBRE>
    y las subcarpetas: 1. Completitud / 1. Documentacion Inicial
    """
    reg = carpeta_regional(territorio)
    if not reg:
        return None
    nc   = nombre_carpeta(sol, nombre, tipo, batch_nombres)
    ruta = os.path.join(reg, nc)
    sub  = Path(ruta) / _SUBCARPETA_DOCS
    if not os.path.exists(ruta):
        sub.mkdir(parents=True, exist_ok=True)
        log(f"  ✔ Carpeta creada: {nc}")
    else:
        log(f"  · Carpeta ya existe: {nc}")
        sub.mkdir(parents=True, exist_ok=True)
    return ruta


# ══════════════════════════════════════════════════════════════
# VALIDACIONES
# ══════════════════════════════════════════════════════════════

def _trafo_unificado(sol: str, df: pd.DataFrame) -> str:
    C_S = col(df, "SOL")
    C_T = col(df, "CODIGO TRANSFORMADOR", "CÓDIGO TRANSFORMADOR", "COD TRANSFORMADOR")
    if not C_S or not C_T or df.empty:
        return ""
    m = df[df[C_S].astype(str).str.strip() == sol]
    return limpio(m.iloc[0][C_T]) if not m.empty else ""


def validar(datos: dict, sol: str, df: pd.DataFrame, tipo: str) -> dict:
    """
    GD > 100  → V1 no aplica | V2 trafo presente y == Unificado | V3 pot_AC == pot_red
    AGPE > 100 → V1 NIC = 0  | V2 trafo == Unificado            | V3 pot_AC >= pot_red
    Otros      → alerta genérica de revisión manual
    """
    r   = {}
    tf  = datos.get("trafo", "").strip()
    tu  = _trafo_unificado(sol, df)
    pac = a_float(datos.get("pot_ac",  ""))
    pre = a_float(datos.get("pot_red", ""))

    if tipo == "GD > 100":
        r["V1_nic"] = "– NIC no aplica (GD)"

        # V2: transformador presente y coincide
        if not tf:
            r["V2_trafo"] = "⚠ ALERTA: Transformador vacío en formulario"
        elif not tu:
            r["V2_trafo"] = f"? Trafo '{tf}' — sin dato en Unificado"
        elif norm(tf) == norm(tu):
            r["V2_trafo"] = f"✔ Transformador coincide: {tf}"
        else:
            r["V2_trafo"] = f"⚠ ALERTA: Trafo formulario '{tf}' ≠ Unificado '{tu}'"

        # V3: exactamente iguales
        if pac is None or pre is None:
            r["V3_potencia"] = "? Potencia no legible"
        elif abs(pac - pre) < 0.01:
            r["V3_potencia"] = f"✔ Potencia AC {pac} kW = {pre} kW entregada"
        else:
            r["V3_potencia"] = f"⚠ ALERTA: Potencia AC {pac} kW ≠ {pre} kW entregada (deben ser iguales)"

    elif tipo == "AGPE > 100":
        nic = datos.get("nic", "").strip()
        nv  = a_float(nic)
        if nic == "" or nic == "0" or nv == 0.0:
            r["V1_nic"] = "✔ NIC en 0"
        else:
            r["V1_nic"] = f"⚠ ALERTA: NIC = '{nic}' (debe ser 0 para AGPE)"

        # V2
        if not tf:
            r["V2_trafo"] = "⚠ ALERTA: Transformador vacío en formulario"
        elif not tu:
            r["V2_trafo"] = f"? Trafo '{tf}' — sin dato en Unificado"
        elif norm(tf) == norm(tu):
            r["V2_trafo"] = f"✔ Transformador coincide: {tf}"
        else:
            r["V2_trafo"] = f"⚠ ALERTA: Trafo formulario '{tf}' ≠ Unificado '{tu}'"

        # V3: instalada >= entregada
        if pac is None or pre is None:
            r["V3_potencia"] = "? Potencia no legible"
        elif pac >= pre:
            r["V3_potencia"] = f"✔ Potencia AC {pac} kW ≥ {pre} kW entregada"
        else:
            r["V3_potencia"] = f"⚠ ALERTA: Potencia AC {pac} kW < {pre} kW entregada"

    else:
        # GD < 100, AGPE < 100, AGGE, vacío
        r["info"] = f"⚠ REVISAR MANUALMENTE — tipo '{tipo}' (validaciones no automáticas)"

    # Veredicto
    alertas = [v for v in r.values() if "ALERTA" in v or v.startswith("?")]
    r["veredicto"] = "✔ OK" if not alertas else "⚠ REVISAR — ver alertas en consola"
    return r


def imprimir_validaciones(sol: str, v: dict):
    log(f"  Validaciones SOL {sol}:")
    for txt in v.values():
        log(f"    {txt}")


# ══════════════════════════════════════════════════════════════
# ACTUALIZAR UNIFICADO
# ══════════════════════════════════════════════════════════════

def actualizar_unificado(sol: str, nombre: str, tipo: str):
    """
    Escribe el nombre del proyecto en NOMBRE DEL PROYECTO del Unificado local
    solo si la celda está vacía. OneDrive sube el cambio automáticamente.
    """
    import time
    from openpyxl import load_workbook

    if not os.path.isfile(_RUTA_UNIFICADO_SYNC):
        log(f"  ⚠ Unificado no disponible para escritura: {_RUTA_UNIFICADO_SYNC}")
        return
    if not nombre or nombre == "SIN_NOMBRE":
        return

    # Para GD el formato es "GD <nombre>", para AGPE solo el nombre
    valor = f"GD {nombre}" if tipo.startswith("GD") and not nombre.startswith("GD") else nombre

    for intento in range(3):
        try:
            wb = load_workbook(_RUTA_UNIFICADO_SYNC)
            ws = wb.active
            ci_sol = ci_nom = None
            for cell in ws[1]:
                if not cell.value:
                    continue
                cn = norm(str(cell.value))
                if cn == "sol":
                    ci_sol = cell.column
                if cn in ("nombre de proyecto", "nombre del proyecto", "nombre proyecto"):
                    ci_nom = cell.column
            if not ci_sol or not ci_nom:
                log("  ⚠ Columnas SOL/NOMBRE no encontradas en Unificado")
                wb.close(); return
            for fila in ws.iter_rows(min_row=2):
                if str(fila[ci_sol - 1].value or "").strip() == sol:
                    celda = ws.cell(row=fila[0].row, column=ci_nom)
                    if celda.value and str(celda.value).strip():
                        log(f"  · Unificado ya tiene nombre: '{celda.value}' — sin cambio")
                        wb.close(); return
                    celda.value = valor
                    break
            wb.save(_RUTA_UNIFICADO_SYNC)
            wb.close()
            log(f"  💾 Unificado actualizado → '{valor}'")
            return
        except PermissionError:
            if intento < 2:
                log(f"  ⏳ Unificado en uso, reintentando ({intento+1}/3)...")
                time.sleep(3)
            else:
                log("  ⚠ Unificado bloqueado — actualiza el nombre manualmente")
        except Exception as e:
            log(f"  ⚠ Error escribiendo Unificado: {e}"); return


# ══════════════════════════════════════════════════════════════
# LOGIN
# ══════════════════════════════════════════════════════════════

def iniciar_sesion(page):
    log("Iniciando sesión...")
    page.goto(PORTAL_URL, wait_until="domcontentloaded")
    page.fill(SEL_USUARIO, PORTAL_USER)
    page.fill(SEL_PASSWORD, PORTAL_PASSWORD)
    page.click(SEL_LOGIN)
    try:
        page.wait_for_load_state("networkidle", timeout=12000)
    except PlaywrightTimeoutError:
        page.wait_for_timeout(4000)
    log(f"  Sesión iniciada | {page.title()}")


# ══════════════════════════════════════════════════════════════
# CONSULTAR SOLICITUD
# ══════════════════════════════════════════════════════════════

def consultar_sol(page, sol: str) -> bool:
    try:
        page.wait_for_selector(SEL_CAMPO_SOL, timeout=12000)
        campo = page.locator(SEL_CAMPO_SOL)
        campo.click()
        campo.press("Control+A")
        campo.fill(str(sol))
        page.click(SEL_BTN_CONSULTAR)
        page.wait_for_selector(SEL_FILAS_TABLA, timeout=12000)
        return True
    except Exception as e:
        log(f"  ✘ Error consultando SOL {sol}: {e}")
        return False


def leer_fila_tabla(page, sol: str) -> dict:
    try:
        filas = page.locator(SEL_FILAS_TABLA)
        for i in range(filas.count()):
            fila = filas.nth(i)
            txt  = fila.inner_text().strip()
            if any(t in txt for t in ("No se encontraron", "No hay datos", "Ningún dato")):
                return {}
            cols = fila.locator("td")
            if cols.count() < 11:
                continue
            return {
                "ciudad":        cols.nth(4).inner_text().strip(),
                "corregimiento": cols.nth(5).inner_text().strip(),
                "vereda":        cols.nth(6).inner_text().strip(),
                "estado":        cols.nth(10).inner_text().strip(),
            }
    except Exception:
        pass
    return {}


# ══════════════════════════════════════════════════════════════
# ABRIR POPUP DEL FORMULARIO
# ══════════════════════════════════════════════════════════════

def abrir_popup(page, sol: str):
    try:
        fila    = page.locator(SEL_FILAS_TABLA).nth(0)
        botones = fila.locator("a, button, input[type='button'], input[type='submit'], img")
        if botones.count() == 0:
            return None
        with page.expect_popup(timeout=6000) as pi:
            botones.last.click()
        popup = pi.value
        popup.wait_for_load_state("domcontentloaded", timeout=12000)
        try:
            popup.wait_for_load_state("networkidle", timeout=8000)
        except PlaywrightTimeoutError:
            popup.wait_for_timeout(2000)
        return popup
    except Exception as e:
        log(f"  ⚠ Error abriendo popup SOL {sol}: {e}")
        return None


# ══════════════════════════════════════════════════════════════
# EXTRAER CAMPOS DEL FORMULARIO
# ══════════════════════════════════════════════════════════════

def _txt(popup, selector: str) -> str:
    try:
        el = popup.locator(selector)
        if el.count() == 0:
            return ""
        tag = el.first.evaluate("e => e.tagName.toLowerCase()")
        if tag in ("input", "textarea", "select"):
            try:
                return el.first.input_value().strip()
            except Exception:
                return el.first.evaluate("e => e.value || ''").strip()
        return (el.first.text_content() or "").strip()
    except Exception:
        return ""


def extraer_campos(popup, sol: str) -> dict:
    popup.wait_for_selector("#formSolicitud", state="attached", timeout=12000)
    popup.wait_for_timeout(400)
    return {
        "nombre_cliente":        _txt(popup, "#TxtNombreCli"),
        # Validación V1
        "nic":                   _txt(popup, "#TxtNCuentaCli"),
        # Nombre del proyecto
        "ciudad_proyecto":       _txt(popup, "#DdlCiudadPro"),
        "corregimiento_proyecto":_txt(popup, "#DdlCorregimientoPro"),
        "vereda_proyecto":       _txt(popup, "#DdlVeredaPro"),
        "observacion_detalle":   _txt(popup, "#TxtObservacion"),
        # Validación V2
        "trafo":                 _txt(popup, "#TxtCodTrafoPro"),
        # Validación V3
        "pot_ac":                _txt(popup, "#TxtPotenciaTotAC"),
        "pot_red":               _txt(popup, "#TxtPotenciaEntregada"),
    }


# ══════════════════════════════════════════════════════════════
# DESCARGAR ANEXOS → DOC. PROYECTOS
# ══════════════════════════════════════════════════════════════

def descargar_anexos(popup, sol: str, carpeta: str | None) -> int:
    """
    Descarga todos los archivos de #ListaAnexos a la carpeta del proyecto.
    Si la carpeta no existe usa descargas/<SOL> como fallback.
    """
    if carpeta and os.path.isdir(carpeta):
        destino = Path(carpeta) / _SUBCARPETA_DOCS
    else:
        destino = Path("descargas") / sol
    destino.mkdir(parents=True, exist_ok=True)
    log(f"  📁 Destino PDFs: {destino}")

    # Buscar #ListaAnexos; si no es visible de inmediato, hacer scroll y reintentar
    try:
        popup.wait_for_selector("#ListaAnexos", state="visible", timeout=10000)
    except Exception:
        try:
            popup.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            popup.wait_for_timeout(1000)
            popup.wait_for_selector("#ListaAnexos", state="visible", timeout=5000)
        except Exception:
            log(f"  ⚠ #ListaAnexos no encontrada para SOL {sol}")
            return 0

    # Scroll suave hasta la tabla para que sea visible en pantalla
    try:
        popup.evaluate(
            "document.querySelector('#ListaAnexos')"
            ".scrollIntoView({behavior:'smooth', block:'center'})"
        )
        popup.wait_for_timeout(600)
    except Exception:
        pass

    filas     = popup.locator(SEL_FILAS_ANEXOS)
    total     = filas.count()
    ok        = 0
    i         = 0
    num_anexo = 0  # posición del anexo según el formulario (para el prefijo de nombre)

    while i < total:
        fila = filas.nth(i)
        try:
            n_cols = fila.locator("td").count()
        except Exception:
            log(f"    · Anexo {i+1}: fila no accesible (popup cerrado?)")
            break
        if n_cols < 4:
            i += 1
            continue
        txt_f = fila.text_content() or ""
        if any(t in txt_f for t in ("No se encontraron", "No hay datos", "Ningún dato")):
            break

        num_anexo += 1  # fila válida de anexo → asignar número de orden

        # Scroll a la fila actual → el usuario ve qué archivo se está descargando
        try:
            fila.scroll_into_view_if_needed()
            popup.wait_for_timeout(200)
        except Exception:
            pass

        botones = fila.locator("a, button, input[type='button'], input[type='submit'], img")
        if botones.count() == 0:
            log(f"    · Anexo {num_anexo}: sin botón de descarga")
            i += 1
            continue

        try:
            with popup.expect_download(timeout=15000) as dl:
                botones.last.click()
            descarga = dl.value
            nombre   = limpiar_fs(descarga.suggested_filename or f"anexo_{num_anexo}", "_")
            # Prefijo de orden + nombre original del portal
            nombre_completo = f"{num_anexo:02d}_{nombre}"
            if len(str(destino / nombre_completo)) > 240:
                ext = Path(nombre).suffix
                nombre_completo = f"{num_anexo:02d}_{nombre[:55]}{ext}"
            ruta = destino / nombre_completo
            descarga.save_as(str(ruta))
            ok += 1
            log(f"    ↓ {ruta.name}")
        except PlaywrightTimeoutError:
            log(f"    · Anexo {i+1}: sin descarga (timeout)")
            # El click puede haber navegado la página — volver atrás y releer la tabla
            try:
                popup.go_back(wait_until="domcontentloaded", timeout=8000)
                # Esperar a que la tabla y sus celdas estén completamente cargadas
                popup.wait_for_selector("#ListaAnexos tbody tr td", state="visible", timeout=8000)
                popup.evaluate(
                    "document.querySelector('#ListaAnexos')"
                    ".scrollIntoView({behavior:'smooth', block:'center'})"
                )
                popup.wait_for_timeout(1000)
                filas = popup.locator(SEL_FILAS_ANEXOS)
                total = filas.count()
                log(f"    · Página restaurada ({total} filas) — continuando desde Anexo {i+2}")
            except Exception as ex:
                log(f"    · No se pudo restaurar la página: {ex}")
                popup.wait_for_timeout(1000)
        except Exception as e:
            log(f"    · Anexo {i+1} error: {e}")
        i += 1

    log(f"  Descargas: {ok}/{total}")
    return ok


# ══════════════════════════════════════════════════════════════
# REVISIÓN DOCUMENTO
# Botón azul al pie del formulario → modal con campo Observacion
# → escribir fecha d-mm-yyyy → CONTINUAR → CERRAR
#
# El portal puede mostrar DOS modales consecutivos:
#   Modal 1: solo CONTINUAR (confirmación, sin fecha)
#   Modal 2: campo Observacion + CONTINUAR → escribir fecha
# ══════════════════════════════════════════════════════════════

def revision_documento(popup, sol: str) -> str:
    """Retorna: "OK" | "SIN_BOTON" | "SIN_MODAL" | "ERROR"."""
    hoy   = date.today()
    fecha = f"{hoy.day}-{hoy.strftime('%m-%Y')}"  # sin cero inicial: "2-03-2026"
    log(f"  Revisión Documento → fecha: {fecha}")

    # ── 1. Localizar botón REVISIÓN DOCUMENTO ────────────────────────────────
    boton = None
    for sel in [
        "input[value='REVISIÓN DOCUMENTO']",
        "input[value='REVISION DOCUMENTO']",
        "button:has-text('REVISIÓN DOCUMENTO')",
        "button:has-text('REVISION DOCUMENTO')",
        "#BtnRevisionDocumento",
        "#btnRevisionDocumento",
        "input[value*='Revis']",
        "button:has-text('Revis')",
    ]:
        try:
            el = popup.locator(sel)
            if el.count() > 0:
                boton = el.first
                log(f"  · Botón encontrado: {sel}")
                break
        except Exception:
            continue

    if not boton:
        log("  ⚠ Botón REVISIÓN DOCUMENTO no encontrado")
        popup.screenshot(path=f"sin_boton_revision_{sol}.png", full_page=True)
        return "SIN_BOTON"

    try:
        boton.scroll_into_view_if_needed()

        # ── 2a. Soporte para window.prompt() (dialog JS nativo) ──────────────
        dialog_info: dict = {"handled": False}

        def _on_dialog(dialog):
            dialog_info["handled"] = True
            log(f"  · Dialog JS detectado: type={dialog.type}")
            dialog.accept(fecha if dialog.type == "prompt" else "")

        popup.once("dialog", _on_dialog)
        boton.click()
        popup.wait_for_timeout(500)

        if dialog_info["handled"]:
            log(f"  · Fecha '{fecha}' enviada vía dialog JS")
            try:
                popup.wait_for_selector("text=correctamente", state="visible", timeout=10000)
                log("  ✔ Portal confirmó cambio (dialog JS)")
            except Exception:
                pass
            for sel_c in ["button:has-text('CERRAR')", "button:has-text('Cerrar')"]:
                try:
                    el_c = popup.locator(sel_c)
                    if el_c.count() > 0 and el_c.first.is_visible():
                        el_c.first.click()
                        break
                except Exception:
                    continue
            popup.wait_for_timeout(500)
            log("  ✔ Revisión Documento completada (vía dialog JS)")
            return "OK"

        # ── 2b. Modal DOM ─────────────────────────────────────────────────────
        SELECTORES_CAMPO = [
            "#TxtObservacionCambio",
            "input[id*='bservac']",
            "textarea[id*='bservac']",
            "input[id*='Observ']",
            "textarea[id*='Observ']",
            "input[name*='bservac']",
            "input[placeholder*='bservac']",
            ".modal-body input[type='text']",
            ".modal input[type='text']",
            "[id*='modal'] input[type='text']",
            "input[type='text']:visible",
        ]

        def _buscar_campo_fecha(timeout_ms: int):
            for sel in SELECTORES_CAMPO:
                try:
                    el = popup.locator(sel)
                    el.first.wait_for(state="visible", timeout=timeout_ms)
                    if el.count() > 0:
                        log(f"  · Campo Observacion encontrado: {sel}")
                        return el.first
                except Exception:
                    continue
            return None

        campo_obs = _buscar_campo_fecha(2000)

        if not campo_obs:
            # Modal 1 de confirmación: solo CONTINUAR
            btn_conf = None
            for sel in [
                "button:has-text('CONTINUAR')",
                "button:has-text('Continuar')",
                "input[value='CONTINUAR']",
                "input[value='Continuar']",
            ]:
                try:
                    el = popup.locator(sel)
                    if el.count() > 0 and el.first.is_visible():
                        btn_conf = el.first
                        break
                except Exception:
                    continue

            if btn_conf:
                log("  · Modal de confirmación detectado → CONTINUAR (sin fecha)")
                btn_conf.click()
                popup.wait_for_timeout(500)
                campo_obs = _buscar_campo_fecha(4000)
            else:
                log("  ⚠ Ni campo Observacion ni botón CONTINUAR encontrados")
                popup.screenshot(path=f"sin_modal_{sol}.png", full_page=True)
                return "SIN_MODAL"

        if not campo_obs:
            log("  ⚠ Campo Observacion del segundo modal no encontrado")
            popup.screenshot(path=f"sin_modal_{sol}.png", full_page=True)
            return "SIN_MODAL"

        # ── 3. Escribir la fecha ──────────────────────────────────────────────
        campo_obs.click()
        campo_obs.press("Control+a")
        campo_obs.fill(fecha)

        try:
            actual = campo_obs.input_value()
            if actual != fecha:
                log(f"  · fill no escribió bien ({actual!r}), usando press_sequentially...")
                campo_obs.triple_click()
                campo_obs.press("Delete")
                campo_obs.press_sequentially(fecha, delay=50)
                actual = campo_obs.input_value()
        except Exception:
            actual = fecha
        log(f"  · Fecha escrita: '{actual}'")

        # ── 4. Clic en CONTINUAR ──────────────────────────────────────────────
        btn_cont = None
        for sel in [
            "button:has-text('CONTINUAR')",
            "button:has-text('Continuar')",
            "input[value='CONTINUAR']",
            "input[value='Continuar']",
        ]:
            try:
                el = popup.locator(sel)
                if el.count() > 0 and el.first.is_visible():
                    btn_cont = el.first
                    break
            except Exception:
                continue

        if not btn_cont:
            log("  ⚠ Botón CONTINUAR no encontrado")
            popup.screenshot(path=f"sin_continuar_{sol}.png", full_page=True)
            return "SIN_MODAL"

        btn_cont.click()
        log("  · CONTINUAR presionado — esperando confirmación...")

        # ── 5. Esperar modal de confirmación ──────────────────────────────────
        modal_confirmado = False
        for texto in [
            "Se ha cambiado correctamente el estado de la solicitud",
            "cambiado correctamente",
            "estado de la solicitud",
        ]:
            try:
                popup.wait_for_selector(f"text={texto}", state="visible", timeout=12000)
                modal_confirmado = True
                log("  ✔ Portal confirmó cambio de estado")
                break
            except Exception:
                continue

        if not modal_confirmado:
            log("  ⚠ Modal de confirmación no apareció — se continúa")
            popup.screenshot(path=f"sin_segundo_modal_{sol}.png", full_page=True)

        # ── 6. Cerrar modal ───────────────────────────────────────────────────
        btn_cerrar = None
        for sel in [
            "button:has-text('CERRAR')",
            "button:has-text('Cerrar')",
            "input[value='CERRAR']",
            "input[value='Cerrar']",
            "#BtnCerrar",
            "#btnCerrar",
        ]:
            try:
                el = popup.locator(sel)
                if el.count() > 0 and el.first.is_visible():
                    btn_cerrar = el.first
                    break
            except Exception:
                continue

        if btn_cerrar:
            btn_cerrar.click()
            popup.wait_for_timeout(500)
            log("  ✔ Modal cerrado")
        else:
            popup.wait_for_timeout(800)

        log("  ✔ Revisión Documento completada")
        return "OK"

    except Exception as e:
        log(f"  ⚠ Error en Revisión Documento: {e}")
        popup.screenshot(path=f"error_revision_{sol}.png", full_page=True)
        return "ERROR"


# ══════════════════════════════════════════════════════════════
# FLUJO PRINCIPAL POR SOLICITUD
# ══════════════════════════════════════════════════════════════

def procesar_solicitudes(page, lista: list, df_uni: pd.DataFrame) -> list:
    resultados   = []
    batch_nombres = {}  # {nombre_base: veces_usado} — determina sufijo numérico por orden del batch

    for sol in lista:
        log(f"\n{'─'*60}")
        log(f"SOL {sol}")
        log(f"{'─'*60}")

        registro = {
            "sol":                sol,
            "tipo_solicitud":     "",
            "nombre_proyecto":    "",
            "nombre_cliente":     "",
            "carpeta_proyecto":   "",
            "veredicto":          "",
            "V1_nic":             "",
            "V2_trafo":           "",
            "V3_potencia":        "",
            "revision_documento": "",
            "anexos_descargados": 0,
            "timestamp":          datetime.now().isoformat(timespec="seconds"),
            "estado_robot":       "ERROR",
        }

        # 1. Consultar
        if not consultar_sol(page, sol):
            resultados.append(registro)
            continue

        fila_tabla = leer_fila_tabla(page, sol)
        if not fila_tabla:
            registro["estado_robot"] = "SIN_RESULTADOS"
            resultados.append(registro)
            continue

        # 2. Abrir popup
        popup = abrir_popup(page, sol)
        if not popup:
            resultados.append(registro)
            continue

        try:
            # 3. Extraer campos
            datos = extraer_campos(popup, sol)

            # 4. Tipo de solicitud
            tipo = tipo_solicitud(sol, df_uni)
            log(f"  Tipo: {tipo or '(sin tipo en Unificado)'}")

            # 5. Nombre del proyecto según tipo
            if tipo.startswith("AGPE"):
                nombre = nombre_agpe(datos)
            else:
                nombre = nombre_gd(popup, datos, fila_tabla)
            log(f"  Nombre: {nombre}")

            # 6. Carpeta en DOC. PROYECTOS
            territorio = regional_de_unificado(sol, df_uni)
            carpeta    = crear_carpeta_proyecto(sol, nombre, tipo, territorio, batch_nombres)

            # 7. Actualizar Unificado con el nombre (solo si celda vacía)
            actualizar_unificado(sol, nombre, tipo)

            # 8. Validaciones → alertas en consola, no bloquean
            vals = validar(datos, sol, df_uni, tipo)
            imprimir_validaciones(sol, vals)

            # 9. Descargar anexos
            n_ok = descargar_anexos(popup, sol, carpeta)

            # Paso 10 (Revisión Documento) desactivado — se hará manualmente

            registro.update({
                "tipo_solicitud":     tipo,
                "nombre_proyecto":    nombre,
                "nombre_cliente":     datos.get("nombre_cliente", ""),
                "carpeta_proyecto":   carpeta or "",
                "veredicto":          vals.get("veredicto", ""),
                "V1_nic":             vals.get("V1_nic", ""),
                "V2_trafo":           vals.get("V2_trafo", ""),
                "V3_potencia":        vals.get("V3_potencia", ""),
                "revision_documento": "pendiente-manual",
                "anexos_descargados": n_ok,
                "timestamp":          datetime.now().isoformat(timespec="seconds"),
                "estado_robot":       "PROCESADO",
            })

        except Exception as e:
            log(f"  ✘ Error procesando SOL {sol}: {e}")
        finally:
            try:
                popup.close()
            except Exception:
                pass

        resultados.append(registro)
        log(f"  → [{registro['tipo_solicitud']}] {registro['nombre_proyecto']} | {registro['veredicto']} | Rev.Doc: {registro['revision_documento']}")
        try:
            page.wait_for_timeout(500)
        except Exception:
            pass

    return resultados


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

def main():
    # ── Solicitudes desde Excel ───────────────────────────────
    log("Cargando solicitudes...")
    solicitudes = cargar_solicitudes()
    if not solicitudes:
        log("Sin solicitudes para procesar."); return

    with sync_playwright() as p:
        # ── 1. Descargar Unificado desde SharePoint ───────────
        ruta_uni_sp = descargar_unificado_sharepoint(p)

        # ── 2. Cargar Unificado (SharePoint → local → OneDrive sync) ─
        log("Cargando Unificado...")
        ruta_uni = ruta_uni_sp or buscar_unificado()
        if not ruta_uni:
            log("⚠ Continuando sin Unificado — validaciones V2 y regional deshabilitadas.")
            df_uni = pd.DataFrame()
        else:
            log(f"  Archivo: {Path(ruta_uni).name}")
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df_uni = pd.read_excel(ruta_uni, dtype={"SOL": str, "NIC": str})
            C_SOL = col(df_uni, "SOL")
            if C_SOL:
                df_uni[C_SOL] = df_uni[C_SOL].astype(str).str.strip()
            log(f"  Unificado: {len(df_uni)} filas")

        # ── 3. Procesar solicitudes en el portal ──────────────
        browser = p.chromium.launch(headless=False, slow_mo=0)
        page    = browser.new_page()

        try:
            iniciar_sesion(page)
            resultados = procesar_solicitudes(page, solicitudes, df_uni)

            # ── Resumen final ─────────────────────────────────
            procesadas = sum(1 for r in resultados if r["estado_robot"] == "PROCESADO")
            ok_val     = sum(1 for r in resultados if r.get("veredicto", "").startswith("✔"))
            revisar    = procesadas - ok_val

            log(f"\n{'═'*60}")
            log("RESUMEN FINAL")
            log(f"{'═'*60}")
            log(f"  Solicitudes procesadas : {procesadas}/{len(solicitudes)}")
            log(f"  ✔ Validaciones OK      : {ok_val}")
            log(f"  ⚠ Revisar              : {revisar}")
            log(f"{'═'*60}")

            log("Cerrando navegador...")

        except Exception as e:
            log(f"Error general: {e}")
            try:
                page.screenshot(path="error_general.png", full_page=True)
            except Exception:
                pass
        finally:
            browser.close()


if __name__ == "__main__":
    main()
