"""
Descarga de anexos — Portal CREG 174 (Air-e SAS ESP)

Por cada SOL:
  1. Inicia sesión en el portal
  2. Consulta la SOL y abre el formulario (popup)
  3. Descarga todos los anexos de #ListaAnexos a descargas/<SOL>/

Uso:
  python descargar_solicitud.py 29106
  python descargar_solicitud.py 29106 29107 29108
  python descargar_solicitud.py            ← pide la SOL por consola
"""

import os
import re
import sys
from datetime import datetime
from pathlib import Path

# Forzar UTF-8 en la consola Windows para soportar tildes
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

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

# Carpeta base donde se crea una subcarpeta por SOL
CARPETA_DESCARGAS = Path(__file__).parent / "descargas"

# ── Selectores del portal ──────────────────────────────────────
SEL_USUARIO       = "#txtUsuario"
SEL_PASSWORD      = "#txtPassword"
SEL_LOGIN         = "#buttonEnviar"
SEL_CAMPO_SOL     = "#TxtIdSolicitud"
SEL_BTN_CONSULTAR = "#BtnConsultarSol"
SEL_FILAS_TABLA   = "#gridSolicitudes tbody tr"
SEL_FILAS_ANEXOS  = "#ListaAnexos tbody tr"

SIN_DATOS = ("No se encontraron", "No hay datos", "Ningún dato")


# ══════════════════════════════════════════════════════════════
# UTILIDADES
# ══════════════════════════════════════════════════════════════

def limpiar_fs(nombre: str, reemplazo: str = "") -> str:
    """Elimina caracteres inválidos para nombres de archivo/carpeta en Windows."""
    return re.sub(r'[\\/*?:"<>|]', reemplazo, str(nombre)).strip(" .") or "SIN_NOMBRE"


def log(msg: str):
    """Print con timestamp HH:MM:SS."""
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def pedir_solicitudes() -> list[str]:
    """SOLs desde argumentos de línea de comandos o, si no hay, desde consola."""
    sols = sys.argv[1:]
    if not sols:
        entrada = input("Número(s) de SOL (separados por espacio o coma): ")
        sols = re.split(r"[\s,;]+", entrada)
    # Limpiar y quitar duplicados manteniendo orden
    return list(dict.fromkeys(s.strip() for s in sols if s.strip()))


# ══════════════════════════════════════════════════════════════
# PORTAL
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
    # Si no aparece el campo de búsqueda, el login falló (credenciales, portal caído...)
    try:
        page.wait_for_selector(SEL_CAMPO_SOL, timeout=12000)
    except PlaywrightTimeoutError:
        raise RuntimeError("No se pudo iniciar sesión: revisa PORTAL_USER/PORTAL_PASSWORD en .env")
    log(f"  Sesión iniciada | {page.title()}")


def consultar_sol(page, sol: str) -> bool:
    """Busca la SOL y verifica que la tabla traiga resultados."""
    try:
        page.wait_for_selector(SEL_CAMPO_SOL, timeout=12000)
        # Texto de la tabla antes de consultar: sirve para detectar que se refrescó
        # y no leer las filas de la SOL anterior
        filas = page.locator(SEL_FILAS_TABLA)
        previo = filas.first.inner_text() if filas.count() else None

        campo = page.locator(SEL_CAMPO_SOL)
        campo.click()
        campo.press("Control+A")
        campo.fill(sol)
        page.click(SEL_BTN_CONSULTAR)
        page.wait_for_selector(SEL_FILAS_TABLA, timeout=12000)
        if previo is not None:
            page.wait_for_function(
                "([sel, prev]) => { const f = document.querySelector(sel);"
                " return f && f.innerText !== prev; }",
                arg=[SEL_FILAS_TABLA, previo],
                timeout=12000,
            )
    except Exception as e:
        log(f"  ✘ Error consultando SOL {sol}: {e}")
        return False

    txt = page.locator(SEL_FILAS_TABLA).first.inner_text()
    if any(t in txt for t in SIN_DATOS):
        log(f"  ✘ SOL {sol} sin resultados en el portal")
        return False
    return True


def abrir_popup(page, sol: str):
    try:
        fila    = page.locator(SEL_FILAS_TABLA).nth(0)
        botones = fila.locator("a, button, input[type='button'], input[type='submit'], img")
        if botones.count() == 0:
            log(f"  ✘ SOL {sol}: no hay botón para abrir el formulario")
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
        log(f"  ✘ Error abriendo formulario SOL {sol}: {e}")
        return None


def descargar_anexos(popup, sol: str, destino: Path) -> int:
    """Descarga todos los archivos de #ListaAnexos a la carpeta destino."""
    destino.mkdir(parents=True, exist_ok=True)
    log(f"  📁 Destino: {destino}")

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

    filas  = popup.locator(SEL_FILAS_ANEXOS)
    total  = filas.count()
    ok     = 0
    usados = set()  # evita que dos anexos con el mismo nombre se sobrescriban

    for i in range(total):
        fila = filas.nth(i)
        try:
            if fila.locator("td").count() < 4:
                continue
        except Exception:
            log(f"    · Anexo {i+1}: fila no accesible (popup cerrado?)")
            break
        if any(t in (fila.text_content() or "") for t in SIN_DATOS):
            break

        botones = fila.locator("a, button, input[type='button'], input[type='submit'], img")
        if botones.count() == 0:
            log(f"    · Anexo {i+1}: sin botón de descarga")
            continue

        try:
            fila.scroll_into_view_if_needed()
            with popup.expect_download(timeout=15000) as dl:
                botones.last.click()
            descarga = dl.value

            nombre = limpiar_fs(descarga.suggested_filename or f"anexo_{i+1}", "_")
            base, ext = Path(nombre).stem, Path(nombre).suffix
            n = 2
            while nombre.lower() in usados:
                nombre = f"{base}_{n}{ext}"
                n += 1
            usados.add(nombre.lower())

            ruta = destino / nombre
            descarga.save_as(str(ruta))
            ok += 1
            log(f"    ↓ {ruta.name}")
        except PlaywrightTimeoutError:
            log(f"    · Anexo {i+1}: sin descarga (timeout)")
            # El clic puede haber navegado la página — volver y releer la tabla
            try:
                popup.go_back(wait_until="domcontentloaded", timeout=8000)
                popup.wait_for_selector("#ListaAnexos tbody tr td", state="visible", timeout=8000)
                filas = popup.locator(SEL_FILAS_ANEXOS)
            except Exception as ex:
                log(f"    · No se pudo restaurar la página: {ex}")
        except Exception as e:
            log(f"    · Anexo {i+1} error: {e}")

    log(f"  Descargas: {ok}/{total}")
    return ok


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

def main():
    solicitudes = pedir_solicitudes()
    if not solicitudes:
        log("Sin solicitudes para procesar.")
        return

    resumen = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page    = browser.new_page()
        try:
            iniciar_sesion(page)
            for sol in solicitudes:
                log(f"\n{'─'*60}\nSOL {sol}\n{'─'*60}")
                resumen[sol] = "ERROR"
                if not consultar_sol(page, sol):
                    continue
                popup = abrir_popup(page, sol)
                if not popup:
                    continue
                try:
                    n = descargar_anexos(popup, sol, CARPETA_DESCARGAS / limpiar_fs(sol))
                    resumen[sol] = f"{n} archivo(s)"
                except Exception as e:
                    log(f"  ✘ Error descargando SOL {sol}: {e}")
                finally:
                    try:
                        popup.close()
                    except Exception:
                        pass
        except Exception as e:
            log(f"Error general: {e}")
        finally:
            browser.close()

    log(f"\n{'═'*60}\nRESUMEN")
    for sol, estado in resumen.items():
        log(f"  SOL {sol}: {estado}")


if __name__ == "__main__":
    main()
