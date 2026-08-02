import sys
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import os
import re
from pathlib import Path

BASE = Path(r"C:\Users\nelson.velasquez\OneDrive - Air-e SAS ESP (1)\RES 174 2021\DOC. PROYECTOS")
PATRON_SOL = re.compile(r"^\d{5}\s+-\s+GD\s+", re.IGNORECASE)
SUB = Path("1. Completitud") / "1. Documentacion Inicial"

# Primera secuencia de 8+ digitos seguida de '_': inicio del nombre real del portal
PATRON_ID = re.compile(r"\d{8,}_")

filtro_sol = sys.argv[1].strip() if len(sys.argv) > 1 else None


def lp(p: Path) -> str:
    s = str(p.resolve())
    return s if s.startswith("\\\\") else "\\\\?\\" + s


renombrados = 0
omitidos    = 0
errores     = []

for regional in sorted(BASE.iterdir()):
    if not regional.is_dir():
        continue
    for proyecto in sorted(regional.iterdir()):
        if not proyecto.is_dir():
            continue
        if not PATRON_SOL.match(proyecto.name):
            continue
        if filtro_sol and not proyecto.name.startswith(filtro_sol):
            continue

        carpeta = proyecto / SUB
        if not carpeta.is_dir():
            continue

        for archivo in sorted(carpeta.iterdir()):
            # Usar os.path con prefijo lp() para manejar rutas > MAX_PATH
            if not os.path.isfile(lp(archivo)):
                continue
            nombre_actual = archivo.name
            m = PATRON_ID.search(nombre_actual)
            if not m:
                omitidos += 1
                continue

            nombre_nuevo = nombre_actual[m.start():]
            if nombre_nuevo == nombre_actual:
                omitidos += 1
                continue

            destino = carpeta / nombre_nuevo
            if os.path.exists(lp(destino)):
                print(f"  [OMITIDO] Ya existe: {nombre_nuevo}")
                omitidos += 1
                continue

            try:
                os.rename(lp(archivo), lp(destino))
                print(f"  OK  {regional.name}/{proyecto.name}:")
                print(f"      {nombre_actual}")
                print(f"   -> {nombre_nuevo}")
                renombrados += 1
            except Exception as e:
                errores.append(f"{nombre_actual}: {e}")

print(f"\nRenombrados : {renombrados}")
print(f"Omitidos    : {omitidos}")
if errores:
    print(f"Errores ({len(errores)}):")
    for e in errores:
        print(f"  ERROR: {e}")
else:
    print("Sin errores.")
