"""
CINSA - Actualizador del Monitor de Proyectos (versión portable, para la nube).

Equivalente a actualizar-web.ps1: lee el CSV exportado del Google Sheet
"Dashboard Proyectos" y reemplaza SOLO el bloque `let proyectos = [...]`
en index.html. No hace git; eso lo hace quien lo llama.

Uso: python actualizar_web.py <ruta_csv> [ruta_index_html]
"""
import csv
import re
import sys
from datetime import datetime
from zoneinfo import ZoneInfo

# Columnas (fila 2 del CSV es el encabezado real, los datos parten en la fila 3):
# 1 Fecha Adjudicada, 2 ID(ncontrato), 3 CC, 4 Proyecto, 5 Tipo, 6 Unidades,
# 7 Monto Contrato, 8 Plazo, 9 Inicio, 10 Termino, 11 Avance Programado,
# 12 % Avance real, 13-16 Facturado Real (por año), 17 Total Facturado, 18 % Facturado


def col(row, n):
    return row[n - 1].strip() if len(row) >= n else ""


def entero(raw):
    raw = re.sub(r"[.$ ,]", "", raw)
    return int(raw) if re.fullmatch(r"\d+", raw) else 0


def fecha(raw):
    # El CSV exporta M/D/YYYY; la web espera DD-MM-YYYY
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", raw)
    if m:
        return f"{int(m.group(2)):02d}-{int(m.group(1)):02d}-{m.group(3)}"
    return raw


def js(s):
    return s.replace("\\", "\\\\").replace("'", "\\'")


def leer_proyectos(csv_path):
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        filas = list(csv.reader(f))[2:]

    proyectos = []
    for r in filas:
        nombre = col(r, 4)
        if not nombre or nombre in ("Total Neto", "Total Bruto"):
            break

        tipo_raw = col(r, 5)
        if "Luminaria" in tipo_raw:
            tipo = "luminaria"
        elif "Televigilancia" in tipo_raw:
            tipo = "camara"
        elif "Postes" in tipo_raw:
            tipo = "poste"
        else:
            tipo = "luminaria"

        avance_raw = re.sub(r"[%\s]", "", col(r, 12))
        if re.fullmatch(r"\d+([.,]\d+)?", avance_raw):
            v = float(avance_raw.replace(",", "."))
            avance = int(v + 0.5)  # redondeo "away from zero" (siempre >= 0)
        else:
            avance = 0

        anio_raw = col(r, 1)
        plazo_raw = col(r, 8)
        inicio_raw, termino_raw = col(r, 9), col(r, 10)

        proyectos.append({
            "anio": int(anio_raw) if re.fullmatch(r"\d{4}", anio_raw) else 2025,
            "ncontrato": col(r, 2) or "—",
            "cc": col(r, 3),
            "nombre": nombre,
            "tipo": tipo,
            "estado": "completado" if avance >= 100 else "activo",
            "contrato": entero(col(r, 7)),
            "facturado": entero(col(r, 17)),
            "avance": avance,
            "plazo": plazo_raw if re.fullmatch(r"\d+", plazo_raw) else "null",
            "inicio": (fecha(inicio_raw) if inicio_raw else "") or "—",
            "termino": (fecha(termino_raw) if termino_raw else "") or "—",
        })
    return proyectos


def generar_array(proyectos):
    items = []
    for i, p in enumerate(proyectos, 1):
        cc = f"cc:'{p['cc']}', " if p["cc"] else ""
        items.append(
            f"  {{ id:{i}, {cc}anio:{p['anio']}, ncontrato:'{js(p['ncontrato'])}', "
            f"nombre:'{js(p['nombre'])}', tipo:'{p['tipo']}', estado:'{p['estado']}', "
            f"ubicacion:'{js(p['nombre'])}', contrato:{p['contrato']}, facturado:{p['facturado']}, "
            f"avance:{p['avance']}, inicio:'{p['inicio']}', termino:'{p['termino']}', plazo:{p['plazo']} }}"
        )
    hoy = datetime.now(ZoneInfo("America/Santiago")).strftime("%Y-%m-%d %H:%M")
    return f"// Actualizado automáticamente el {hoy}\nlet proyectos = [\n" + ",\n".join(items) + "\n];"


def main():
    if len(sys.argv) < 2:
        sys.exit("Uso: python actualizar_web.py <ruta_csv> [ruta_index_html]")
    csv_path = sys.argv[1]
    html_path = sys.argv[2] if len(sys.argv) > 2 else "index.html"

    proyectos = leer_proyectos(csv_path)
    if not proyectos:
        sys.exit("ERROR: no se leyó ningún proyecto del CSV. No se modifica index.html.")

    with open(html_path, encoding="utf-8", newline="") as f:
        html = f.read()

    nuevo = generar_array(proyectos)
    patron = re.compile(r"(// Actualizado automáticamente[^\n]*\n)?let proyectos = \[.*?\];", re.S)
    if len(patron.findall(html)) != 1:
        sys.exit("ERROR: no se encontró exactamente un bloque `let proyectos = [...]` en index.html.")
    html = patron.sub(lambda _: nuevo, html)

    with open(html_path, "w", encoding="utf-8", newline="") as f:
        f.write(html)
    print(f"OK: {len(proyectos)} proyectos escritos en {html_path}")


if __name__ == "__main__":
    main()
