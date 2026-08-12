"""
CINSA - Actualizador de datos de Licitaciones (Monitor de Proyectos)
Lee la copia local cacheada (CSV) de la pestaña "Listado licitaciones" del
Google Sheet "Dashboard Proyectos" (descargada por separado antes de correr
este script, igual que "Dashboard Proyectos (cache).csv" para Proyectos),
recalcula KPIs / evolucion por anio / rubro / ranking de competidores
(incluyendo el puesto de CINSA), y reemplaza el bloque `const LIC = {...};`
dentro de index.html.

Columnas esperadas en el CSV (fila 1 = titulo, fila 2 = encabezados, datos
desde fila 3): Nombre, Fecha Publicacion, Resultado CINSA, Presupuesto Bruto,
Monto Neto Adjudicado, Empresa Adjudicada, Tipo.

No toca nada mas del archivo. El commit/push lo hace actualizar-licitaciones.ps1.
"""
import json
import os
import re
import sys
import unicodedata
import pandas as pd

CSV_PATH = r"C:\Users\tira1\Listado Licitaciones (cache).csv"
HTML_PATH = r"C:\Users\tira1\index.html"
YEARS = [2023, 2024, 2025, 2026]
RUBROS = ["Luminaria", "Televigilancia", "Luminaria Solar", "Postes Inteligentes"]
CSV_COLUMNS = ["Nombre", "FechaPub", "ResultadoCINSA", "PresupuestoBruto", "MontoNetoAdj", "EmpresaAdj", "Tipo"]


def strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def build_ranking(subset):
    """Full sorted list (by n desc, by monto desc), each entry con empresa/n/monto/isCinsa/rank/total.
    CINSA siempre aparece, incluso con 0 adjudicaciones, para mostrar su ultimo lugar."""
    if len(subset) == 0:
        g = pd.DataFrame([{"empresa": "CINSA", "n": 0, "monto": 0.0}])
    else:
        g = subset.groupby("Winner").agg(n=("Winner", "count"), monto=("MontoNetoAdj", "sum")).reset_index()
        g = g.rename(columns={"Winner": "empresa"})
        if "CINSA" not in g["empresa"].values:
            g = pd.concat([g, pd.DataFrame([{"empresa": "CINSA", "n": 0, "monto": 0.0}])], ignore_index=True)
    g["isCinsa"] = g["empresa"] == "CINSA"

    by_n = g.sort_values(["n", "monto"], ascending=[False, False]).reset_index(drop=True)
    by_n["rank"] = by_n.index + 1
    by_monto = g[(g["monto"] > 0) | g["isCinsa"]].sort_values(["monto", "n"], ascending=[False, False]).reset_index(drop=True)
    by_monto["rank"] = by_monto.index + 1

    def to_list(d):
        return [
            {"empresa": r.empresa, "n": int(r.n), "monto": float(r.monto), "isCinsa": bool(r.isCinsa),
             "rank": int(r.rank), "total": int(len(d))}
            for r in d.itertuples()
        ]
    return to_list(by_n), to_list(by_monto)


def main():
    if not os.path.exists(CSV_PATH):
        print(f"ERROR: no se encontro el CSV en {CSV_PATH}")
        print("Descarga la pestana 'Listado licitaciones' de Dashboard Proyectos como CSV a esa ruta antes de correr este script.")
        sys.exit(1)

    raw = pd.read_csv(CSV_PATH, skiprows=2, header=None, names=CSV_COLUMNS, encoding="utf-8-sig",
                       dtype=str, keep_default_na=False)

    # cortar en la primera fila sin nombre (o fila de totales, por si se pega de mas)
    nombre_stripped = raw["Nombre"].str.strip()
    stop_mask = nombre_stripped.eq("") | nombre_stripped.isin(["Total Neto", "Total Bruto"])
    stop_idx = stop_mask.idxmax() if stop_mask.any() else len(raw)
    df = raw.iloc[:stop_idx].copy() if stop_mask.any() else raw.copy()

    def limpiar_monto(s):
        s = s.astype(str).str.replace(r"[\.\$\s,]", "", regex=True)
        return pd.to_numeric(s, errors="coerce").fillna(0)

    df["PresupuestoBruto"] = limpiar_monto(df["PresupuestoBruto"])
    df["MontoNetoAdj"] = limpiar_monto(df["MontoNetoAdj"])

    # el export de Google Sheets entrega las fechas como M/D/YYYY
    df["FechaPub"] = pd.to_datetime(df["FechaPub"], format="%m/%d/%Y", errors="coerce")
    df["Anio"] = df["FechaPub"].dt.year
    df = df[df["Anio"].between(2023, 2026)].copy()

    # fillna() antes de astype(str): en pandas 3.x, astype(str) sobre una
    # columna con NaN reales no siempre los convierte al texto 'nan', lo que
    # dejaba pasar filas sin resultado como si hubieran participado.
    df["ResultadoCINSA"] = df["ResultadoCINSA"].fillna("Sin dato").astype(str).str.strip()
    df.loc[df["ResultadoCINSA"].isin(["nan", "None", ""]), "ResultadoCINSA"] = "Sin dato"
    df.loc[df["ResultadoCINSA"].str.lower() == "no ofertada", "ResultadoCINSA"] = "No Ofertada"

    df["Tipo"] = df["Tipo"].fillna("Sin clasificar").astype(str).str.strip()
    df.loc[df["Tipo"].isin(["nan", "None", "", "450"]), "Tipo"] = "Sin clasificar"
    df["Tipo"] = df["Tipo"].replace({"Luminaria solar/ Televigilancia": "Luminaria/Televigilancia"})

    df["EmpresaAdj"] = df["EmpresaAdj"].astype(str).str.strip()
    df.loc[df["EmpresaAdj"].isin(["nan", "None", ""]), "EmpresaAdj"] = None
    df["EmpresaAdj"] = df["EmpresaAdj"].str.replace(r"\s+", " ", regex=True).str.strip()

    # unificar variantes de un mismo nombre que solo difieren en tildes
    mask = df["EmpresaAdj"].notna()
    norm_key = df.loc[mask, "EmpresaAdj"].apply(lambda s: strip_accents(s).upper())
    counts_per_variant = df.loc[mask, "EmpresaAdj"].value_counts()
    canonical = {}
    for _, grp in df.loc[mask, "EmpresaAdj"].groupby(norm_key):
        variants = grp.unique()
        best = max(variants, key=lambda v: counts_per_variant.get(v, 0))
        for v in variants:
            canonical[v] = best
    df.loc[mask, "EmpresaAdj"] = df.loc[mask, "EmpresaAdj"].map(canonical)
    df["MontoNetoAdj"] = pd.to_numeric(df["MontoNetoAdj"], errors="coerce").fillna(0)

    df["Winner"] = None
    df.loc[df["ResultadoCINSA"] == "Adjudicada", "Winner"] = "CINSA"
    df.loc[(df["ResultadoCINSA"] != "Adjudicada") & df["EmpresaAdj"].notna(), "Winner"] = df["EmpresaAdj"]

    participamos = df[~df["ResultadoCINSA"].isin(["No Ofertada", "Sin dato"])].copy()
    decididas = participamos[participamos["Winner"].notna()].copy()
    ganamos = df["ResultadoCINSA"] == "Adjudicada"

    kpi = {
        "total_registros": int(len(df)),
        "total_participadas": int(len(participamos)),
        "total_adjudicadas": int(ganamos.sum()),
        "pct_adjudicacion": round(float(ganamos.sum()) / len(participamos) * 100, 1),
        "presupuesto_total_participadas": float(participamos["PresupuestoBruto"].fillna(0).sum()),
        "monto_adjudicado_cinsa": float(df.loc[ganamos, "MontoNetoAdj"].fillna(0).sum()),
    }

    by_year = []
    for y in YEARS:
        dy = df[df["Anio"] == y]
        py = dy[~dy["ResultadoCINSA"].isin(["No Ofertada", "Sin dato"])]
        ay = (dy["ResultadoCINSA"] == "Adjudicada").sum()
        by_year.append({
            "anio": y, "total": int(len(dy)), "participadas": int(len(py)), "adjudicadas": int(ay),
            "pct": round(float(ay) / len(py) * 100, 1) if len(py) > 0 else 0,
            "monto_adjudicado": float(dy.loc[dy["ResultadoCINSA"] == "Adjudicada", "MontoNetoAdj"].fillna(0).sum()),
        })

    def cat_result(r):
        if r == "Adjudicada":
            return "Adjudicada CINSA"
        if r in ["Aceptada", "Ofertada"]:
            return "En evaluacion / Aceptada"
        if r in ["No Adjudicada", "Rechazada"]:
            return "No Adjudicada"
        return "Otros"

    participamos["cat"] = participamos["ResultadoCINSA"].apply(cat_result)
    stacked = participamos.groupby(["Anio", "cat"]).size().unstack(fill_value=0)
    cats = ["Adjudicada CINSA", "En evaluacion / Aceptada", "No Adjudicada", "Otros"]
    stacked_by_year = []
    for y in YEARS:
        row = {"anio": y}
        for c in cats:
            row[c] = int(stacked.loc[y, c]) if y in stacked.index and c in stacked.columns else 0
        stacked_by_year.append(row)

    by_tipo = []
    for t in participamos["Tipo"].value_counts().index.tolist():
        dt = participamos[participamos["Tipo"] == t]
        full_t = df[df["Tipo"] == t]
        aj = (full_t["ResultadoCINSA"] == "Adjudicada").sum()
        by_tipo.append({"tipo": t, "participadas": int(len(dt)), "adjudicadas": int(aj),
                         "pct": round(float(aj) / len(dt) * 100, 1) if len(dt) > 0 else 0})
    by_tipo = sorted(by_tipo, key=lambda x: -x["participadas"])

    top_n, top_monto = build_ranking(decididas)

    comp_by_tipo_n, comp_by_tipo_monto = {}, {}
    for t in RUBROS:
        n_list, m_list = build_ranking(decididas[decididas["Tipo"] == t])
        comp_by_tipo_n[t] = n_list
        comp_by_tipo_monto[t] = m_list

    comp_by_year_n, comp_by_year_monto = {}, {}
    comp_by_year_tipo_n, comp_by_year_tipo_monto = {}, {}
    for y in YEARS:
        n_list, m_list = build_ranking(decididas[decididas["Anio"] == y])
        comp_by_year_n[str(y)] = n_list
        comp_by_year_monto[str(y)] = m_list
        comp_by_year_tipo_n[str(y)] = {}
        comp_by_year_tipo_monto[str(y)] = {}
        for t in RUBROS:
            n_list2, m_list2 = build_ranking(decididas[(decididas["Anio"] == y) & (decididas["Tipo"] == t)])
            comp_by_year_tipo_n[str(y)][t] = n_list2
            comp_by_year_tipo_monto[str(y)][t] = m_list2

    lic = {
        "kpi": kpi, "by_year": by_year, "stacked_by_year": stacked_by_year, "stacked_categories": cats,
        "by_tipo": by_tipo, "top_competitors": top_n, "top_competitors_monto": top_monto,
        "comp_by_tipo": comp_by_tipo_n, "comp_by_tipo_monto": comp_by_tipo_monto,
        "comp_by_year": comp_by_year_n, "comp_by_year_monto": comp_by_year_monto,
        "comp_by_year_tipo": comp_by_year_tipo_n, "comp_by_year_tipo_monto": comp_by_year_tipo_monto,
    }

    # utf-8-sig / newline='\r\n' preservan el BOM y los finales de linea CRLF
    # que ya tiene el archivo, para que el diff de git muestre solo el bloque
    # de datos que realmente cambio.
    with open(HTML_PATH, encoding="utf-8-sig", newline="") as f:
        html = f.read()

    new_line = "const LIC = " + json.dumps(lic, ensure_ascii=False, separators=(",", ":")) + ";"
    pattern = re.compile(r"const LIC = \{.*?\};", re.S)
    if not pattern.search(html):
        print("ERROR: no se encontro el bloque 'const LIC = {...};' en index.html")
        sys.exit(1)
    html2 = pattern.sub(new_line, html, count=1)

    with open(HTML_PATH, "w", encoding="utf-8-sig", newline="") as f:
        f.write(html2)

    cinsa_rank = next((i["rank"] for i in top_n if i["isCinsa"]), "N/A")
    print(f"OK: {kpi['total_participadas']} licitaciones participadas, {kpi['total_adjudicadas']} adjudicadas "
          f"({kpi['pct_adjudicacion']}%), CINSA puesto {cinsa_rank} de {top_n[0]['total'] if top_n else 0} en el ranking general.")


if __name__ == "__main__":
    main()
