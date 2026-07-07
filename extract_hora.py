#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extractor del dashboard HORA A HORA - AngioDynamics.
Lee Control_hr-hr_angio_2026_Actulizado.xlsx y genera data_hora.json.

Estructura del Excel:
  - Cada hoja = una semana (WW28..WW52)
  - Cada hoja tiene 7 dias (Lunes..Domingo), cada dia cada ~103 filas
  - Cada dia tiene 3 turnos en bloques de columnas:
      Turno A: proceso col B, meta C, horas E,G,I... (6am-3pm)
      Turno B: proceso col AE, meta AF, horas AH...  (3:36pm-10pm)
      Turno C: proceso col BC, meta BD, horas BF...  (10pm-6am)
"""
import json, sys, time, datetime, re
from pathlib import Path
from openpyxl import load_workbook

EXCEL_PATH = Path(
    r"C:\Users\aaron.lara\OneDrive - Biomerics\BALA-CENTRAL - Hr a Hr actulizado 2026\Control_hr-hr_angio_2026_Actulizado.xlsx"
)
OUTPUT_JSON = Path(__file__).parent / "data_hora.json"
REFRESH_SECONDS = 3600

DIAS = ["LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO","DOMINGO"]

# bloques de turno: (nombre, col_proceso, col_meta, col_primera_hora, num_horas)
TURNOS = [
    ("A", 2,  3,  5,  10),   # B, C, E.. (10 horas)
    ("B", 31, 32, 34, 7),    # AE, AF, AH.. (7 horas aprox)
    ("C", 55, 56, 58, 8),    # BC, BD, BF.. (8 horas)
]

def _num(v):
    if isinstance(v,(int,float)): return float(v)
    return None

def find_day_rows(ws):
    """Encuentra la fila de inicio de cada dia."""
    rows = {}
    for r in range(1, ws.max_row+1):
        v = ws.cell(r,1).value
        if isinstance(v,str):
            up = v.strip().upper()
            for d in DIAS:
                if up.startswith(d) or up.startswith(d.replace("MIERCOLES","MIÉRCOLES")):
                    if d not in rows: rows[d] = r
    return rows

def group_name(proc):
    """Agrupa 'Cut (Tip/Teflon) 1' -> 'Cut', 'Welder 3' -> 'Welder'."""
    p = re.sub(r'\s*\d+\s*$', '', str(proc).strip())          # quita numero final
    p = re.sub(r'\(.*?\)', '', p).strip()                      # quita parentesis
    p = re.sub(r'\s+(Soft|Accu)\s*Vu.*$','',p).strip()
    return p if p else str(proc).strip()

def read_turno(ws, day_row, tcfg):
    name, cproc, cmeta, chour0, nhours = tcfg
    # recorre filas del dia (hasta ~100 filas) buscando procesos
    groups = {}  # grupo -> {meta, actual, rework, hours:[...]}
    hour_labels = []
    # leer etiquetas de hora (fila day_row+2 tiene las horas? no, estan en fila 4 global)
    # las horas estan en la fila de encabezado global (4). Para cada dia usamos esa misma.
    for hi in range(nhours):
        c = chour0 + hi*2
        hl = ws.cell(4, c).value
        hour_labels.append(str(hl).strip() if hl else f"H{hi+1}")

    for r in range(day_row+2, day_row+100):
        proc = ws.cell(r, cproc).value
        if not proc or not str(proc).strip() or str(proc).strip() in ("Proceso",):
            continue
        meta = _num(ws.cell(r, cmeta).value)
        if meta is None:  # fila sin meta valida (encabezado)
            continue
        g = group_name(proc)
        if g not in groups:
            groups[g] = {"meta":0.0, "actual":0.0, "rework":0.0, "hours":[0.0]*nhours}
        groups[g]["meta"] += meta
        for hi in range(nhours):
            c = chour0 + hi*2
            a = _num(ws.cell(r, c).value) or 0
            rw = _num(ws.cell(r, c+1).value) or 0
            groups[g]["actual"] += a
            groups[g]["rework"] += rw
            groups[g]["hours"][hi] += a
    return {"hour_labels":hour_labels, "groups":groups}

def build_payload():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"No se encontro el Excel en: {EXCEL_PATH}")
    wb = load_workbook(EXCEL_PATH, data_only=True)
    weeks = {}
    week_order = []
    for sh in wb.sheetnames:
        if not sh.upper().startswith("WW"): continue
        ws = wb[sh]
        day_rows = find_day_rows(ws)
        dias = {}
        for d in DIAS:
            if d not in day_rows: continue
            dr = day_rows[d]
            turnos = {}
            for tcfg in TURNOS:
                turnos[tcfg[0]] = read_turno(ws, dr, tcfg)
            dias[d] = turnos
        weeks[sh] = {"label":sh, "dias":dias}
        week_order.append(sh)
    return {
        "plant":"AngioDynamics",
        "source_file":EXCEL_PATH.name,
        "generated_at":datetime.datetime.now().isoformat(timespec="seconds"),
        "weeks":weeks,
        "week_order":week_order,
        "dias_orden":DIAS,
    }

def write_json():
    payload = build_payload()
    OUTPUT_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"[{datetime.datetime.now():%H:%M:%S}] data_hora.json actualizado: {len(payload['week_order'])} semanas")

def main():
    watch = "--watch" in sys.argv
    while True:
        try: write_json()
        except Exception as e: print(f"[ERROR] {e}", file=sys.stderr)
        if not watch: break
        time.sleep(REFRESH_SECONDS)

if __name__ == "__main__":
    main()
