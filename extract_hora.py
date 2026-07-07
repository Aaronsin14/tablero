#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extractor del dashboard HORA A HORA - AngioDynamics.
Lee Control_hr-hr_angio_2026_Actulizado.xlsx y genera data_hora.json.

Estructura real:
  - Cada hoja = una semana (WW28..WW52)
  - Cada dia empieza en filas: Lunes=1, Martes=104, Miercoles=207... (cada +103)
  - La tabla de METAS+PRODUCCION de cada dia esta 68 filas despues del inicio.
    Esa tabla tiene por proceso: Meta/hora + Actual/Rework por cada hora.
  - 3 turnos en bloques de columnas:
      Turno A: proceso col B(2),  meta D(4),  horas E,G,I... (10 horas)
      Turno B: proceso col AE(31),meta AG(33),horas AH...    (7 horas)
      Turno C: proceso col BC(55),meta BE(57),horas BF...    (8 horas)
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
META_OFFSET = 68  # filas desde el inicio del dia hasta su tabla de metas

# turnos: (nombre, col_proceso, col_meta, col_primera_hora, num_horas)
TURNOS = [
    ("A", 2,  4,  5,  10),
    ("B", 31, 33, 34, 7),
    ("C", 55, 57, 58, 8),
]

def _num(v):
    if isinstance(v,(int,float)): return float(v)
    return None

def clean(proc):
    return re.sub(r'\s+',' ',str(proc).strip())

def find_day_rows(ws):
    rows={}
    for r in range(1, ws.max_row+1):
        v=ws.cell(r,1).value
        if isinstance(v,str):
            up=v.strip().upper()
            for d in DIAS:
                if up.startswith(d) and d not in rows: rows[d]=r
    return rows

def read_turno(ws, meta_row, tcfg):
    name,cproc,cmeta,chour0,nhours = tcfg
    # etiquetas de hora (fila meta_row-1 tiene las horas)
    hour_labels=[]
    for hi in range(nhours):
        c=chour0+hi*2
        hl=ws.cell(meta_row-1,c).value
        hour_labels.append(str(hl).strip() if hl else f"H{hi+1}")
    procesos=[]
    for r in range(meta_row+1, meta_row+40):
        proc=ws.cell(r,cproc).value
        meta=_num(ws.cell(r,cmeta).value)
        if not proc or not str(proc).strip() or str(proc).strip()=="Proceso":
            # fin de la tabla al encontrar vacio despues de haber leido algo
            if procesos and (proc is None or str(proc).strip()==""):
                break
            continue
        if meta is None: 
            continue
        hours=[]
        rework=0.0
        for hi in range(nhours):
            c=chour0+hi*2
            a=_num(ws.cell(r,c).value) or 0
            rw=_num(ws.cell(r,c+1).value) or 0
            hours.append(a); rework+=rw
        procesos.append({
            "name":clean(proc),
            "meta_hora":meta,
            "hours":hours,
            "rework":rework,
        })
    return {"hour_labels":hour_labels,"procesos":procesos}

def build_payload():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"No se encontro el Excel en: {EXCEL_PATH}")
    wb=load_workbook(EXCEL_PATH, data_only=True)
    weeks={}; week_order=[]
    for sh in wb.sheetnames:
        if not sh.upper().startswith("WW"): continue
        ws=wb[sh]
        day_rows=find_day_rows(ws)
        dias={}
        for d in DIAS:
            if d not in day_rows: continue
            meta_row=day_rows[d]+META_OFFSET
            turnos={}
            for tcfg in TURNOS:
                turnos[tcfg[0]]=read_turno(ws, meta_row, tcfg)
            dias[d]=turnos
        weeks[sh]={"label":sh,"dias":dias}
        week_order.append(sh)
    return {
        "plant":"AngioDynamics",
        "source_file":EXCEL_PATH.name,
        "generated_at":datetime.datetime.now().isoformat(timespec="seconds"),
        "weeks":weeks,"week_order":week_order,"dias_orden":DIAS,
    }

def write_json():
    p=build_payload()
    OUTPUT_JSON.write_text(json.dumps(p,ensure_ascii=False,indent=1),encoding="utf-8")
    print(f"[{datetime.datetime.now():%H:%M:%S}] data_hora.json: {len(p['week_order'])} semanas")

def main():
    watch="--watch" in sys.argv
    while True:
        try: write_json()
        except Exception as e: print(f"[ERROR] {e}",file=sys.stderr)
        if not watch: break
        time.sleep(REFRESH_SECONDS)

if __name__=="__main__": main()
