#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extractor del dashboard HORA A HORA - AngioDynamics (v2, nueva estructura).
Lee Control_hr-hr_angio_2026_Actulizado.xlsx y genera data_hora.json.

Estructura (nueva):
  - Cada hoja = una semana (WW28..WW50). Ojo: pueden traer espacios (ej "WW28 ").
  - Cada dia empieza en col A: Lunes=1, Martes=102, Miercoles=202... (cada +100)
  - Tabla de produccion en el mismo bloque del dia (encabezados en fila dia+4).
  - 3 turnos en bloques de columnas. Cada proceso trae:
      Meta * Dia, Meta UPH (meta por hora), horas (Actual + Scrap por hora),
      DELTA y DOWN TIME al final del turno.
  - Cada celda hora se colorea vs Meta UPH; el total del proceso vs Meta*Dia.
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

# turnos: nombre, col_proc, col_metaDia, col_metaUPH, col_hora0, num_horas, col_delta, col_downtime
TURNOS = [
    ("A", 2,  3,  4,  7,  10, 30, 31),
    ("B", 34, 35, 36, 39, 7,  56, 57),
    ("C", 60, 61, 62, 65, 8,  84, 85),
]
HDR_OFFSET = 4  # filas desde inicio del dia hasta la fila de encabezados (Actual, etc.)

def _num(v):
    if isinstance(v,(int,float)): return float(v)
    return None

def clean(s): return re.sub(r'\s+',' ',str(s).strip())

def group_name(proc):
    """'Cut (Tip/Teflon) 1' -> 'Cut', 'Welder 3' -> 'Welder'."""
    p=re.sub(r'\s*\d+\s*$','',str(proc).strip())   # quita numero de maquina
    p=re.sub(r'\(.*?\)','',p).strip()               # quita parentesis
    p=re.sub(r'\s+(Soft|Accu)\s*Vu.*$','',p).strip()
    return p if p else clean(proc)

def group_name(proc):
    """Agrupa 'Cut (Tip/Teflon) 1' -> 'Cut', 'Welder 3' -> 'Welder'."""
    p=re.sub(r'\s*\d+\s*$','',str(proc).strip())     # quita numero de maquina
    p=re.sub(r'\(.*?\)','',p).strip()                 # quita parentesis
    p=re.sub(r'\s+(Soft|Accu)\s*Vu.*$','',p).strip()
    return p if p else str(proc).strip()

def group_name(proc):
    """'Cut (Tip/Teflon) 1' -> 'Cut', 'Welder 3' -> 'Welder', etc."""
    p=re.sub(r'\s*\d+\s*$','',str(proc).strip())   # quita numero de maquina
    p=re.sub(r'\(.*?\)','',p).strip()               # quita parentesis
    p=re.sub(r'\s+(Soft|Accu)\s*Vu.*$','',p).strip()
    return p if p else clean(proc)

def find_day_rows(ws):
    rows={}
    for r in range(1, ws.max_row+1):
        v=ws.cell(r,1).value
        if isinstance(v,str):
            up=v.strip().upper()
            for d in DIAS:
                if up.startswith(d) and d not in rows: rows[d]=r
    return rows

def read_turno(ws, day_row, tcfg):
    name,cproc,cmetaDia,cuph,chour0,nhours,cdelta,cdown = tcfg
    hdr = day_row + HDR_OFFSET
    # etiquetas de hora estan en la fila hdr-1
    hour_labels=[]
    for hi in range(nhours):
        c=chour0+hi*2
        hl=ws.cell(hdr-1,c).value
        hour_labels.append(clean(hl) if hl else f"H{hi+1}")
    groups={}  # nombre_grupo -> datos acumulados
    order=[]
    blancos=0
    for r in range(hdr+1, hdr+80):
        proc=ws.cell(r,cproc).value
        pstr=str(proc).strip() if proc else ""
        if pstr=="" :
            blancos+=1
            if blancos>=6 and groups: break  # fin de tabla tras varios blancos
            continue
        if pstr=="Proceso":  # sub-encabezado repetido, saltar
            continue
        blancos=0
        metaDia=_num(ws.cell(r,cmetaDia).value)
        uph=_num(ws.cell(r,cuph).value)
        if metaDia is None and uph is None: continue
        hours=[]; scrap=0.0
        for hi in range(nhours):
            c=chour0+hi*2
            a=_num(ws.cell(r,c).value) or 0
            sc=_num(ws.cell(r,c+1).value) or 0
            hours.append(a); scrap+=sc
        delta=_num(ws.cell(r,cdelta).value) or 0
        down=_num(ws.cell(r,cdown).value) or 0
        g=group_name(proc)
        if g not in groups:
            groups[g]={"name":g,"maquinas":0,"meta_dia":0.0,"meta_hora":0.0,
                       "hours":[0.0]*nhours,"scrap":0.0,"delta":0.0,"downtime":0.0}
            order.append(g)
        G=groups[g]
        G["maquinas"]+=1
        G["meta_dia"]+=(metaDia or 0)
        G["meta_hora"]+=(uph or 0)
        for hi in range(nhours): G["hours"][hi]+=hours[hi]
        G["scrap"]+=scrap
        G["delta"]+=delta
        G["downtime"]+=down
    procesos=[groups[g] for g in order]
    return {"hour_labels":hour_labels,"procesos":procesos}

def build_payload():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"No se encontro el Excel en: {EXCEL_PATH}")
    wb=load_workbook(EXCEL_PATH, data_only=True)
    weeks={}; week_order=[]
    for sh in wb.sheetnames:
        key=sh.strip()  # quitar espacios ("WW28 " -> "WW28")
        if not key.upper().startswith("WW"): continue
        ws=wb[sh]
        day_rows=find_day_rows(ws)
        dias={}
        for d in DIAS:
            if d not in day_rows: continue
            turnos={}
            for tcfg in TURNOS:
                turnos[tcfg[0]]=read_turno(ws, day_rows[d], tcfg)
            dias[d]=turnos
        weeks[key]={"label":key,"dias":dias}
        week_order.append(key)
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
