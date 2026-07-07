#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extractor del dashboard HORA A HORA - AngioDynamics.
Lee Control_hr-hr_angio_2026_Actulizado.xlsx y genera data_hora.json.

Cada celda = lo que produjo un proceso en una hora, comparado con su META POR HORA
(tabla fija en fila 69: Cut=200, Welder=480, etc.).
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

# turnos: (nombre, col_proceso_captura, col_primera_hora, num_horas, col_proceso_metas, col_meta)
# metas estan en fila 69+ ; captura en fila dia+2..
TURNOS = [
    ("A", 2,  5,  10, 2,  4),    # Turno A
    ("B", 31, 34, 7,  31, 33),   # Turno B
    ("C", 55, 58, 8,  55, 57),   # Turno C
]
META_ROW_HDR = 69  # fila del encabezado de la tabla de metas

def _num(v):
    if isinstance(v,(int,float)): return float(v)
    return None

def norm(proc):
    """Normaliza nombre para hacer match entre captura y tabla de metas."""
    p=str(proc).strip().lower()
    p=re.sub(r'\s+',' ',p)
    p=re.sub(r'\s*\d+\s*$','',p)      # quita numero de maquina al final
    p=re.sub(r'\(.*?\)','',p).strip()
    return p

def group_name(proc):
    p=re.sub(r'\s*\d+\s*$','',str(proc).strip())
    p=re.sub(r'\(.*?\)','',p).strip()
    p=re.sub(r'\s+(Soft|Accu)\s*Vu.*$','',p).strip()
    return p if p else str(proc).strip()

def find_day_rows(ws):
    rows={}
    for r in range(1, ws.max_row+1):
        v=ws.cell(r,1).value
        if isinstance(v,str):
            up=v.strip().upper()
            for d in DIAS:
                if up.startswith(d) and d not in rows: rows[d]=r
    return rows

def read_metas(ws, cproc, cmeta):
    """Lee la tabla de metas por hora (fila 69+). Devuelve {grupo: meta_por_hora}."""
    metas={}
    for r in range(META_ROW_HDR+1, META_ROW_HDR+40):
        proc=ws.cell(r,cproc).value
        meta=_num(ws.cell(r,cmeta).value)
        if proc and meta is not None:
            metas[group_name(proc)]=meta
    return metas

def read_turno(ws, day_row, tcfg):
    name,cproc,chour0,nhours,cmproc,cmeta = tcfg
    metas=read_metas(ws,cmproc,cmeta)
    hour_labels=[]
    for hi in range(nhours):
        c=chour0+hi*2
        hl=ws.cell(4,c).value
        hour_labels.append(str(hl).strip() if hl else f"H{hi+1}")
    # agrupar produccion por hora
    groups={}  # grupo -> {meta_hora, hours:[act por hora], rework:total}
    for r in range(day_row+2, day_row+63):
        proc=ws.cell(r,cproc).value
        if not proc or not str(proc).strip() or str(proc).strip()=="Proceso": continue
        g=group_name(proc)
        if g not in groups:
            groups[g]={"meta_hora":metas.get(g,0), "hours":[0.0]*nhours, "rework":0.0}
        for hi in range(nhours):
            c=chour0+hi*2
            a=_num(ws.cell(r,c).value) or 0
            rw=_num(ws.cell(r,c+1).value) or 0
            groups[g]["hours"][hi]+=a
            groups[g]["rework"]+=rw
    return {"hour_labels":hour_labels,"groups":groups}

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
            turnos={}
            for tcfg in TURNOS:
                turnos[tcfg[0]]=read_turno(ws, day_rows[d], tcfg)
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
