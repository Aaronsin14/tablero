#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extractor del dashboard HORA A HORA - AngioDynamics (v5).
Lee la TABLA RESUMIDA de cada dia (ya agrupada por proceso, con Total/DELTA/
Meta*Dia/Downtime/Proficiency ya calculados por el Excel).

Estructura de la tabla resumida (una por dia, se detecta por 'Cantidad de estaciones'):
  Turno A: proc B(2),  estac C(3),  meta E(5),  horas F(6)..   10h,
           Total Z(26), DELTA AA(27), MetaDia AB(28), Downtime AC(29), Prof AD(30)
  Turno B: proc AK(37),estac AL(38),meta AN(40),horas AO(41)..  7h,
           Total BC(55),DELTA BD(56),MetaDia BE(57),Downtime BF(58),Prof BG(59)
  Turno C: proc BN(66),estac BO(67),meta BQ(69),horas BR(70)..  8h,
           Total CH(86),DELTA CI(87),MetaDia CJ(88),Downtime CK(89),Prof CL(90)

Nota: el turno C tiene sus encabezados una fila mas abajo que A y B.
El UPH (Meta) y Meta*Dia se leen directo, asi que si cambian en el Excel,
el dashboard los refleja automaticamente.
"""
import json, sys, time, datetime, re
from pathlib import Path
from openpyxl import load_workbook

EXCEL_PATH = Path(
    r"C:\Users\aaron.lara\OneDrive - Biomerics\BALA-CENTRAL - Hr a Hr actulizado 2026\Control_hr-hr_angio_2026_Actulizado.xlsx"
)
OUTPUT_JSON = Path(__file__).parent / "data_hora.json"
REFRESH_SECONDS = 900  # 15 min

DIAS = ["LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO","DOMINGO"]

# turno: (nombre, col_proc, col_estac, col_meta, col_hora0, n_horas,
#         col_total, col_delta, col_metaDia, col_downtime, col_prof, offset_fila)
TURNOS = [
    ("A", 2,  3,  5,  6,  10, 26, 27, 28, 29, 30, 0),
    ("B", 37, 38, 40, 41, 7,  55, 56, 57, 58, 59, 0),
    ("C", 66, 67, 69, 70, 8,  86, 87, 88, 89, 90, 1),  # C: encabezados 1 fila abajo
]

def _num(v):
    return float(v) if isinstance(v,(int,float)) else None

def clean(s): return re.sub(r'\s+',' ',str(s).strip())

def find_day_rows(ws):
    rows={}
    for r in range(1, ws.max_row+1):
        v=ws.cell(r,1).value
        if isinstance(v,str):
            up=v.strip().upper()
            for d in DIAS:
                if up.startswith(d) and d not in rows: rows[d]=r
    return rows

def find_summary_row(ws, day_row, next_day_row):
    """Fila del encabezado 'Cantidad de estaciones' de la tabla resumida del dia."""
    fin = next_day_row if next_day_row else ws.max_row
    for r in range(day_row, min(fin, ws.max_row+1)):
        v=ws.cell(r,3).value   # col C
        if isinstance(v,str) and 'cantidad' in v.lower():
            return r
    return None

def read_turno(ws, hdr, tcfg):
    (name, cproc, cestac, cmeta, chour0, nhours,
     ctot, cdelta, cmetaDia, cdown, cprof, off) = tcfg
    h = hdr + off                       # el turno C corre una fila
    # etiquetas de hora: en la misma fila del encabezado del turno
    hour_labels=[]
    for hi in range(nhours):
        c=chour0+hi*2
        hl=ws.cell(h,c).value
        hour_labels.append(clean(hl) if hl else f"H{hi+1}")

    procesos=[]
    for r in range(h+2, h+30):          # +2: saltar sub-encabezado Actual/Re-work
        proc=ws.cell(r,cproc).value
        if not proc or not str(proc).strip():
            if procesos: break
            continue
        if str(proc).strip() in ("Proceso","PROCESO"): continue

        estac   = _num(ws.cell(r,cestac).value)
        meta    = _num(ws.cell(r,cmeta).value)      # meta por hora (UPH)
        metaDia = _num(ws.cell(r,cmetaDia).value)
        total   = _num(ws.cell(r,ctot).value)
        delta   = _num(ws.cell(r,cdelta).value)
        down    = _num(ws.cell(r,cdown).value)
        prof    = _num(ws.cell(r,cprof).value)
        if meta is None and metaDia is None: continue

        hours=[]; rework=0.0
        for hi in range(nhours):
            c=chour0+hi*2
            hours.append(_num(ws.cell(r,c).value) or 0)
            rework += _num(ws.cell(r,c+1).value) or 0

        procesos.append({
            "name": clean(proc),
            "maquinas": int(estac) if estac else 0,
            "meta_hora": meta or 0,
            "meta_dia": metaDia or 0,
            "hours": hours,
            "rework": rework,
            "total": total if total is not None else sum(hours),
            "delta": delta if delta is not None else 0,
            "downtime": down if down is not None else 0,     # en minutos
            "proficiency": prof if prof is not None else 0,  # fraccion (1.25 = 125%)
        })
    return {"hour_labels":hour_labels,"procesos":procesos}

def build_payload():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"No se encontro el Excel en: {EXCEL_PATH}")
    wb=load_workbook(EXCEL_PATH, data_only=True)
    weeks={}; week_order=[]
    for sh in wb.sheetnames:
        key=sh.strip()
        if not key.upper().startswith("WW"): continue
        ws=wb[sh]
        day_rows=find_day_rows(ws)
        ordenados=sorted(day_rows.items(), key=lambda kv: kv[1])
        siguiente={d:(ordenados[i+1][1] if i+1<len(ordenados) else None)
                   for i,(d,_) in enumerate(ordenados)}
        dias={}
        for d in DIAS:
            if d not in day_rows: continue
            sr=find_summary_row(ws, day_rows[d], siguiente[d])
            if sr is None: continue
            turnos={}
            for tcfg in TURNOS:
                turnos[tcfg[0]]=read_turno(ws, sr, tcfg)
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
