#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Actualiza AMBOS dashboards y sube todo a GitHub.
  - extract.py       -> data.json       (produccion / legacy)
  - extract_hora.py  -> data_hora.json  (hora a hora)

USO:
  python actualizar_todo.py            (una vez)
  python actualizar_todo.py --watch    (cada 15 minutos)
"""
import subprocess, sys, time, datetime
from pathlib import Path
import extract, extract_hora

REFRESH_SECONDS = 900
HERE = Path(__file__).parent
LOG = HERE / "actualizador_log.txt"

# En Windows, cada subprocess abre su propia ventana de consola aunque el
# proceso padre corra oculto con pythonw. Esta bandera lo evita.
if sys.platform == "win32":
    NO_WINDOW = subprocess.CREATE_NO_WINDOW
else:
    NO_WINDOW = 0

def log(m):
    line=f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {m}"
    print(line)
    try:
        with open(LOG,"a",encoding="utf-8") as f: f.write(line+"\n")
    except: pass

def run(cmd):
    return subprocess.run(
        cmd,
        cwd=str(HERE),
        capture_output=True,
        text=True,
        creationflags=NO_WINDOW,   # <- sin ventana de consola
    )

def ciclo():
    # 1) generar los dos json
    try: extract.write_json()
    except Exception as e: log(f"ERROR produccion: {e}")
    try: extract_hora.write_json()
    except Exception as e: log(f"ERROR hora a hora: {e}")
    # 2) subir a github
    run(["git","pull","--no-edit","-X","ours"])
    run(["git","add","data.json","data_hora.json"])
    st=run(["git","status","--porcelain"])
    if not st.stdout.strip():
        log("Sin cambios."); return
    stamp=datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    run(["git","commit","-m",f"Actualiza datos {stamp}"])
    r=run(["git","push"])
    log("Subido a GitHub OK." if r.returncode==0 else f"ERROR push: {r.stderr.strip()}")

def main():
    watch="--watch" in sys.argv
    log("=== Actualizador DOBLE iniciado "+("(watch)" if watch else "(una vez)")+" ===")
    while True:
        try: ciclo()
        except Exception as e: log(f"ERROR: {e}")
        if not watch: break
        time.sleep(REFRESH_SECONDS)

if __name__=="__main__": main()
