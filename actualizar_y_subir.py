#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera data.json desde el Excel y lo SUBE a GitHub automaticamente.
Pensado para correr cada 15 hora en la compu que tiene el Excel.

Requisitos:
  - git instalado y configurado (git config user.name / user.email)
  - el repo ya clonado en esta carpeta (ver LEEME_GITHUB.txt)
  - pip install openpyxl

USO:
  python actualizar_y_subir.py            (una vez)
  python actualizar_y_subir.py --watch    (cada hora, deja la ventana abierta)
"""

import subprocess
import sys
import time
import datetime
from pathlib import Path

import extract  # tu logica de extraccion (lee el Excel -> data.json)

REFRESH_SECONDS = 3600  # 15 hora
HERE = Path(__file__).parent


def run(cmd):
    return subprocess.run(cmd, cwd=str(HERE), capture_output=True, text=True)


def push_to_github():
    # 1) regenerar data.json desde el Excel
    extract.write_json()

    # 2) commit + push solo si cambio algo
    run(["git", "add", "data.json"])
    status = run(["git", "status", "--porcelain"])
    if not status.stdout.strip():
        print(f"[{datetime.datetime.now():%H:%M:%S}] Sin cambios, nada que subir.")
        return

    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    run(["git", "commit", "-m", f"Actualiza datos {stamp}"])
    r = run(["git", "push"])
    if r.returncode == 0:
        print(f"[{datetime.datetime.now():%H:%M:%S}] Datos subidos a GitHub OK.")
    else:
        print(f"[ERROR git push] {r.stderr}", file=sys.stderr)


def main():
    watch = "--watch" in sys.argv
    while True:
        try:
            push_to_github()
        except Exception as e:
            print(f"[ERROR] {e}", file=sys.stderr)
        if not watch:
            break
        time.sleep(REFRESH_SECONDS)


if __name__ == "__main__":
    main()
