@echo off
title Subir Tablero a GitHub
cd /d "%~dp0"
python actualizar_y_subir.py --watch
pause
