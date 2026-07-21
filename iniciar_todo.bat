@echo off
title Tableros AngioDynamics - Actualizador
cd /d "%~dp0"
python actualizar_todo.py --watch
pause
