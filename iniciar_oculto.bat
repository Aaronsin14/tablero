@echo off
REM ============================================================
REM  Actualizador de tableros AngioDynamics - arranque oculto
REM  Lanza actualizar_todo.py --watch con pythonw (sin ventana)
REM  y se cierra. El proceso queda vivo en segundo plano.
REM ============================================================

cd /d "%~dp0"

REM pythonw.exe corre Python SIN ventana de consola
set PYW=%LOCALAPPDATA%\Microsoft\WindowsApps\pythonw.exe

if not exist "%PYW%" set PYW=%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe
if not exist "%PYW%" set PYW=%LOCALAPPDATA%\Programs\Python\Python311\pythonw.exe
if not exist "%PYW%" set PYW=pythonw.exe

start "" /b "%PYW%" "%~dp0actualizar_todo.py" --watch
exit
