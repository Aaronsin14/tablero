' ============================================================
'  Actualizador de tableros AngioDynamics - arranque silencioso
'  Corre actualizar_todo.py --watch en segundo plano, sin ventana.
'  Se dispara desde el Programador de tareas de Windows.
' ============================================================

Option Explicit

Dim shell, fso, carpeta, pythonw, script, comando

Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

' carpeta donde vive este .vbs (asi no importa desde donde se llame)
carpeta = fso.GetParentFolderName(WScript.ScriptFullName)

' pythonw.exe corre Python SIN abrir ventana de consola
pythonw = "pythonw.exe"
script  = carpeta & "\actualizar_todo.py"

If Not fso.FileExists(script) Then
    MsgBox "No se encontro actualizar_todo.py en:" & vbCrLf & carpeta, 16, "Error"
    WScript.Quit 1
End If

comando = """" & pythonw & """ """ & script & """ --watch"

' 0 = ventana oculta, False = no esperar a que termine
shell.CurrentDirectory = carpeta
shell.Run comando, 0, False
