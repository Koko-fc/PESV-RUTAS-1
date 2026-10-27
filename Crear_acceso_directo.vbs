' Crear_acceso_directo.vbs
' Crea el acceso directo "App PESV" (con icono) en el Escritorio. Se ejecuta una sola vez.
Option Explicit
Dim fso, sh, carpeta, escritorio, lnk

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
carpeta = fso.GetParentFolderName(WScript.ScriptFullName)
escritorio = sh.SpecialFolders("Desktop")

Set lnk = sh.CreateShortcut(escritorio & "\App PESV.lnk")
lnk.TargetPath = "wscript.exe"
lnk.Arguments = Chr(34) & carpeta & "\Iniciar_PESV.vbs" & Chr(34)
lnk.WorkingDirectory = carpeta
lnk.IconLocation = carpeta & "\icon.ico,0"
lnk.Description = "App PESV - Rutas frecuentes y riesgos viales"
lnk.Save

MsgBox "Listo. Se creo el acceso directo 'App PESV' en tu Escritorio." & vbCrLf & _
       "Haz doble clic en el para abrir la app.", vbInformation, "App PESV"
