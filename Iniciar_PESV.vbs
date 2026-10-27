' Iniciar_PESV.vbs
' Abre la app PESV sin mostrar la ventana negra de CMD y luego abre el navegador.
' (Mensajes sin tildes a proposito: los archivos .vbs no leen bien caracteres especiales.)
Option Explicit
Dim fso, sh, carpeta, intentos

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
carpeta = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = carpeta

Function ServidorListo()
    Dim http
    ServidorListo = False
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    http.setTimeouts 1000, 1000, 1000, 1000
    http.open "GET", "http://127.0.0.1:8501/_stcore/health", False
    http.send
    If Err.Number = 0 Then
        If http.Status = 200 Then ServidorListo = True
    End If
    On Error GoTo 0
End Function

If Not ServidorListo() Then
    If Not fso.FileExists(carpeta & "\venv\Scripts\activate.bat") Then
        MsgBox "No se encontro la carpeta 'venv' junto a este archivo." & vbCrLf & _
               "Este archivo debe estar dentro de la carpeta de la app (donde esta app.py).", vbExclamation, "App PESV"
        WScript.Quit 1
    End If
    ' 0 = ventana oculta, False = no esperar a que termine
    sh.Run "cmd /c " & Chr(34) & Chr(34) & carpeta & "\_iniciar_servidor.bat" & Chr(34) & Chr(34), 0, False
    intentos = 0
    Do While (Not ServidorListo()) And intentos < 60
        WScript.Sleep 500
        intentos = intentos + 1
    Loop
    If Not ServidorListo() Then
        MsgBox "La app no arranco en 30 segundos." & vbCrLf & _
               "Abre el archivo pesv_log.txt (en la carpeta de la app) para ver el motivo.", vbCritical, "App PESV"
        WScript.Quit 1
    End If
End If

sh.Run "http://localhost:8501"
