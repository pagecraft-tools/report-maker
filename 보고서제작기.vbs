' Report maker launcher - updates itself (installed copy only), then opens the window with no console.
Set fso = CreateObject("Scripting.FileSystemObject")
appDir = fso.GetParentFolderName(WScript.ScriptFullName)
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = appDir
If fso.FileExists(appDir & "\installed.txt") Then
  sh.Run "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & appDir & "\update.ps1""", 0, True
End If
py = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python312\pythonw.exe"
If Not fso.FileExists(py) Then py = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python313\pythonw.exe"
If Not fso.FileExists(py) Then py = "pythonw"
sh.Run """" & py & """ """ & appDir & "\gui.py""", 0, False
