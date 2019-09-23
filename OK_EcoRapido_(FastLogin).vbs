Set objshell = createobject("wscript.shell")
Set WshShell = WScript.CreateObject("WScript.Shell")

objshell.run "Z:\Eco2019.exe",vbhide
WshShell.AppActivate "Iniciando..."
WScript.Sleep(15000)
WshShell.AppActivate "Ingreso al Sistema Eco9"
WshShell.SendKeys "servicios16"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "unacontra"
WshShell.SendKeys "{ENTER}"
