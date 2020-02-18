Set objshell = createobject("wscript.shell")
Set WshShell = WScript.CreateObject("WScript.Shell")

objshell.run "Z:\Eco2019.exe",vbhide
WshShell.AppActivate "Iniciando..."
WScript.Sleep(16000)
WshShell.AppActivate "Ingreso al Sistema Eco9"
WScript.Sleep(300)
WshShell.SendKeys "servicios16" 'Usuario "servicios16"
WScript.Sleep(300)
WshShell.SendKeys "{TAB}"
WScript.Sleep(300)
WshShell.SendKeys "unacontra" 'Contraseña "unacontra"
WScript.Sleep(300)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(7000)
WshShell.AppActivate "Eco 9"
WScript.Sleep(200)
WshShell.SendKeys "%s"
WScript.Sleep(300)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(300)
WshShell.SendKeys "{ENTER}" 'Contrato de servicios
WScript.Sleep(7000)
WshShell.AppActivate "Eco 9"
WScript.Sleep(300)
WshShell.SendKeys "%s"
WScript.Sleep(300)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(200)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}" 'Confirmacion de servicios
WScript.Sleep(300)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(7000)
WshShell.AppActivate "Eco 9"
WScript.Sleep(300)
WshShell.SendKeys "%s"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}" 'Confirmacion por fuera de hoja de ruta
WScript.Sleep(300)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(7000)
WshShell.AppActivate "Eco 9"
WScript.Sleep(300)
WshShell.SendKeys "%s"
WScript.Sleep(300)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{RIGHT}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{ENTER}" 'Listar servicios
WScript.Sleep(17000)
WshShell.AppActivate "Eco 9"
WScript.Sleep(300)
WshShell.SendKeys "%s"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(2000)
msgbox("Ya se cargo el Eco."+vbLf+vbLf+vbTab+"Guille")