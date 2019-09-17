Set WshShell = WScript.CreateObject("WScript.Shell")

'ACTIVAR VENTANA DE OUTLOOK
WScript.Sleep(300)
WshShell.AppActivate "Bandeja de entrada - Microsoft Outlook"
WScript.Sleep(300)
'REENVIO EL MAIL SELECCIONADO
WshShell.SendKeys "^f"
WScript.Sleep(300)
WshShell.SendKeys "rjusto@ecosan.com.ar"
WScript.Sleep(300)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(300)
WshShell.SendKeys "^{ENTER}"
WScript.Sleep(400)
'MARCAR COMO REALIZADO
WshShell.SendKeys "+{F10}"
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
WScript.Sleep(30)
WshShell.SendKeys "{RIGHT}"
WScript.Sleep(30)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(30)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(30)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(30)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(30)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(30)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(30)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(30)
WshShell.SendKeys "{ENTER}"
'SIGO PARA ARRIBA
WScript.Sleep(30)
WshShell.SendKeys "{UP}"
