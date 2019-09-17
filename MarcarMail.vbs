Set WshShell = WScript.CreateObject("WScript.Shell")

'ACTIVAR VENTANA DE OUTLOOK
WScript.Sleep(100)
WshShell.AppActivate "Bandeja de entrada - Microsoft Outlook"
WScript.Sleep(100)
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
