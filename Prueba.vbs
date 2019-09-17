Sub prueba
Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.SendKeys "%{TAB}"
WScript.Sleep(500)
WshShell.SendKeys "%{LEFT}"
WshShell.SendKeys "%"

End Sub
'=========================

Call prueba




