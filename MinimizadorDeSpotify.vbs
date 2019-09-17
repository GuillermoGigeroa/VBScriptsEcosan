Set WshShell = WScript.CreateObject("WScript.Shell")
Set oShell = CreateObject("WScript.Shell")
Dim verdadero
verdadero = true
'While verdadero
	WScript.Sleep(5000)
	WshShell.AppActivate "Spotify"
	WScript.Sleep(500)
	oShell.SendKeys "% n"
	WScript.Sleep(300)
	oShell.SendKeys "n"
	'oShell.SendKeys "{DOWN}"
	'oShell.SendKeys "{DOWN}"
	'oShell.SendKeys "{DOWN}"
	'WScript.Sleep(300)
	'oShell.SendKeys "{ENTER}"
	WScript.Sleep(300000)
'Wend