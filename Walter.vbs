Set WshShell = WScript.CreateObject("WScript.Shell")
Set oShell = CreateObject("WScript.Shell")
Dim cantidad
cantidad = inputbox("Ingrese la cantidad de repeticiones:")
Dim i
	WshShell.AppActivate "Firefox"
	WScript.Sleep(2000)
	for i = 1 to cantidad
	WScript.Sleep(100)
	oShell.SendKeys " "
	WScript.Sleep(100)
	oShell.SendKeys "{TAB}"
	next