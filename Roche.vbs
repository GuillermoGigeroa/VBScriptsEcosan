Set WshShell = WScript.CreateObject("WScript.Shell")

	WScript.Sleep(3000)
'Espera 3 segundos para el cambio de ventana
for i = 1 to 10
	WshShell.SendKeys "{ENTER}"
	WshShell.SendKeys "lumberjack"
	WshShell.SendKeys "{ENTER}"
	WScript.Sleep(200)
	WshShell.SendKeys "{ENTER}"
	WshShell.SendKeys "cheese steak jimmy's"
	WshShell.SendKeys "{ENTER}"
	WScript.Sleep(200)
	WshShell.SendKeys "{ENTER}"
	WshShell.SendKeys "rock on"
	WshShell.SendKeys "{ENTER}"
	WScript.Sleep(200)
	WshShell.SendKeys "{ENTER}"
	WshShell.SendKeys "robin hood"
	WshShell.SendKeys "{ENTER}"
	WScript.Sleep(200)
next