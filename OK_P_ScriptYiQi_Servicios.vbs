Set WshShell = WScript.CreateObject("WScript.Shell")

Dim serie

serie = inputbox("Numero de serie: ")
if isnumeric(serie) then
	if serie > 0 and serie < 1000 then

WshShell.AppActivate "Firefox"

WScript.Sleep(200)
WshShell.SendKeys "+{TAB}"
WScript.Sleep(200)
WshShell.SendKeys "{UP}"
WScript.Sleep(200)
WshShell.SendKeys "{UP}"
WScript.Sleep(200)
WshShell.SendKeys "{UP}"
WScript.Sleep(200)
WshShell.SendKeys "{UP}"
WScript.Sleep(200)
WshShell.SendKeys "{UP}"
WScript.Sleep(200)
WshShell.SendKeys "{UP}"
WScript.Sleep(200)
WshShell.SendKeys "{TAB}"

'WshShell.SendKeys "Alquiler"
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
WshShell.SendKeys "{ENTER}"
WScript.Sleep(800)
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WScript.Sleep(100)
WshShell.SendKeys "gs" 'SELECCION DE BAÑO
WScript.Sleep(100)
WshShell.SendKeys "-"
WScript.Sleep(100)
WshShell.SendKeys "eq"
WScript.Sleep(200)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(800)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(500)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(1000)
WshShell.SendKeys "{TAB}"
WScript.Sleep(100)
if serie < 10 then
	WshShell.SendKeys "B000"
else 	if serie > 9 and serie < 100 then
			WshShell.SendKeys "B00"
		else
			WshShell.SendKeys "B0"
		end if
end if
WScript.Sleep(100)
WshShell.SendKeys serie
WScript.Sleep(100)
WshShell.SendKeys "{TAB}"

WScript.Sleep(200)
WshShell.SendKeys "s"
WScript.Sleep(200)
WshShell.SendKeys "s"
WScript.Sleep(600)
WshShell.SendKeys "{TAB}"
WScript.Sleep(100)
WshShell.SendKeys "1" 'Cantidad de servicios semanales
WScript.Sleep(100)
WshShell.SendKeys "{DELETE}"
WScript.Sleep(300)
WshShell.SendKeys "{TAB}"
WScript.Sleep(400)
WshShell.SendKeys "1" 'Precio unitario
WScript.Sleep(100)
WshShell.SendKeys "{TAB}"
WScript.Sleep(400)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(100)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(100)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(100)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(100)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(100)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(100)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(100)
WshShell.SendKeys "e"
WScript.Sleep(100)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(300)
WshShell.SendKeys "{TAB}"
WScript.Sleep(300)
WshShell.SendKeys "22"
WScript.Sleep(100)
WshShell.SendKeys "-"
WScript.Sleep(100)
WshShell.SendKeys "08"
WScript.Sleep(100)
WshShell.SendKeys "-"
WScript.Sleep(100)
WshShell.SendKeys "2019"
WScript.Sleep(100)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(300)
WshShell.SendKeys "{TAB}"
WScript.Sleep(300)
WshShell.SendKeys "{TAB}"
WScript.Sleep(300)
WshShell.SendKeys "1"

	end if
end if