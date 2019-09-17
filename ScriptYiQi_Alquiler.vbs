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

WScript.Sleep(100)
WshShell.SendKeys "Alquiler"
WScript.Sleep(100)
WshShell.SendKeys "{TAB}"
WScript.Sleep(100)
WshShell.SendKeys "{TAB}"
WScript.Sleep(100)
WshShell.SendKeys "{TAB}"
WScript.Sleep(100)
WshShell.SendKeys "gs"
WScript.Sleep(100)
WshShell.SendKeys "-"
WScript.Sleep(100)
WshShell.SendKeys "eq"
WScript.Sleep(300)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(700)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(300)
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
WshShell.SendKeys "{ENTER}"
'WshShell.SendKeys "alquiler" '32
for i = 1 to 32
	WScript.Sleep(100)
	WshShell.SendKeys "{DOWN}"
next
WshShell.SendKeys "{TAB}"
WScript.Sleep(100)
WshShell.SendKeys "{TAB}"
WScript.Sleep(100)
WshShell.SendKeys "{TAB}"
WScript.Sleep(100)
WshShell.SendKeys "{TAB}"
WScript.Sleep(100)
WshShell.SendKeys "1"
WScript.Sleep(100)
WshShell.SendKeys "{TAB}"
WScript.Sleep(200)
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{ENTER}"
WScript.Sleep(300)
WshShell.SendKeys "{TAB}"
WScript.Sleep(300)
WshShell.SendKeys "21"
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

	end if
end if