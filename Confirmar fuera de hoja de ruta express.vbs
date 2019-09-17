Set WshShell = WScript.CreateObject("WScript.Shell")
Dim f
Dim t
Dim tipoDeBanio


f = 1
'f = 1 si es arriba, f = 3 si es abajo
t = 1
't = 1 si es B0, t = 2 si es DPF00, t = 3 si es GCBACMA00, t = 4 si es CMAEHU00

select case t
	case 1
		tipoDeBanio = "B0"
	case 2
		tipoDeBanio = "DPF00"
	case 3
		tipoDeBanio = "GCBACMA00"
	case 4
		tipoDeBanio = "CMAEHU00"
end select

WshShell.AppActivate "Confirmacion de Servicios fuera de la hoja de ruta"
WScript.Sleep(300)

for i = 1 to f
WshShell.SendKeys "{TAB}"
next
WshShell.SendKeys "{ENTER}"
WshShell.SendKeys "{LEFT}"
WshShell.SendKeys "{ENTER}"
WScript.Sleep(300)
WshShell.SendKeys "{ENTER}"
WshShell.SendKeys "{ENTER}"
WshShell.SendKeys tipoDeBanio