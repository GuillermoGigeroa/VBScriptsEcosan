Sub ponerNoPasoElChofer
Set WshShell = WScript.CreateObject("WScript.Shell")
Dim opcion
Dim observacion
WshShell.SendKeys " "

'x Ecosan Otros - 22
'x Cliente Otros - 11
'x Cliente Lugar Inaccesible - 7
'x Cliente No habia nadie - 8
'x Ecosan Suspendido - 26
'x Cliente No permiten ingreso - 9

opcion = 22
observacion = "CAMION ROTO"
For x = 1 to opcion
	WshShell.SendKeys "{DOWN}"
Next
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys observacion 'ACA SE INGRESA LA OBSERVACION DE LA CONFIRMACION
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WScript.Sleep(200)
WshShell.SendKeys " "
WshShell.SendKeys "{DOWN}"

End sub
'============================================================================================================================================================
dim cantidad

cantidad = inputbox("Ingrese la cantidad de veces")
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.AppActivate "Confirmacion de Entregas, Servicios y Retiros"
WScript.Sleep(300)

if isnumeric(cantidad) then
	if cantidad > 0 AND cantidad < 100 then
		FOR I = 1 TO cantidad
			Call ponerNoPasoElChofer
		NEXT
	else
		msgbox("Numero incorrecto")
	end if
else
	Call ponerNoPasoElChofer
end if