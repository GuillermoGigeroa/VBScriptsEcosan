'Acceso directo que se utiliza: CTRL + ALT + A

Sub ponerNoPasoElChofer

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.SendKeys " "
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
	WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{TAB}"
WScript.Sleep(200)
WshShell.SendKeys " "
WshShell.SendKeys "{DOWN}"

End sub
'==============================================================================================================================================================
Set WshShell = WScript.CreateObject("WScript.Shell")
Dim cantidad
Dim salir
salir = False

cantidad = inputbox("Ingrese la cantidad de veces"+vbLf+vbLf+"X ECOSAN - NO PASO EL CHOFER")
Select Case cantidad
	Case "Salir"
		salir = True
	Case "salir"
		salir = True
	Case "SALIR"
		salir = True
	Case "Exit"
		salir = True
	Case "exit"
		salir = True
	Case "EXIT"
		salir = True
	Case "S"
		salir = True
End Select
If salir = False Then
	WshShell.AppActivate "Confirmacion de Entregas, Servicios y Retiros"
	WScript.Sleep(300)

	If isnumeric(cantidad) Then
		If cantidad > 0 Then
			For I = 1 TO cantidad
				Call ponerNoPasoElChofer
			Next
		Else
			msgbox("Numero incorrecto")
		End If
	Else
		Call ponerNoPasoElChofer
	End If
End If



