'Acceso directo que se utiliza: CTRL + ALT + D
Sub eliminarSeleccion
Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.SendKeys " "
WScript.Sleep(100)
WshShell.SendKeys "%"
WScript.Sleep(100)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(100)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(100)
WshShell.SendKeys "{ENTER}"
WScript.Sleep(100)
WshShell.SendKeys "{DOWN}"
WScript.Sleep(100)

End sub
'==============================================================================================================================================================
Set WshShell = WScript.CreateObject("WScript.Shell")
Dim cantidad
Dim salir
salir = False

cantidad = inputbox("Ingrese la cantidad de veces")
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
				Call eliminarSeleccion
			Next
		Else
			msgbox("Numero incorrecto")
		End If
	Else
		Call eliminarSeleccion
	End If
End If



