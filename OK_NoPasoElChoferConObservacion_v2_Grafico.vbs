Sub ponerNoPasoElChofer (op, obs)
'01 CANCELAR ENTREGA
'02 CANCELAR RETIRO
'03 RECHAZA ENTREGA
'04 RECHAZA RETIRO
'05 X CLIENTE CERRADO CON CANDADO
'06 X CLIENTE HORMIGONANDO
'07 X CLIENTE LLUVIAS
'08 X CLIENTE LUGAR INACCESIBLE
'09 X CLIENTE NO HABIA NADIE
'10 X CLIENTE NO PERMITEN INGRESO
'11 X CLIENTE OBRA CERRADA
'12 X CLIENTE OTRO
'13 X CLIENTE RECHAZA SERVICIO
'14 X CLIENTE TRASLADO NO NOTIFICADO
'15 X CLIENTE ZONA PELIGROSA
'16 X ECOSAN CAMION ROTO
'17 X ECOSAN FALTA DE PAGO
'18 X ECOSAN FUERA DE HORARIO
'19 X ECOSAN LLUVIA
'20 X ECOSAN LUGAR INACCESIBLE
'21 X ECOSAN NO PASO EL CHOFER
'22 X ECOSAN NO PERMITEN INGRESO
'23 X ECOSAN OTRO
'24 X ECOSAN PLAZO 96 HS
'25 X ECOSAN REALIZADO OTRO DIA
'26 X ECOSAN RECHAZA SERVICIO
'27 X ECOSAN SUSPENDIDO
'28 X ECOSAN ZONA PELIGROSA
WshShell.SendKeys " "
For x = 1 To op-1
	WshShell.SendKeys "{DOWN}"
Next
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys obs 'ACA SE INGRESA LA OBSERVACION DE LA CONFIRMACION
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WScript.Sleep(200)
WshShell.SendKeys " "
WshShell.SendKeys "{DOWN}"

End Sub
'============================================================================================================================================================
Set WshShell = WScript.CreateObject("WScript.Shell")
Dim cantidad
Dim opcion
Dim observacion
Dim salir
salir = False

opcion = inputbox("Ingrese alguna de las siguientes opciones numericas:"+vbLf+vbLf+"( 05 ) X CLIENTE CERRADO CON CANDADO"+vbLf+"( 06 ) X CLIENTE HORMIGONANDO"+vbLf+"( 07 ) X CLIENTE LLUVIAS"+vbLf+"( 08 ) X CLIENTE LUGAR INACCESIBLE"+vbLf+"( 09 ) X CLIENTE NO HABIA NADIE"+vbLf+"( 10 ) X CLIENTE NO PERMITEN INGRESO"+vbLf+"( 11 ) X CLIENTE OBRA CERRADA"+vbLf+"( 12 ) X CLIENTE OTRO"+vbLf+"( 13 ) X CLIENTE RECHAZA SERVICIO"+vbLf+"( 14 ) X CLIENTE TRASLADO NO NOTIFICADO"+vbLf+"( 15 ) X CLIENTE ZONA PELIGROSA"+vbLf+vbLf+"( 16 ) X ECOSAN CAMION ROTO"+vbLf+"( 17 ) X ECOSAN FALTA DE PAGO"+vbLf+"( 18 ) X ECOSAN FUERA DE HORARIO"+vbLf+"( 19 ) X ECOSAN LLUVIA"+vbLf+"( 20 ) X ECOSAN LUGAR INACCESIBLE"+vbLf+"( 21 ) X ECOSAN NO PASO EL CHOFER"+vbLf+"( 22 ) X ECOSAN NO PERMITEN INGRESO"+vbLf+"( 23 ) X ECOSAN OTRO"+vbLf+"( 24 ) X ECOSAN PLAZO 96 HS"+vbLf+"( 25 ) X ECOSAN REALIZADO OTRO DIA"+vbLf+"( 26 ) X ECOSAN RECHAZA SERVICIO"+vbLf+"( 27 ) X ECOSAN SUSPENDIDO"+vbLf+"( 28 ) X ECOSAN ZONA PELIGROSA")
Select Case opcion
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
If salir = False then
	While Not isnumeric(opcion) And Not opcion > 4 And Not opcion < 29
		opcion = inputbox("Ingrese alguna de las siguientes opciones numericas:"+vbLf+vbLf+"( 05 ) X CLIENTE CERRADO CON CANDADO"+vbLf+"( 06 ) X CLIENTE HORMIGONANDO"+vbLf+"( 07 ) X CLIENTE LLUVIAS"+vbLf+"( 08 ) X CLIENTE LUGAR INACCESIBLE"+vbLf+"( 09 ) X CLIENTE NO HABIA NADIE"+vbLf+"( 10 ) X CLIENTE NO PERMITEN INGRESO"+vbLf+"( 11 ) X CLIENTE OBRA CERRADA"+vbLf+"( 12 ) X CLIENTE OTRO"+vbLf+"( 13 ) X CLIENTE RECHAZA SERVICIO"+vbLf+"( 14 ) X CLIENTE TRASLADO NO NOTIFICADO"+vbLf+"( 15 ) X CLIENTE ZONA PELIGROSA"+vbLf+vbLf+"( 16 ) X ECOSAN CAMION ROTO"+vbLf+"( 17 ) X ECOSAN FALTA DE PAGO"+vbLf+"( 18 ) X ECOSAN FUERA DE HORARIO"+vbLf+"( 19 ) X ECOSAN LLUVIA"+vbLf+"( 20 ) X ECOSAN LUGAR INACCESIBLE"+vbLf+"( 21 ) X ECOSAN NO PASO EL CHOFER"+vbLf+"( 22 ) X ECOSAN NO PERMITEN INGRESO"+vbLf+"( 23 ) X ECOSAN OTRO"+vbLf+"( 24 ) X ECOSAN PLAZO 96 HS"+vbLf+"( 25 ) X ECOSAN REALIZADO OTRO DIA"+vbLf+"( 26 ) X ECOSAN RECHAZA SERVICIO"+vbLf+"( 27 ) X ECOSAN SUSPENDIDO"+vbLf+"( 28 ) X ECOSAN ZONA PELIGROSA")
	Wend
	observacion = inputbox("Ingrese la observacion a agregar:"+vbLf)
	cantidad = inputbox("Ingrese la cantidad de veces")
	Set WshShell = WScript.CreateObject("WScript.Shell")
	WshShell.AppActivate "Confirmacion de Entregas, Servicios y Retiros"
	WScript.Sleep(300)

	If isnumeric(cantidad) Then
		If cantidad > 0 And cantidad < 100 Then
			For i = 1 To cantidad
				Call ponerNoPasoElChofer (opcion, observacion)
			Next
		Else
			msgbox("Numero incorrecto")
		End If
	Else
		Call ponerNoPasoElChofer (opcion, observacion)
	End If
End If