'Acceso directo que se utiliza: CTRL + ALT + L

Set WshShell = WScript.CreateObject("WScript.Shell")
Dim f
Dim empresa
Dim salir
salir = False

empresa = inputbox("Ingrese la empresa:")
Select Case empresa
	Case "SINDICATO"
		f = 25
	Case "sindicato"
		f = 25
	Case "CONTRERAS"
		f = 35
	Case "contreras"
		f = 35
	Case "ANDRADE"
		f = 40
	Case "andrade"
		f = 40
	Case "EXOLGAN"
		f = 36
	Case "exolgan"
		f = 36
	Case "T4"
		f = 23
	Case "t4"
		f = 23
	Case "BACFSSA"
		f = 24
	Case "BACFSA"
		f = 24
	Case "bacfssa"
		f = 24
	Case "bacfsa"
		f = 24
	Case "TRP"
		f = 25
	Case "trp"
		f = 25
	Case "NASA"
		f = 12
	Case "nasa"
		f = 12
	Case "PAN"
		f = 45
	Case "pan"
		f = 45
	Case "TECNA"
		f = 13
	Case "tecna"
		f = 13
	Case "RINALDI"
		f = 14
	Case "rinaldi"
		f = 14
	Case "SOLCAN"
		f = 13
	Case "solcan"
		f = 13
	Case "AGP"
		f = 14
	Case "agp"
		f = 14
	Case "TECHINT"
		f = 20
	Case "techint"
		f = 20
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
	If Not f = "" Then
		For i = 1 To f
			WshShell.SendKeys " "
			WshShell.SendKeys "{DOWN}"
		Next
	End If
End If