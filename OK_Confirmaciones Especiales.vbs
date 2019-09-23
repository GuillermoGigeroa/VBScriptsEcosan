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
		f = 47
	Case "contreras"
		f = 47
	Case "ANDRADE"
		f = 40
	Case "andrade"
		f = 40
	Case "EXOLGAN"
		f = 36
	Case "exolgan"
		f = 36
	Case "T4"
		f = 19
	Case "t4"
		f = 19
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
		f = 15
	Case "nasa"
		f = 15
	Case "PAN12"
		f = 12
	Case "pan12"
		f = 12
	Case "PAN13"
		f = 13
	Case "pan13"
		f = 13
	Case "TECNA"
		f = 13
	Case "tecna"
		f = 13
	Case "RINALDI"
		f = 22
	Case "rinaldi"
		f = 22
	Case "SOLCAN"
		f = 13
	Case "solcan"
		f = 13
	Case "AGP"
		f = 14
	Case "agp"
		f = 14
	Case "TECHINT"
		f = 25
	Case "techint"
		f = 25
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
			WshShell.SendKeys "{down}"
		Next
	End If
End If