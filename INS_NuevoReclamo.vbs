Option Explicit
Dim objFSO
Dim sobreescribir
Dim nombre
Dim numero
Dim nombrenumero
dim i


nombre = "FOR-PRO-15 FORMULARIO RECLAMOS REV 56 - Nº "
numero = "9999"

nombrenumero = nombre&numero&".xlsx"

sobreescribir = False

Set objFSO = CreateObject("Scripting.FileSystemObject")

If (objFSO.FolderExists("C:\Users\servicios11\Documents\Reclamos 2018\")) Then
objFSO.CopyFile "C:\Users\servicios11\Documents\Reclamos 2018\FOR-PRO-15 FORMULARIO RECLAMOS REV 56 - Nº 0RIGINAL.xlsx", ("C:\Users\servicios11\Documents\Reclamos 2018\"&nombrenumero), sobreescribir
End If

Dim WshShell 

Set WshShell = WScript.CreateObject("WScript.Shell")

Dim dteWait
dteWait = DateAdd("s", 1, Now())
Do Until (Now() > dteWait)
Loop

for i = 1 to 30
WshShell.SendKeys "{DOWN}"
next

'WshShell.SendKeys "{UP}"

WshShell.SendKeys "{F2}"

for i = 1 to 15
WshShell.SendKeys "{RIGHT}"
next

for i = 1 to 4
WshShell.SendKeys "{BACKSPACE}"
next


 