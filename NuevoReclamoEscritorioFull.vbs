Dim objShell

Set objShell = CreateObject("Wscript.Shell")

objShell.Run "C:\Reclamos2018.lnk"

Dim dteWait
dteWait = DateAdd("s", 1, Now())
Do Until (Now() > dteWait)
Loop

objShell.Run "C:\NuevoReclamo.vbs" ,vbhide
