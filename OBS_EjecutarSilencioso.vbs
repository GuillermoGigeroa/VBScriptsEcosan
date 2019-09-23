set objshell = createobject("wscript.shell")

objshell.run "C:\Ecosan",vbhide

Dim dteWait
dteWait = DateAdd("s", 20, Now())
Do Until (Now() > dteWait)
Loop

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.SendKeys "servicios16"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "melaolvide025."
WshShell.SendKeys "{ENTER}"


objshell.run "C:\OUTLOOK.lnk",vbhide

