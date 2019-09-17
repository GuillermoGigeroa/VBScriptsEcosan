' Crear el objeto shell
Set objShell = CreateObject("Wscript.Shell")

' Ejecutamos el vbs, usamos true para que no continue el principal hasta que termine
objShell.Run("c:\scripts\script.vbs"),,True
