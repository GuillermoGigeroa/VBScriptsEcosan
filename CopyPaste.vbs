Option Explicit
Dim objFSO         ‘objeto del tipo FileSystemObject para manejar el fichero
Dim sobreescribir         ‘ variable que contendra true o false.

sobreescribir = False    ‘ En mi caso no quiero que sobreescriba el fichero si existe.
Set objFSO = CreateObject(“Scripting.FileSystemObject”)

‘comprobar si existe el directorio c:\destino existe y si es asi, copiar el fichero.
If (objFSO.FolderExists(“c:\destino\”)) Then
objFSO.CopyFile “x:\fichero.bat”, “C:\destino\”, sobreescribir
End If