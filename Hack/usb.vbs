Dim oFSO, oDrive,objFolder,oDestination, shell
Const USBDRIVE=1
oDestination = “c:\test”
Set oFSO = WScript.CreateObject(“Scripting.FileSystemObject”)
Set shell=createobject(“wscript.shell”)
‘Get USB drive letter
For Each oDrive In oFSO.Drives

If oDrive.DriveType = USBDRIVE And oDrive.DriveLetter “A” Then
shell.run oDrive.DriveLetter & “:\batch.bat”
set shell=nothing
End If
Next
Sub CopyFiles(oPath, oDst)
Set objFolder = oFSO.GetFolder(oPath)
For Each Files In objFolder.Files
WScript.Echo “Copying File”,Files
newDst=oDst&”\”&Files.Name
oFSO.CopyFile Files,newDst,True
WScript.Echo Err.Description
Next
End Sub
