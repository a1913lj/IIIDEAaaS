dim fso
dim ws
set fso = CreateObject("Scripting.FileSystemObject")
Set ws  = CreateObject("Wscript.Shell")
dim bat
bat = fso.BuildPath(fso.getParentFolderName(WScript.ScriptFullName), "IIIDEAaaMGUI.bat")
ws.run "cmd /c " & """" & bat & """", vbhide
