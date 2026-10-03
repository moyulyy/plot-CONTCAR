' Launch the CONTCAR -> PPT GUI with no console (black) window.
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(WScript.ScriptFullName)
Set ws = CreateObject("WScript.Shell")
ws.CurrentDirectory = base
pyw = "D:\miniconda3\envs\chem_env\pythonw.exe"
If Not fso.FileExists(pyw) Then pyw = "pythonw.exe"
ws.Run """" & pyw & """ """ & base & "\gui_launcher.pyw""", 0, False
