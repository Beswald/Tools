' Restarts the workstation

Option Explicit

Dim objShell
Set objShell = wscript.CreateObject("wscript.shell")

objShell.SendKeys "{NUMLOCK}"

if(MsgBox("Reboot PC?",vbOKCancel) = vbOK) then	
	objShell.Run "shutdown.exe /R /T 5 /C ""Reboot Initiated"" "
end if

