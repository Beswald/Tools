' Opens administrative command prompt

Option Explicit

Dim objShell,exShell
Set objShell = WScript.CreateObject( "WScript.Shell" )

Set exShell = CreateObject("Shell.Application")

objShell.SendKeys "{NUMLOCK}"
exShell.ShellExecute "cmd.exe", , , "runas", 1
