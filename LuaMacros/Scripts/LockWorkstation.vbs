' Locks the workstation

Option Explicit

Dim objShell
Set objShell = CreateObject("Wscript.Shell")

dim sWindowsDir
sWindowsDir = objShell.ExpandEnvironmentStrings("%windir%")

objShell.Run sWindowsDir & "\System32\rundll32.exe user32.dll,LockWorkStation"