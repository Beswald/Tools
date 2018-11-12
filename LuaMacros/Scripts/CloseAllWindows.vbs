' Visual Studio: Closes all windows currently open.

Option Explicit

Dim objShell

Set objShell = wscript.CreateObject("Wscript.shell")
objShell.SendKeys "{NUMLOCK}"
objShell.SendKeys "%wl"