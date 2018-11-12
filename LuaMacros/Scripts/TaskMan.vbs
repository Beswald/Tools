'Launches task manager

Option Explicit

Dim objShell

Set objShell = wscript.CreateObject("wscript.shell")

objShell.Run "taskmgr.exe"
objShell.SendKeys "{NUMLOCK}"

WScript.Sleep (600)

objShell.AppActivate  "Windows Task Manager"