' Displays the last TimeTrack entry

Option Explicit

Const ForReading = 1

Dim objShell
Set objShell = wscript.CreateObject("Wscript.shell")
objShell.SendKeys "{NUMLOCK}"

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim objFile

' Current Drive
Dim sDrive
sDrive = objFSO.GetDriveName(objShell.CurrentDirectory)

' Monday's Date (Current Week)
Dim MondaysDate 
MondaysDate = DateAdd("d",(WeekDay(Date())-2) * -1,Date())

' Build File Name from Date and Drive
Dim sTrackingFileName 
sTrackingFileName = sDrive & "\TimeTrack\" & Year(MondaysDate) & Right(String(2, "0") & Month(MondaysDate), 2) &  Right(String(2, "0") & Day(MondaysDate), 2)  & "_TimeLog.csv"

Set objFile = objFSO.OpenTextFile(sTrackingFileName, ForReading)

Dim strLine

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
Loop

objFile.Close



msgbox strLine

