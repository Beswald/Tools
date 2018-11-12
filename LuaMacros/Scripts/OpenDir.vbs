' Opens directory in windows explorer.
' Full path is either passed as a string or two arguments are provided to open pre-defined locations.

Option Explicit

Dim FSO
Dim strDirectory 
Dim intDirIdx

strDirectory = WScript.Arguments.Item(0)

If strDirectory = "TABLE" Then
	Dim wshShell
	Set wshShell = CreateObject( "WScript.Shell" )

	Dim strComputerName
	strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

	' Predefined Directory Table (Work and Home)
	Dim strWorkDirs
	strWorkDirs = Array("C:\HomeDirectory","C:\GitFolder","\\File-Server\SharedDocuments")

	Dim strHomeDirs
	strHomeDirs = Array("D:\","C:\GitFolder","D:\Documents")	

	' Directory Index in the Lookup Table
	intDirIdx = WScript.Arguments.Item(1)

	If strComputerName = "WORK-PC" Then
		strDirectory = strWorkDirs(intDirIdx)
	Else
		strDirectory = strHomeDirs(intDirIdx)
	End If
End If

Set FSO = CreateObject("Scripting.FileSystemObject")

If FSO.FolderExists(strDirectory) Then
	Dim Shell
	Set Shell = CreateObject("Shell.Application")

	Shell.open "file:///" & strDirectory
Else
	MsgBox "Cannot Find Path" & vbCrlf & vbCrlf & strDirectory
End If