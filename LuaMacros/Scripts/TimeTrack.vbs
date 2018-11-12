' Basic time tracking script, creates (or appends) to csv to track time spent on projects

Option Explicit

Const ForReading = 1, ForAppending = 8

Dim sProjNameArg
Dim sProjectName
Dim sComments
Dim sDrive
Dim sTrackingFileName 
Dim oFSO
Dim oShell
Dim oOutPutFile

Main

Function Main
	Dim wshShell
	Set wshShell = CreateObject( "WScript.Shell" )
	Dim strComputerName
	strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

	If WScript.Arguments.Count > 0 Then
		
		sProjNameArg = WScript.Arguments.Item(0)
		sComments = ""
		
		If( (sProjNameArg = "IN") Or (sProjNameArg = "OUT") ) Then		
		
			Dim CurrentHour			
			CurrentHour = Hour(Now())
			
			If ((CurrentHour > 8) And (CurrentHour < 16)) Then
				Exit Function
			End If
		End If
	End If

	If sProjNameArg <> "" Then
		'Project Name Passed as Argument
		sProjectName = sProjNameArg
	Else
		'Get Project Name From Input
		sProjectName = InputBox("Project Name?", "Time Tracking")
		
		' Special cases
		If UCase(sProjectName) = "ADMIN" Then
			sProjectName = "Admin"
			sComments = InputBox("Entry Comments", "Comments") 
		ElseIf UCase(sProjectName) = "MS" Then
			sProjectName = "Mfr Support"
			sComments = InputBox("Entry Comments", "Comments") 
		End If
	End If


	If sProjectName <> "" Then
		'Determine File Name and Location
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set oShell = CreateObject("Wscript.Shell")

		If strComputerName = "ENG-WKS-10" Then
			sDrive = oFSO.GetDriveName(oShell.CurrentDirectory)
		Else
			sDrive = "D:\"
		End If

		If (Not oFSO.FolderExists(sDrive & "\TimeTrack\")) Then
			 oFSO.CreateFolder(sDrive & "\TimeTrack\")
		End If
		
		Dim MondaysDate 
		MondaysDate = DateAdd("d",(WeekDay(Date())-2) * -1,Date())
		
		sTrackingFileName = sDrive & "\TimeTrack\" & Year(MondaysDate) & Right(String(2, "0") & Month(MondaysDate), 2) &  Right(String(2, "0") & Day(MondaysDate), 2)  & "_TimeLog.csv"

		' Write Time stamp and File Name to tracking File
		dim sLogEntry
		sLogEntry =  Now() & "," & sProjectName
		
		If (sComments <> "") Then
			sLogEntry = sLogEntry & ",""" & sComments & """"
		End If
		
		Set oOutPutFile = oFSO.OpenTextFile( sTrackingFileName, ForAppending, True )
		oOutPutFile.WriteLine sLogEntry
		oOutPutFile.Close
	End If

End Function