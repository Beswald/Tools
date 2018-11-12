' Pops up a file open dialog box, opens the selected file in Notepad++

Option Explicit

Dim FSO
Dim strNotePadPath
Dim strFile
Dim sArg

sArg = ""

If WScript.Arguments.Count > 0 Then			
	sArg = WScript.Arguments.Item(0)
End If
	
Main(sArg)


Function Main(ByVal sMode)
	Dim objShell
	Set objShell = WScript.CreateObject( "WScript.Shell" )
	Set FSO = CreateObject("Scripting.FileSystemObject")	
	strNotePadPath = "C:\Program Files (x86)\Notepad++\notepad++.exe"
	
	If (Not fso.FileExists(strNotePadPath)) Then
		
		strNotePadPath = "C:\Program Files\Notepad++\notepad++.exe" 
		
		If (Not fso.FileExists(strNotePadPath)) Then
			MsgBox "NotePad++ Not Found"
			Exit Function
		End If
	End If
	
	If (sMode = "BLANK") Then			
		objShell.SendKeys "{NUMLOCK}"
		objShell.Run("""" & strNotePadPath & """")
	Else
		strFile = SelectFile( )
	
		If strFile <> "" Then 			
			objShell.Run("""" & strNotePadPath & """ " & strFile)
			Set objShell = Nothing	
		End If
	End If
	


End Function

Function SelectFile( )
    ' File Browser via HTA
    ' Author:   Rudi Degrande, modifications by Denis St-Pierre and Rob van der Woude
    ' Features: Works in Windows Vista and up (Should also work in XP).
    '           Fairly fast.
    '           All native code/controls (No 3rd party DLL/ XP DLL).
    ' Caveats:  Cannot define default starting folder.
    '           Uses last folder used with MSHTA.EXE stored in Binary in [HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32].
    '           Dialog title says "Choose file to upload".
    ' Source:   https://social.technet.microsoft.com/Forums/scriptcenter/en-US/a3b358e8-15ae-4ba3-bca5-ec349df65ef6/windows7-vbscript-open-file-dialog-box-fakepath?forum=ITCG

    Dim objExec, strMSHTA, wshShell

    SelectFile = ""

    ' For use in HTAs as well as "plain" VBScript:
    strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
             & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
             & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
    ' For use in "plain" VBScript only:
    ' strMSHTA = "mshta.exe ""about:<input type=file id=FILE>" _
    '          & "<script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
    '          & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>"""

    Set wshShell = CreateObject( "WScript.Shell" )
    Set objExec = wshShell.Exec( strMSHTA )

    SelectFile = objExec.StdOut.ReadLine( )

    Set objExec = Nothing
    Set wshShell = Nothing
End Function