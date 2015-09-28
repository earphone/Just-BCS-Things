'Debugging
	Dim debug
	debug = False
	
Do
'Global Variables
	Dim deletedFiles, count, resumeLoop
	count = 0
	resumeLoop = 0
	
'Get current filepath
	Dim WshShell, curDir
	Set WshShell = CreateObject("WScript.Shell")
	curDir = WshShell.CurrentDirectory
	
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set Folder = FSO.GetFolder(curDir)
	
	removalInput=InputBox("Enter the Extension of the Files you would like to REMOVE without the '.'" + vbNewLine + "Ex: pdf","Remove Files with ...")
	
	'Check if canceled or blank input
		If removalInput = "2" Then
			WScript.Quit
		ElseIf removalInput = "" Then
			MsgBox "There was no input" + vbNewLine + "Exiting . . ."
			WScript.Quit
		End If
		removalInput = "." + removalInput
		removalInputLength = Len(removalInput)
	deletedFiles = "DELETED THE FOLLOWING:" + vbNewLine
	For Each targetfile In Folder.files
		If Right(targetfile.name, removalInputLength) = removalInput Then
			count = count + 1
			deletedFiles = deletedFiles & targetfile.name & vbNewLine
			FSO.DeleteFile(curDir + "\" + targetfile.name)
		End If
	Next
	
	If count = 0 Then
		resumeLoop = MsgBox ("There were NO files to delete" + vbNewLine + vbNewLine + "Try Again?", 4)
	Else
		resumeLoop = MsgBox (deletedFiles + vbNewLine + vbNewLine + "Run Again?", 4)
	End If
	
	Loop While resumeLoop = 6