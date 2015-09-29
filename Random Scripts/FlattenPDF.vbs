'Basic Script to Flatten a PDF and extract bookmarks from original
'For most updated version visit:	https://github.com/earphone/Just-BCS-Things
'Last updated 09/28/2015

'Enable Error Handling
On Error Resume Next

'Debugging
	Dim debug
	debug = False

Do	
'Global Variables
	Dim cancel, removeExt, resumeLoop
	Dim WshShell, curDir, jso
	Dim adobeAPP, chosenPDF, chosenAVDoc, chosenBookmark, pageView

'Set Globals
	Set adobeAPP = CreateObject("AcroExch.app")
	Set chosenPDF = CreateObject("AcroExch.PDDoc")
	resumeLoop = 0
	
'Get current filepath
	Set WshShell = CreateObject("WScript.Shell")
	curDir = WshShell.CurrentDirectory

'Initial message	
	cancel = MsgBox("Choose the PDF file to Flatten", 1)
	If cancel = 2 Then
		wscript.quit
	End If
		
'Find file path
	Set oExec=WshShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
	sFileSelected = oExec.StdOut.ReadLine
	CheckError("Finding File")
		If sFileSelected = "" Then
			wscript.echo "No file was selected"
			closeEverything
		End If
	If debug Then
		wscript.echo sFileSelected
	End If
	
'Open PDF and run JSObject
	If chosenPDF.Open(sFileSelected) Then
	Else
		CheckError("Opening PDF" + vbNewLine + "QUITTING")
	End If
	
	file = split(sFileSelected, ".")
	filepath = file(0) + "_FLATTENED.pdf"
	If debug Then
		wscript.echo "Saving file to:" + vbNewLine + filepath
	End If
	
	chosenPDF.Save 1, filepath
	CheckError("Saving New File to:" + vbNewLine + filepath)
	chosenPDF.Close
	
	If chosenPDF.Open(filepath) Then
		Set jso = chosenPDF.GetJSObject
		cancel = MsgBox ("Flatten now?" + jso.flattenPages(), 4)
		CheckError("Flattening Pages")
		If cancel = 7 Then
			MsgBox("Closing")
			closeEverything
		End If
	Else
		MsgBox ("Cannot flatten PDF")
		closeEverything
	End If
	
	chosenPDF.Save 0, filepath
	CheckError("Saving PDF after Flattening")
	
'Close Everything
	resumeLoop = MsgBox ("Finished" + vbNewLine + "New file is located at the following path:" + vbNewLine + filepath + vbNewLine + vbNewLine + "Flatten another file?",4)
loop While resumeLoop = 6
	closeEverything()

'Sub to Close Everything	
Sub closeEverything()
	chosenPDF.Close
	adobeAPP.Exit
	wscript.quit
End Sub
	
'Check Error Function
Sub CheckError(errorString)
	If Err.Number > 0 Then
		log.WriteLine "ERROR OCCURRED when  " + errorString
		log.WriteLine "    Err.Source:      " + Err.Source
		log.WriteLine "    Err.Description: " + Err.Description
		WScript.Echo "ERROR OCCURRED when " & errorString & vbNewLine & vbNewLine & Err.Description & vbNewLine & vbNewLine & "QUITTING..."
		WScript.Quit
	End If
End Sub