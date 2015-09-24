'Basic Script to Flatten a PDF and extract bookmarks from original
'For most updated version visit:	https://github.com/earphone/Just-BCS-Things
'Last updated 09/24/2015

'Debugging
	Dim debug
	debug = False
	
'Global Variables
	Dim cancel, removeExt
	Dim WshShell, curDir, jso
	Dim adobeAPP, chosenPDF, chosenAVDoc, chosenBookmark, pageView

'Set Globals
	Set adobeAPP = CreateObject("AcroExch.app")
	Set chosenPDF = CreateObject("AcroExch.PDDoc")
	
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
		MsgBox ("Cannot open PDF" + vbNewLine + "Quitting")
		closeEverything
	End If
	
	file = split(sFileSelected, ".")
	filepath = file(0) + "_FLATTENED.pdf"
	If debug Then
		wscript.echo "Saving file to:" + vbNewLine + filepath
	End If
	
	chosenPDF.Save 1, filepath
	chosenPDF.Close
	
	If chosenPDF.Open(filepath) Then
		Set jso = chosenPDF.GetJSObject
		cancel = MsgBox ("Flatten now?" + jso.flattenPages(), 1)
		If cancel = 2 Then
			MsgBox("Closing")
			closeEverything
		End If
	Else
		MsgBox ("Cannot flatten PDF")
		closeEverything
	End If
	
	chosenPDF.Save 0, filepath
	
'Close Everything
	wscript.echo "Finished" + vbNewLine + "New file is located at the following path:" + vbNewLine + filepath
	closeEverything
	Sub closeEverything()
		chosenPDF.Close
		adobeAPP.Exit
		wscript.quit
	End Sub