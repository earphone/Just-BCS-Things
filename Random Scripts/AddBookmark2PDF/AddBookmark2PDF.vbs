'Basic Script to Add Bookmarks to PDF
'For most updated version visit:	https://github.com/earphone/Just-BCS-Things

'Debugging
	Dim debug
	debug = False

'Global Variables
	Dim initialClick
	Dim WshShell, curDir
	Dim excelFile, bookmarkWorkbook, bookmarkWorksheet
	Dim adobeAPP, chosenPDF, chosenAVDoc, chosenBookmark, pageView
	Dim row

'Set Globals
	Set adobeAPP = CreateObject("AcroExch.app")
	Set chosenPDF = CreateObject("AcroExch.PDDoc")

'Get current filepath
	Set WshShell = CreateObject("WScript.Shell")
	curDir = WshShell.CurrentDirectory

'Initial message	
	initialClick = MsgBox("Make sure that the bookmarks are input into the Excel file" + vbNewLine + "Choose the PDF file", 1)
		If initialClick = 2 Then
			wscript.quit
		End If
		
'Find file path
	Set oExec=WshShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
	sFileSelected = oExec.StdOut.ReadLine
		If sFileSelected = "" Then
			wscript.echo "No file was selected"
			wscript.quit
		End If
	If debug Then
		wscript.echo sFileSelected
	End If
	chosenPDF.Open sFileSelected
	
'Set up PDF for adding bookmarks
	Set chosenAVDoc = chosenPDF.OpenAVDoc(sFileSelected)
	Set chosenBookmark = CreateObject("AcroExch.PDBookmark")
	Set pageView = chosenAVDoc.GetAVPageView()
	
'Create Excel Object
	Set excelFile = CreateObject("Excel.Application")
	excelFile.Visible = False
	
'Setup Excel
	Set bookmarkWorkbook = excelFile.Workbooks.Open(curDir + "\AddBookmark2PDFExcel.xlsx")
	CheckError("Opening Bookmark Workbook")
	Set bookmarkWorksheet = bookmarkWorkbook.Worksheets("Bookmarks")
	CheckError("Opening Bookmark Worksheet")

'Find and create bookmarks
	row = 2
	If bookmarkWorksheet.Cells(row,1) = "" Then
		wscript.echo "Excel File was not filled in properly"
		wscript.quit
	End If
	Do
		pageView.GoTo(bookmarkWorksheet.Cells(row,1) - 1)
		adobeAPP.MenuItemExecute "NewBookmark"
		chosenBookmark.GetByTitle chosenPDF, "Untitled"
		chosenBookmark.SetTitle bookmarkWorksheet.Cells(row,2)
		row = row + 1
		CheckError("Setting bookmark for " + bookmarkWorksheet.Cells(row,2))
	Loop Until bookmarkWorksheet.Cells(row,1) = ""
	
'Close Everything
	chosenPDF.Save 0, sFileSelected
	CheckError("Saving PDF")
	chosenPDF.Close
	excelFile.Quit
	adobeAPP.Exit
	wscript.echo "Finished"
	
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