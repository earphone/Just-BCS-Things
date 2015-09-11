'Debugging
	Dim debug
	debug = False

'Global Variables
	Dim rowNumber
	Dim itemNumber
	Dim files, fileCount
	Dim currentPage
'Completed PDF Setup
	Dim completedAPP, completedPDF
	Set completedAPP = CreateObject("AcroExch.app")
	Set completedPDF = CreateObject("AcroExch.PDDoc")
	Set tempPDF = CreateObject("AcroExch.PDDoc")

'Initialize Variables
	itemNumber = 5
	rowNumber = 12
	currentPage = 2

'Get the initial information
	'1 if ok; 2 is cancel
	shortTitle=InputBox("Enter the Short Project Title","Short Title","Short Title")
	longTitle=InputBox("Enter the Full Project Title","Long Title","Full Title")
	address=InputBox("Enter the Full Address", "Project Address", "Address")
	specSection=InputBox("Enter the Spec Section", "Spec Section","Section")
	version=InputBox("What Version of Submittal is this?", "Version", "1")
	todaysDate=InputBox("Enter the Date for the Submittal", "Date", Date)
	splitDate=split(todaysDate,"/")
		Select case splitDate(0)
			case "1"
				splitDate(0) = "01"
			case "2"
				splitDate(0) = "02"				
			case "3"
				splitDate(0) = "03"				
			case "4"
				splitDate(0) = "04"				
			case "5"
				splitDate(0) = "05"				
			case "6"
				splitDate(0) = "06"				
			case "7"
				splitDate(0) = "07"				
			case "8"
				splitDate(0) = "08"				
			case "9"
				splitDate(0) = "09"				
		End select
		Select case splitDate(1)
			case "1"
				splitDate(1) = "01"
			case "2"
				splitDate(1) = "02"				
			case "3"
				splitDate(1) = "03"				
			case "4"
				splitDate(1) = "04"				
			case "5"
				splitDate(1) = "05"				
			case "6"
				splitDate(1) = "06"				
			case "7"
				splitDate(1) = "07"				
			case "8"
				splitDate(1) = "08"				
			case "9"
				splitDate(1) = "09"				
		End select
	singleDate=CStr(splitDate(0))+CStr(splitDate(1))+CStr(splitDate(2))
	todaysDate=CStr(splitDate(0))+"/"+CStr(splitDate(1))+"/"+CStr(splitDate(2))
		'6 is yes, 7 is no
	shopDrawings=MsgBox("Include Shop Drawings?", 4, "Shop Drawings")

'Get current filepath
	Dim WshShell, curDir
	Set WshShell = CreateObject("WScript.Shell")
	curDir = WshShell.CurrentDirectory
	
'Create log file
	if debug Then
		Dim log
		Set log = CreateObject("Scripting.FileSystemObject").OpenTextFile(curDir + "\log.txt", 2, true)
		log.WriteLine("Log File")
	End If
	
'Get all folders and their paths
	'Get Cat-Cut folder
	Set ccFSO = CreateObject("Scripting.FileSystemObject")
	ccPath = curDir + "\Cat-Cuts"

	'Get Certificate folder
	Set certFSO = CreateObject("Scripting.FileSystemObject")
	certPath = curDir + "\Certificates"

	'Get Misc Document folder
	Set miscFSO = CreateObject("Scripting.FileSystemObject")
	miscPath = curDir + "\Misc Documents"

	'Get Completed Submittals folder
	Set completedFSO = CreateObject("Scripting.FileSystemObject")
	completedPath = curDir + "\Completed Submittals"

'Debug the Initializations
	If debug Then
		log.WriteLine "Current Directory: " + curDir
		log.WriteLine "Short Title: " + shortTitle
		log.WriteLine "Long Title: " + longTitle
		log.WriteLine "Address: " + address
		log.WriteLine "Spec Section: " + specSection
		log.WriteLine "Version: " + version
		log.WriteLine "Date: " + todaysDate
		log.WriteLine "Shop Drawings: " + CStr(shopDrawings)
		If ccFSO.FolderExists(curDir + "\Cat-Cuts") Then
			log.WriteLine "Cat-Cut Folder exists."
		End If
		If certFSO.FolderExists(curDir + "\Certificates") Then
			log.WriteLine "Cert Folder exists."
		End If
		If miscFSO.FolderExists(curDir + "\Misc Documents") Then
			log.WriteLine "Misc Folder exists."
		End If
		If completedFSO.FolderExists(curDir + "\Completed Submittals") Then
			log.WriteLine "Completed Folder exists."
		End If
		log.WriteLine ""
	End If
	
'Word Documents
	'Title Sheet
		'Setup Word for TS
		Set allWord = CreateObject("Word.Application")
		allWord.Visible = False
		Set tsDocument = allWord.Documents.Open(miscPath + "\Title Sheet.docx")
		
		'Find and Replace Specific Words
		SearchAndReplace "LONG", longTitle, allWord
		SearchAndReplace "SHORT", shortTitle, allWord
		SearchAndReplace "DATE", todaysDate, allWord
		SearchAndReplace "ADDRESS", address, allWord
		SearchAndReplace "SECTIONNO", specSection, allWord
		
		'Save, Print to PDF, and Quit Word TS
		tsDocument.Save
		tsDocument.saveas miscPath + "\Title Sheet.pdf", 17
		completedPath = completedPath + "\" + shortTitle + " Completed_" + singleDate + ".pdf"
		tsDocument.saveas completedPath, 17		
		tsDocument.Close
		allWord.Quit
		
		'Remove unneeded pages based upon shop drawing choice
'Open completed PDF doc and add in bookmarks for it
		completedPDF.Open completedPath
		'If include shop drawings
		If shopDrawings = 6 Then
			completedPDF.DeletePages 1,1
		'If don't include shop drawings
		Else:
			completedPDF.DeletePages 2, 2
			completedPDF.DeletePages 7, 7
		End If
		
		*****
		'Add in Bookmarks for each section
		WshShell.SendKeys "{HOME}"
		WshShell.SendKeys "^b" & "Title Page" & "{ENTER}"
		*****
		
	'Telecommunications Contractor
		'Setup Word for TC
		Set allWord = CreateObject("Word.Application")
		allWord.Visible = False
		Set tcDocument = allWord.Documents.Open(miscPath + "\Telecommunications Contractor.docx")
		
		'Find and Replace Specific Words
		SearchAndReplace "SHORT", shortTitle, allWord
		SearchAndReplace "DATE", todaysDate, allWord		
		
		'Save, Print to PDF, and Quit Word TC
		tcDocument.Save
		tcDocument.saveas miscPath + "\Telecommunications Contractor.pdf", 17
		tcDocument.Close
		allWord.Quit
		
	'Test Plan
		'Setup Word for TP
		Set allWord = CreateObject("Word.Application")
		allWord.Visible = False
		Set tpDocument = allWord.Documents.Open(miscPath + "\Test Plan.docx")
		
		'Find and Replace Specific Words
		SearchAndReplace "SHORT", shortTitle, allWord
		SearchAndReplace "DATE", todaysDate, allWord		
		
		'Save, Print to PDF, and Quit Word TP
		tpDocument.Save
		tpDocument.saveas miscPath + "\Test Plan.pdf", 17
		tpDocument.Close
		allWord.Quit
		
	'Word Functions
		Sub SearchAndReplace(find, replace, wordDoc)
			If debug Then
				log.WriteLine "Replacing " + find + " with " + replace
			End If
			Const wdReplaceAll = 2
			Set selection = wordDoc.Selection
			selection.Find.Text = find
			selection.Find.Forward = True
			selection.Find.MatchWholeWord = True
			selection.Find.Replacement.Text = replace
			selection.Find.Execute ,,,,,,,,,,wdReplaceAll
		End Sub
	
	If debug Then
		log.WriteLine "Finished Word"
	End If
	
'Excel Documents
	'Create Object for All Excel
	Set allExcel = CreateObject("Excel.Application")
	allExcel.Visible = False
	
	'Table of Contents
		'Setup Excel for ToC
		Set tocWorkbook = allExcel.Workbooks.Open(miscPath + "\Table of Contents.xml")
		Set tocWorksheet = tocWorkbook.Worksheets("Table 1")

		'Fill in ToC main info
		tocWorksheet.Cells(1,1) = longTitle
		tocWorksheet.Cells(2,1) = address
		tocWorksheet.Cells(3,6) = version
		tocWorksheet.Cells(4,6) = specSection
		
		'Fill in Product Info
		GetFileNames ccFSO, ccPath
		For Each targetfile In files
			removeExt = Left(targetfile.name, InStrRev(targetfile.name,".") - 1)
			splitPath = Split(removeExt,"_")
			If debug Then
				log.WriteLine "Size of " + removeExt + " " + CStr(ubound(splitPath))
			End If
			'Item Number
			tocWorksheet.Cells(rowNumber,1) = itemNumber
			'Submittal Type
			tocWorksheet.Cells(rowNumber,2) = "Product"
			'Spec Ref
			If ubound(splitPath) = 4 Then
				tocWorksheet.Cells(rowNumber,4) = splitPath(4)
			End If
			'Description
			tocWorksheet.Cells(rowNumber,6) = splitPath(2)
			If debug Then
				log.WriteLine "Row Size of *" + CStr(tocWorksheet.Cells(rowNumber,6))+ "* is *" + CStr(Len(tocWorksheet.Cells(rowNumber,6))) + "*"
			End If
			If Len(tocWorksheet.Cells(rowNumber,6)) > 45 Then
				If Len(tocWorksheet.Cells(rowNumber,6)) > 90 Then
					tocWorksheet.Cells(rowNumber,6).EntireRow.RowHeight = 30
				Else: tocWorksheet.Cells(rowNumber,6).EntireRow.RowHeight = 20
				End If
			End If
			'tocWorksheet.Cells(rowNumber,6).EntireRow.AutoFit
			'Model/type/color
			tocWorksheet.Cells(rowNumber,8) = splitPath(3)
			If debug Then
				log.WriteLine "Row Size of *" + CStr(tocWorksheet.Cells(rowNumber,8))+ "* is *" + CStr(Len(tocWorksheet.Cells(rowNumber,8))) + "*"
			End If
			If Len(tocWorksheet.Cells(rowNumber,8)) > 45 Then
				If Len(tocWorksheet.Cells(rowNumber,8)) > 90 Then
					tocWorksheet.Cells(rowNumber,8).EntireRow.RowHeight = 30
				Else: tocWorksheet.Cells(rowNumber,8).EntireRow.RowHeight = 20
				End If
			End If
			'Part Number
			tocWorksheet.Cells(rowNumber,10) = splitPath(1)
			'Manufacturer
			tocWorksheet.Cells(rowNumber,12) = splitPath(0)
			rowNumber = rowNumber + 1
			itemNumber = itemNumber + 1
		Next
		
		'Add in field for shop drawings if adding in
		If (shopDrawings = 6) Then
			tocWorksheet.Cells(rowNumber,1) = itemNumber
			tocWorksheet.Cells(rowNumber,2) = "Document"
			tocWorksheet.Cells(rowNumber,6) = "Shop Drawings"
		End If
		
		'Save, Print to PDF, and Quit Excel ToC
		allExcel.ActiveWorkbook.Save
		tocWorkbook.ActiveSheet.ExportAsFixedFormat EXCEL_PDF, miscPath & "\Table of Contents.pdf", EXCEL_QualityStandard, TRUE,FALSE,,,False
		allExcel.ActiveWorkbook.Close
		
	'Key Personnel List
		'Setup Excel for KPL
		Set kplWorkbook = allExcel.Workbooks.Open(miscPath + "\Key Personnel List")
		Set kplWorksheet = kplWorkbook.Worksheets("Sheet1")
		
		'Fill in KPL info
		kplWorksheet.Cells(2,1) = longTitle
		kplWorksheet.Cells(3,1) = address
		
		'Save, Print to PDF and Quit Excel KPL
		allExcel.ActiveWorkbook.Save
		kplWorkbook.ActiveSheet.ExportAsFixedFormat EXCEL_PDF, miscPath & "\Key Personnel List.pdf", EXCEL_QualityStandard, TRUE,FALSE,,,False
		allExcel.ActiveWorkbook.Close
		allExcel.Quit
	
	'Excel Functions
		Sub GetFileNames(targetFSO, targetPath)
			Set TargetFolder = targetFSO.GetFolder(targetPath)
			Set files = TargetFolder.Files
			For Each targetfile In files
				If debug Then
					log.WriteLine "     " + targetfile.name
				End If
				fileCount = fileCount + 1
			Next
		End Sub
		
	'Add Pages to PDF
		'Telecommunications Contractor
			If debug Then
				log.WriteLine("ADDING TELECOMMUNICATIONS CONTRACTOR TO PDF")
			End If
			tempPDF.Open miscPath + "\Telecommunications Contractor.pdf"
			completedPDF.InsertPages currentPage, tempPDF, 0, tempPDF.GetNumPages(), 0
			currentPage = currentPage + tempPDF.GetNumPages() + 1
			tempPDF.Close
		'Key Personnel
			If debug Then
				log.WriteLine("ADDING KEY PERSONNEL TO PDF")
			End If
			tempPDF.Open miscPath + "\Key Personnel List.pdf"
			completedPDF.InsertPages currentPage, tempPDF, 0, tempPDF.GetNumPages(), 0
			currentPage = currentPage + tempPDF.GetNumPages()
			tempPDF.Close
			GetFileNames certFSO, certPath
			For Each targetfile In files
				splitPath = Split(targetfile.name, " ")
				If splitPath(0) = "cert" Then
					tempPDF.Open targetfile
					completedPDF.InsertPages currentPage, tempPDF, 0, tempPDF.GetNumPages(), 0
					currentPage = currentPage + tempPDF.GetNumPages()
					tempPDF.Close
				End If
			Next
			currentPage = currentPage + 1
		'Minimum Manufacturer Qualifications
			If debug Then
				log.WriteLine("ADDING MMQ TO PDF")
			End If
			GetFileNames certFSO, certPath
			For Each targetfile In files
				splitPath = Split(targetfile.name, " ")
				If splitPath(0) = "letter" Then
					tempPDF.Open targetfile
					completedPDF.InsertPages currentPage, tempPDF, 0, tempPDF.GetNumPages(), 0
					currentPage = currentPage + tempPDF.GetNumPages()
					tempPDF.Close
				End If
			Next
			currentPage = currentPage + 1
		'Test Plan
			If debug Then
				log.WriteLine("ADDING TEST PLAN TO PDF")
			End IF
			GetFileNames miscFSO, miscPath
			For Each targetfile In files
				splitPath = Split(targetfile.name, " ")
				If splitPath(0) = "Test" Then 
					tempPDF.Open targetfile
					completedPDF.InsertPages currentPage, tempPDF, 0, tempPDF.GetNumPages(), 0
					currentPage = currentPage + tempPDF.GetNumPages() + 1
					tempPDF.Close
				End If
			Next
		'Products
			If debug Then
				log.WriteLine("ADDING PRODUCTS TO PDF")
			End If
			GetFileNames ccFSO, ccPath
			For Each targetfile In files
				splitPath = Split(targetfile.name, " ")
				tempPDF.Open targetfile
				completedPDF.InsertPages currentPage, tempPDF, 0, tempPDF.GetNumPages(), 0
				currentPage = currentPage + tempPDF.GetNumPages()
				tempPDF.Close
			Next
'Close completed PDF app
	completedPDF.Save 0, completedPath
	completedPDF.Close
	completedAPP.Exit
'Done
If debug Then
	log.WriteLine("Finished")
	log.Close
End If
Wscript.Echo "Finished"