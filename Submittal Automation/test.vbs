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
	currentPage = 1

'Reminder to close all excel, word, and PDF documents
	warningString = "Please Close All Excel, Word, and PDF documents before continuing." + vbNewLine + "Failure to do so will cause unexpected problems." + vbNewLine + "Hit CANCEL to exit this script!"
	warningMsg = MsgBox(warningString)
	If warningMsg = 2 Then
		WScript.Quit
	End If
	
'Get the initial information
	'1 if ok; 2 is cancel
	shortTitle=InputBox("Enter the Short Project Title","Short Title","Short Title")
		If shortTitle = "2"  Then
			WScript.Quit
		End If
	longTitle=InputBox("Enter the Full Project Title","Long Title","Full Title")
		If longTitle = "2"  Then
			WScript.Quit
		End If
	address=InputBox("Enter the Full Address", "Project Address", "Address")
		If address = "2"  Then
			WScript.Quit
		End If
	specSection=InputBox("Enter the Spec Section", "Spec Section","Section")
		If specSection = "2"  Then
			WScript.Quit
		End If
	version=InputBox("What Version of Submittal is this?", "Version", "1")
		If version = "2" Then
			WScript.Quit
		End If
	todaysDate=InputBox("Enter the Date for the Submittal", "Date", Date)
		If todaysDate = "2"  Then
			WScript.Quit
		End If
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
	If debug Then
		Dim log
		Set log = CreateObject("Scripting.FileSystemObject").OpenTextFile(curDir + "\log.txt", 2, true)
		log.WriteLine("Log File")
		log.WriteLine(Now)
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
	'Setup Word for All documents
	GetFileNames miscFSO, miscPath
		For Each targetfile In files
			splitPath = Split(targetfile.name, " ")
			Select case splitPath(0)
				case "Telecommunications"
				'Setup Word for TC
					Set tcWord = CreateObject("Word.Application")
					tcWord.Visible = False
					Set tcDocument = tcWord.Documents.Open(miscPath + "\" + targetfile.name)
				case "Title"
				'Setup Word for TS
					Set tsWord = CreateObject("Word.Application")
					tsWord.Visible = False
					Set tsDocument = tsWord.Documents.Open(miscPath + "\" + targetfile.name)
				case "Test"
				'Setup Word for TP
					Set tpWord = CreateObject("Word.Application")
					tpWord.Visible = False
					Set tpDocument = tpWord.Documents.Open(miscPath + "\" + targetfile.name)
			End Select
		Next
	'Title Sheet
		'Find and Replace Specific Words
		SearchAndReplace "`LONG~", longTitle, tsWord
		SearchAndReplace "`SHORT~", shortTitle, tsWord
		SearchAndReplace "`DATE~", todaysDate, tsWord
		SearchAndReplace "`ADDRESS~", address, tsWord
		SearchAndReplace "`SECTIONNO~", specSection, tsWord
		
		'Save, Print to PDF, and Quit Word TS
		tsDocument.Save
		tsDocument.saveas miscPath + "\Title Sheet.pdf", 17
		completedPath = completedPath + "\" + shortTitle + " Completed_" + singleDate + ".pdf"
		tsDocument.saveas completedPath, 17		
		tsDocument.Close
		tsWord.Quit
		
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
	'*****
		'WshShell.AppActivate "Acrobat.exe"
		'Add in Bookmarks for each section
		'WshShell.SendKeys "{HOME}"
		'WshShell.SendKeys "^b"
		'log.WriteLine "Created new bookmark"
		'WshShell.SendKeys "Title Page"
		'log.WriteLine "Typed in title"
		'WshShell.SendKeys "{ENTER}"
		'log.WriteLine "Pressed Enter"
		'log.WriteLine ""
	'*****
	'Telecommunications Contractor		
		'Find and Replace Specific Words
		SearchAndReplace "`SHORT~", shortTitle, tcWord
		SearchAndReplace "`DATE~", todaysDate, tcWord		
		
		'Save, Print to PDF, and Quit Word TC
		tcDocument.Save
		tcDocument.saveas miscPath + "\Telecommunications Contractor.pdf", 17
		tcDocument.Close
		tcWord.Quit
		
	'Test Plan
		'Find and Replace Specific Words
		SearchAndReplace "`SHORT~", shortTitle, tpWord
		SearchAndReplace "`DATE~", todaysDate, tpWord		
		
		'Save, Print to PDF, and Quit Word TP
		tpDocument.Save
		tpDocument.saveas miscPath + "\Test Plan.pdf", 17
		tpDocument.Close
		tpWord.Quit
		
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
		
		'Fill in spec ref for Misc Documents
		GetFileNames miscFSO, miscPath
		For Each targetfile In files
			removeExt = Left(targetfile.name, InStrRev(targetfile.name,".") - 1)
			splitPath = Split(removeExt,"_")
			If splitPath(0) = "Telecommunications Contractor" Then
				If ubound(splitPath) > 0 Then
					tocWorksheet.Cells(8,4) = splitPath(1)
				End If
			End If
			If splitPath(0) = "Key Personnel List" Then
				If ubound(splitPath) > 0 Then
					tocWorksheet.Cells(9,4) = splitPath(1)
				End If
			End If
			If splitPath(0) = "Minimum Manufacturer Qualifications" Then
				If ubound(splitPath) > 0 Then
					tocWorksheet.Cells(10,4) = splitPath(1)
				End If
			End If
			If splitPath(0) = "Test Plan" Then
				If ubound(splitPath) > 0 Then
					tocWorksheet.Cells(11,4) = splitPath(1)
				End If
			End If
		Next
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
		'Setup Excel for KPL	GetFileNames miscFSO, miscPath
		Dim kplWorkbook
		For Each targetfile In files
			splitPath = Split(targetfile.name, " ")
			If splitPath(0) = "Key" Then
				Set kplWorkbook = allExcel.Workbooks.Open(miscPath + targetfile.name)
				Set kplWorksheet =  kplWorkbook.Worksheets("Sheet1")
				log.WriteLine "		Key Personnel List file found"
				
				'Fill in KPL info
				kplWorksheet.Cells(2,1) = longTitle
				kplWorksheet.Cells(3,1) = address
				
				'Save, Print to PDF and Quit Excel KPL
				allExcel.ActiveWorkbook.Save
				kplWorkbook.ActiveSheet.ExportAsFixedFormat EXCEL_PDF, miscPath & "\Key Personnel List.pdf", EXCEL_QualityStandard, TRUE,FALSE,,,False
				allExcel.ActiveWorkbook.Close
			End If
		Next
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
		'Table of Contents
			If debug Then
				log.WriteLine("ADDING TABLE OF CONTENTS TO PDF")
			End If
			tempPDF.Open miscPath + "\Table of Contents.pdf"
			completedPDF.InsertPages currentPage, tempPDF, 0, tempPDF.GetNumPages(), 0
			currentPage = currentPage + tempPDF.GetNumPages() + 1
			tempPDF.Close
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