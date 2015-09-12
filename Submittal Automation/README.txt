Basic Setup for Automating Job Submittals

To Run:
	Double click setup.vbs	
	Fill in information needed
	Will say "Finished" when completed
	New Submittal file will be found in "Completed Submittals" Folder
	
	Heirarcheral Standard
		Approved Submittals
		Cat-Cuts
			Maufacturer_Part#_Description_Model/Type/Color_SpecRef
		Certificates
			"cert "
			"letter "
		Completed Submittals
		Misc Documents
			Key Personnel List_SpecRef
			Table of Contents
			Telecommunications Contractor_SpecRef
			Test Plan_SpecRef
			Title Sheet_SpecRef

ToDo:
	Add annotation of item number to pdfs
		Write batch for all scripts
	Possible save-all-to-pdf script	
		Choice of whether to combine or not
			
Useful:
	use for batch:	http://www.online-tech-tips.com/computer-tips/create-windows-batch-files/
	possible: 		http://www.mrexcel.com/forum/excel-questions/302970-task-scheduler-vbulletin-script-auto-open-excel.html
	get file names:	http://spreadsheetpage.com/index.php/tip/getting_a_list_of_file_names_using_vba/
	selecting text:	https://msdn.microsoft.com/EN-US/library/office/ff191718.aspx
	selecting text:	https://technet.microsoft.com/en-us/library/Ee692875.aspx
	search n rep:	http://stackoverflow.com/questions/6128880/vb-script-to-find-and-replace-text-in-word-document
	get array size:	http://www.access-programmers.co.uk/forums/showthread.php?t=130453
	split function:	http://www.tutorialspoint.com/vbscript/vbscript_split_function.htm
	list all files:	https://manojsawant.wordpress.com/2013/02/05/vbscript-list-all-the-files-in-folder-and-subfolders/
	merge pdfs:		http://stackoverflow.com/questions/4154110/merge-multiple-pdf-files-with-vbscript
	pdf reference:	http://www.adobe.com/content/dam/Adobe/en/devnet/acrobat/pdfs/iac_api_reference.pdf
	send keys:		http://ss64.com/vb/sendkeys.html
	batch bookmark:	https://forums.adobe.com/thread/613362
	
Log:
	***09/11/2015***
	Rewrote some areas to have better readability
	Added in ToC pdf into main pdf
	Added ability to add spec refs to end of MISC file names after "_"
	Changed words to be Replaced in some MISC files
	***09/10/2015***
	Stuck on Inserting bookmarks
	Implemented a log file if debugging is enabled
	Insert pages into main completedPDF file, completed all:
		Telecommunications Contractor
		Key Personnel
		Minimum Manufacturer Qualifications
		Test Plan
		Product
	***09/09/2015***
	Found out how to merge pdfs and created basis to start
	***09/08/2015***
	Writes cat-cut file names to ToC
		Strips file extension
		If model/type/color field is too long then increases row size up to 2 times
		Still looking for way to merge pdf's
	***09/04/2015***
	Fixed SearchAndReplace to work for all word documents
	Created Blank Templates in BLANK_MISC folder
	***09/03/2015***
	Wrote out prompts to get the following:
		Short Project Name
		Long Project Name
		Address
		Spec Section
		Version
		Date
		Shop drawings (yes or no)
	Open Excel Sheet
	Saves Word and Excel documents then saves them to PDF
	Closes each document as it opens it
	