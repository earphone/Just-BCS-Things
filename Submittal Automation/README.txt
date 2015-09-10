Basic Setup for Automating Job Submittals

ToDo:
	Add annotation of item number to pdfs
	Figure out if can all be run in one script or if individual ones need to be written
		Write batch for all scripts
	Put spec number into each file name
	Possible save-all-to-pdf script	
		Choice of whether to combine or not
	Combine all pdf's into one
		and specific pages (for shop drawings)
	Create a heirarcheral standard
		Approved Submittals
		Cat-Cuts
			Maufacturer_Part#_Description_Model/Type/Color_SpecRef
		Certificates
			"cert "
			"letter "
		Completed Submittals
		Misc Documents
			Key Personnel List
			Table of Contents
			Telecommunications Contractor
			Test Plan
			Title Sheet
			
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
	
Log:
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
	