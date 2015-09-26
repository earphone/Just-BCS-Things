Basic Setup for Automating Job Submittals
For most updated version visit:	https://github.com/earphone/Just-BCS-Things
Last Updated:	09/25/2015
===============================================================================
To Run:

    A:	Make sure all spec sheets are in the cat-cut folder and named correctly
    B:	Edit all Misc Documents per job 
    C:	Double click setup.vbs
    D:	Fill in information needed when prompted
		1:	Short Title
		2:	Full Title
		3:	Address
		4:	Section
		5:	Version
		6:	Date
    E:	Will say "Finished" when completed
    F:	New Submittal file will be found in "Completed Submittals" Folder
    G:	Proofread and Edit Submittal
    		1:	Remove Unneeded Minimum manufacturer qualification letters
		2:	Check if section title pages are in the right places
		3:	Circle/bubble product number
		4:	Check if bookmarks are to the correct pages
===============================================================================
Heirarcheral Standard:

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
==============================================================================
Troubleshooting if an error occurs:

	Snip error and send log to brandon.higashi@bcshawaii.com
		Then close all word and excel.exe open in task manager
==============================================================================
BUGS:
	Chance where PDF won't come to front
		Click on PDF to bring to foreground
		Redo title page bookmark
	    Possible fix: move create title page bookmark with others
==============================================================================
Notes:
	If Running again, Clear all PDF's from Misc Documents
	"Clear Files with Extension" deletes files in the same folder
		Gets user prompt for file extension
	Can't run from a zip file
==============================================================================
ToDo:
	Add annotation of item number to pdfs
		Write batch for all scripts
	Possible save-all-to-pdf script	
		Choice of whether to combine or not
==============================================================================			
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
		Possilby Show pdf and use send keys
	PDBookmark:	http://forums.planetpdf.com/cant-title-bookmarks_topic958.html
	flattenPages:	http://help.adobe.com/livedocs/acrobat_sdk/11/Acrobat11_HTMLHelp/wwhelp/wwhimpl/common/html/wwhelp.htm?context=Acrobat11_HTMLHelp&file=JS_API_AcroJS.89.472.html
	jso docs:	http://help.adobe.com/livedocs/acrobat_sdk/11/Acrobat11_HTMLHelp/wwhelp/wwhimpl/js/html/wwhelp.htm?href=JS_API_AcroJS.89.1.html#1515776&accessible=true
==============================================================================	
Log:
	***09/25/2015***
	Added in "Item #" annotations and addItem sub
	***09/24/2015***
	Changed some formatting and made it ready to go
	***09/23/2015***
	Changed bookmarks to be implemented using a Sub that is called
	***09/21/2015***
	Changed how bookmarks are implemented (not using sendKeys anymore)
	***09/18/2015***
	Able to add bookmarks but cannot work unless one useless bookmark is made first
		Used PDF API AVDoc Object
		Used sendKeys function in script
		User should not use keyboard or click things when running script
	***09/15/2015***
	Resolved an issue where Key Personnel List was not being turned into a PDF
		and therefore not being added into final file
	Changed some wording for warning string
	Added to "To Run" instructions and made top half more readability
	***09/14/2015***
	Added in "Clear Files with Extension" script into Misc Documents folder
	Added in support for spec ref to ToC for Key Personnel List_SpecRef
	Added in reminder for closing all word, PDF, and excel documents before start
		Helps with clearing all PDFs to rerun "setup"
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
	