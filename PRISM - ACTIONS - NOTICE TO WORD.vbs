'Script structure 
' 1 >> CSO runs script. 
' 2 >> Script determines if CSO is in CORD or DORD.
' 3 >> Script displays "You are trying to generate document <blank>. If this is correct, press OK. If this is incorrect, press STOP SCRIPT."
' 4 >> Script converts the document to Word.


'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================


'==== CUSTOM FUNCTION ====
'This custom function is going to be used to determine if the worker is viewing the DORD doc or if the worker is viewing the general information for the document.
'If the worker is viewing the document itself, the function backs PRISM out to the general information screen and then re-enters the DORD document.
'This will assure that the script starts reading from the top of the document every time.
FUNCTION find_start_of_dord_doc(doc_id, case_number)
	EMReadScreen in_dord_doc_now, 7, 3, 50
	in_dord_doc_now = trim(in_dord_doc_now)
	IF in_dord_doc_now = "Request" THEN 
		PF3
		EMReadScreen doc_id, 16, 4, 50
		doc_id = replace(doc_id, " ", "")
		EMReadScreen case_number, 13, 4, 15
		case_number = replace(case_number, " ", "-")
		PF21
	ELSEIF in_dord_doc_now = "" THEN 
		EMReadScreen where_in_dord, 30, 2, 25
		where_in_dord = trim(where_in_dord)
		IF where_in_dord = "- Document Request List -" THEN
			script_end_procedure("You must be viewing a specific DORD document for this script to work. The script will now stop.")
		ELSEIF where_in_dord = "Document Request Detail" THEN 
			EMReadScreen doc_id, 16, 4, 50
			doc_id = replace(doc_id, " ", "")
			EMReadScreen case_number, 13, 4, 15
			case_number = replace(case_number, " ", "-")
			PF21
		ELSE
			script_end_procedure("The script is not able to determine where you are in PRISM. The script will now stop." & vbCr & vbCr & "For this script to work properly, you will need to run it while viewing the DORD document you wish to export to Word.")
		END IF
	END IF
END FUNCTION

'===== CUSTOM FUNCTION =====
'This function reads every line of the DORD doc and puts it into an array.
FUNCTION copy_dord_screen_to_array(output_array)
	output_array = "" 'resetting array
	Dim screenarray(15)	'16 line array
	row = 5
	For each line in screenarray
		EMReadScreen reading_line, 80, row, 2
		output_array = output_array & reading_line & "UUDDLRLRBA"
		row = row + 1
	Next
	output_array = split(output_array, "UUDDLRLRBA")
END FUNCTION

'==== CUSTOM FUNCTION ====
'This custom function is going to be used to determine if the worker is viewing the DORD doc or if the worker is viewing the general information for the document.
'If the worker is viewing the document itself, the function backs PRISM out to the general information screen and then re-enters the DORD document.
'This will assure that the script starts reading from the top of the document every time.
FUNCTION find_start_of_cord_doc
	EMReadScreen where_in_cord, 30, 2, 28
	where_in_cord = trim(where_in_cord)
	
	IF where_in_cord = "CODO Request Detail" THEN 
		PF21
	ELSEIF where_in_cord = "CODO Online View" THEN
		PF3
		PF21
	ELSE
		script_end_procedure("The script cannot determine where you are in PRISM. In order for the script to work, you must be viewing a CORD document. The script will now stop.")
	END IF
END FUNCTION

'===== CUSTOM FUNCTION =====
'This function reads every line of the DORD doc and puts it into an array.
FUNCTION copy_cord_screen_to_array(output_array)
	output_array = "" 'resetting array
	Dim screenarray(15)	'16 line array
	row = 5
	For each line in screenarray
		EMReadScreen reading_line, 78, row, 2
		output_array = output_array & reading_line & "UUDDLRLRBA"
		row = row + 1
	Next
	output_array = split(output_array, "UUDDLRLRBA")
END FUNCTION



'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Checking for PRISM
CALL check_for_PRISM(True)

EMReadScreen PRISM_screen, 4, 21, 75
IF PRISM_screen = "DORD" THEN 

	CALL find_start_of_dord_doc(doc_id, case_number)

	'Creates the Word doc
	
	'These two lines are for generating text files if we go with text files instead of Word docs.
	'SET objTxtFile = CreateObject("Scripting.FileSystemObject")
	'SET objDord = objTxtFile.CreateTextFile("Q:\Blue Zone Scripts\Child Support\Sandbox\Test Text Files\Doc ID " & doc_id & " Case Number " & case_number & ".txt", 1)

	Set objWord = CreateObject("Word.Application")
	objWord.Visible = True
	
	'Formatting the Word document.
	Set objDoc = objWord.Documents.Add()
	Set objSelection = objWord.Selection
	objSelection.PageSetup.LeftMargin = 25
	objSelection.PageSetup.RightMargin = 25
	objSelection.PageSetup.TopMargin = 0
	objSelection.PageSetup.BottomMargin = 0
	objSelection.ParagraphFormat.LineSpacing = 12
	objSelection.Range.ParagraphFormat.SpaceAfter = 0
	objSelection.Paragraphs.Alignment = wdAlignParagraphCenter
	objSelection.Font.Name = "Courier New"
	objSelection.Font.Size = "11"
	
	DO
		call copy_dord_screen_to_array(screentest)
		'Adds current screen to Word doc
		CALL find_variable("Page: ", page_number, 2)
		For each line in screentest
			IF line <> "" THEN objSelection.TypeText line & Chr(13)
			'objDord.WriteLine line & chr(13)
		Next
		transmit
		CALL find_variable("Page: ", confirm_page_number, 2)
		page_number = trim(page_number)
		confirm_page_number = trim(confirm_page_number)
		'IF page_number <> confirm_page_number THEN objSelection.InsertBreak(wdPageBreak)
		EMReadScreen end_of_display, 14, 24, 2
	LOOP UNTIL end_of_display = "End of Display"

ELSEIF PRISM_screen = "CORD" OR PRISM_screen = "ORD " THEN 

	'Creates the Word doc
	Set objWord = CreateObject("Word.Application")
	objWord.Visible = True
	
	'Formatting the Word document.
	Set objDoc = objWord.Documents.Add()
	Set objSelection = objWord.Selection
	objSelection.PageSetup.LeftMargin = 25
	objSelection.PageSetup.RightMargin = 25
	objSelection.PageSetup.TopMargin = 0
	objSelection.PageSetup.BottomMargin = 0
	objSelection.ParagraphFormat.LineSpacing = 12
	objSelection.Range.ParagraphFormat.SpaceAfter = 0
	selection.Paragraphs.Alignment = wdAlignParagraphCenter
	objSelection.Font.Name = "Courier New"
	objSelection.Font.Size = "11"
	
	find_start_of_cord_doc
	
	DO
		call copy_cord_screen_to_array(screentest)
		'Adds current screen to Word doc
		For each line in screentest
			IF line <> "" THEN objSelection.TypeText line & Chr(13)
			'objSelection.TypeParagraph()
		Next
		PF8
		EMReadScreen last_page, 21, 24, 2
	LOOP UNTIL last_page = "This is the last page"

ELSE
	script_end_procedure("You do not appear to be in either DORD or CORD. Please navigate to the exact document you are attempting to duplicate and run the script again.")
END IF

script_end_procedure("Success!!")
