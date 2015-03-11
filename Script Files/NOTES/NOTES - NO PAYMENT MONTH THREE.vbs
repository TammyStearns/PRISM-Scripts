Option Explicit
'option Explicit every time you start a new Script
'this part to END IF is part of Veronica's routine function that should be added to every script that interfaces with PRISM
'second line is DIM
DIM beta_agency

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF


'DIM for dialog
'DIM is all the info showin in black lettering with the underscore's between the words

DIM Month_Three_No_Payment_Dialog, Case_Number, Date_NCP_called, NCP_No_Payment_Reason, NCP_Receiving_PA_dropdownlist, Month_one_no_payment_add_worklist_CheckBox, _
New_Hire_checkbox, worker_number, Sent_Pay2_letter_checkbox, ButtonPressed, List3, Social_Security_benefits_droplist



'This document is being used by Stearns County for no payment of support for three months. 
'dialog box
BeginDialog Month_Three_No_Payment_Dialog, 0, 0, 376, 330, "Month Three of No Payment"

 Text 5, 10, 50, 15, "Case Number"  'it would be really nice if this would auto fill
  EditBox 60, 10, 75, 15, Case_Number
  Text 5, 35, 130, 10, "Date NCP was called for status update:"
  EditBox 135, 35, 70, 15, Date_NCP_called
  Text 5, 60, 120, 15, "Reason no payment has been made:"
  EditBox 130, 60, 165, 15, NCP_No_Payment_Reason
  Text 5, 110, 280, 15, "Confirmed via MAXIS, NCP is receiving public assistance and coded NCDE panel 2:"
  DropListBox 30, 120, 110, 15, "Select one:"+chr(9)+"Yes receiving public assistance"+chr(9)+"No public assistance case", NCP_Receiving_PA_dropdownlist
  CheckBox 5, 250, 180, 15, "Create a worker worklist note for 30 days from today.", Month_one_no_payment_add_worklist_CheckBox
  CheckBox 5, 225, 205, 15, "Confirmed all employer information is current in NEW HIRE", New_Hire_checkbox
  Text 5, 300, 50, 10, "Worker Name:"
  EditBox 60, 295, 65, 15, worker_number
  ButtonGroup ButtonPressed
    OkButton 250, 315, 50, 15
    CancelButton 310, 315, 50, 15
  Text 5, 195, 150, 10, "Case appears to be moving toward contempt action:"
  DropListBox 5, 205, 150, 15, "Select one:"+chr(9)+"Yes payment history created"+chr(9)+"No payment history created", List3
  CheckBox 5, 85, 150, 10, "Sent Pay 2 Letter to NCP", Sent_Pay2_letter_checkbox
  Text 10, 150, 285, 15, "Confirmed via SSTD and SSSD, NCP is receiving Social Security Benefits:"
  DropListBox 10, 165, 225, 20, "Select One:"+chr(9)+"Yes Social Security Benefits"+chr(9)+"No Social Security Benefits", Social_Security_benefits_droplist
EndDialog


DIM main_number 
'Connect to PRISM

EMConnect ""


'run Dialog
DO
	DO

		Dialog Month_Three_No_Payment_Dialog 
		IF ButtonPressed = 0 THEN StopScript
		IF worker_number = "" THEN MsgBox "You must sign."
	Loop UNTIL worker_number <> ""
	IF case_number = "" THEN MsgBox "You must enter case number."
Loop UNTIL case_number <> ""



'Write Script in CAAD

EMWriteScreen "CAAD", 21, 18
EMSendKey "<enter>"
EMWaitReady 0,0
 
PF5


'Code
EMWriteScreen "T0055", 4, 54	
EMSetCursor 16, 4

CALL write_variable_in_CAAD("NCP stated no payments made due to " & NCP_No_Payment_Reason)
'a blob of text in between quotes are a called a string

EMSendKey "<enter>" 
EMWaitReady 0,0

'write worklist
'always add EMSendKey and EMWaitReady together
EMWriteScreen "CAWD", 21, 18
EMSendKey "<enter>"

EMSendKey "<PF5>"
EMWaitReady 0,0

EMWriteScreen "Free", 4, 37
EMWriteScreen "Check to see if NCP made payment", 10, 4
EMWriteScreen "30", 17, 52
EMWriteScreen "CAST", 21, 18
EMSendKey "<enter>"
EMWaitReady 0,0



