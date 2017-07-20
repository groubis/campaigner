#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         George Roubis (georgeroubis@gmail.com)

 Script Function:
	A simple tool to deploy email campaings.

#ce ----------------------------------------------------------------------------

	#include <File.au3>
	#include <Array.au3>
	#include <GUIConstantsEx.au3>
	#include <WindowsConstants.au3>
	#include <EditConstants.au3>
	#include <StaticConstants.au3>
	#include <ComboConstants.au3>
	#include <GuiDateTimePicker.au3>
	#include <GuiListView.au3>
	#include <FileConstants.au3>
	#include "_IsValidEmail.au3"
	#include "_INetSmtpMailCom.au3"
	#include <GuiEdit.au3>

	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	;;Check ini file									;;
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	;Declare the ini file with default values
	$INI_File = @ScriptDir&'\Campaigner.ini'

	;If the ini file was not found
	If FileExists($INI_File) = 0 Then

		MsgBox(4096, "Ini file error", "Ini file '"&$INI_File&"' was not found.", 10)

		Exit

	EndIf

    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    ;;Declare global variables						   	;;
    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	$MailServer 				=	''
	$MailPort					=	''
	$MailSSL					=	0
	$AccountUsername			=	''
	$AccountPassword 			=	''
	$FromName					=	''
	$FromAddress				=	''
	$Importance					=	''
	$HTMLTemplate				=	''
	$MandatoryHeaderElements 	= 	2
	$ToAddressColumnNumber 		= 	0
	$SubjectColumnNumber		= 	0
	$FULL_LOG_RESULTS			=  ""

    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    ;;Declare ini file predeclared variables of Server 	;;
    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	;Read all Server section elements to an array
	Local $INI_Server_Array = IniReadSection($INI_File, "MailServer")

	;If an error occured
	If @error Then

		MsgBox(4096, "Ini file error", "Error while reading MailServer section from file '"&$INI_File&"'.", 10)

		Exit

	;If the section was succesfully read
	Else

		;Loop through all Server elements
		For $i = 1 To $INI_Server_Array[0][0]

			;Dynamically assign a variable
			Assign($INI_Server_Array[$i][0], $INI_Server_Array[$i][1], 2)

		Next

	EndIf

    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    ;;Declare ini file predeclared variables of Mail 	;;
    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	;Read all Mail section elements to an array
	Local $INI_Mail_Array = IniReadSection($INI_File, "MailDefaults")

	;If an error occured
	If @error Then

		MsgBox(4096, "Ini file error", "Error while reading MailDefaults section from file '"&$INI_File&"'.", 10)

		Exit

	;If the section was succesfully read
	Else

		;Loop through all mail elements
		For $i = 1 To $INI_Mail_Array[0][0]

			;Dynamically assign a variable
			Assign($INI_Mail_Array[$i][0], $INI_Mail_Array[$i][1], 2)

		Next

	EndIf

    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    ;;Process INI variables 							;;
    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	;If ssl is off
	If $MailSSL = 0  Then

		$HumanReadableMailSSL	= 'No'

	;If SSL is on
	Else

		$HumanReadableMailSSL	= 'Yes'

	EndIf

    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    ;;Create the GUI		 							;;
    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	;Create the GUI
	$CampaignerGUI = GUICreate ("Campaigner v.1.6", 990 , 450)

	;Set status to visible
	GUISetState(@SW_SHOW)

	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	;Create the group with mail server configuration elements
	GUICtrlCreateGroup("Mail Server Configuration", 10, 10, 480, 150)

	;Create label and input for mail server ip/hostname
	GUICtrlCreateLabel("Server Name/Ip Address", 25, 30, 160, 20)
	$MailServerInput	=	GUICtrlCreateInput ( $MailServer, 150, 25 , 325 , 20)

	;Create label and input for mail server port
	GUICtrlCreateLabel("Server Port", 25, 53, 160, 20)
	$MailPortInput		=	GUICtrlCreateInput ( $MailPort, 150, 50 , 230 , 20,$ES_NUMBER)

	;Create a telnet button
	$TelnetTestButton	=	GUICtrlCreateButton("Telnet", 390, 48, 85, 23)

	;Create label and input for mail SSL mode
	GUICtrlCreateLabel("Server SSL", 25, 80, 120, 20)
	$MailSSLInput		=	GUICtrlCreateCombo ('', 151, 75, 70, 20,$CBS_DROPDOWNLIST )
	GUICtrlSetData($MailSSLInput, "Yes|No", $HumanReadableMailSSL)

	;Create label and input for mail username
	GUICtrlCreateLabel("Account Username", 25, 105, 160, 20)
	$MailAcountUsernameInput	=	GUICtrlCreateInput ( $AccountUsername, 150, 102 , 325 , 20)

	;Create label and input for mail password
	GUICtrlCreateLabel("Account Password", 25, 130, 160, 20)
	$MailAcountPasswordInput	=	GUICtrlCreateInput ( $AccountPassword, 150, 127 , 325 , 20, $ES_PASSWORD)

	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	;Create the group with mail content configuration elements
	GUICtrlCreateGroup("Mail Content Configuration", 10, 170, 480, 230)

	;Create label and input for From Name
	GUICtrlCreateLabel("From Name", 25, 194, 160, 20)
	$FromNameInput		=	GUICtrlCreateInput ( $FromName, 150, 190 , 325 , 20)

	;Create label and input for From address
	GUICtrlCreateLabel("From Address", 25, 219, 160, 20)
	$FromAddressInput	=	GUICtrlCreateInput ( $FromAddress, 150, 215 , 325 , 20)

	;Create label and input for mail importance
	GUICtrlCreateLabel("Mail Importance", 25, 245, 120, 20)
	$MailImportanceInput	=	GUICtrlCreateCombo ('', 151, 241, 70, 20,$CBS_DROPDOWNLIST )
	GUICtrlSetData($MailImportanceInput, "High|Normal|Low", $Importance)

	;Create template file input label, input and button
	GUICtrlCreateLabel("Template File", 25, 270, 100, 20)
	$HTMLTemplateInput 		=	GUICtrlCreateInput ( $HTMLTemplate, 150, 267 , 260 , 20 , $ES_READONLY)
	$HTMLTemplateButton		=	GUICtrlCreateButton ( "Select", 416, 267, 60, 20)

	;Create template file input label, input and button
	GUICtrlCreateLabel("Data File", 25, 295, 100, 20)
	$DataFileInput 			=	GUICtrlCreateInput ( "", 150, 293 , 260 , 20 , $ES_READONLY)
	$DataFileButton			=	GUICtrlCreateButton ( "Select", 416, 293, 60, 20)

	;Create template file input label, input and button
	GUICtrlCreateLabel("Attachment(s) (optional)", 25, 320, 120, 20)
	$AttachmentFileInput 	=	GUICtrlCreateInput ( "", 150, 319 , 130 , 20 , $ES_READONLY)
	$AttachmentShowButton	=	GUICtrlCreateButton ( "Show", 286, 319, 60, 20)
	$AttachmentDeleteButton	=	GUICtrlCreateButton ( "Delete", 351, 319, 60, 20)
	$AttachmentFileButton	=	GUICtrlCreateButton ( "Select", 416, 319, 60, 20)

	;Create CC input label and input
	GUICtrlCreateLabel("CC (optional)", 25, 345, 100, 20)
	$CCInput			 	=	GUICtrlCreateInput ( "", 150, 345 , 325 , 20)

	;Create BCC input label and input
	GUICtrlCreateLabel("BCC (optional)", 25, 371, 100, 20)
	$BCCInput			 	=	GUICtrlCreateInput ( "", 150, 371 , 325 , 20)

	;Create operation buttons
	$ExitButton 			= GUICtrlCreateButton("Exit", 106, 410, 85, 25)
	$UnlockUIButton			= GUICtrlCreateButton("Unlock", 206, 410, 85, 25)
	$CheckButton 			= GUICtrlCreateButton("Check", 306, 410, 85, 25)
	$RunButton 				= GUICtrlCreateButton("Launch", 406, 410, 85, 25)


	;Pre-disable the run button
	GUICtrlSetState ( $RunButton, $GUI_DISABLE  )

	;Pre-disable the unlock button
	GUICtrlSetState ( $UnlockUIButton, $GUI_DISABLE  )

	;Create the group with preview listview
	GUICtrlCreateGroup("Campaign Preview", 500, 10, 480, 390)

	;Create the list view structure
	$PreviewerListview = GUICtrlCreateListView("Recipient|Subject|Status", 510, 25, 460, 340)

	;Set list view column widths
	_GUICtrlListView_SetColumnWidth ( $PreviewerListview, 0, 150 )
    _GUICtrlListView_SetColumnWidth ( $PreviewerListview, 1, 240 )

	;Create the log preview button
	$LogPreviewButton	= GUICtrlCreateButton("View Log", 796, 410, 85, 25)

	;Pre-disable the log preview button
	GUICtrlSetState ( $LogPreviewButton, $GUI_DISABLE  )

	;Create the html preview button
	$HTMLPreviewButton	= GUICtrlCreateButton("Preview", 896, 410, 85, 25)

	;Pre-disable the preview button
	GUICtrlSetState ( $HTMLPreviewButton, $GUI_DISABLE)

	;Create a label for
	$TotalCampaignMailsCounter = GUICtrlCreateLabel("", 771, 380, 200, 20, $SS_RIGHT )

	;Set color to transparent
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)

    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    ;;Start GUI operation	 							;;
    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	;Run the GUI until the dialog is closed
	While 1

		;Get current action
		$msg = GUIGetMsg()

		;Handle changes
		Select

			;If X was pressed
			Case $msg = $GUI_EVENT_CLOSE

				;Exit
				ExitLoop

			;If Exit was pressed
			Case $msg = $ExitButton

				;Exit
				ExitLoop

			;If Telnet button was pressed
			Case $msg = $TelnetTestButton

				;Get server
				$TelnetHost = 	GUICtrlRead($MailServerInput)

				;Get port
				$TelnetPort	=	GUICtrlRead($MailPortInput)

				$GlobalTelnetErrorText = ''
				$TelnetErrorStatus = 0

				;If SMTP mail server is blank
				If StringStripWS($TelnetHost,8) = '' Then

					;Update the error
					$GlobalTelnetErrorText &= "Server Name/Ip Address can not be blank."&@CRLF

					;ExitLoop
					$TelnetErrorStatus+=1

				EndIf

				;If SMTP mail server port is blank
				If StringStripWS($TelnetPort,8) = '' Then

					;Update the error
					$GlobalTelnetErrorText &= "Server port can not be blank."&@CRLF

					;ExitLoop
					$TelnetErrorStatus+=1

				EndIf

				;If an error was found
				If $TelnetErrorStatus > 0 Then

					;Notify
					MsgBox(4096, "Error", $TelnetErrorStatus&" error(s) identified:"&@CRLF&$GlobalTelnetErrorText, 0, $CampaignerGUI)

				Else

					 ;If file exists
					 If (FileExists ( @SCRIPTDIR&"\telnet\telnet.exe" ) = 1) Then

					   ;Launch telnet
					   Run(@SCRIPTDIR&"\telnet\telnet.exe "&$TelnetHost&" "&$TelnetPort)

					 Else

						;Notify
						MsgBox(4096, "Error", "File "&@ScriptDir&"\telnet\telnet.exe not found!", 0, $CampaignerGUI)

					 EndIf

				EndIf

			;If the HTML template file selection button was clicked
			Case $msg = $HTMLTemplateButton

				;Open the file selection dialog
				Dim $HTMLTemplate = FileOpenDialog ( "HTML Template Selection", @ScriptDir, "HTML files (*.html;*.htm)" ,$FD_FILEMUSTEXIST + $FD_PATHMUSTEXIST, "", $CampaignerGUI)

				;If an error occured
				If @error Then


				Else

					;Update the file input box
					GUICtrlSetData($HTMLTemplateInput, $HTMLTemplate)

				EndIf

			;If the Data file selection button was clicked
			Case $msg = $DataFileButton

				;Open the file selection dialog
				Dim $DataFile = FileOpenDialog ( "Data File Selection", @ScriptDir, "Data files (*.csv;*.txt)" ,$FD_FILEMUSTEXIST + $FD_PATHMUSTEXIST, "", $CampaignerGUI)

				;If an error occured
				If @error Then


				Else

					;Update the file input box
					GUICtrlSetData($DataFileInput, $DataFile)

				EndIf

			Case $msg = $CheckButton

				;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
				;;Disable GUI elements								;;
				;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

				;Lock UI elements to avoid changes or errors
				GUICtrlSetState ($MailServerInput, $GUI_DISABLE)
				GUICtrlSetState ($MailPortInput, $GUI_DISABLE)
				GUICtrlSetState ($MailSSLInput, $GUI_DISABLE)
				GUICtrlSetState ($MailSSLInput, $GUI_DISABLE)
				GUICtrlSetState ($MailAcountUsernameInput, $GUI_DISABLE)
				GUICtrlSetState ($MailAcountPasswordInput, $GUI_DISABLE)
				GUICtrlSetState ($FromNameInput, $GUI_DISABLE)
				GUICtrlSetState ($FromAddressInput, $GUI_DISABLE)
				GUICtrlSetState ($MailImportanceInput, $GUI_DISABLE)
				GUICtrlSetState ($DataFileButton, $GUI_DISABLE)
				GUICtrlSetState ($HTMLTemplateButton, $GUI_DISABLE)
				GUICtrlSetState ($CheckButton, $GUI_DISABLE)
				GUICtrlSetState ($AttachmentDeleteButton, $GUI_DISABLE)
				GUICtrlSetState ($AttachmentShowButton, $GUI_DISABLE)
				GUICtrlSetState ($AttachmentFileButton, $GUI_DISABLE)
				GUICtrlSetState ($CCInput, $GUI_DISABLE)
				GUICtrlSetState ($BCCInput, $GUI_DISABLE)
				GUICtrlSetState ($ExitButton, $GUI_DISABLE)
				GUICtrlSetState ($TelnetTestButton, $GUI_DISABLE)


				;disable  unlock button
				GUICtrlSetState ( $UnlockUIButton, $GUI_DISABLE)

				;Update the counter label
				GUICtrlSetData ($TotalCampaignMailsCounter, 'Checking Data...')

				;Clear the listview
				_GUICtrlListView_DeleteAllItems($PreviewerListview)

				;Disable the preview button
				GUICtrlSetState ( $HTMLPreviewButton, $GUI_DISABLE)

				;Disable the launch button
				GUICtrlSetState ($RunButton, $GUI_DISABLE)

				;Declare an operation error handler
				$CheckComplete 	= 0
				$ErrorStatus	= 0

				;Reset arrays and strings
				Dim $RawDataArray
				Dim $TemplateFileRead = ''

				;If no error was detected
				While $CheckComplete = 0

					;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
					;;Get GUI data										;;
					;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

					;Update the counter label
					GUICtrlSetData ($TotalCampaignMailsCounter, 'Checking Mail Server configuration.')

					;Get Mail Server
					$SMTPMailServer = GUICtrlRead($MailServerInput)

					;Get Mail Port
					$SMTPPort		= GUICtrlRead($MailPortInput)

					;Get SSL mode and convert it to boolean
					If GUICtrlRead($MailSSLInput) = 'Yes' Then

						$SMTPSSL = 1

					Else

						$SMTPSSL = 0

					EndIf

					;Get SMTP Account Username
					$SMTPAccountUsername = GUICtrlRead($MailAcountUsernameInput)

					;Get SMTP Account Password
					$SMTPAccountPassword = GUICtrlRead($MailAcountPasswordInput)

					;Update the counter label
					GUICtrlSetData ($TotalCampaignMailsCounter, 'Checking Mail Content configuration.')

					;Get SMTP From Name
					$SMTPFromName 	= GUICtrlRead($FromNameInput)

					;Get SMTP From Address
					$SMTPFromAddress	= GUICtrlRead($FromAddressInput)

					;Get SMTP importance
					$SMTPImportance		= GUICtrlRead($MailImportanceInput)

					;Get SMTP Template file
					$SMTPTemplateFile	= GUICtrlRead($HTMLTemplateInput)

					;Get SMTP Data file
					$SMTPDataFile		= GUICtrlRead($DataFileInput)

					;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
					;;Validate Data										;;
					;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

					$GlobalErrorText = ''

					;If SMTP mail server is blank
					If StringStripWS($SMTPMailServer,8) = '' Then

						;Update the error
						$GlobalErrorText &= "Server Name/Ip Address can not be blank."&@CRLF

						;ExitLoop
						$ErrorStatus+=1

					EndIf

					;If SMTP mail server port is blank
					If StringStripWS($SMTPPort,8) = '' Then

						;Update the error
						$GlobalErrorText &= "Server port can not be blank."&@CRLF

						;ExitLoop
						$ErrorStatus+=1

					EndIf
					#cs
					;If SMTP account username is blank
					If StringStripWS($SMTPAccountUsername,8) = '' Then

						;Update the error
						$GlobalErrorText &= "Server account username can not be blank."&@CRLF

						;ExitLoop
						$ErrorStatus+=1

					EndIf

					;If SMTP account password is blank
					If StringStripWS($SMTPAccountPassword,8) = '' Then

						;Update the error
						$GlobalErrorText &= "Server account password can not be blank."&@CRLF

						;ExitLoop
						$ErrorStatus+=1

					EndIf
					#ce
					;If From name is blank
					If StringStripWS($SMTPFromName,8) = '' Then

						;Update the error
						$GlobalErrorText &= "From name can not be blank."&@CRLF

						;ExitLoop
						$ErrorStatus+=1

					EndIf

					;If From address is invalid
					If _IsValidEmail($SMTPFromAddress) = 0 Then

						;Update the error
						$GlobalErrorText &= "From address is invalid."&@CRLF

						;ExitLoop
						$ErrorStatus+=1

					EndIf

					;Open the template file
					Local $TemplateFileOpen = FileOpen($SMTPTemplateFile, $FO_READ)

					;If unable to open
					If $TemplateFileOpen = -1 Then

						;Update the error
						$GlobalErrorText &= "Unable to open template file for reading."&@CRLF

						;ExitLoop
						$ErrorStatus+=1

					;If file was opened, read data
					Else

						;Read the file
						$TemplateFileRead = FileRead($TemplateFileOpen)

						If @error = 1 Then

							;Update the error
							$GlobalErrorText &= "Unable to read data from template file."&@CRLF

							;ExitLoop
							$ErrorStatus+=1

						Else

							;close the file
							FileClose($TemplateFileOpen)

						EndIf

					EndIf

					;If file contents are blank
					If $TemplateFileRead = '' Then

						;Update the error
						$GlobalErrorText &= "Template file is empty."&@CRLF

						;ExitLoop
						$ErrorStatus+=1

					EndIf


					;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
					;;Check attachments															;;
					;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

					;Get attachments string
					$SMTPAttachments = GUICtrlRead($AttachmentFileInput)

					;If the string is empty
					If StringStripWS($SMTPAttachments,8) = '' Then

						$MailAttachments = ''

					;Else
					Else

						;Convert string to array
						$TempAttachmentsRawArray	= StringSplit($SMTPAttachments,'|')

						If @error = 1 then

							Dim $TempAttachmentsFPArray[2] = [1,$SMTPAttachments]

						Else

							;Get the attachments directory
							$TempAttachmentsPath	 	= $TempAttachmentsRawArray[1]

							;Declare an path to store all attachments with their full path
							Dim $TempAttachmentsFPArray[1] = [0]

							;Loop through all attachments to get full paths
							For $i = 2 To $TempAttachmentsRawArray[0]

								;Add element
								_ArrayAdd($TempAttachmentsFPArray,$TempAttachmentsPath&'\'&$TempAttachmentsRawArray[$i])

								;Increase index
								$TempAttachmentsFPArray[0]+=1

							Next

						EndIf

						;Loop through all attachments and attempt to open them
						For $i = 1 To $TempAttachmentsFPArray[0]

							;Open the attachment file
							Local $CurrentAttachmentOpen = FileOpen($TempAttachmentsFPArray[$i], $FO_READ)

							;If unable to open
							If $CurrentAttachmentOpen = -1 Then

								;Update the error
								$GlobalErrorText &= 'Unable to open attachment file "'&$TempAttachmentsFPArray[$i]&'".'&@CRLF

								;ExitLoop
								$ErrorStatus+=1

							;If file was opened, read data
							Else

								FileClose($CurrentAttachmentOpen)

							EndIf

						Next

						;Convert the attachments string to valid email ready string
						$MailAttachments = _ArrayToString($TempAttachmentsFPArray,';',1,$TempAttachmentsFPArray[0])

					EndIf

					;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
					;;Check CC address															;;
					;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

					;Get SMTP CC address
					$SMTPCCAddress		= StringStripWS(GUICtrlRead($CCInput),8)

					;Convert address  to array
					$SMTPCCAddressArray = StringSplit($SMTPCCAddress,';')

					;Check if blank
					If _ArrayToString($SMTPCCAddressArray,'',1) <> '' Then

						;Loop through addresses
						For $i = 1 To $SMTPCCAddressArray[0]

							;If CC address is invalid
							If _IsValidEmail($SMTPCCAddressArray[$i]) = 0 Then

								;Update the error
								$GlobalErrorText &= 'CC address "'&$SMTPCCAddressArray[$i]&'" is invalid.'&@CRLF

								;ExitLoop
								$ErrorStatus+=1

								ExitLoop

							EndIf

						Next

					EndIf

					;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
					;;Check BCC address															;;
					;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

					;Get SMTP CC address
					$SMTPBCCAddress		= StringStripWS(GUICtrlRead($BCCInput),8)

					;Convert address  to array
					$SMTPBCCAddressArray = StringSplit($SMTPBCCAddress,';')

					;Check if blank
					If _ArrayToString($SMTPBCCAddressArray,'',1) <> '' Then

						;Loop through addresses
						For $i = 1 To $SMTPBCCAddressArray[0]

							;If CC address is invalid
							If _IsValidEmail($SMTPBCCAddressArray[$i]) = 0 Then

								;Update the error
								$GlobalErrorText &= 'BCC address "'&$SMTPBCCAddressArray[$i]&'" is invalid.'&@CRLF

								;ExitLoop
								$ErrorStatus+=1

								ExitLoop

							EndIf

						Next

					EndIf


					;Read the data file to an array
					_FileReadToArray ( $SMTPDataFile, $RawDataArray, 1 )

					;If an error occured
					If @error Then

						;1 - Error opening specified file
						If @error = 1 Then

							;Update the error
							$GlobalErrorText &= "Unable to open data file for reading."&@CRLF

							;ExitLoop
							$ErrorStatus+=1

						;2 - Unable to Split the file
						ElseIf @error = 2 Then

							;Update the error
							$GlobalErrorText &= "Unable to convert data file to a memory array (spliting issue)."&@CRLF

							;ExitLoop
							$ErrorStatus+=1

						EndIf

					Else

						;Check for blank file
						If	StringLen(StringStripWS(_ArrayToString($RawDataArray,"",1),8)) = 0 Then

							;Update the error
							$GlobalErrorText &= "Data file is empty."&@CRLF

							;ExitLoop
							$ErrorStatus+=1

						;If not blank check data
						Else

							;Loop through lines to check data
							For $i = 1 To $RawDataArray[0]

								;Convert the line to an array
								$CurrentLineArray = StringSplit($RawDataArray[$i],';')

								;If the length is at least 2
								If $CurrentLineArray[0] <= 2 Then

									;Update the error
									$GlobalErrorText &= "Line "&$i&" contains less than 3 data columns."&@CRLF

									;ExitLoop
									$ErrorStatus+=1

									ExitLoop

								Else

								EndIf

							Next

						EndIf

					EndIf

					;If no error
					If $ErrorStatus = 0 Then

						;Get Headers
						$HeaderLineArray = StringSplit($RawDataArray[1],';')

						;Set a variable to compare with mandatory elements
						$ToAddressHeadersFound 	= 0
						$SubjectHeadersFound 	= 0

						;Declare an array to store replacement rules
						Dim $ReplaceRulesArray[1][2]=[["ColumnName", "ColumnPosition"]]

						;Loop through headers
						For $i = 1 To $HeaderLineArray[0]

							;If a mandatory column was found
							If $HeaderLineArray[$i] = 'ToAddress' Then

								;Increase the counter
								$ToAddressHeadersFound+=1

								;Update the column position
								$ToAddressColumnNumber = $i

							EndIf

						Next

						;Loop through headers
						For $i = 1 To $HeaderLineArray[0]

							;If a mandatory column was found
							If $HeaderLineArray[$i] = 'Subject' Then

								;Increase the counter
								$SubjectHeadersFound+=1

								;Update the column position
								$SubjectColumnNumber = $i

							EndIf

						Next

						;Check if everything is fine
						If $ToAddressHeadersFound = 1 And $SubjectHeadersFound = 1 Then

							;Loop through headers
							For $i = 1 To $HeaderLineArray[0]

								;If a dynamic column was matched
								If $HeaderLineArray[$i] <> 'Subject' And $HeaderLineArray[$i] <> 'ToAddress' Then

									;Increase lengths
									ReDim $ReplaceRulesArray[Ubound($ReplaceRulesArray)+1][2]

									;Update rules
									$ReplaceRulesArray[Ubound($ReplaceRulesArray)-1][0]	= "[{"&$HeaderLineArray[$i]&"}]"
									$ReplaceRulesArray[Ubound($ReplaceRulesArray)-1][1]	= $i

								EndIf

							Next

						;If a header was missing
						Else

							;Update the error
							$GlobalErrorText &= "Data file does not contain all mandatory columns."&@CRLF

							;ExitLoop
							$ErrorStatus+=1

						EndIf

					EndIf

					;If no error
					If $ErrorStatus = 0 Then

						Dim $ReadyToEmailArray[1][3]=[["To Address", "Subject", "HTML Content"]]

						;Loop through data
						For $i = 2 To $RawDataArray[0]

							;Update the counter label
							GUICtrlSetData ($TotalCampaignMailsCounter, 'Generating dynamic email '&$i&' of '&$RawDataArray[0]&'.')

							;Convert the line to an array
							$CurrentLine = StringSplit($RawDataArray[$i],';')

							$CurrentHTMLString = $TemplateFileRead

							;Increase lengths
							ReDim $ReadyToEmailArray[Ubound($ReadyToEmailArray)+1][3]

							;Loop through line elements
							For $y = 1 To $CurrentLine[0]

								;If To was found
								If $y = $ToAddressColumnNumber Then

									$ReadyToEmailArray[Ubound($ReadyToEmailArray)-1][0] = $CurrentLine[$y]

								;If Subject was found
								ElseIf $y = $SubjectColumnNumber Then

									$ReadyToEmailArray[Ubound($ReadyToEmailArray)-1][1] = $CurrentLine[$y]

								;If another column was matched
								Else

									;Loop through rules
									For $j = 1 To Ubound($ReplaceRulesArray)-1

										;If a rule was matched
										If $ReplaceRulesArray[$j][1] = $y then

											;Replace the keyword
											$CurrentHTMLString = StringReplace($CurrentHTMLString,$ReplaceRulesArray[$j][0],$CurrentLine[$y])

											;Update the array
											$ReadyToEmailArray[Ubound($ReadyToEmailArray)-1][2] = $CurrentHTMLString

										EndIf

									Next

								EndIf

							Next

						Next

					EndIf

					;If no error - check for invalid emails
					If $ErrorStatus = 0 Then

						For $i = 1 To Ubound($ReadyToEmailArray)-1

							;Update the counter label
							GUICtrlSetData ($TotalCampaignMailsCounter, 'Checking recipient '&$i&' of '&(Ubound($ReadyToEmailArray)-1)&'.')

							If _IsValidEmail($ReadyToEmailArray[$i][0]) = 0 Then

								;Update the error
								$GlobalErrorText &= "Line "&($i+1)&" contains an invalid email recipient("&$ReadyToEmailArray[$i][0]&")."&@CRLF

								;ExitLoop
								$ErrorStatus+=1

							EndIf

						Next

					EndIf

					;If no error - update list view
					If $ErrorStatus = 0 Then

						$TotalOutgoing = 0

						;Update list view elements
						For $i = 1 To Ubound($ReadyToEmailArray) -1

							;Update the counter label
							GUICtrlSetData ($TotalCampaignMailsCounter, 'Updating preview element '&$i&' of '&(Ubound($ReadyToEmailArray) -1)&'.')

							GUICtrlCreateListViewItem($ReadyToEmailArray[$i][0]&"|"&$ReadyToEmailArray[$i][1], $PreviewerListview)

							$TotalOutgoing+=1

						Next

						;enable the preview button
						GUICtrlSetState ($HTMLPreviewButton, $GUI_ENABLE)

						;enable the launch button
						GUICtrlSetState ($RunButton, $GUI_ENABLE)

						;Update the counter label
						GUICtrlSetData ($TotalCampaignMailsCounter, 'Total emails in campaign: '& $TotalOutgoing)

						;Lock UI elements to avoid changes or errors
						GUICtrlSetState ($MailServerInput, $GUI_DISABLE)
						GUICtrlSetState ($MailPortInput, $GUI_DISABLE)
						GUICtrlSetState ($MailSSLInput, $GUI_DISABLE)
						GUICtrlSetState ($MailSSLInput, $GUI_DISABLE)
						GUICtrlSetState ($MailAcountUsernameInput, $GUI_DISABLE)
						GUICtrlSetState ($MailAcountPasswordInput, $GUI_DISABLE)
						GUICtrlSetState ($FromNameInput, $GUI_DISABLE)
						GUICtrlSetState ($FromAddressInput, $GUI_DISABLE)
						GUICtrlSetState ($MailImportanceInput, $GUI_DISABLE)
						GUICtrlSetState ($DataFileButton, $GUI_DISABLE)
						GUICtrlSetState ($HTMLTemplateButton, $GUI_DISABLE)
						GUICtrlSetState ($CheckButton, $GUI_DISABLE)
						GUICtrlSetState ($AttachmentDeleteButton, $GUI_DISABLE)
						GUICtrlSetState ($AttachmentShowButton, $GUI_DISABLE)
						GUICtrlSetState ($AttachmentFileButton, $GUI_DISABLE)
						GUICtrlSetState ($CCInput, $GUI_DISABLE)
						GUICtrlSetState ($BCCInput, $GUI_DISABLE)
						GUICtrlSetState ($TelnetTestButton, $GUI_DISABLE)

						;enablethe unlock button
						GUICtrlSetState ( $UnlockUIButton, $GUI_ENABLE  )

					EndIf






					;If an error was found
					If $ErrorStatus > 0 Then

						;Notify
						MsgBox(4096, "Error", $ErrorStatus&" error(s) identified:"&@CRLF&$GlobalErrorText, 0, $CampaignerGUI)

						;Stop with error
						$CheckComplete	= 1

						GUICtrlSetState ($MailServerInput, $GUI_ENABLE)
						GUICtrlSetState ($MailPortInput, $GUI_ENABLE)
						GUICtrlSetState ($MailSSLInput, $GUI_ENABLE)
						GUICtrlSetState ($MailSSLInput, $GUI_ENABLE)
						GUICtrlSetState ($MailAcountUsernameInput, $GUI_ENABLE)
						GUICtrlSetState ($MailAcountPasswordInput, $GUI_ENABLE)
						GUICtrlSetState ($FromNameInput, $GUI_ENABLE)
						GUICtrlSetState ($FromAddressInput, $GUI_ENABLE)
						GUICtrlSetState ($MailImportanceInput, $GUI_ENABLE)
						GUICtrlSetState ($DataFileButton, $GUI_ENABLE)
						GUICtrlSetState ($HTMLTemplateButton, $GUI_ENABLE)
						GUICtrlSetState ($CheckButton, $GUI_ENABLE)
						GUICtrlSetState ($AttachmentDeleteButton, $GUI_ENABLE)
						GUICtrlSetState ($AttachmentShowButton, $GUI_ENABLE)
						GUICtrlSetState ($AttachmentFileButton, $GUI_ENABLE)
						GUICtrlSetState ($CCInput, $GUI_ENABLE)
						GUICtrlSetState ($BCCInput, $GUI_ENABLE)
						GUICtrlSetState ($ExitButton, $GUI_ENABLE)
						GUICtrlSetState ($TelnetTestButton, $GUI_ENABLE)

						;Update the counter label
						GUICtrlSetData ($TotalCampaignMailsCounter, 'Operation stopped due to error(s).')

					;If everything went fine
					Else

						;Continue
						$CheckComplete 	= 100

						;disable  unlock button
						GUICtrlSetState ( $UnlockUIButton, $GUI_ENABLE)
						GUICtrlSetState ($ExitButton, $GUI_ENABLE)

					EndIf

				WEnd

			;If log preview button was pressed
			Case $msg = $LogPreviewButton

				;Create the GUI
				$CampaignerLogGUI = GUICreate ("Campaign Log Viewer", 1200 , 800,-1,-1,-1 ,$WS_EX_TOOLWINDOW,$CampaignerGUI)

				GUISetFont(9, 100, 0, "Consolas", $CampaignerLogGUI)

				Local $LogArea = GUICtrlCreateEdit($FULL_LOG_RESULTS, 10, 10, 1180, 780, $ES_MULTILINE+$ES_AUTOVSCROLL+$WS_VSCROLL+$ES_READONLY)

				;Set status to visible
				GUISetState(@SW_SHOW)

				_GUICtrlEdit_SetSel ( $LogArea,0,0)

				While 1

					 ; We can only get messages from the second GUI
					 Switch GUIGetMsg()
						 Case $GUI_EVENT_CLOSE
							 GUIDelete($CampaignerLogGUI)
							 ExitLoop

					EndSwitch
				WEnd


			;If the preview item was pressed
			Case $msg = $HTMLPreviewButton

				;Get current item to an array
				$SelectedListItem = _GUICtrlListView_GetItemTextArray($PreviewerListview)

				;If no item selected
				If StringLen(StringStripWS(_ArrayToString($SelectedListItem,"",1,$SelectedListItem[0]),8)) = "" Then

					;Notify
					MsgBox(4096, "Notification", 'Please select an item from the list.', 0, $CampaignerGUI)

				Else

					; Create a constant variable in Local scope of the filepath that will be read/written to.
					$sFilePath = @TempDir & "\"&@YEAR&@MON&@MDAY&@HOUR&@MIN&@SEC&@MSEC&random(0,10000000)&"CampaignerTemplate.html"

					; Create a temporary file to write data to.
					If Not _FileCreate($sFilePath) Then

						;Notify
						MsgBox(4096, "Error", 'Unable to create html preview file "'&$sFilePath&'".', 0, $CampaignerGUI)

					Else

						; Open the file for writing (append to the end of a file) and store the handle to a variable.
						Local $hFileOpen = FileOpen($sFilePath, 258)

						If $hFileOpen = -1 Then

							;Notify
							MsgBox(4096, "Error", 'Unable to open file "'&$sFilePath&'" for writing.', 0, $CampaignerGUI)

						Else

							;Get selected index
							$SelectedIndex = _GUICtrlListView_GetSelectedIndices ($PreviewerListview) + 1

							;Write data to the file using the handle returned by FileOpen.
							If FileWrite($hFileOpen, $ReadyToEmailArray[$SelectedIndex][2]) = 0 Then

								;Notify
								MsgBox(4096, "Error", 'Unable to write html code to file "'&$sFilePath&'".', 0, $CampaignerGUI)

							Else

								;Close the handle returned by FileOpen.
								FileClose($hFileOpen)

								;Launch default browser
								ShellExecute($sFilePath)

							EndIf

						EndIf

					EndIf

				EndIf

			;If launch was selected
			Case $msg = $RunButton

				;Reset logs
				$FULL_LOG_RESULTS = ""

				GUICtrlSetState ( $UnlockUIButton, $GUI_DISABLE)
				GUICtrlSetState ( $RunButton, $GUI_DISABLE)
				GUICtrlSetState ( $PreviewerListview, $GUI_DISABLE)
				GUICtrlSetState ( $HTMLPreviewButton, $GUI_DISABLE)
				GUICtrlSetState ( $LogPreviewButton, $GUI_DISABLE)

				$FULL_LOG_RESULTS &= "----------------------------------------------------------------------------------------------------------------------"&@CRLF
				$FULL_LOG_RESULTS &= "   _____ _      ____  ____          _        _____        _____            __  __ ______ _______ ______ _____   _____ "&@CRLF
				$FULL_LOG_RESULTS &= "  / ____| |    / __ \|  _ \   /\   | |      |  __ \ /\   |  __ \     /\   |  \/  |  ____|__   __|  ____|  __ \ / ____|"&@CRLF
				$FULL_LOG_RESULTS &= " | |  __| |   | |  | | |_) | /  \  | |      | |__) /  \  | |__) |   /  \  | \  / | |__     | |  | |__  | |__) | (___  "&@CRLF
				$FULL_LOG_RESULTS &= " | | |_ | |   | |  | |  _ < / /\ \ | |      |  ___/ /\ \ |  _  /   / /\ \ | |\/| |  __|    | |  |  __| |  _  / \___ \ "&@CRLF
				$FULL_LOG_RESULTS &= " | |__| | |___| |__| | |_) / ____ \| |____  | |  / ____ \| | \ \  / ____ \| |  | | |____   | |  | |____| | \ \ ____) |"&@CRLF
				$FULL_LOG_RESULTS &= "  \_____|______\____/|____/_/    \_\______| |_| /_/    \_\_|  \_\/_/    \_\_|  |_|______|  |_|  |______|_|  \_\_____/ "&@CRLF
				$FULL_LOG_RESULTS &= @CRLF
				$FULL_LOG_RESULTS &= "----------------------------------------------------------------------------------------------------------------------"&@CRLF


				$FULL_LOG_RESULTS &= @YEAR&'-'&@MON&'-'&@MDAY&' '&@HOUR&':'&@MIN&':'&@SEC&'.'&@MSEC&"->Transmission Daemon Started."&@CRLF
				$FULL_LOG_RESULTS &= "                         Daemon Configuration -> SMTP Server       -> "&$SMTPMailServer&@CRLF
				$FULL_LOG_RESULTS &= "                                                 SMTP Server Port  -> "&$SMTPPort&@CRLF
				$FULL_LOG_RESULTS &= "                                                 SMTP SSL (Bool)   -> "&$SMTPSSL&@CRLF
				$FULL_LOG_RESULTS &= "                                                 SMTP Username     -> "&$SMTPAccountUsername&@CRLF
				$FULL_LOG_RESULTS &= "                                                 SMTP Password     -> "&$SMTPAccountPassword&@CRLF
				$FULL_LOG_RESULTS &= "                                                 SMTP Importance   -> "&$SMTPImportance&@CRLF
				$FULL_LOG_RESULTS &= "                                                 SMTP From Name    -> "&$SMTPFromName&@CRLF
				$FULL_LOG_RESULTS &= "                                                 SMTP From Address -> "&$SMTPFromAddress&@CRLF
				$FULL_LOG_RESULTS &= "                                                 SMTP CC Address   -> "&$SMTPCCAddress&@CRLF
				$FULL_LOG_RESULTS &= "                                                 SMTP BCC Address  -> "&$SMTPBCCAddress&@CRLF



				If $MailAttachments <> '' Then

					$FULL_LOG_RESULTS &= "                                                 SMTP Attachments  -> "&$TempAttachmentsFPArray[1]&@CRLF

					if $TempAttachmentsFPArray[0] > 1 Then

						For $i = 2 To $TempAttachmentsFPArray[0]

							$FULL_LOG_RESULTS &= "                                                                   -> "&$TempAttachmentsFPArray[$i]&@CRLF

						Next

					EndIf

				EndIf

				$FULL_LOG_RESULTS &= "----------------------------------------------------------------------------------------------------------------------"&@CRLF
				$FULL_LOG_RESULTS &= "   _____          __  __ _____        _____ _____ _   _     _              _    _ _   _  _____ _    _ "&@CRLF
				$FULL_LOG_RESULTS &= "  / ____|   /\   |  \/  |  __ \ /\   |_   _/ ____| \ | |   | |        /\  | |  | | \ | |/ ____| |  | |"&@CRLF
				$FULL_LOG_RESULTS &= " | |       /  \  | \  / | |__) /  \    | || |  __|  \| |   | |       /  \ | |  | |  \| | |    | |__| |"&@CRLF
				$FULL_LOG_RESULTS &= " | |      / /\ \ | |\/| |  ___/ /\ \   | || | |_ | . ` |   | |      / /\ \| |  | | . ` | |    |  __  |"&@CRLF
				$FULL_LOG_RESULTS &= " | |____ / ____ \| |  | | |  / ____ \ _| || |__| | |\  |   | |____ / ____ \ |__| | |\  | |____| |  | |"&@CRLF
				$FULL_LOG_RESULTS &= "  \_____/_/    \_\_|  |_|_| /_/    \_\_____\_____|_| \_|   |______/_/    \_\____/|_| \_|\_____|_|  |_|"&@CRLF
				$FULL_LOG_RESULTS &= @CRLF

				;Clear the listview
				_GUICtrlListView_DeleteAllItems($PreviewerListview)

				$TotalSend = 0

				For $i = 1 To Ubound($ReadyToEmailArray)-1

					$FULL_LOG_RESULTS &= "----------------------------------------------------------------------------------------------------------------------"&@CRLF
					$FULL_LOG_RESULTS &= @YEAR&'-'&@MON&'-'&@MDAY&' '&@HOUR&':'&@MIN&':'&@SEC&'.'&@MSEC&"->Sending email "&$i&" of "&(Ubound($ReadyToEmailArray)-1)&"."&@CRLF
					$FULL_LOG_RESULTS &= "                         Recepient set to -> "&$ReadyToEmailArray[$i][0]&@CRLF
					$FULL_LOG_RESULTS &= "                         Subject set to   -> "&$ReadyToEmailArray[$i][1]&@CRLF

					$LogMailBodyArray = StringSplit(StringRegExpReplace($ReadyToEmailArray[$i][2], "\r\n|\r|\n", @LF),@LF)

					$LogMailBodyString = ''


					$test123  = StringReplace ($ReadyToEmailArray[$i][2], @TAB, "")
					$test123  = StringRegExpReplace($test123, "\r\n|\r|\n", '')
					$test123  = _StringChop($test123, 73)

					For $l = 1 To $test123[0]

						if $l = 1  Then

							$LogMailBodyString &= "                         Body set to      -> "&$test123[$l]&@CRLF

						Else

							$LogMailBodyString &= "                                          -> "&$test123[$l]&@CRLF

						EndIf

					Next

					$FULL_LOG_RESULTS &= $LogMailBodyString

					$rc = _INetSmtpMailCom($SMTPMailServer, $SMTPFromName, $SMTPFromAddress, $ReadyToEmailArray[$i][0], $ReadyToEmailArray[$i][1], $ReadyToEmailArray[$i][2], $MailAttachments, $SMTPCCAddress, $SMTPBCCAddress, $SMTPImportance, $SMTPAccountUsername, $SMTPAccountPassword, $SMTPPort, $SMTPSSL)

					If @error Then

						$TempStatus =  "                         Status           -> Error code "&@error&" returned from _INetSmtpMailCom."&@CRLF
						$TempStatus &= "                                             Error description -> "&$oMyRet[1]&@CRLF
						$TempStatus &= "                                             Hex code          -> "&$oMyRet[0]&@CRLF


						$FULL_LOG_RESULTS &= $TempStatus

						GUICtrlCreateListViewItem($ReadyToEmailArray[$i][0]&"|"&$ReadyToEmailArray[$i][1]&"|Failed", $PreviewerListview)

					Else

						$FULL_LOG_RESULTS &= "                         Status           -> Transmission was successful."&@CRLF

						$TotalSend+=1

						GUICtrlCreateListViewItem($ReadyToEmailArray[$i][0]&"|"&$ReadyToEmailArray[$i][1]&"|OK", $PreviewerListview)

					EndIf

				Next

				$FULL_LOG_RESULTS &= "----------------------------------------------------------------------------------------------------------------------"&@CRLF
				$FULL_LOG_RESULTS &= "  _____  ______  _____ _    _ _   _______ _____ "&@CRLF
				$FULL_LOG_RESULTS &= " |  __ \|  ____|/ ____| |  | | | |__   __/ ____|"&@CRLF
				$FULL_LOG_RESULTS &= " | |__) | |__  | (___ | |  | | |    | | | (___  "&@CRLF
				$FULL_LOG_RESULTS &= " |  _  /|  __|  \___ \| |  | | |    | |  \___ \ "&@CRLF
				$FULL_LOG_RESULTS &= " | | \ \| |____ ____) | |__| | |____| |  ____) |"&@CRLF
				$FULL_LOG_RESULTS &= " |_|  \_\______|_____/ \____/|______|_| |_____/ "&@CRLF
				$FULL_LOG_RESULTS &= ""&@CRLF
				$FULL_LOG_RESULTS &= "----------------------------------------------------------------------------------------------------------------------"&@CRLF

				$FULL_LOG_RESULTS &= "Total Planned                             -> "&(Ubound($ReadyToEmailArray)-1)&@CRLF
				$FULL_LOG_RESULTS &= "Total Sent                                -> "&$TotalSend&@CRLF

				GUICtrlSetState ( $UnlockUIButton, $GUI_ENABLE)
				GUICtrlSetState ( $RunButton, $GUI_ENABLE)
				GUICtrlSetState ( $PreviewerListview, $GUI_ENABLE)
				GUICtrlSetState ( $HTMLPreviewButton, $GUI_ENABLE)
				GUICtrlSetState ( $LogPreviewButton, $GUI_ENABLE)

			;If unlock button was pressed
			Case $msg = $UnlockUIButton

				;Lock UI elements to avoid changes or errors
				GUICtrlSetState ($MailServerInput, $GUI_ENABLE)
				GUICtrlSetState ($MailPortInput, $GUI_ENABLE)
				GUICtrlSetState ($MailSSLInput, $GUI_ENABLE)
				GUICtrlSetState ($MailSSLInput, $GUI_ENABLE)
				GUICtrlSetState ($MailAcountUsernameInput, $GUI_ENABLE)
				GUICtrlSetState ($MailAcountPasswordInput, $GUI_ENABLE)
				GUICtrlSetState ($FromNameInput, $GUI_ENABLE)
				GUICtrlSetState ($FromAddressInput, $GUI_ENABLE)
				GUICtrlSetState ($MailImportanceInput, $GUI_ENABLE)
				GUICtrlSetState ($DataFileButton, $GUI_ENABLE)
				GUICtrlSetState ($HTMLTemplateButton, $GUI_ENABLE)
				GUICtrlSetState ($CheckButton, $GUI_ENABLE)
				GUICtrlSetState ($AttachmentDeleteButton, $GUI_ENABLE)
				GUICtrlSetState ($AttachmentShowButton, $GUI_ENABLE)
				GUICtrlSetState ($AttachmentFileButton, $GUI_ENABLE)
				GUICtrlSetState ($CCInput, $GUI_ENABLE)
				GUICtrlSetState ($BCCInput, $GUI_ENABLE)
				GUICtrlSetState ($TelnetTestButton, $GUI_ENABLE)

				;disable  unlock button
				GUICtrlSetState ( $UnlockUIButton, $GUI_DISABLE)

				;disable launch button
				GUICtrlSetState ( $RunButton, $GUI_DISABLE)

				;Clear the listview
				_GUICtrlListView_DeleteAllItems($PreviewerListview)

				;Disable the log preview button
				GUICtrlSetState ( $LogPreviewButton, $GUI_DISABLE)

				;Disable the preview button
				GUICtrlSetState ( $HTMLPreviewButton, $GUI_DISABLE)

				;Update the counter label
				GUICtrlSetData ($TotalCampaignMailsCounter, '')

			;If attachment selection button was pressed
			Case $msg = $AttachmentFileButton

				;Open the file selection dialog
				Dim $AttachmentFile = FileOpenDialog ( "Attachment File Selection", @ScriptDir, "All (*.*)" ,$FD_MULTISELECT + $FD_PROMPTCREATENEW, "", $CampaignerGUI)

				;If an error occured
				If @error Then


				Else

					;Update the file input box
					GUICtrlSetData($AttachmentFileInput, $AttachmentFile)

				EndIf

			;If attachment removal button was selected
			Case $msg = $AttachmentDeleteButton

				;Update the file input box
				GUICtrlSetData($AttachmentFileInput, "")

			;If attachment show button was selected
			case $msg = $AttachmentShowButton

				;Get Files
				$TempAttachedFiles = GUICtrlRead($AttachmentFileInput)

				;If nothing was selected
				If $TempAttachedFiles = '' Then

					;Notify
					MsgBox(4096, "Notification", 'No attachments have been selected.', 0, $CampaignerGUI)

				;If something was selected
				Else

					;Convert the string to an array
					$TempAttachedFilesArray = StringSplit($TempAttachedFiles,'|')

					If @error = 1 Then

						;Generate a blank string to fill with attachments
						$AttachmentsString = 'The following file has been selected:'&@CRLF&@CRLF&$TempAttachedFiles

						;Notify
						MsgBox(4096, "Notification", $AttachmentsString, 0, $CampaignerGUI)

					Else

						;Get path
						$AttachmentsPath = $TempAttachedFilesArray[1]

						;Generate a blank string to fill with attachments
						$AttachmentsString = 'The following files have been selected:'&@CRLF&@CRLF

						;Loop through files and generate data
						For $i = 2 To $TempAttachedFilesArray[0]

							;Update the string
							$AttachmentsString&=$AttachmentsPath&'\'&$TempAttachedFilesArray[$i]&@CRLF

						Next

						;Notify
						MsgBox(4096, "Notification", $AttachmentsString, 0, $CampaignerGUI)

					EndIf

				EndIf

		EndSelect

	Wend

	Func _StringChop($string, $size)
	$count = Ceiling(StringLen($string)/$size)
	Dim $array[$count+1], $start = 1
	For $i = 1 To $count
		$array[$i] = StringMid($string, $start, $size)
		$start += $size
	Next
	$array[0] = $count
	Return $array
	EndFunc
