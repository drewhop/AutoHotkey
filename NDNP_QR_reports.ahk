; ********************************
; NDNP_QR_reports.ahk
; required file for NDNP_QR.ahk
; ********************************

; ======HELPER FUNCTIONS
Create_Divider(path)
{
		divider =
		StringLen, length, path
		length++
		Loop, %length%
		{
			divider .= "-"
		}
		return divider
}

BatchFolderPath(reelfolderpath)
{
	StringGetPos, batchpathpos, reelfolderpath, \, R2
	StringLeft, batchfolderpath, reelfolderpath, batchpathpos
	return batchfolderpath
}

ReelFolderName(reelfolderpath)
{
	StringGetPos, foldernamepos, reelfolderpath, \, R1
	foldernamepos++
	StringTrimLeft, reelfoldername, reelfolderpath, foldernamepos
	return reelfoldername
}

Format_Issue_Data(identifier, volume, issue, date, questionable, questionabledisplay, numpages, missing, missingdisplay, editionlabel)
{			
	; tab insert after identifier
	StringLen, length, identifier
	if (length < 24)
	{
		identitab := "`t`t"
	}
	else
	{
		identitab := "`t"
	}

	; tab insert after volume number	
	if volume is integer
	{
		volumetab := "`t"
	}
	else ; alphabetic volume number
	{
		StringLen, length, volume
		if (length < 8)
		{
			if (volume = "") ; no volume number
			{
				volumetab := "`t"
			}
			else
			{
				volumetab := "`t`t"
			}
		}
		else
		{
			volumetab := "`t"
		}		
	}
	
	; tab insert after issue number
	if issue is integer
	{
		issuetab := "`t"
	}
	else ; alphabetic issue number
	{
		StringLen, length, issue
		if (length < 8)
		{
			if (issue = "") ; no issue number
			{
				issuetab := "`t"
			}
			else
			{
				issuetab := "`t`t"
			}
		}
		else
		{
			issuetab := "`t"
		}		
	}

	; tab insert after date
	if (questionable == "")
	{
		datetab := "`t`t"
	}
	else
	{
		questionabledisplay = (%questionable%)
		datetab := "`t"
	}

	if (missing > 0)
	{
		missingdisplay = (%missing%)
	}
	else
	{
		missingdisplay =
	}
	
	if Mod((numpages+missing), 2)=0
	{    
		oddpages := ""
	}
	else
	{
		oddpages := " --"
	}
	
	issuedata = %identifier%%identitab%%volume%%volumetab%%issue%%issuetab%%date% %questionabledisplay%%datetab%%numpages% %missingdisplay%%oddpages%`t%editionlabel%
	
	return issuedata
}
; ======HELPER FUNCTIONS

; ======ISSUE INFO
IssueInfo:
	; issue folder path
	StringTrimRight, issuepath, A_LoopField, 15
	StringRight, issuefoldername, issuepath, 10
	issuefolderpath = %batchfolderpath%\%issuepath%

	; LCCN and Reel #
	StringTrimRight, identifier, A_LoopField, 26
	Sleep, 100	
Return
; ======ISSUE INFO

; ======BATCH PARSE REPORTS
BatchParseReports:
	; remove </issue> tag after the file path
	StringTrimRight, rawissueXMLpath, A_LoopField, 8

	; remove <issue> tag before the file path
	StringGetPos, pos, rawissueXMLpath, >
	pos++
	StringTrimLeft, issueXMLpath, rawissueXMLpath, %pos%
			
	; append issue.xml path to issuefile
	issuefile .= issueXMLpath
	issuefile .= "`n"
	issuecount++				
Return
; ======BATCH PARSE REPORTS

; ======BATCH REPORT
; creates a report of all the issues in a batch
; text file will be created in the batch folder
BatchReport:
	ControlSetText, Static3, BATCH, NDNP_QR
 
	; initialize variables
	issuefile =
	batchxml =
	metaloopflag = 0

	FileSelectFolder, batchfolderpath, %batchdrive%, 2, BATCH REPORT`n`nSelect a BATCH folder:
	if ErrorLevel
		Return
	else
	{
		; batch name
		StringGetPos, foldernamepos, batchfolderpath, \, R1
		foldernamepos++
		StringTrimLeft, batchname, batchfolderpath, foldernamepos

		; notification window
		Gui, 15:+ToolWindow
		Gui, 15:Add, Text,, Processing:  %batchname%
		Gui, 15:Add, Text,, This may take awhile . . .
		SetTitleMatchMode 1
		WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
		Gui, 15:Show, x%winX% y%winY%, Batch Report
	
		start = %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% 

		divider := Create_Divider(batchfolderpath)
		
		; create report file in batch folder
		FileAppend, %divider%`n%batchfolderpath%`n%divider%`n`n, %batchfolderpath%\%batchname%-report.txt
		FileAppend, %reportheader%`n`n, %batchfolderpath%\%batchname%-report.txt

		; read in batch.xml file to batchxml variable
		FileRead, batchxml, %batchfolderpath%\batch.xml

		; initialize counters
		issuecount = 0
		totalpages = 0
		runcount = 1
		
		; parse batchxml and store issue paths
		Loop, parse, batchxml, `n, `r%A_Space%%A_Tab%
		{
			IfInString, A_LoopField, issue
			{
				Gosub, BatchParseReports
			}
		}
		
		Sort, issuefile

		; *******************************
		; parse sorted issuefile
		; *******************************
		Loop, parse, issuefile, `n
		{
			; *************************************************
			; AUTO EXIT when issuefile is empty
			if A_LoopField =
			{
				FileAppend, `nBatch: %batchname%`nIssues: %issuecount%`n Pages: %totalpages%, %batchfolderpath%\%batchname%-report.txt
				FileAppend, `n`nSTART: %start%, %batchfolderpath%\%batchname%-report.txt
				FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %batchfolderpath%\%batchname%-report.txt

				Gui, 15:Destroy ; notification window
				
				MsgBox, 4, Batch Report, Batch: %batchname%`n`nIssues: %issuecount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nThe report is complete. Would you like to open it?
				IfMsgBox, Yes
					Run, %batchfolderpath%\%batchname%-report.txt
				Return
			}
			; *************************************************

			Gosub, IssueInfo
			
			; first identifiers for report dividers
			if (runcount == 1)
			{
				startidentifier = %identifier%
				StringTrimRight, startlccn, identifier, 12
				FileAppend, %lccndivider%`n%lccnurl%%startlccn%`n%lccndivider%`n, %batchfolderpath%\%batchname%-report.txt
			}

			; lccn
			StringTrimRight, lccn, identifier, 12

			; initialize variables
			issuexml =
			volume =
			issue =
			date =
			editionlabel =
			questionable =
			questionabledisplay =
			volumeflag = 0
			issueflag = 0
			missing = 0
			missingdisplay =

			; read in isue.xml file to issuexml variable
			FileRead, issuexml, %batchfolderpath%\%A_LoopField%

			Loop, parse, issuexml, `n, `r%A_Space%%A_Tab%
			{
				Gosub, IssueMetadata ; NDNP_QR_metadata.ahk
			}

			numpages = 0
			Loop, %issuefolderpath%\*.tif ; count tiffs
			{
				numpages++
			}

			totalpages += numpages					

			if (identifier <> startidentifier) ; dividers between reels
			{
				if (lccn <> startlccn) ; lccn divider
				{
					StringTrimRight, startlccn, identifier, 12			
					FileAppend, `n%lccndivider%`n%lccnurl%%startlccn%`n%lccndivider%`n, %batchfolderpath%\%batchname%-report.txt
				}
				else ; reel divider
				{
					FileAppend, %reeldivider%`n, %batchfolderpath%\%batchname%-report.txt
				}
				startidentifier = %identifier%
			}

			issuedata := Format_Issue_Data(identifier, volume, issue, date, questionable, questionabledisplay, numpages, missing, missingdisplay, editionlabel)

			FileAppend, %issuedata%`n, %batchfolderpath%\%batchname%-report.txt
			
			runcount++
		}
	}
Return
; ======BATCH REPORT

; ======BATCH PARSE OCR
BatchParseOCR:
	IfInString, A_LoopField, issue
	{
		IfInString, A_LoopField, %reelnumber%
		{
			IfInString, A_LoopField, %lccn%
			{			
				; remove </issue> tag after the folder path
				StringTrimRight, rawissuefolderpath, A_LoopField, 23

				; remove <issue> tag before the folder path
				StringGetPos, pos, rawissuefolderpath, >
				pos++
				StringTrimLeft, issuefolderpath, rawissuefolderpath, %pos%

				; append issue folder path to issuefile
				issuefile .= issuefolderpath
				issuefile .= "`n"
			}
		}
	}
Return
; ======BATCH PARSE OCR

; ======LANGUAGE CODE REPORT
EditLanguageCode:
	if (languagecode == "eng")
		languagechoice = 1
	else if (languagecode == "fre")
		languagechoice = 2
	else if (languagecode == "ger")
		languagechoice = 3
	else if (languagecode == "ita")
		languagechoice = 4
	else if (languagecode == "spa")
		languagechoice = 5

	Gui, 18:+ToolWindow
	Gui, 18:Add, Text,, Choose a code:
	Gui, 18:Add, DropDownList, Choose%languagechoice% R5 vlanguagecode, eng|fre|ger|ita|spa
	Gui, 18:Add, Button, w40 x10 y55 gLanguageGo default, OK
	Gui, 18:Add, Button, x65 y55 gLanguageCancel, Cancel	
	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	Gui, 18:Show, x%winX% y%winY%, Language Code
Return

; OK button
LanguageGo:
	Gui, 18:Submit
	Gui, 18:Destroy
Return
		
; Cancel button
LanguageCancel:
	Gui, 18:Destroy
	cancelbutton = 1
Return

LanguageCodeReport:
	ControlSetText, Static3, LANGUAGE, NDNP_QR
 
	Gosub, EditLanguageCode
	SetTitleMatchMode 1
	WinWaitActive, Language Code
	WinWaitClose, Language Code

	if (cancelbutton == 1)
	{
		cancelbutton = 0 ; reset
		Return
	}
	
	FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
	if ErrorLevel
		Return
	else
	{
		Loop ; check for valid reel folder
		{
			reelfoldername := ReelFolderName(reelfolderpath)
			
			StringLen, length, reelfoldername ; check for valid reel folder
			if (length != 11)
			{		
				MsgBox, 0, Language Code Report, %reelfoldername% does not appear to be a REEL folder.`n`n`tClick OK to continue.
				reelfolderpath = _
				reelfoldername = _

				FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
				if ErrorLevel
					Return
				else Continue
			}
				
			else Break
		}
	
		Gosub, GetReportInfo

		Gui, 19:+ToolWindow
		Gui, 19:Add, Text,, Processing:  %lccn%-%reelfoldername%-%languagecode%-report.txt
		Gui, 19:Add, Text,, Please wait . . .
		SetTitleMatchMode 1
		WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
		Gui, 19:Show, x%winX% y%winY%, Language Code Report
	
		divider := Create_Divider(reelfolderpath)

		; create report file
		FileAppend, %divider%`n%reelfolderpath%`n%divider%`n`n, %reportpath%\%lccn%-%reelnumber%-%languagecode%-report.txt
		FileAppend, Language Code: %languagecode%`n`n, %reportpath%\%lccn%-%reelnumber%-%languagecode%-report.txt
			
		; initialize variables
		issuefile =
		issuecount = 0
			
		batchfolderpath := BatchFolderPath(reelfolderpath)
			
		; read in batch.xml file to batchxml variable
		FileRead, batchxml, %batchfolderpath%\batch.xml

		; parse batchxml and store issue paths
		Loop, parse, batchxml, `n, `r%A_Space%%A_Tab%
		{
			Gosub, BatchParseOCR
		}
			
		Sort, issuefile

		; initialize counters
		totalcodes = 0
		issuecodes = 0
		
		; *******************************
		; parse sorted issuefile
		; *******************************
		Loop, parse, issuefile, `n
		{
			; *************************************************
			; AUTO EXIT when issuefile is empty
			if A_LoopField =
			{
				FileAppend, `n___________________________`n`n, %reportpath%\%lccn%-%reelnumber%-%languagecode%-report.txt
				FileAppend, LCCN: %lccn%`nReel: %reelnumber%`nCode: %languagecode%`n`nIssues: %issuecount%`nBlocks: %totalcodes%, %reportpath%\%lccn%-%reelnumber%-%languagecode%-report.txt
				FileAppend, `n`nSTART: %start%, %reportpath%\%lccn%-%reelnumber%-%languagecode%-report.txt
				FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %reportpath%\%lccn%-%reelnumber%-%languagecode%-report.txt

				Gui, 19:Destroy ; notification window
				
				MsgBox, 4, Language Code Report, LCCN: %lccn%`nReel: %reelnumber%`nCode: %languagecode%`n`nIssues: %issuecount%`nCodes: %totalcodes%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nThe report is complete. Would you like to open it?
				IfMsgBox, Yes
					Run, %reportpath%\%lccn%-%reelnumber%-%languagecode%-report.txt

				return
			}
			; *************************************************

			; issue date
			StringRight, issuedate, A_LoopField, 10
			Sleep, 100	

			; issue folder path for page count
			issuefolderpath = %batchfolderpath%\%A_LoopField%

			; initialize variables and counters
			ALTOxml =
			issueresults =
			codecount = 0
			issuecodes = 0
			
			; parse the ALTO XML file
			Loop, %issuefolderpath%\*.xml
			{
				codecount = 0
				
				StringLen, length, A_LoopFileName
				if (length < 11)
				{
					; read in file to ALTOxml variable
					FileRead, ALTOxml, %issuefolderpath%\%A_LoopFileName%
					
					; parse the textblocks
					Loop, parse, ALTOxml, >, %A_Space%%A_Tab%
					{
						IfInString, A_LoopField, language="%languagecode%"
						{
							codecount++
						}
					}				

					if (codecount > 0) ; add line to issueresults
					{
						issueresults .= A_LoopFileName
						issueresults .= "`t"
						issueresults .= codecount
						issueresults .= "`n"
					}

					issuecodes += codecount
					totalcodes += codecount
				}
			}

			if (issuecodes > 0) ; add to report
			{
				FileAppend, ____________________`n`n, %reportpath%\%lccn%-%reelnumber%-%languagecode%-report.txt			
				FileAppend, %issuedate%`n`n, %reportpath%\%lccn%-%reelnumber%-%languagecode%-report.txt			
				FileAppend, %issueresults%, %reportpath%\%lccn%-%reelnumber%-%languagecode%-report.txt
				FileAppend, `n`t Total: %issuecodes%`n, %reportpath%\%lccn%-%reelnumber%-%languagecode%-report.txt				
				issuecount++
			}
		}
	}
Return
; ======LANGUAGE CODE REPORT

; ======OCR SEARCH
OCRSearch:
	ControlSetText, Static3, OCR, NDNP_QR
 
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	InputBox, ocrterm, OCR Search, Enter a search term:,, 250, 120, %winX%, %winY%,,, %ocrterm%
	if ErrorLevel
		Return
	else
	{
		FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
		if ErrorLevel
			Return
		else
		{
			Loop ; valid reel folder check
			{
				reelfoldername := ReelFolderName(reelfolderpath)
				
				StringLen, length, reelfoldername ; valid reel folder check
				if (length != 11)
				{		
					MsgBox, 0, Language Code Report, %reelfoldername% does not appear to be a REEL folder.`n`n`tClick OK to continue.

					reelfolderpath = _
					reelfoldername = _

					FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
					if ErrorLevel
						Return
					else Continue
				}
					
				else Break
			}
		
			Gosub, GetReportInfo

			Gui, 20:+ToolWindow
			Gui, 20:Add, Text,, Processing:  %lccn%-%reelnumber%-%ocrterm%-report.txt
			Gui, 20:Add, Text,, Please wait . . .
			SetTitleMatchMode 1
			WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
			Gui, 20:Show, x%winX% y%winY%, OCR Search
		
			divider := Create_Divider(reelfolderpath)

			; create report file
			FileAppend, %divider%`n%reelfolderpath%`n%divider%`n`n, %reportpath%\%lccn%-%reelnumber%-%ocrterm%-report.txt
			FileAppend, Search Term: %ocrterm%`n`n, %reportpath%\%lccn%-%reelnumber%-%ocrterm%-report.txt
				
			; initialize variables
			issuefile =
			issuecount = 0
				
			batchfolderpath := BatchFolderPath(reelfolderpath)
				
			; read in batch.xml file to batchxml variable
			FileRead, batchxml, %batchfolderpath%\batch.xml

			; parse batchxml and store issue paths
			Loop, parse, batchxml, `n, `r%A_Space%%A_Tab%
			{
				Gosub, BatchParseOCR
			}
				
			Sort, issuefile

			; initialize counters
			totalterms = 0
			issueterms = 0
			
			; *******************************
			; parse sorted issuefile
			; *******************************
			Loop, parse, issuefile, `n
			{
				; *************************************************
				; AUTO EXIT when issuefile is empty
				if A_LoopField =
				{
					FileAppend, `n___________________________`n`n, %reportpath%\%lccn%-%reelnumber%-%ocrterm%-report.txt
					FileAppend, LCCN: %lccn%`nReel: %reelnumber%`nTerm: %ocrterm%`n`nIssues: %issuecount%`n  Hits: %totalterms%, %reportpath%\%lccn%-%reelnumber%-%ocrterm%-report.txt
					FileAppend, `n`nSTART: %start%, %reportpath%\%lccn%-%reelnumber%-%ocrterm%-report.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %reportpath%\%lccn%-%reelnumber%-%ocrterm%-report.txt

					Gui, 20:Destroy ; notification window
					
					MsgBox, 4, OCR Search, LCCN: %lccn%`nReel: %reelnumber%`nTerm: %ocrterm%`n`nIssues: %issuecount%`nHits: %totalterms%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nThe report is complete. Would you like to open it?
					IfMsgBox, Yes
						Run, %reportpath%\%lccn%-%reelnumber%-%ocrterm%-report.txt

					return
				}
				; *************************************************

				; issue date
				StringRight, issuedate, A_LoopField, 10
				Sleep, 100	

				; issue folder path for page count
				issuefolderpath = %batchfolderpath%\%A_LoopField%

				; initialize variables and counters
				ALTOxml =
				issueresults =
				termcount = 0
				issueterms = 0
				
				; parse the ALTO XML files
				Loop, %issuefolderpath%\*.xml
				{
					termcount = 0
					
					StringLen, length, A_LoopFileName
					if (length < 11)
					{
						; read in file to ALTOxml variable
						FileRead, ALTOxml, %issuefolderpath%\%A_LoopFileName%
						
						; parse the textblocks
						Loop, parse, ALTOxml, >, %A_Space%%A_Tab%
						{
							IfInString, A_LoopField, %ocrterm%
							{
								termcount++
							}
						}				

						; add line to issueresults if code found
						if (termcount > 0)
						{
							issueresults .= A_LoopFileName
							issueresults .= "`t"
							issueresults .= termcount
							issueresults .= "`n"
						}

						issueterms += termcount
						totalterms += termcount
					}
				}

				if (issueterms > 0) ; add to report
				{
					FileAppend, ____________________`n`n, %reportpath%\%lccn%-%reelnumber%-%ocrterm%-report.txt			
					FileAppend, %issuedate%`n`n, %reportpath%\%lccn%-%reelnumber%-%ocrterm%-report.txt			
					FileAppend, %issueresults%, %reportpath%\%lccn%-%reelnumber%-%ocrterm%-report.txt
					FileAppend, `n`t Total: %issueterms%`n, %reportpath%\%lccn%-%reelnumber%-%ocrterm%-report.txt				
								
					issuecount++
					issueterms = 0 ; reset
				}
			}
		}
	}
Return
; ======OCR SEARCH
