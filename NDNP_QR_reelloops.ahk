; ********************************
; NDNP_QR_reelloops.ahk
; required file for NDNP_QR.ahk
; ********************************

; ======REEL REPORT
; displays first TIFF and metadata
; and creates a report for the issues in a reel folder
; text file will be created in the LCCN folder
ReelReport:
	metaloopname = Reel Report
	
	ControlSetText, Static3, REEL, NDNP_QR
	ControlSetText, Static4, Reel Report, NDNP_QR

	metaloopflag = 1
	reelreportflag = 1

	Gosub, ReelLoopFunction
Return
; ======REEL REPORT

; ======METADATA VIEWER
; opens the first TIFF and displays the issue metadata
; for each issue in a reel folder
MetaViewer:
	metaloopname = Metadata Viewer
	
	ControlSetText, Static3, VIEWER, NDNP_QR
	ControlSetText, Static4, Meta Viewer, NDNP_QR

	metaloopflag = 1

	Gosub, ReelLoopFunction
Return
; ======METADATA VIEWER

; ======GET REPORT INFO
GetReportInfo:
	start = %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% 

	; display variables for reel folder path
	StringGetPos, reelfolderpos, reelfolderpath, \, R2
	StringLeft, reelfolder1, reelfolderpath, reelfolderpos
	StringTrimLeft, reelfolder2, reelfolderpath, reelfolderpos

	; report path (LCCN folder)
	StringGetPos, foldernamepos, reelfolderpath, \, R1
	StringLeft, reportpath, reelfolderpath, foldernamepos

	; reel number
	foldernamepos++
	StringTrimLeft, reelnumber, reelfolderpath, foldernamepos
		
	; lccn
	StringGetPos, lccnpos, reportpath, \, R1
	lccnpos++
	StringTrimLeft, lccn, reportpath, lccnpos
Return

; ======REEL LOOP FUNCTION
; Reel Report & Metadata Viewer
ReelLoopFunction:
	; initialize variables
	issuefile =
	batchxml =

	; initialize counters
	loopcount = 0
	totalpages = 0
	notecount = 0

	FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
	if ErrorLevel
		Return
	else
	{
		Loop ; valid reel folder check
		{
			reelfoldername := ReelFolderName(reelfolderpath) ; NDNP_QR_reports.ahk
			
			StringLen, length, reelfoldername ; valid reel folder check
			if (length != 11)
			{		
				MsgBox, 0, Reel Loop, %reelfoldername% does not appear to be a REEL folder.`n`n`tClick OK to continue.

				reelfolderpath = _
				reelfoldername = _

				FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
				if ErrorLevel
					Return
				else Continue
			}
				
			else Break
		}
		
		Gosub, DelayDialog
		SetTitleMatchMode 1
		WinWaitActive, Delay Time
		WinWaitClose, Delay Time

		if (cancelbutton == 1)
		{
			cancelbutton = 0 ; reset
			Return
		}

		Gosub, CloseFirstImpressionWindows			
			
		Gosub, GetReportInfo

		divider := Create_Divider(reelfolderpath) ; NDNP_QR_reports.ahk
		
		if (reelreportflag == 1) ; create report file
		{
			FileAppend, %divider%`n%reelfolderpath%`n%divider%`n`n%reportheader%`n`n, %reportpath%\%lccn%-%reelnumber%-report.txt
		}

		batchfolderpath := BatchFolderPath(reelfolderpath) ; NDNP_QR_reports.ahk
			
		; read in batch.xml file to batchxml variable
		FileRead, batchxml, %batchfolderpath%\batch.xml

		; parse batchxml and store issue paths
		Loop, parse, batchxml, `n, `r%A_Space%%A_Tab%
		{
			IfInString, A_LoopField, issue
			{
				IfInString, A_LoopField, %reelnumber%
				{
					IfInString, A_LoopField, %lccn%
					{			
						Gosub, BatchParseReports ; NDNP_QR_reports.ahk
					}
				}
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
				; close metadata windows
				Gui, 4:Destroy
				Gui, 5:Destroy
				Gui, 6:Destroy
				Gui, 7:Destroy
				Gui, 9:Destroy

				; clear Metadata window
				ControlSetText, Static6,, Metadata
				ControlSetText, Static7,, Metadata		
				ControlSetText, Static8,, Metadata		
				ControlSetText, Static9,, Metadata
				ControlSetText, Static10,, Metadata
				
				if (reelreportflag == 1) ; Reel Report
				{
					FileAppend, `nIssues: %loopcount%`n Pages: %totalpages%, %reportpath%\%lccn%-%reelnumber%-report.txt
					FileAppend, `n`nSTART: %start%, %reportpath%\%lccn%-%reelnumber%-report.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %reportpath%\%lccn%-%reelnumber%-report.txt

					MsgBox, 4, Reel Report, LCCN: %lccn%`nReel: %reelnumber%`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nThe report is complete. Would you like to open it?
					IfMsgBox, Yes
						Run, %reportpath%\%lccn%-%reelnumber%-report.txt

					reelreportflag = 0
					loopcount = 0
					return
				}
					
				else if (notecount > 0) ; Metadata Viewer with Notes
				{
					FileAppend, `nSTART: %start%, %reportpath%\%lccn%-%reelnumber%-notes.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %reportpath%\%lccn%-%reelnumber%-notes.txt

					MsgBox, 4, Metadata Viewer, LCCN: %lccn%`nReel: %reelnumber%`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nYou added %notecount% notes. Would you like to open the Notes file?
					IfMsgBox, Yes
						Run, %reportpath%\%lccn%-%reelnumber%-notes.txt
					loopcount = 0

					return						
				}
					
				else ; Metadata Viewer
				{
					MsgBox, 0, Metadata Viewer, LCCN: %lccn%`nReel: %reelnumber%`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%
					loopcount = 0

					return
				}
			}
			; *************************************************

			Gosub, IssueInfo ; NDNP_QR_reports.ahk
			
			Gosub, ExtractMeta ; NDNP_QR_metadata.ahk

			if (numpages > 0)
			{
				totalpages += %numpages%

				; open first TIF file
				Loop, %issuefolderpath%\*.tif
				{
					Run, %issuefolderpath%\%A_LoopFileName%
					Break
				}
					
				; First Impression window id#
				SetTitleMatchMode 2
				WinWaitActive, First Impression
				Sleep, 100
				WinGet, firstid, ID, First Impression

				; activate First Impression
				SetTitleMatchMode 1
				WinActivate, ahk_id %firstid%
				WinWaitActive, ahk_id %firstid%
				Sleep, 100
					
				; zoom out
				Send, {NumpadSub %imagewidth%}

				; send First Impression backward
				WinSet, Bottom,, ahk_id %firstid%
				Sleep, 300
			}
			else
			{
				Msgbox, 0, This issue has no images.
			}
				
			loopcount++

			if (reelreportflag == 1) ; Reel Report / Format_Issue_Data in NDNP_QR_reports.ahk
			{	
				issuedata := Format_Issue_Data(identifier, volume, issue, date, questionable, questionabledisplay, numpages, missing, missingdisplay, editionlabel, sectionlabels)
				FileAppend, %issuedata%`n, %reportpath%\%lccn%-%reelnumber%-report.txt
				Sleep, 100
			}
		
			; update scoreboard
			ControlSetText, Static15, %loopcount%, NDNP_QR	
					
			; bring metadata windows forward
			WinSet, Top,, Metadata
			WinSet, Top,, Date
			WinSet, Top,, Questionable
			WinSet, Top,, Volume
			WinSet, Top,, Issue
			WinSet, Top,, Edition

			if (loopcount == 1) ; Timer screen coordinates
			{
				WinGetPos, timerX, timerY, winWidth, winHeight, Metadata
				timerX+=%winWidth%
			}
				
			Gosub, DelayTimer

			; *************************************************
			; NOTES FUNCTION
			; hold down Right Shift during page display delay
			getKeyState, state, RShift
			if state = D
			{
				if (reelreportflag == 1) ; Reel Report
				{
					WinGetPos, winX, winY, winWidth, winHeight, Metadata
					winX+=%winWidth%
					InputBox, note, %reelnumber% Report, Enter a note for %date%:,, 450, 120, %winX%, %winY%,,,
					if ErrorLevel
					{
						Gui, 7:Destroy ; Questionable Date
						Gui, 9:Destroy ; Edition Label							
						Gosub, CloseFirstImpressionWindows
						Continue ; next issue
					}

					else ; add report note
					{
						if (note != "")
						{
							FileAppend, %date%: %note%`n, %reportpath%\%lccn%-%reelnumber%-report.txt
						}
						Gui, 7:Destroy ; Questionable Date
						Gui, 9:Destroy ; Edition Label							
						Gosub, CloseFirstImpressionWindows
						Continue ; next issue
					}
				}
					
				else ; Metadata Viewer
				{
					if (notecount == 0) ; inactive notes
					{
						MsgBox, 4, Metadata Viewer, Would you like to save a note for %date% ?`n`nThe notes file for reel %reelnumber%`nwill be created if it does not already exist.
						IfMsgBox, Yes
						{
							WinGetPos, winX, winY, winWidth, winHeight, Metadata
							winX+=%winWidth%
							InputBox, note, %reelnumber% Notes, Enter a note for %date%:,, 450, 120, %winX%, %winY%,,,
							if ErrorLevel
							{
								Gui, 7:Destroy ; Questionable Date
								Gui, 9:Destroy ; Edition Label							
								Gosub, CloseFirstImpressionWindows
								Continue ; next issue
							}

							else ; add note
							{
								if (note != "")
								{
									notecount++
									FileAppend, -----------------------------------------`nNotes for Reel: %lccn%\%reelnumber%`n`nDate`t`tNote`n`n, %reportpath%\%lccn%-%reelnumber%-notes.txt									
									FileAppend, %date%`t%note%`n, %reportpath%\%lccn%-%reelnumber%-notes.txt
								}
								Gui, 7:Destroy ; Questionable Date
								Gui, 9:Destroy ; Edition Label							
								Gosub, CloseFirstImpressionWindows
								Continue ; next issue
							}
						}
					}
						
					else ; active notes
					{
						WinGetPos, winX, winY, winWidth, winHeight, Metadata
						winX+=%winWidth%
						InputBox, note, %reelnumber% Notes, Enter a note for %date%:,, 450, 120, %winX%, %winY%,,,
						if ErrorLevel
						{
							Gui, 7:Destroy ; Questionable Date
							Gui, 9:Destroy ; Edition Label							
							Gosub, CloseFirstImpressionWindows
							Continue ; next issue
						}
						if (note != "") ; add note
						{
							notecount++
							FileAppend, %date%`t%note%`n, %reportpath%\%lccn%-%reelnumber%-notes.txt
						}
						Gui, 7:Destroy ; Questionable Date
						Gui, 9:Destroy ; Edition Label							
						Gosub, CloseFirstImpressionWindows
						Continue ; next issue
					}
				}
			}
			; *************************************************
				
			; *************************************************
			; MANUAL EXIT
			; hold down Left Shift during page display delay
			getKeyState, state, LShift
			if state = D
			{
				; close metadata windows
				Gui, 4:Destroy
				Gui, 5:Destroy
				Gui, 6:Destroy
				Gui, 7:Destroy
				Gui, 9:Destroy

				Gosub, CloseFirstImpressionWindows

				if (reelreportflag == 1) ; Reel Report
				{
					FileAppend, `nReport for reel %reelnumber% was cancelled.`n`nIssues: %loopcount%`n Pages: %totalpages%, %reportpath%\%lccn%-%reelnumber%-report.txt
					FileAppend, `n`nSTART: %start%, %reportpath%\%lccn%-%reelnumber%-report.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %reportpath%\%lccn%-%reelnumber%-report.txt
						
					MsgBox, 4, Reel Report, The report for reel %reelnumber% was cancelled.`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nWould you like to open the report file?
					IfMsgBox, Yes
						Run, %reportpath%\%lccn%-%reelnumber%-report.txt
						
					reelreportflag = 0
					loopcount = 0
					return
				}
				else if (notecount > 0) ; Metadata Viewer with Notes
				{
					FileAppend, `nSTART: %start%, %reportpath%\%lccn%-%reelnumber%-notes.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %reportpath%\%lccn%-%reelnumber%-notes.txt

					MsgBox, 4, Metadata Viewer, The metadata viewer for reel %reelnumber% was cancelled.`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nYou added %notecount% notes. Would you like to open the notes file?
					IfMsgBox, Yes
						Run, %reportpath%\%lccn%-%reelnumber%-notes.txt
					loopcount = 0

					return						
				}
				else ; Metadata Viewer
				{
					MsgBox, 0, Metadata Viewer, The metadata viewer for reel %reelnumber% was cancelled.`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%						
					loopcount = 0
					return
				}
			}
			; *************************************************						

			Gui, 7:Destroy ; Questionable Date
			Gui, 9:Destroy ; Edition Label							
			Gosub, CloseFirstImpressionWindows
		}
	}
Return
; ======REEL LOOP FUNCTION

; ======DELAY TIMER
DelayDialog:
	Gui, 16:+ToolWindow
	Gui, 16:Add, Text,, Seconds:
	; delaychoice value (1-7) pre-selected, assigns delay (3-9 seconds)
	Gui, 16:Add, DropDownList, Choose%delaychoice% R7 vdelay, 3|4|5|6|7|8|9
	Gui, 16:Add, Button, w40 x10 y55 gDelayGo default, OK
	Gui, 16:Add, Button, x65 y55 gDelayCancel, Cancel
	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	Gui, 16:Show, x%winX% y%winY%, Delay Time
Return

; OK button
DelayGo:
	Gui, 16:Submit ; delay value
	delaychoice := delay - 2
	Gui, 16:Destroy
Return
		
; Cancel button
DelayCancel:
	Gui, 16:Destroy
	cancelbutton = 1
Return

DelayTimer: ; delay timer window

	seconds = %delay%
	
	Gui, 8:+AlwaysOnTop
	Gui, 8:+ToolWindow
	Gui, 8:Font, cGreen s25 bold, Arial
	Gui, 8:Add, Text, x15 y8 w30 h35, %seconds%
	Gui, 8:Font, cBlack s8 norm, Arial
	Gui, 8:Add, Button, x5 y55 w40 gReelPause default, Pause
	Gui, 8:Add, Button, x5 y85 w18 gDelayMinus, -
	Gui, 8:Add, Button, x26 y85 gDelayPlus, +
	Gui, 8:Show, x%timerX% y%timerY% h115 w50, Timer

	; hide plus or minus if max or min delay time
	if (seconds == 9)
		GuiControl, 8:Hide, +
	if (seconds == 3)
		GuiControl, 8:Hide, -
	
	Loop
	{
		; pause button loop
		if (ispaused == 1)
		{
			Loop
			{
				; check for unpause every half second
				Sleep, 500
				if (ispaused == 0)
					break
				}
		}

		; wait one second
		Sleep, 1000
		
		; pause button loop
		if (ispaused == 1)
		{
			Loop
			{
				; check for unpause every half second
				Sleep, 500
				if (ispaused == 0)
					break
			}
		}

		seconds--
		ControlSetText, Static1, %seconds%, Timer
		
		; stop timer when it reaches 0
		if (seconds == 0)
		{
			WinGetPos, timerX, timerY,,, Timer
			Gui, 8:Destroy
			Break
		}
	}
Return

; toggles Pause/Start
ReelPause:
	if (ispaused == 0)
	{
		GuiControl,, Pause, Start
		ispaused = 1
	}
	else
	{
		GuiControl,, Start, Pause
		ispaused = 0
	}
Return

; increase delay time
DelayPlus:
	delay++
	if (delay > 9) ; max 9
		delay = 9
	delaychoice := delay - 2
Return

; decrease delay time
DelayMinus:
	delay--
	if (delay < 3) ; min 3
		delay = 3
	delaychoice := delay - 2
Return
; ======DELAY TIMER
