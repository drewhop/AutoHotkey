; ********************************
; NDNP_QR_tools.ahk
; required file for NDNP_QR.ahk
; ********************************

; ======BATCH REPORT
; creates a report of all the issues in a batch
; text file will be created in the batch folder
BatchReport:
	; update last hotkey
	ControlSetText, Static3, BATCH, NDNP_QR
 
	; initialize the file variables
	issuefile =
	batchxml =

	; create dialog to select a batch folder
	FileSelectFolder, batchfolderpath, %batchdrive%, 2, BATCH REPORT`n`nSelect a BATCH folder:
	if ErrorLevel
		Return
	else
	{
		; extract batch name from path
		StringGetPos, foldernamepos, batchfolderpath, \, R1
		foldernamepos++
		StringTrimLeft, batchname, batchfolderpath, foldernamepos

		; create notification window
		Gui, 15:+ToolWindow
		Gui, 15:Add, Text,, Processing:  %batchname%
		Gui, 15:Add, Text,, This may take awhile . . .

		; position in upper left corner of the NDNP_QR window
		SetTitleMatchMode 1
		WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
		Gui, 15:Show, x%winX% y%winY%, Batch Report
	
		; store the start time
		start = %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% 
			
		; create the report file in the batch folder
		FileAppend, --------------------------------------------------------------`n%batchfolderpath%`n`nIdentifier`t`tPages`tDate`t`tVolume`tIssue`n`n, %batchfolderpath%\%batchname%-report.txt

		; read in the batch.xml file to the batchxml variable
		FileRead, batchxml, %batchfolderpath%\batch.xml

		; initialize the counters
		issuecount = 0
		totalpages = 0
		

		; loop to parse the batchxml and store issue paths
		Loop, parse, batchxml, `n, `r%A_Space%%A_Tab%
		{
			; if the line contains an <issue> tag
			IfInString, A_LoopField, issue
			{
				; trim the <issue> line before the file path
				StringTrimLeft, rawissueXMLpath, A_LoopField, 66
			
				; remove the </issue> tag and store the issue.xml path
				StringTrimRight, issueXMLpath, rawissueXMLpath, 8

				; append issue.xml path to issuefile
				issuefile .= issueXMLpath
						
				; append new line to issuefile
				issuefile .= "`n"
					
				; update the issue count
				issuecount++
			}
		}
		
		; sort the issuefile variable
		Sort, issuefile

		; *******************************
		; loop through sorted issuefile
		; *******************************
		Loop, parse, issuefile, `n
		{
			; *************************************************
			; AUTO EXIT if issuefile is empty
			if A_LoopField =
			{
				; add the issue count to the report
				FileAppend, `nBatch: %batchname%`nIssues: %issuecount%`n Pages: %totalpages%, %batchfolderpath%\%batchname%-report.txt

				; print the start and end times
				FileAppend, `n`nSTART: %start%, %batchfolderpath%\%batchname%-report.txt
				FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %batchfolderpath%\%batchname%-report.txt

				; close the notification window
				Gui, 15:Destroy
				
				; create a message box to indicate that the script ended
				MsgBox, 4, Batch Report, Batch: %batchname%`n`nIssues: %issuecount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nThe report is complete. Would you like to open it?
				IfMsgBox, Yes
					; open the report if Yes
					Run, %batchfolderpath%\%batchname%-report.txt
				
				; exit the script
				Return
			}
			; *************************************************

			; remove the issue.xml file name and store the partial issue folder path
			StringTrimRight, issuepath, A_LoopField, 15

			; remove the dates and store the LCCN and Reel #
			StringTrimRight, identifier, A_LoopField, 26
			Sleep, 100	

			; create issue folder path variable for page count
			issuefolderpath = %batchfolderpath%\%issuepath%

			; initialize the issuexml variable
			issuexml =
			
			; initialize the metadata values
			volume =
			issue =
			date =
			editionlabel =
			questionable =
			questionabledisplay =
			
			; initialize the metadata flags
			volumeflag = 0
			issueflag = 0

			; read in the issue.xml file to the issuexml variable
			FileRead, issuexml, %batchfolderpath%\%A_LoopField%

			; ************************
			; extract issue metadata
			; ************************
			Loop, parse, issuexml, `n, `r%A_Space%%A_Tab%
			{
				; if the line contains a volume number tag
				; set the volume number flag
				IfInString, A_LoopField, type="volume"
				{
					; set the volume number flag
					volumeflag = 1
					
					; continue the parsing loop
					Continue
				}

				; if the volumeflag has been set
				if (volumeflag == 1)
				{
					; assign the volume number
					StringTrimLeft, volume, A_LoopField, 13
					StringTrimRight, volume, volume, 14
					
					; reset the volumeflag
					volumeflag = 0
				}

				; if the line contains an issue number tag
				IfInString, A_LoopField, type="issue"
				{
					; set the issue number flag
					issueflag = 1
					
					; continue the parsing loop
					Continue
				}

				; if the issueflag has been set
				if (issueflag == 1)
				{
					; assign the issue number
					StringTrimLeft, issue, A_LoopField, 13
					StringTrimRight, issue, issue, 14
					
					; reset the issueflag
					issueflag = 0
				}

				; if the line contains an edition label
				IfInString, A_LoopField, mods:caption
				{
					; assign the edition label
					StringTrimLeft, editionlabel, A_LoopField, 14
					StringTrimRight, editionlabel, editionlabel, 15
				}

				; if the line contains an issue date
				IfInString, A_LoopField, <mods:dateIssued encoding="iso8601">
				{
					; assign the issue date
					StringMid, date, A_LoopField, 37, 10
				}

				; if the line contains a questionable issue date
				IfInString, A_LoopField, qualifier="questionable"
				{
					; assign the report display variable
					StringMid, questionable, A_LoopField, 62, 10
					questionabledisplay = Questionable: %questionable%
				}
			}

			; initialize the number of pages variable
			numpages = 0
							
			; count the number of .TIF files
			Loop, %issuefolderpath%\*.tif
			{
				numpages++
			}

			; add the page count to the total
			totalpages += numpages					

			; add the issue data to the report
			FileAppend, %identifier%`t%numpages%`t%date%`t%volume%`t%issue%`t%editionlabel%`t%questionabledisplay%`n, %batchfolderpath%\%batchname%-report.txt
		}
	}
Return
; ======BATCH REPORT

; ======REEL REPORT
; displays first TIFF and metadata
; and creates a report for the issues in a reel folder
; text file will be created in the LCCN folder
ReelReport:
	metaloopname = Reel Report
	
	; update last hotkey & loop label
	ControlSetText, Static3, REEL, NDNP_QR
	ControlSetText, Static4, Reel Report, NDNP_QR

	; set the reel report flag
	reelreportflag = 1

	; enter the loop function
	Gosub, ReelLoopFunction
Return
; ======REEL REPORT

; ======METADATA VIEWER
; opens the first TIFF and displays the issue metadata
; for each issue in a reel folder
MetaViewer:
	metaloopname = Metadata Viewer
	
	; update last hotkey & loop label
	ControlSetText, Static3, VIEWER, NDNP_QR
	ControlSetText, Static4, Meta Viewer, NDNP_QR

	; enter the loop function
	Gosub, ReelLoopFunction
Return
; ======METADATA VIEWER

; ======REEL LOOP FUNCTION
; Reel Report & Metadata Viewer
ReelLoopFunction:
	; initialize the file variables
	issuefile =
	batchxml =

	; initialize the counters
	loopcount = 0
	totalpages = 0
	notecount = 0

	; create dialog to select a reel folder
	FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
	if ErrorLevel
		Return
	else
	{
		; loop to check for a valid reel folder
		Loop
		{
			; create display variables for reel folder path
			StringGetPos, reelfolderpos, reelfolderpath, \, R2
			StringLeft, reelfolder1, reelfolderpath, reelfolderpos
			StringTrimLeft, reelfolder2, reelfolderpath, reelfolderpos

			; assign the new reel folder name
			StringGetPos, foldernamepos, reelfolderpath, \, R1
			foldernamepos++
			StringTrimLeft, reelfoldername, reelfolderpath, foldernamepos
			
			; check to see if it is a reel folder
			StringLen, length, reelfoldername
			if (length != 11)
			{		
				; print error message
				MsgBox, 0, Reel Loop, %reelfoldername% does not appear to be a REEL folder.`n`n`tClick OK to continue.

				; reset the variables
				reelfolderpath = _
				reelfoldername = _

				; create dialog to select a reel folder
				FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
				if ErrorLevel
					Return
				else Continue
			}
				
			else Break
		}
		
		; create delay time dialog
		Gosub, DelayDialog
		
		; wait until Delay Time window is closed
		SetTitleMatchMode 1
		WinWaitActive, Delay Time
		WinWaitClose, Delay Time

		if (cancelbutton == 1)
		{
			; reset cancelbutton
			cancelbutton = 0
			
			; exit the script
			Return
		}

		; close any First Impression windows
		Gosub, CloseFirstImpressionWindows			
			
		; store the start time
		start = %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% 

		; create a variable for the report path (LCCN folder)
		StringGetPos, foldernamepos, reelfolderpath, \, R1
		StringLeft, reportpath, reelfolderpath, foldernamepos

		; extract reel number from path
		foldernamepos++
		StringTrimLeft, reelnumber, reelfolderpath, foldernamepos

		; create report divider
		divider =
		StringLen, length, reelfolderpath
		length++
		Loop, %length%
		{
			divider .= "-"
		}
		
		if (reelreportflag == 1)
		{
			; create the report file
			FileAppend, %divider%`n%reelfolderpath%`n`nPages`tDate`t`tVolume`tIssue`n`n, %reportpath%\%reelnumber%-report.txt
		}
			
		; create a variable for the batch folder path
		StringGetPos, batchpathpos, reelfolderpath, \, R2
		StringLeft, batchfolderpath, reelfolderpath, batchpathpos
			
		; read in the batch.xml file to the batchxml variable
		FileRead, batchxml, %batchfolderpath%\batch.xml

		; loop to parse the batchxml and store issue paths for the reel
		Loop, parse, batchxml, `n, `r%A_Space%%A_Tab%
		{
			; if the line contains the reel number
			IfInString, A_LoopField, %reelnumber%
			{
				; trim the <issue> line before the file path
				StringTrimLeft, rawissueXMLpath, A_LoopField, 66
				
				; remove the </issue> tag and store the issue.xml path
				StringTrimRight, issueXMLpath, rawissueXMLpath, 8

				; check to see that it was an <issue> line
				StringLen, length, issueXMLpath
				if (length > 20)
				{
					; append issue.xml path to issuefile
					issuefile .= issueXMLpath
					
					; append new line to issuefile
					issuefile .= "`n"
							
					; update the issue count
					issuecount++
				}
			}
		}
			
		; sort the issuefile variable
		Sort, issuefile

		; *******************************
		; loop through sorted issuefile
		; *******************************
		Loop, parse, issuefile, `n
		{
			; *************************************************
			; AUTO EXIT if issuefile is empty
			if A_LoopField =
			{
				; close the metadata windows
				Gui, 4:Destroy
				Gui, 5:Destroy
				Gui, 6:Destroy
				Gui, 7:Destroy
				Gui, 9:Destroy

				; clear the Metadata window
				ControlSetText, Static6,, Metadata
				ControlSetText, Static7,, Metadata		
				ControlSetText, Static8,, Metadata		
				ControlSetText, Static9,, Metadata
				ControlSetText, Static10,, Metadata
				
				; Reel Report
				if (reelreportflag == 1)
				{
					; add the issue count to the report
					FileAppend, `nIssues: %loopcount%`n Pages: %totalpages%, %reportpath%\%reelnumber%-report.txt

					; print the start and end times
					FileAppend, `n`nSTART: %start%, %reportpath%\%reelnumber%-report.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %reportpath%\%reelnumber%-report.txt

					; create a message box to indicate that the script ended
					MsgBox, 4, Reel Report, Reel: %reelnumber%`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nThe report is complete. Would you like to open it?
					IfMsgBox, Yes
						; open the report if Yes
						Run, %reportpath%\%reelnumber%-report.txt

					; reset the report flag
					reelreportflag = 0
					
					;exit the script
					return
				}
					
				; Metadata Viewer with Notes
				else if (notecount > 0)
				{
					; print the start and end times
					FileAppend, `nSTART: %start%, %reportpath%\%reelnumber%-notes.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %reportpath%\%reelnumber%-notes.txt

					; create a message box to indicate that the script ended
					MsgBox, 4, Metadata Viewer, Reel: %reelnumber%`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nYou added %notecount% notes. Would you like to open the Notes file?
					IfMsgBox, Yes
						; open the notes file if Yes
						Run, %reportpath%\%reelnumber%-notes.txt

					;exit the script
					return						
				}
					
				; Metadata Viewer
				else
				{
					; create a message box to indicate that the script ended
					MsgBox, 0, Metadata Viewer, Reel: %reelnumber%`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%

					;exit the script
					return
				}
			}
			; *************************************************

			; remove the issue.xml file name and store the partial issue folder path
			StringTrimRight, issuepath, A_LoopField, 15

			; get the issue folder name
			StringRight, issuefoldername, issuepath, 10

			; create issue folder path variable
			issuefolderpath = %batchfolderpath%\%issuepath%

			; find and open the first TIF file
			Loop, %issuefolderpath%\*.tif
			{
				Run, %issuefolderpath%\%A_LoopFileName%
				Break
			}
				
			; get the First Impression window id#
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
			Send, {NumpadSub 20}

			; send First Impression to bottom of stack
			WinSet, Bottom,, ahk_id %firstid%
			Sleep, 300
				
			; increment the counter
			loopcount++

			; extract and display the issue metadata
			Gosub, ExtractMeta

			; update total pages count
			totalpages+=%numpages%

			; if the report flag has been set
			if (reelreportflag == 1)
			{
				; add the issue data to the report
				FileAppend, %numpages%`t%date%`t%volume%`t%issue%`t%editionlabel%`t%questionabledisplay%`n, %reportpath%\%reelnumber%-report.txt
				Sleep, 100
			}
		
			; update the scoreboard
			ControlSetText, Static15, %loopcount%, NDNP_QR	
					
			; bring all metadata windows to front
			WinSet, Top,, Metadata
			WinSet, Top,, Date
			WinSet, Top,, Questionable
			WinSet, Top,, Volume
			WinSet, Top,, Issue
			WinSet, Top,, Edition

			; get screen coordinates for Timer if first run
			if (loopcount == 1)
			{
				WinGetPos, timerX, timerY, winWidth, winHeight, Metadata
				timerX+=%winWidth%
			}
				
			; create Delay Timer
			Gosub, DelayTimer

			; *************************************************
			; NOTES FUNCTION
			; hold down Right Shift during page display delay
			getKeyState, state, RShift
			if state = D
			{
				; Reel Report
				if (reelreportflag == 1)
				{
					WinGetPos, winX, winY, winWidth, winHeight, Metadata
					winX+=%winWidth%
					InputBox, note, %reelnumber% Report, Enter a note for %date%:,, 450, 120, %winX%, %winY%,,,
					if ErrorLevel
					{
						; close the Questionable Date & Edition Label windows
						Gui, 7:Destroy
						Gui, 9:Destroy
							
						; close any First Impression windows
						Gosub, CloseFirstImpressionWindows

						; next issuefile loop iteration
						Continue
					}

					; add to report and continue loop
					else
					{
						; check if the note is blank
						if (note != "")
						{
							; add the note to the report
							FileAppend, %date%: %note%`n, %reportpath%\%reelnumber%-report.txt
						}
						
						; close the Questionable Date & Edition Label windows
						Gui, 7:Destroy
						Gui, 9:Destroy
							
						; close any First Impression windows
						Gosub, CloseFirstImpressionWindows

						; next issuefile loop iteration
						Continue
					}
				}
					
				; Metadata Viewer
				else
				{
					; case for inactive notes function
					if (notecount == 0)
					{
						MsgBox, 4, Metadata Viewer, Would you like to save a note for %date% ?`n`nThe notes file for reel %reelnumber%`nwill be created if it does not already exist.
						IfMsgBox, Yes
						{
							; activate notes function
							WinGetPos, winX, winY, winWidth, winHeight, Metadata
							winX+=%winWidth%
							InputBox, note, %reelnumber% Notes, Enter a note for %date%:,, 450, 120, %winX%, %winY%,,,
							if ErrorLevel
							{
								; close the Questionable Date & Edition Label windows
								Gui, 7:Destroy
								Gui, 9:Destroy
									
								; close any First Impression windows
								Gosub, CloseFirstImpressionWindows

								; next issuefile loop iteration
								Continue
							}

							; add note to notes file
							else
							{
								; check if the note is blank
								if (note != "")
								{
									; increment the notes counter
									notecount++
									
									; format the notes file
									FileAppend, ----------------------------`nNotes for Reel: %reelnumber%`n`nDate`t`tNote`n`n, %reportpath%\%reelnumber%-notes.txt									
										
									; add the note to the notes file
									FileAppend, %date%`t%note%`n, %reportpath%\%reelnumber%-notes.txt
								}
								
								; close the Questionable Date & Edition Label windows
								Gui, 7:Destroy
								Gui, 9:Destroy
									
								; close any First Impression windows
								Gosub, CloseFirstImpressionWindows

								; next issuefile loop iteration
								Continue
							}
						}
					}
						
					; add note to notes file
					else
					{
						WinGetPos, winX, winY, winWidth, winHeight, Metadata
						winX+=%winWidth%
						InputBox, note, %reelnumber% Notes, Enter a note for %date%:,, 450, 120, %winX%, %winY%,,,
						if ErrorLevel
						{
							; close the Questionable Date & Edition Label windows
							Gui, 7:Destroy
							Gui, 9:Destroy
									
							; close any First Impression windows
							Gosub, CloseFirstImpressionWindows

							; next issuefile loop iteration
							Continue
						}

						; check if the note is blank
						if (note != "")
						{
							; increment the notes counter
							notecount++
									
							; add the note to the report
							FileAppend, %date%`t%note%`n, %reportpath%\%reelnumber%-notes.txt
						}

						; close the Questionable Date & Edition Label windows
						Gui, 7:Destroy
						Gui, 9:Destroy
									
						; close any First Impression windows
						Gosub, CloseFirstImpressionWindows

						; next issuefile loop iteration
						Continue
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
				; close the metadata windows
				Gui, 4:Destroy
				Gui, 5:Destroy
				Gui, 6:Destroy
				Gui, 7:Destroy
				Gui, 9:Destroy

				; close any First Impression windows
				Gosub, CloseFirstImpressionWindows

				; if the report flag has been set
				if (reelreportflag == 1)
				{
					; add the issue count to the report
					FileAppend, `nReport for reel %reelnumber% was cancelled.`n`nIssues: %loopcount%`n Pages: %totalpages%, %reportpath%\%reelnumber%-report.txt

					; print the start and end times
					FileAppend, `n`nSTART: %start%, %reportpath%\%reelnumber%-report.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %reportpath%\%reelnumber%-report.txt
						
					; create a message box to indicate the script was cancelled
					MsgBox, 4, Reel Report, The report for reel %reelnumber% was cancelled.`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nWould you like to open the report file?
					; open the report if Yes
					IfMsgBox, Yes
						; open the report if Yes
						Run, %reportpath%\%reelnumber%-report.txt
						
					; reset the report flag
					reelreportflag = 0

					;exit the loop
					return
				}
					
				; if notes have been added during this run
				else if (notecount > 0)
				{
					; print the start and end times
					FileAppend, `nSTART: %start%, %reportpath%\%reelnumber%-notes.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %reportpath%\%reelnumber%-notes.txt

					; create a message box to indicate that the script was cancelled
					MsgBox, 4, Metadata Viewer, The metadata viewer for reel %reelnumber% was cancelled.`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nYou added %notecount% notes. Would you like to open the notes file?
					IfMsgBox, Yes
						; open the notes file if Yes
						Run, %reportpath%\%reelnumber%-notes.txt

					;exit the script
					return						
				}
					
				else
				{
					; create a message box to indicate the script was cancelled
					MsgBox, 0, Metadata Viewer, The metadata viewer for reel %reelnumber% was cancelled.`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%
						
					;exit the loop
					return
				}
			}
			; *************************************************						

			; close the Questionable Date & Edition Label windows
			Gui, 7:Destroy
			Gui, 9:Destroy
				
			; close any First Impression windows
			Gosub, CloseFirstImpressionWindows
		}
	}
Return
; ======REEL LOOP FUNCTION

; ======DELAY TIMER
; Delay Time dialog
DelayDialog:
	Gui, 16:+ToolWindow
	Gui, 16:Add, Text,, Seconds:
	; delaychoice value (1-7) pre-selected, assigns delay (3-9 seconds)
	Gui, 16:Add, DropDownList, Choose%delaychoice% R7 vdelay, 3|4|5|6|7|8|9
	Gui, 16:Add, Button, w40 x10 y55 gDelayGo default, OK
	Gui, 16:Add, Button, x65 y55 gDelayCancel, Cancel
	
	; position below the NDNP_QR window
	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	winY+=%winHeight%
	Gui, 16:Show, x%winX% y%winY%, Delay Time
Return

; OK button function for delay assignment
DelayGo:
	; assign the delay variable value
	Gui, 16:Submit
		
	; assign the delaychoice variable value
	delaychoice := delay - 2
	
	; close the Folder Navigation GUI
	Gui, 16:Destroy
Return
		
; Cancel button function for delay assignment
DelayCancel:
	; close the Delay Time GUI
	Gui, 16:Destroy
		
	; set the cancelbutton
	cancelbutton = 1
Return

; delay timer window
DelayTimer:
	; assign the number of seconds
	seconds = %delay%
	
	; create the timer GUI
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
	
	; start the timer
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

		; decrement the seconds
		seconds--
		
		; update the display
		ControlSetText, Static1, %seconds%, Timer
		
		; stop timer when it reaches 0
		if (seconds == 0)
		{
			; get window position for next run
			WinGetPos, timerX, timerY,,, Timer
			; close the window
			Gui, 8:Destroy
			Break
		}
	}
Return

; pause button function toggles Pause/Start
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

; increase the delay time
DelayPlus:
	; increment the delay value
	delay++
	
	; ensure maximum of 9
	if (delay > 9)
		delay = 9
		
	; set the delaychoice value
	delaychoice := delay - 2
Return

; decrease the delay time
DelayMinus:
	; decrement the delay value
	delay--
	
	; ensure minimum of 3
	if (delay < 3)
		delay = 3
		
	; set the delaychoice value
	delaychoice := delay - 2
Return
; ======DELAY TIMER

; ======LANGUAGE CODE REPORT
EditLanguageCode:
	; determine the language setting
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
	
	; position below the NDNP_QR window
	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	winY+=%winHeight%
	Gui, 18:Show, x%winX% y%winY%, Language Code
Return

; OK button function
LanguageGo:
	; assign the language code
	Gui, 18:Submit
			
	; close the GUI
	Gui, 18:Destroy
Return
		
; Cancel button function
LanguageCancel:
	; close the GUI
	Gui, 18:Destroy
		
	; set the cancelbutton
	cancelbutton = 1
Return

LanguageCodeReport:
	; update last hotkey
	ControlSetText, Static3, LANGUAGE, NDNP_QR
 
	Gosub, EditLanguageCode
	
	; wait until Language Code window is closed
	SetTitleMatchMode 1
	WinWaitActive, Language Code
	WinWaitClose, Language Code

	if (cancelbutton == 1)
	{
		; reset cancelbutton
		cancelbutton = 0
			
		; exit the script
		Return
	}
	
	; create dialog to select a reel folder
	FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
	if ErrorLevel
		Return
	else
	{
		; loop to check for a valid reel folder
		Loop
		{
			; create display variables for reel folder path
			StringGetPos, reelfolderpos, reelfolderpath, \, R2
			StringLeft, reelfolder1, reelfolderpath, reelfolderpos
			StringTrimLeft, reelfolder2, reelfolderpath, reelfolderpos

			; assign the new reel folder name
			StringGetPos, foldernamepos, reelfolderpath, \, R1
			foldernamepos++
			StringTrimLeft, reelfoldername, reelfolderpath, foldernamepos
			
			; check to see if it is a reel folder
			StringLen, length, reelfoldername
			if (length != 11)
			{		
				; print error message
				MsgBox, 0, Language Code Report, %reelfoldername% does not appear to be a REEL folder.`n`n`tClick OK to continue.

				; reset the variables
				reelfolderpath = _
				reelfoldername = _

				; create dialog to select a reel folder
				FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
				if ErrorLevel
					Return
				else Continue
			}
				
			else Break
		}
	
		; create notification window
		Gui, 19:+ToolWindow
		Gui, 19:Add, Text,, Processing:  %reelfoldername%-%languagecode%-report.txt
		Gui, 19:Add, Text,, Please wait . . .

		; position in upper left corner of the NDNP_QR window
		SetTitleMatchMode 1
		WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
		Gui, 19:Show, x%winX% y%winY%, Language Code Report
	
		; store the start time
		start = %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% 

		; create a variable for the report path (LCCN folder)
		StringGetPos, foldernamepos, reelfolderpath, \, R1
		StringLeft, reportpath, reelfolderpath, foldernamepos

		; extract reel number from path
		foldernamepos++
		StringTrimLeft, reelnumber, reelfolderpath, foldernamepos

		; create report divider
		divider =
		StringLen, length, reelfolderpath
		length++
		Loop, %length%
		{
			divider .= "-"
		}

		; create the report file
		FileAppend, %divider%`n%reelfolderpath%`n`n, %reportpath%\%reelnumber%-%languagecode%-report.txt
			
		; create a variable for the batch folder path
		StringGetPos, batchpathpos, reelfolderpath, \, R2
		StringLeft, batchfolderpath, reelfolderpath, batchpathpos
			
		; read in the batch.xml file to the batchxml variable
		FileRead, batchxml, %batchfolderpath%\batch.xml

		; loop to parse the batchxml and store issue paths for the reel
		Loop, parse, batchxml, `n, `r%A_Space%%A_Tab%
		{
			; if the line contains the reel number
			IfInString, A_LoopField, %reelnumber%
			{
				; trim the <issue> line before the file path
				StringTrimLeft, rawissuefolderpath, A_LoopField, 66
				
				; remove the </issue> tag and issue.xml file name
				StringTrimRight, issuefolderpath, rawissuefolderpath, 23

				; check to see that it was an <issue> line
				StringLen, length, issuefolderpath
				if (length > 20)
				{
					; append issue folder path to issuefile
					issuefile .= issuefolderpath
					
					; append new line to issuefile
					issuefile .= "`n"
				}
			}
		}
			
		; sort the issuefile variable
		Sort, issuefile

		; initialize the total codes counters
		totalcodes = 0
		issuecodes = 0
		
		; *******************************
		; loop through sorted issuefile
		; *******************************
		Loop, parse, issuefile, `n
		{
			; *************************************************
			; AUTO EXIT if issuefile is empty
			if A_LoopField =
			{
				; add the issue count to the report
				FileAppend, `n___________________________`n`n, %reportpath%\%reelnumber%-%languagecode%-report.txt
				FileAppend, Reel: %reelnumber%`nCode: %languagecode%`n`nIssues: %issuecount%`nBlocks: %totalcodes%, %reportpath%\%reelnumber%-%languagecode%-report.txt

				; print the start and end times
				FileAppend, `n`nSTART: %start%, %reportpath%\%reelnumber%-%languagecode%-report.txt
				FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %reportpath%\%reelnumber%-%languagecode%-report.txt

				; close the notification window
				Gui, 19:Destroy
				
				; create a message box to indicate that the script ended
				MsgBox, 4, Language Code Report, Reel: %reelnumber%`nCode: %languagecode%`n`nIssues: %issuecount%`nCodes: %totalcodes%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nThe report is complete. Would you like to open it?
				IfMsgBox, Yes
					; open the report if Yes
					Run, %reportpath%\%reelnumber%-%languagecode%-report.txt

				;exit the script
				return
			}
			; *************************************************

			; get the issue date
			StringRight, issuedate, A_LoopField, 10
			Sleep, 100	

			; create issue folder path variable for page count
			issuefolderpath = %batchfolderpath%\%A_LoopField%

			; initialize the ALTOxml and issueresults variable and page code counters
			ALTOxml =
			issueresults =
			codecount = 0
			issuecodes = 0
			
			; loop to find and parse the ALTO XML files
			Loop, %issuefolderpath%\*.xml
			{
				; initialize the code counter
				codecount = 0
				
				; check to see that the file is an ALTOxml
				StringLen, length, A_LoopFileName
				if (length < 11)
				{
					; read in the file to the ALTOxml variable
					FileRead, ALTOxml, %issuefolderpath%\%A_LoopFileName%
					
					; loop to parse the textblocks
					Loop, parse, ALTOxml, >, %A_Space%%A_Tab%
					{
						; if the substring contains the language code
						IfInString, A_LoopField, language="%languagecode%"
						{
							codecount++
						}
					}				

					; add line to issueresults if code found
					if (codecount > 0)
					{
						issueresults .= A_LoopFileName
						issueresults .= "`t"
						issueresults .= codecount
						issueresults .= "`n"
					}

					; add the code count to the totals
					issuecodes += codecount
					totalcodes += codecount
				}
			}

			; create issue entry in report if results found
			if (issuecodes > 0)
			{
				; add the issue date to the report
				FileAppend, ____________________`n`n, %reportpath%\%reelnumber%-%languagecode%-report.txt			
				FileAppend, %issuedate%`n`n, %reportpath%\%reelnumber%-%languagecode%-report.txt			

				; add the results to the report
				FileAppend, %issueresults%, %reportpath%\%reelnumber%-%languagecode%-report.txt
				
				; add the issue total to the report
				FileAppend, `n`t Total: %issuecodes%`n, %reportpath%\%reelnumber%-%languagecode%-report.txt				
							
				; update the issue count
				issuecount++
			}
		}
	}
Return
; ======LANGUAGE CODE REPORT

; ======OCR SEARCH
OCRSearch:
	; update last hotkey
	ControlSetText, Static3, OCR, NDNP_QR
 
	WinGetPos, winX, winY, winWidth, winHeight, Metadata
	winX+=%winWidth%
	InputBox, ocrterm, OCR Search, Enter a search term:,, 250, 120, %winX%, %winY%,,, %ocrterm%
	if ErrorLevel
		Return
	else
	{
		; create dialog to select a reel folder
		FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
		if ErrorLevel
			Return
		else
		{
			; loop to check for a valid reel folder
			Loop
			{
				; create display variables for reel folder path
				StringGetPos, reelfolderpos, reelfolderpath, \, R2
				StringLeft, reelfolder1, reelfolderpath, reelfolderpos
				StringTrimLeft, reelfolder2, reelfolderpath, reelfolderpos

				; assign the new reel folder name
				StringGetPos, foldernamepos, reelfolderpath, \, R1
				foldernamepos++
				StringTrimLeft, reelfoldername, reelfolderpath, foldernamepos
				
				; check to see if it is a reel folder
				StringLen, length, reelfoldername
				if (length != 11)
				{		
					; print error message
					MsgBox, 0, Language Code Report, %reelfoldername% does not appear to be a REEL folder.`n`n`tClick OK to continue.

					; reset the variables
					reelfolderpath = _
					reelfoldername = _

					; create dialog to select a reel folder
					FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
					if ErrorLevel
						Return
					else Continue
				}
					
				else Break
			}
		
			; create notification window
			Gui, 20:+ToolWindow
			Gui, 20:Add, Text,, Processing:  %reelfoldername%-%ocrterm%-report.txt
			Gui, 20:Add, Text,, Please wait . . .

			; position in upper left corner of the NDNP_QR window
			SetTitleMatchMode 1
			WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
			Gui, 20:Show, x%winX% y%winY%, OCR Search
		
			; store the start time
			start = %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% 

			; create a variable for the report path (LCCN folder)
			StringGetPos, foldernamepos, reelfolderpath, \, R1
			StringLeft, reportpath, reelfolderpath, foldernamepos

			; extract reel number from path
			foldernamepos++
			StringTrimLeft, reelnumber, reelfolderpath, foldernamepos

			; create report divider
			divider =
			StringLen, length, reelfolderpath
			length++
			Loop, %length%
			{
				divider .= "-"
			}

			; create the report file
			FileAppend, %divider%`n%reelfolderpath%`n`n, %reportpath%\%reelnumber%-%ocrterm%-report.txt
				
			; initialize the issuefile variable
			issuefile =
			
			; create a variable for the batch folder path
			StringGetPos, batchpathpos, reelfolderpath, \, R2
			StringLeft, batchfolderpath, reelfolderpath, batchpathpos
				
			; read in the batch.xml file to the batchxml variable
			FileRead, batchxml, %batchfolderpath%\batch.xml

			; loop to parse the batchxml and store issue paths for the reel
			Loop, parse, batchxml, `n, `r%A_Space%%A_Tab%
			{
				; if the line contains the reel number
				IfInString, A_LoopField, %reelnumber%
				{
					; trim the <issue> line before the file path
					StringTrimLeft, rawissuefolderpath, A_LoopField, 66
					
					; remove the </issue> tag and issue.xml file name
					StringTrimRight, issuefolderpath, rawissuefolderpath, 23

					; check to see that it was an <issue> line
					StringLen, length, issuefolderpath
					if (length > 20)
					{
						; append issue folder path to issuefile
						issuefile .= issuefolderpath
						
						; append new line to issuefile
						issuefile .= "`n"
					}
				}
			}
				
			; sort the issuefile variable
			Sort, issuefile

			; initialize the total codes counters
			totalterms = 0
			issueterms = 0
			
			; *******************************
			; loop through sorted issuefile
			; *******************************
			Loop, parse, issuefile, `n
			{
				; *************************************************
				; AUTO EXIT if issuefile is empty
				if A_LoopField =
				{
					; add the issue count to the report
					FileAppend, `n___________________________`n`n, %reportpath%\%reelnumber%-%ocrterm%-report.txt
					FileAppend, Reel: %reelnumber%`nTerm: %ocrterm%`n`nIssues: %issuecount%`n  Hits: %totalterms%, %reportpath%\%reelnumber%-%ocrterm%-report.txt

					; print the start and end times
					FileAppend, `n`nSTART: %start%, %reportpath%\%reelnumber%-%ocrterm%-report.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`n, %reportpath%\%reelnumber%-%ocrterm%-report.txt

					; close the notification window
					Gui, 20:Destroy
					
					; create a message box to indicate that the script ended
					MsgBox, 4, OCR Search, Reel: %reelnumber%`Term: %ocrterm%`n`nIssues: %issuecount%`nHits: %totalterms%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%`n`nThe report is complete. Would you like to open it?
					IfMsgBox, Yes
						; open the report if Yes
						Run, %reportpath%\%reelnumber%-%ocrterm%-report.txt

					;exit the script
					return
				}
				; *************************************************

				; get the issue date
				StringRight, issuedate, A_LoopField, 10
				Sleep, 100	

				; create issue folder path variable for page count
				issuefolderpath = %batchfolderpath%\%A_LoopField%

				; initialize the ALTOxml and issueresults variable and page code counters
				ALTOxml =
				issueresults =
				termcount = 0
				issueterms = 0
				
				; loop to find and parse the ALTO XML files
				Loop, %issuefolderpath%\*.xml
				{
					; initialize the code counter
					termcount = 0
					
					; check to see that the file is an ALTOxml
					StringLen, length, A_LoopFileName
					if (length < 11)
					{
						; read in the file to the ALTOxml variable
						FileRead, ALTOxml, %issuefolderpath%\%A_LoopFileName%
						
						; loop to parse the textblocks
						Loop, parse, ALTOxml, >, %A_Space%%A_Tab%
						{
							; if the substring contains the language code
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

						; add the code count to the totals
						issueterms += termcount
						totalterms += termcount
					}
				}

				; create issue entry in report if results found
				if (issueterms > 0)
				{
					; add the issue date to the report
					FileAppend, ____________________`n`n, %reportpath%\%reelnumber%-%ocrterm%-report.txt			
					FileAppend, %issuedate%`n`n, %reportpath%\%reelnumber%-%ocrterm%-report.txt			

					; add the results to the report
					FileAppend, %issueresults%, %reportpath%\%reelnumber%-%ocrterm%-report.txt
					
					; add the issue total to the report
					FileAppend, `n`t Total: %issueterms%`n, %reportpath%\%reelnumber%-%ocrterm%-report.txt				
								
					; update the issue count
					issuecount++
					
					; reset the issueterms variable
					issueterms = 0
				}
			}
		}
	}
Return
; ======OCR SEARCH
