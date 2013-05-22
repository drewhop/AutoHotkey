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
		Gui, 15:Show,, Batch Report
	
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
				FileAppend, `nBatch: %batchname%`n`nIssues: %issuecount%`n Pages: %totalpages%, %batchfolderpath%\%batchname%-report.txt

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
	; update last hotkey & loop label
	ControlSetText, Static3, VIEWER, NDNP_QR
	ControlSetText, Static4, Meta Viewer, NDNP_QR

	; enter the loop function
	Gosub, ReelLoopFunction
Return
; ======METADATA VIEWER

; ======DELAY TIMER
; for Reel Loop Function

; creates Delay Time dialog
DelayDialog:
	; create the Delay assignment dialog
	Gui, 16:+ToolWindow
	Gui, 16:Add, Text,, Seconds:
	; delaychoice value (1-7) pre-selected, assigns delay (3-9 seconds)
	Gui, 16:Add, DropDownList, Choose%delaychoice% R7 vdelay, 3|4|5|6|7|8|9
	Gui, 16:Add, Button, w40 x10 y55 gDelayGo default, OK
	Gui, 16:Add, Button, x65 y55 gDelayCancel, Cancel
	Gui, 16:Show,, Delay Time
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
	WinGetPos, winX, winY, winWidth, winHeight, Metadata
	winX+=%winWidth%
	Gui, 8:+AlwaysOnTop
	Gui, 8:+ToolWindow
	Gui, 8:Font, cGreen s25 bold, Arial
	Gui, 8:Add, GroupBox, x0 y-18 w50 h74,       
	Gui, 8:Add, Text, x15 y8 w30 h35, %seconds%
	Gui, 8:Show, x%winX% y%winY% h55 w50, Timer

	; start the timer
	Loop
	{
		; wait one second
		Sleep, 1000
		
		; decrement the seconds
		seconds--
		
		; update the display
		ControlSetText, Static1, %seconds%, Timer
		
		; stop timer when it reaches 0
		if (seconds == 0)
		{
			Gui, 8:Destroy
			Break
		}
	}
Return
; ======DELAY TIMER

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

		if (reelreportflag == 1)
		{
			; create the report file
			FileAppend, --------------------------------------`n%reelfolderpath%`n`nPages`tDate`t`tVolume`tIssue`n`n, %reportpath%\%reelnumber%-report.txt
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
				ControlSetText, Static5,, Metadata
				ControlSetText, Static6,, Metadata
				ControlSetText, Static7,, Metadata		
				ControlSetText, Static8,, Metadata		
				ControlSetText, Static9,, Metadata
				
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
						; add the note to the report
						FileAppend, `tNote: %note%`n, %reportpath%\%reelnumber%-report.txt

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
								; increment the notes counter
								notecount++
								
								; format the notes file
								FileAppend, ----------------------------`nNotes for Reel: %reelnumber%`n`nDate`t`tNote`n`n, %reportpath%\%reelnumber%-notes.txt									
									
								; add the note to the report
								FileAppend, %date%`t%note%`n, %reportpath%\%reelnumber%-notes.txt

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

						; increment the notes counter
						notecount++
								
						; add the note to the report
						FileAppend, %date%`t%note%`n, %reportpath%\%reelnumber%-notes.txt

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

; ======DVV PAGES LOOP
; loop to view individual pages in the DVV
; Pause to pause/restart
; hold down F1 to cancel
DVVpages:
	; delay time dialog
	Gosub, DVVDelay
	
	; wait for window to close
	SetTitleMatchMode 1
	WinWaitActive, DVV Delay
	WinWaitClose, DVV Delay
	
	if (cancelbutton == 1)
	{
		; reset cancelbutton
		cancelbutton = 0
		
		; exit the script
		Return
	}
	
	; update last hotkey
	ControlSetText, Static3, PAGES, NDNP_QR
	ControlSetText, Static4, DVV Pages, NDNP_QR
		
	; initialize the page counter
	pagecount = 0
		
	; initialize the mini-counter
	; if the thumbs loop has been activated
	if (thumbscount != 0)
	{
		count = 0
		thumbscount = 0
	}

	; create the counter GUI
	; this window may be closed without exiting the loop
	Gui, 3:+AlwaysOnTop
	Gui, 3:+ToolWindow
	Gui, 3:Font, cGreen s25 bold, Arial
	Gui, 3:Add, Text, x15 y3 w75 h35, 0
	Gui, 3:Add, GroupBox, x0 y-19 w111 h65,
	Gui, 3:Show, h45 w110, DVV_Pages

	; set the loop to begin as paused
	Pause, On

	Loop
	{
		; open the next page
		Send, {Down}
		Send, {Enter}
			
		; update the scoreboard
		count++
		pagecount++
		ControlSetText, Static1, %pagecount%, DVV_Pages
		ControlSetText, Static15, %count%, NDNP_QR
			
		; delay for the specified time
		Sleep, %dvvdelayms%
			
		; end the loop if F1 is held down
		if getKeyState("F1")
		break
	}
		
	; close the counter window
	Gui, 3:Destroy

	; print exit message
	MsgBox, 0, DVV Pages, The loop has ended.
Return
; ======DVV PAGES LOOP

; ======DVV THUMBS LOOP
; loop to view the issue thumbs in the DVV
; Pause to pause/restart
; hold down F1 to cancel
DVVthumbs:
	; delay time dialog
	Gosub, DVVDelay
	
	; wait for window to close
	SetTitleMatchMode 1
	WinWaitActive, DVV Delay
	WinWaitClose, DVV Delay
	
	if (cancelbutton == 1)
	{
		; reset cancelbutton
		cancelbutton = 0
		
		; exit the script
		Return
	}

	; update last hotkey
	ControlSetText, Static3, THUMBS, NDNP_QR
	ControlSetText, Static4, DVV Thumbs, NDNP_QR
		
	; initialize the thumbs counter
	thumbscount = 0
		
	; initialize the mini-counter
	; if the pages loop has been activated
	if pagecount != 0
	{
		count = 0
		pagecount = 0
	}
		
	; create the counter GUI
	; this window may be closed without exiting the loop
	Gui, 3:+AlwaysOnTop
	Gui, 3:+ToolWindow
	Gui, 3:Font, cGreen s25 bold, Arial
	Gui, 3:Add, Text, x15 y3 w75 h35, 0
	Gui, 3:Add, GroupBox, x0 y-19 w111 h65,
	Gui, 3:Show, h45 w110, DVV_Thumbs
		
	; start the loop as paused
	Pause, On

	Loop
	{
		; move to next issue
		Send, {Down}
		Sleep,100
			
		; close the issue tree
		Send, {Left}
		Sleep, 100
			
		; open the issue thumbs
		Send, {AltDown}s{AltUp}
			
		; update scoreboard
		count++
		thumbscount++
		ControlSetText, Static1, %thumbscount%, DVV_Thumbs
		ControlSetText, Static15, %count%, NDNP_QR

		; delay for the specified time
		Sleep, %dvvdelayms%
			
		; end the loop if F1 is held down
		if getKeyState("F1")
		break
	}
		
	; close the counter window
	Gui, 3:Destroy
	
	; print exit message
	MsgBox, 0, DVV Thumbs, The loop has ended.
Return
; ======DVV THUMBS LOOP

; ======DVV DELAY DIALOG
; create the GUI
DVVDelay:
	Gui, 17:+ToolWindow
	Gui, 17:Add, Text,, Number seconds:
	; navchoice value (1-11) pre-selected, assigns value (2-12) to dvvdelay
	Gui, 17:Add, DropDownList, Choose%dvvdelaychoice% R11 vdvvdelay, 2|3|4|5|6|7|8|9|10|11|12
	; run DVVDelayGo if OK
	Gui, 17:Add, Button, w40 x10 y55 gDVVDelayGo default, OK
	; run DVVDelayCancel if Cancel
	Gui, 17:Add, Button, x65 y55 gDVVDelayCancel, Cancel
	Gui, 17:Show,, DVV Delay
Return

; OK button function
DVVDelayGo:
	; assign the dvvdelay variable value
	Gui, 17:Submit
	
	; assign the dvvdelaychoice variable value
	dvvdelaychoice := dvvdelay - 1
	
	; set the delay in milliseconds
	dvvdelayms := dvvdelay * 1000
	
	; close the DVV Delay GUI
	Gui, 17:Destroy
Return

; Cancel button function
DVVDelayCancel:
	; close the DVV Delay GUI
	Gui, 17:Destroy
	
	; set the CancelButton
	cancelbutton = 1
Return
; ======DVV DELAY DIALOG
