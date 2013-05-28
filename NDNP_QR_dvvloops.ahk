; ********************************
; NDNP_QR_dvvloops.ahk
; required file for NDNP_QR.ahk
; ********************************

; ======DVV PAGES LOOP
; loop to view individual pages in the DVV
; Pause to pause/restart
; hold down F1 to cancel
DVVpages:
	dvvloopname = Pages
	notecount = 0
	
	; delay time dialog
	Gosub, DVVDelay
	
	; wait for window to close
	SetTitleMatchMode 1
	WinWaitActive, DVV %dvvloopname% Delay
	WinWaitClose, DVV %dvvloopname% Delay
	
	if (cancelbutton == 1)
	{
		; reset cancelbutton
		cancelbutton = 0
		
		; reset loopname
		dvvloopname =
		
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
	Gui, 3:+AlwaysOnTop
	Gui, 3:+ToolWindow
	Gui, 3:Add, Button, w45 gDVVpause default, Start
	Gui, 3:Add, Button, gDVVcancel, Cancel
	Gui, 3:Font, cGreen s25 bold, Arial
	Gui, 3:Add, Text, x75 y13 w75 h35, 0
	
	; position in upper left corner of DVV window
	SetTitleMatchMode 1
	WinGetPos, winX, winY,,, LOC Digital Viewer
	Gui, 3:Show, w170 x%winX% y%winY%, DVV_Pages

	; start loop as paused
	ispaused = 1
	
	Loop
	{
		; pause loop
		if (ispaused == 1)
		{
			Loop
			{
				; check for unpause every half second
				Sleep, 500
				if (ispaused == 0)
					break
					
				; end the Pages loop if Cancel button
				if (cancelbutton == 1)
				{
					cancelbutton = 0

					; close the counter window
					Gui, 3:Destroy

					; print exit message
					if (notecount > 0)
					{
						MsgBox, 4, DVV_%dvvloopname% Ended, You created %notecount% notes.`n`nWould you like to open the DVVnotes.txt file?
						IfMsgBox, Yes
						{
							Run, %DVVpath%\DVVnotes.txt
						}
					}
					
					else
						MsgBox, 0, DVV_Pages, The loop has ended.

					; reset loopname
					dvvloopname =
	
					Return
				}
			}
		}
		
		; activate DVV window
		SetTitleMatchMode 1
		WinActivate, LOC Digital Viewer
		Sleep, 100
		
		; open the next page
		Send, {Down}
		Send, {Enter}
		
		; activate the counter window
		IfWinExist, DVV_Pages
			WinActivate, DVV_Pages
		
		; update the scoreboard
		count++
		pagecount++
		ControlSetText, Static1, %pagecount%, DVV_Pages
		ControlSetText, Static15, %count%, NDNP_QR
			
		; delay for the specified time
		Sleep, %dvvdelayms%
			
		; hold down Right Shift for note
		getKeyState, state, RShift
		if state = D
			Gosub, DVVnotes

		; end the loop if F1 is held down
		if getKeyState("F1")
			break

		; end the loop if Cancel button
		if (cancelbutton == 1)
		{
			cancelbutton = 0
			break
		}
	}
		
	; close the counter window
	Gui, 3:Destroy

	; print exit message
	if (notecount > 0)
	{
		MsgBox, 4, DVV_%dvvloopname% Ended, You created %notecount% notes.`n`nWould you like to open the DVVnotes.txt file?
		IfMsgBox, Yes
		{
			Run, %DVVpath%\DVVnotes.txt
		}
	}
	
	else
		MsgBox, 0, DVV_Pages, The loop has ended.

	; reset loopname
	dvvloopname =
Return
; ======DVV PAGES LOOP

; ======DVV THUMBS LOOP
; loop to view thumbnail images in the DVV
DVVthumbs:
	dvvloopname = Thumbs
	notecount = 0

	; delay time dialog
	Gosub, DVVDelay
	
	; wait for window to close
	SetTitleMatchMode 1
	WinWaitActive, DVV %dvvloopname% Delay
	WinWaitClose, DVV %dvvloopname% Delay
	
	if (cancelbutton == 1)
	{
		; reset cancelbutton
		cancelbutton = 0
		
		; reset loopname
		dvvloopname =
		
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
	Gui, 3:+AlwaysOnTop
	Gui, 3:+ToolWindow
	Gui, 3:Add, Button, w45 gDVVpause default, Start
	Gui, 3:Add, Button, gDVVcancel, Cancel
	Gui, 3:Font, cGreen s25 bold, Arial
	Gui, 3:Add, Text, x75 y13 w75 h35, 0
	
	; position in upper left corner of DVV window
	SetTitleMatchMode 1
	WinGetPos, winX, winY,,, LOC Digital Viewer	
	Gui, 3:Show, x%winX% y%winY%, DVV_Thumbs
		
	; start the loop as paused
	ispaused = 1

	Loop
	{
		; pause loop
		if (ispaused == 1)
		{
			Loop
			{
				; check for unpause every half second
				Sleep, 500
				if (ispaused == 0)
					break
					
				; end the Pages loop if Cancel button
				if (cancelbutton == 1)
				{
					cancelbutton = 0

					; close the counter window
					Gui, 3:Destroy

					; print exit message
					if (notecount > 0)
					{
						MsgBox, 4, DVV_%dvvloopname% Ended, You created %notecount% notes.`n`nWould you like to open the DVVnotes.txt file?
						IfMsgBox, Yes
						{
							Run, %DVVpath%\DVVnotes.txt
						}
					}
					
					else
						MsgBox, 0, DVV_Thumbs, The loop has ended.

					; reset loopname
					dvvloopname =
	
					Return
				}
			}
		}
		
		; activate DVV window
		SetTitleMatchMode 1
		WinActivate, LOC Digital Viewer
		Sleep, 100
				
		; move to next issue
		Send, {Down}
		Sleep, 100
			
		; close the issue tree
		Send, {Left}
		Sleep, 100
			
		; open the issue thumbs
		Send, {AltDown}s{AltUp}
			
		; activate the counter window
		IfWinExist, DVV_Thumbs
			WinActivate, DVV_Thumbs
		
		; update scoreboard
		count++
		thumbscount++
		ControlSetText, Static1, %thumbscount%, DVV_Thumbs
		ControlSetText, Static15, %count%, NDNP_QR

		; delay for the specified time
		Sleep, %dvvdelayms%
		
		; hold down Right Shift for note
		getKeyState, state, RShift
		if state = D
			Gosub, DVVnotes

		; end the loop if F1 is held down
		if getKeyState("F1")
			break
			
		; end the loop if Cancel button
		if (cancelbutton == 1)
		{
			cancelbutton = 0
			break
		}
	}
	
	; close the counter window
	Gui, 3:Destroy
	
	; print exit message
	if (notecount > 0)
	{
		MsgBox, 4, DVV_%dvvloopname% Ended, You created %notecount% notes.`n`nWould you like to open the DVVnotes.txt file?
		IfMsgBox, Yes
		{
			Run, %DVVpath%\DVVnotes.txt
		}
	}
	
	else
		MsgBox, 0, DVV_Thumbs, The loop has ended.

	; reset loopname
	dvvloopname =
Return
; ======DVV THUMBS LOOP

; ======DVV NOTES
DVVnotes:
	; case for inactive notes function
	if (notecount == 0)
	{
		MsgBox, 4, DVV_%dvvloopname% Notes, Would you like to save a note?`n`nThe DVVnotes.txt file will be`ncreated if it does not already exist.
		IfMsgBox, Yes
		{
			; activate notes function
			IfWinExist, DVV_%dvvloopname%
			{
				WinGetPos, winX, winY, winWidth, winHeight, DVV_%dvvloopname%
				winX+=%winWidth%
				InputBox, note, DVV_%dvvloopname% Notes, Enter a note:,, 450, 120, %winX%, %winY%,,,
				if ErrorLevel
					Return				
			}
			
			else
			{
				WinGetPos, winX, winY, winWidth, winHeight, LOC Digital Viewer
				InputBox, note, DVV_%dvvloopname% Notes, Enter a note:,, 450, 120, %winX%, %winY%,,,
				if ErrorLevel
					Return
			}

			; check if the note is blank
			if (note != "")
			{
				; increment the notes counter
				notecount++
								
				; format the notes file
				FileAppend, -------------------------------------`nDVV_%dvvloopname% Notes: %A_YYYY%-%A_MM%-%A_DD% (%A_Hour%:%A_Min%)`n`n, %DVVpath%\DVVnotes.txt									
										
				; add the note to the DVVnotes.txt file
				FileAppend, %note%`n`n, %DVVpath%\DVVnotes.txt
			}		
		}
	}
						
	; add note to notes file
	else
	{
		IfWinExist, DVV_%dvvloopname%
		{
			WinGetPos, winX, winY, winWidth, winHeight, DVV_%dvvloopname%
			winX+=%winWidth%
			InputBox, note, DVV_%dvvloopname% Notes, Enter a note:,, 450, 120, %winX%, %winY%,,,
			if ErrorLevel
				Return				
		}
			
		else
		{
			WinGetPos, winX, winY, winWidth, winHeight, LOC Digital Viewer
			InputBox, note, DVV_%dvvloopname% Notes, Enter a note:,, 450, 120, %winX%, %winY%,,,
			if ErrorLevel
				Return
		}

		; check if the note is blank
		if (note != "")
		{
			; increment the notes counter
			notecount++
									
			; add the note to the DVVnotes.txt file
			FileAppend, %note%`n`n, %DVVpath%\DVVnotes.txt
		}
	}
Return
; ======DVV NOTES

; ======DVV COUNTER BUTTONS
; toggles Pause/Start
DVVpause:
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

; cancels the dvv loops
DVVcancel:
	cancelbutton = 1
	GuiControl, Hide, Pause
	GuiControl, Hide, Start
	GuiControl, Hide, Cancel	
	ControlSetText, Static1, ..., DVV_Pages
	ControlSetText, Static1, ..., DVV_Thumbs
Return
; ======DVV COUNTER BUTTONS

; ======DVV DELAY DIALOG
; create the GUI
DVVDelay:
	Gui, 17:+ToolWindow
	Gui, 17:Add, Text,, Number of seconds:
	; navchoice value (1-11) pre-selected, assigns value (2-12) to dvvdelay
	Gui, 17:Add, DropDownList, Choose%dvvdelaychoice% R11 vdvvdelay, 2|3|4|5|6|7|8|9|10|11|12
	; run DVVDelayGo if OK
	Gui, 17:Add, Button, w40 x10 y55 gDVVDelayGo default, OK
	; run DVVDelayCancel if Cancel
	Gui, 17:Add, Button, x65 y55 gDVVDelayCancel, Cancel
	
	; position below the NDNP_QR window
	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	winY += %winHeight%
	Gui, 17:Show, x%winX% y%winY%, DVV %dvvloopname% Delay
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
