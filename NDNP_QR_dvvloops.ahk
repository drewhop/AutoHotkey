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
	
	Gosub, DVVDelay
	SetTitleMatchMode 1
	WinWaitActive, DVV %dvvloopname% Delay
	WinWaitClose, DVV %dvvloopname% Delay
	
	if (cancelbutton == 1)
	{
		cancelbutton = 0 ; reset
		dvvloopname =
		Return
	}
	
	ControlSetText, Static3, PAGES, NDNP_QR
	ControlSetText, Static4, DVV Pages, NDNP_QR
		
	pagecount = 0
		
	; initialize the mini-counter
	; if the thumbs loop has been activated
	if (thumbscount != 0)
	{
		count = 0
		thumbscount = 0
	}

	; counter window
	Gui, 3:+AlwaysOnTop
	Gui, 3:+ToolWindow
	Gui, 3:Add, Button, w45 gDVVpause default, Start
	Gui, 3:Add, Button, gDVVcancel, Cancel
	Gui, 3:Font, cGreen s25 bold, Arial
	Gui, 3:Add, Text, x75 y13 w75 h35, 0
	SetTitleMatchMode 1
	WinGetPos, winX, winY,,, LOC Digital Viewer
	Gui, 3:Show, w170 x%winX% y%winY%, DVV_Pages

	ispaused = 1 ; start loop as paused
	
	Loop
	{
		if (ispaused == 1) ; paused
		{
			Loop
			{
				Sleep, 500 ; check for unpaused every half second
				if (ispaused == 0)
					break
					
				if (cancelbutton == 1) ; cancel button
				{
					cancelbutton = 0
					Gui, 3:Destroy ; counter
					
					if (notecount > 0) ; notes added
					{
						MsgBox, 4, DVV_%dvvloopname% Ended, You created %notecount% notes.`n`nWould you like to open the DVVnotes.txt file?
						IfMsgBox, Yes
						{
							Run, %DVVpath%\DVVnotes.txt
						}
					}
					else
						MsgBox, 0, DVV_Pages, The loop has ended.

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
		
	Gui, 3:Destroy ; counter
	if (notecount > 0) ; notes added
	{
		MsgBox, 4, DVV_%dvvloopname% Ended, You created %notecount% notes.`n`nWould you like to open the DVVnotes.txt file?
		IfMsgBox, Yes
		{
			Run, %DVVpath%\DVVnotes.txt
		}
	}
	else
		MsgBox, 0, DVV_Pages, The loop has ended.

	dvvloopname =
Return
; ======DVV PAGES LOOP

; ======DVV THUMBS LOOP
; loop to view thumbnail images in the DVV
DVVthumbs:
	dvvloopname = Thumbs
	notecount = 0

	Gosub, DVVDelay
	SetTitleMatchMode 1
	WinWaitActive, DVV %dvvloopname% Delay
	WinWaitClose, DVV %dvvloopname% Delay
	
	if (cancelbutton == 1)
	{
		cancelbutton = 0 ; reset
		dvvloopname =
		Return
	}

	ControlSetText, Static3, THUMBS, NDNP_QR
	ControlSetText, Static4, DVV Thumbs, NDNP_QR
		
	thumbscount = 0
		
	; initialize the mini-counter
	; if the pages loop has been activated
	if pagecount != 0
	{
		count = 0
		pagecount = 0
	}
		
	; counter window
	Gui, 3:+AlwaysOnTop
	Gui, 3:+ToolWindow
	Gui, 3:Add, Button, w45 gDVVpause default, Start
	Gui, 3:Add, Button, gDVVcancel, Cancel
	Gui, 3:Font, cGreen s25 bold, Arial
	Gui, 3:Add, Text, x75 y13 w75 h35, 0
	SetTitleMatchMode 1
	WinGetPos, winX, winY,,, LOC Digital Viewer	
	Gui, 3:Show, x%winX% y%winY%, DVV_Thumbs
		
	ispaused = 1 ; start loop as paused

	Loop
	{
		if (ispaused == 1) ; pause loop
		{
			Loop
			{
				Sleep, 500 ; check for unpaused every half second
				if (ispaused == 0)
					break
					
				if (cancelbutton == 1) ; cancel button
				{
					cancelbutton = 0
					Gui, 3:Destroy ; counter

					if (notecount > 0) ; notes added
					{
						MsgBox, 4, DVV_%dvvloopname% Ended, You created %notecount% notes.`n`nWould you like to open the DVVnotes.txt file?
						IfMsgBox, Yes
						{
							Run, %DVVpath%\DVVnotes.txt
						}
					}
					else
						MsgBox, 0, DVV_Thumbs, The loop has ended.

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
	
	Gui, 3:Destroy ; counter
	
	if (notecount > 0) ; notes added
	{
		MsgBox, 4, DVV_%dvvloopname% Ended, You created %notecount% notes.`n`nWould you like to open the DVVnotes.txt file?
		IfMsgBox, Yes
		{
			Run, %DVVpath%\DVVnotes.txt
		}
	}
	else
		MsgBox, 0, DVV_Thumbs, The loop has ended.

	dvvloopname =
Return
; ======DVV THUMBS LOOP

; ======DVV NOTES
DVVnotes:
	if (notecount == 0) ; inactive notes
	{
		MsgBox, 4, DVV_%dvvloopname% Notes, Would you like to save a note?`n`nThe DVVnotes.txt file will be`ncreated if it does not already exist.
		IfMsgBox, Yes
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

			if (note != "") ; check for blank note
			{
				notecount++
				FileAppend, -------------------------------------`nDVV_%dvvloopname% Notes: %A_YYYY%-%A_MM%-%A_DD% (%A_Hour%:%A_Min%)`n`n, %DVVpath%\DVVnotes.txt									
				FileAppend, %note%`n`n, %DVVpath%\DVVnotes.txt
			}		
		}
	}
						
	else ; add to Notes file
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

		if (note != "") ; check for blank note
		{
			notecount++
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
DVVDelay:
	Gui, 17:+ToolWindow
	Gui, 17:Add, Text,, Number of seconds:
	; navchoice value (1-11) pre-selected, assigns value (2-12) to dvvdelay
	Gui, 17:Add, DropDownList, Choose%dvvdelaychoice% R11 vdvvdelay, 2|3|4|5|6|7|8|9|10|11|12
	; run DVVDelayGo if OK
	Gui, 17:Add, Button, w40 x10 y55 gDVVDelayGo default, OK
	; run DVVDelayCancel if Cancel
	Gui, 17:Add, Button, x65 y55 gDVVDelayCancel, Cancel
	
	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	Gui, 17:Show, x%winX% y%winY%, DVV %dvvloopname% Delay
Return

; OK button
DVVDelayGo:
	Gui, 17:Submit
	dvvdelaychoice := dvvdelay - 1
	dvvdelayms := dvvdelay * 1000 ; milliseconds
	Gui, 17:Destroy
Return

; Cancel button
DVVDelayCancel:
	Gui, 17:Destroy
	cancelbutton = 1
Return
; ======DVV DELAY DIALOG
