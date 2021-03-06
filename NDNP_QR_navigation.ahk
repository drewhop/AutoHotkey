; ********************************
; NDNP_QR_navigation.ahk
; required file for NDNP_QR.ahk
; ********************************

; ======================REEL FOLDER CHECK
ReelFolderCheck:
	if (reelfolderpath == "_")
	{
		; create dialog to select the reel folder
		FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect the REEL folder:
		if ErrorLevel
		{
			; reset the variables on Cancel
			reelfolderpath = _
			reelfoldername = _
			Exit
		}

		; create display variables for reel folder path
		StringGetPos, reelfolderpos, reelfolderpath, \, R2
		StringLeft, reelfolder1, reelfolderpath, reelfolderpos
		StringTrimLeft, reelfolder2, reelfolderpath, reelfolderpos

		; extract reel folder name from path
		StringGetPos, foldernamepos, reelfolderpath, \, R1
		foldernamepos++
		StringTrimLeft, reelfoldername, reelfolderpath, foldernamepos
		
		; loop checks to see if it is a reel folder
		Loop
		{
			StringLen, length, reelfoldername
			if (length != 11)
			{		
				; print error message
				MsgBox, 0, Reel Folder, %reelfoldername% does not appear to be a REEL folder.`n`n`tPlease try again.

				; reset the variables
				reelfolderpath = _
				reelfoldername = _

				; create dialog to select the reel folder
				FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect the REEL folder:
				if ErrorLevel
				{
					; reset the variables on Cancel
					reelfolderpath = _
					reelfoldername = _
					Exit
				}

				; create display variables for reel folder path
				StringGetPos, reelfolderpos, reelfolderpath, \, R2
				StringLeft, reelfolder1, reelfolderpath, reelfolderpos
				StringTrimLeft, reelfolder2, reelfolderpath, reelfolderpos

				; extract reel folder name from path
				StringGetPos, foldernamepos, reelfolderpath, \, R1
				foldernamepos++
				StringTrimLeft, reelfoldername, reelfolderpath, foldernamepos
			}
				
			else Break
		}
	}
Return
; ======================REEL FOLDER CHECK

; ======================RESET REEL
ResetReel:
	; loop to close any issue folders
	SetTitleMatchMode RegEx
	Loop
	{
		IfWinExist, ^[1-2][0-9]{9}$, , , ,
		{
			WinClose, ^[1-2][0-9]{9}$, , , ,
			Sleep, 200
		}
		else Break
	}

	; activate current reel folder if it exists
	SetTitleMatchMode 1
	IfWinExist, %reelfoldername%,,report
		WinActivate, %reelfoldername%,,,report

	; otherwise open the reel folder
	else
	{
		Run, %reelfolderpath%
		
		; wait for the folder to load
		WinWaitActive, %reelfoldername%,,,report
		
		; select first issue
		Send, {Down}
		Sleep, 50
		Send, {Up}
	}
Return
; ======================RESET REEL

; ======================CLOSE FIRST IMPRESSION WINDOWS
CloseFirstImpressionWindows:
	Loop
	{
		SetTitleMatchMode 2
		IfWinExist, First Impression
		{
			; get the unique window id#
			WinGet, firstid, ID, First Impression

			; close First Impression
			SetTitleMatchMode 1
			WinClose, ahk_id %firstid%
					
			; look for next window
			Continue
		}
				
		; exit loop if no more FI windows
		else Break
	}
Return
; ======================CLOSE FIRST IMPRESSION WINDOWS

; ======================OPEN FIRST TIFF
OpenFirstTIFF:
	; set the script flag
	tiffopenflag = 1

	; get the issue folder path
	Gosub, IssueFolderPath
	
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

	; reset the script flag
	tiffopenflag = 0
Return
; ======================OPEN FIRST TIFF

; ======================GO TO ISSUE
GoToIssue:
	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	winY+=%winHeight%
	InputBox, input, GoTo Issue, Issue Folder Name:,, 150, 125, %winX%, %winY%,,,%issuefolder%
	if ErrorLevel
		Return
	else
	{
		Loop ; valid title folder check
		{
			if RegExMatch(input, "\d\d\d\d\d\d\d\d\d\d")
			{
				issuefolder = %input%

				IfNotExist, %reelfolderpath%\%issuefolder%
				{
					MsgBox, 0, Error, GoToIssue`nNDNP_QR_navigation.ahk`n`n%issuefolder% does not exist in this directory:`n`n`t%reelfolder1%`n`t%reelfolder2%
					Return
				}
					
				; activate reel folder
				SetTitleMatchMode 1
				WinActivate, %reelfoldername%, , , report,
				WinWaitActive, %reelfoldername%, , 5, report,
				if ErrorLevel
				{
					; print error message after 5 seconds
					MsgBox, 0, Error, GoToIssue`nNDNP_QR_navigation.ahk`n`nCannot find window: %reelfoldername%`n`nNew Window:`n`n`t"File > Open Reel Folder"`n`nExisting Window:`n`n`t"Edit > Reel Folder > Set Path"
					Exit
				}
				Sleep, 100

				; go to the issue
				SetKeyDelay, 150
				Send, %issuefolder%
				SetKeyDelay, 10
				Sleep, 100

				Break
			}
				
			; if the issue folder name entered is not valid
			; print error message and re-enter loop
			else
			{
				MsgBox, 0, GoTo Issue, Please enter a folder name in the format: YYYYMMDDEE`n`nExample: 1942061901
				InputBox, input, GoTo Issue, Issue Folder Name:,, 150, 125, %winX%, %winY%,,,
					if ErrorLevel
						Return
			}
		}
	}
Return
; ======================GO TO ISSUE

; ======================ISSUE TO REEL
; legacy navigation script (versions 1.0-1.7)
; useful model for navigating Explorer windows
; handles inadvertent issue folder opens
; for OpenFirstTiff
IssueToReel:
	; activate the current NDNP issue folder
	SetTitleMatchMode RegEx
	WinWait, ^1[0-9]{9}$, , , ,
	IfWinNotActive, ^1[0-9]{9}$, , , ,
	WinActivate, ^1[0-9]{9}$, , , ,
	WinWaitActive, ^1[0-9]{9}$, , , ,
	Sleep, 100

	; up one directory
	Send, {AltDown}v{AltUp}
	Sleep, 100
	Send, g
	Sleep, 100
	Send, u
  
	; wait for the reel folder
	SetTitleMatchMode 1
	WinWaitActive, %reelfoldername%, , , report,
	Sleep, 100
Return
; ======================ISSUE TO REEL
