; ********************************
; NDNP_QR_navigation.ahk
; required file for NDNP_QR.ahk
; ********************************

; ======================REEL FOLDER CHECK
ReelFolderCheck:
  if (reelfolderpath == "_")
	{
		; create dialog to select the reel folder
		FileSelectFolder, reelfolderpath, %batchdrive%, 0, `nSelect the REEL folder:
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
Return
; ======================REEL FOLDER CHECK

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
	; create input box to enter the folder name
	InputBox, input, GoTo Issue, Issue Folder Name:,, 150, 125,,,,,%issuefolder%
	if ErrorLevel
		Return
	else
	{
		; loop checks for valid title folder name
		Loop
		{
			; if title folder name is valid (10 digit number)
			if RegExMatch(input, "\d\d\d\d\d\d\d\d\d\d")
			{
				; accept the input
				issuefolder = %input%

				; exit script if folder does not exist
				IfNotExist, %reelfolderpath%\%issuefolder%
				{
					; print an error message
					MsgBox, 0, Error, GoToIssue`nNDNP_QR_navigation.ahk`n`n%issuefolder% does not exist in this directory:`n`n`t%reelfolder1%`n`t%reelfolder2%
					
					; exit the script
					Return
				}
					
				; activate reel folder
				SetTitleMatchMode 1
				WinActivate, %reelfoldername%, , , ,
				WinWaitActive, %reelfoldername%, , 5, ,
				if ErrorLevel
				{
					; print error message after 5 seconds
					MsgBox, 0, Error, GoToIssue`nNDNP_QR_navigation.ahk`n`nCannot find folder %reelfoldername%`n`nNew Folder:`n`n`t"File > Open Reel Folder"`n`nExisting Folder:`n`n`t"Edit > Reel Folder > Select"
						
					; exit the script
					Exit
				}
				Sleep, 100

				; go to the issue
				SetKeyDelay, 150
				Send, %issuefolder%
				SetKeyDelay, 10
				Sleep, 100

				; end title check loop
				Break
			}
				
			; if the issue folder name entered is not valid
			; print error message and re-enter loop
			else
			{
				MsgBox, 0, GoTo Issue, Please enter a folder name in the format: YYYYMMDDEE`n`nExample: 1942061901
				InputBox, input, GoTo Issue, Issue Folder Name:,, 150, 125,,,,,
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
	WinWaitActive, %reelfoldername%, , , ,
	Sleep, 100
Return
; ======================ISSUE TO REEL
