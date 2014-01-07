#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#Persistent  ; keeps the script running until explicitly ended
#SingleInstance force  ; skips reload dialog

; variables
Qpath = Q:\Greensheet\Greensheet_Archives\
titlefolderpath = _
renameflag = 0
renameadd = 0
renamesubtract = 0
dayadjust = 0

; GUI
Gui, 11:Color, d0d0d0, 912206
Gui, 11:Show, h0 w0, Greensheet

; File menu
Menu, FileMenu, Add, Greensheet &Archives, GreensheetArchives
Menu, FileMenu, Add, Greensheet &Processing, GreensheetProcessing
Menu, FileMenu, Add
Menu, FileMenu, Add, Create &Report, IssueCheckPY
Menu, FileMenu, Add
Menu, FileMenu, Add, Reloa&d, Reload
Menu, FileMenu, Add, E&xit, Exit
; Edit menu
Menu, EditMenu, Add, &Rename Setup, RenameSetup
; Help menu
Menu, HelpMenu, Add, &Documentation, Documentation
Menu, HelpMenu, Add, &About, About
; create menus
Menu, MenuBar, Add, &File, :FileMenu
Menu, MenuBar, Add, &Edit, :EditMenu
Menu, MenuBar, Add, &Help, :HelpMenu
; create menu toolbar
Gui, 11:Menu, MenuBar

; GUI elements
Gui, 11:Add, Text, x75 y15 w300 h15, %Qpath%
Gui, 11:Add, Button, x10 y10 w55 h25 gFolderSelect, Select 

; row 2
Gui, 11:Add, Button, x10 y45 w105 h35 gRename, Rename Issues
Gui, 11:Add, Button, x120 y45 w105 h35 gMove, Move PDFs
Gui, 11:Add, Button, x230 y45 w105 h35 gProcess, Process Images


WinGetPos, winX, winY, winWidth, winHeight, Greensheet
winX+=%winWidth%
Gui, 11:Show, x%winX% y%winY% h90 w345, Greensheet
winactivate, Greensheet

Pause::Pause  ; pause key pauses any script function

; ==============================
; DISPLAY META
; display the issue metadata
; Hotkey: Alt + m
!m::
	; get the selected folder path
	Gosub, IssueFolderPath

	Sleep, 500

	; metadata harvest subroutine
	Gosub, MetaHarvest

Return
; ==============================

; ==============================
; DISPLAY NEXT META
; display the issue metadata
; Hotkey: Alt + ,
!,::
	; activate the year folder
	WinActivate, %titlefoldername%
	Sleep, 200
	
	; down 1 folder
	Send, {Down}
	Sleep, 200

	; get the selected folder path
	Gosub, IssueFolderPath

	Sleep, 500

	; metadata harvest subroutine
	Gosub, MetaHarvest

Return
; ==============================

; ==============================
; EDIT META
; opens the metadata.txt file for editing
; Hotkey: Alt + 0
!0::
	; harvest the issue folder path
	Gosub, IssueFolderPath

	Sleep, 500
	
	; if there is a metadata.txt file in the folder
	IfExist, %issuefolderpath%\metadata.txt
	{
		; open the metadata.txt document
		Run, metadata.txt, %issuefolderpath%

	Return
	}

	; if there is no metadata.txt file in the folder
	else
	{
		; create display variables for issue folder path
		StringGetPos, issuefolderpos, issuefolderpath, \, R3
		StringLeft, issuefolder1, issuefolderpath, issuefolderpos
		StringTrimLeft, issuefolder2, issuefolderpath, issuefolderpos		

		; print an error message
		MsgBox, 0, Edit Meta, There is no metadata.txt file in this folder.`n`n%issuefolder1%`n%issuefolder2%`n`nUse New Meta (Alt + n) to create a new file.
	}
Return
; ==============================

; ==================SAVE
; Save and Close Notepad document (Ctrl + . )
^.::
	SetTitleMatchMode 2
	WinMenuSelectItem, Notepad,, File, Save,,,,,, Notepad++,
	WinMenuSelectItem, Notepad,, File, Exit,,,,,, Notepad++,
Return
; ==================SAVE

; ===================SELECT FOLDER
FolderSelect:
	; destroy the metadata display windows
	Gui, 6:Destroy
	Gui, 7:Destroy
	Gui, 8:Destroy
	Gui, 9:Destroy
	
	FileSelectFolder, titlefolderpath, %Qpath%, 2, `nSelect the folder to process:
	if ErrorLevel
	{
		titlefolderpath = _
		Return
	}
	else
	{
		; set working directory to selected folder
		SetWorkingDir %titlefolderpath%
		
		; reset the rename setup flag
		renameflag = 0
		
		; create display path
		StringTrimLeft, displaypath, titlefolderpath, 43

		; extract year folder name from path
		StringRight, titlefoldername, titlefolderpath, 4
		
		; update GUI
		ControlSetText, Static1, Folder: %displaypath%, Greensheet
	}
Return
; ===================SELECT FOLDER

; ===================RENAME ISSUES
Rename:
/*
		Loop, *, 2, 0
		{
			; get the folder name
			foldername = %A_LoopFileName%
			
			if RegExMatch(A_LoopFileName, "\d\d\d\d\d\d\d\d\d\d")
			{
				; get the year
				StringLeft, year, foldername, 4
					
				; get the month
				StringMid, month, foldername, 5, 2
					
				; get the day
				StringMid, day, foldername, 7, 2
					
				; increase the day by one
				day++
					
				; rename the folder
				if (day < 10)
				{
					FileMoveDir, %A_LoopFileName%, %year%%month%0%day%01, R
				}
				else
				{
					FileMoveDir, %A_LoopFileName%, %year%%month%%day%01, R		
				}

				; pause
				Sleep, 100
			}
		}
*/
	if (titlefolderpath == "_")
	{
		MsgBox, Please select a folder and try again.
		Return
	}
	if (renameflag == 0)
	{
		MsgBox, Please select "Edit > Rename Setup" and try again.
		Return
	}

	; adding days - loop to process all folders without recursion
	if (radiobutton == 2)
	{
		Loop, *, 2, 0
		{
			; get the folder name
			foldername = %A_LoopFileName%
			
			if RegExMatch(A_LoopFileName, "\d\d\d\d\d\d\d\d\d\d")
				Continue
				
			else if RegExMatch(A_LoopFileName, "\d\d-\d\d-\d\d")
			{
				; get the year
				StringRight, year, foldername, 2
					
				; get the month
				StringLeft, month, foldername, 2
					
				; get the day
				StringMid, day, foldername, 4, 2
					
				; increase the day by one
				day += %dayadjust%
					
				; rename the folder
				if (day < 10)
				{
					FileMoveDir, %A_LoopFileName%, 20%year%%month%0%day%01, R
				}
				else
				{
					FileMoveDir, %A_LoopFileName%, 20%year%%month%%day%01, R		
				}

				; pause
				Sleep, 100
			}
			
			else Continue
		}
	}
	
	; subtracting days - loop to process all folders without recursion
	else if (radiobutton == 3)
	{
		Loop, *, 2, 0
		{
			if RegExMatch(A_LoopFileName, "\d\d\d\d\d\d\d\d\d\d")
				Continue
				
			else if RegExMatch(A_LoopFileName, "\d\d-\d\d-\d\d")
			{
				; get the folder name
				foldername = %A_LoopFileName%
					
				; get the year
				StringRight, year, foldername, 2
					
				; get the month
				StringLeft, month, foldername, 2
					
				; get the day
				StringMid, day, foldername, 4, 2
					
				; reduce the day by one
				day -= %dayadjust%
					
				; rename the folder
				if (day < 10)
				{
					FileMoveDir, %A_LoopFileName%, 20%year%%month%0%day%01, R
				}
				else
				{
					FileMoveDir, %A_LoopFileName%, 20%year%%month%%day%01, R		
				}

				; pause
				Sleep, 100
			}
			
			else Continue
		}
	}

	; default - loop to process all folders without recursion
	else
	{
		Loop, *, 2, 0
		{
			if RegExMatch(A_LoopFileName, "\d\d\d\d\d\d\d\d\d\d")
				Continue
				
			else if RegExMatch(A_LoopFileName, "\d\d-\d\d-\d\d")
			{
				; get the folder name
				foldername = %A_LoopFileName%
					
				; get the year
				StringRight, year, foldername, 2
					
				; get the month
				StringLeft, month, foldername, 2
					
				; get the day
				StringMid, day, foldername, 4, 2
					
				; rename the folder
				FileMoveDir, %A_LoopFileName%, 20%year%%month%%day%01, R
				
				; pause
				Sleep, 100
			}
			
			else Continue
		}
	}

	; check for invalid dates
	Loop, *, 2, 0
	{
		if RegExMatch(A_LoopFileName, "\d\d\d\d\d\d\d\d\d\d")
		{
			; get the folder name
			foldername = %A_LoopFileName%
					
			; get the year
			StringLeft, year, foldername, 4
					
			; get the month
			StringMid, month, foldername, 5, 2
					
			; get the day
			StringMid, day, foldername, 7, 2
								
			; for day 32
			if (day == 32)
			{
				; set the day to 1
				day = 01
				
				; rename the folder
				if (month == 12)
				{
					; increase the year
					year++
					
					; set the month to 01
					month = 01
					
					FileMoveDir, %A_LoopFileName%, %year%%month%%day%01, R					
				}
				else
				{
					; increment the month
					month++
					
					if (month < 10)
					{				
						FileMoveDir, %A_LoopFileName%, %year%0%month%%day%01, R
					}
					else
					{				
						FileMoveDir, %A_LoopFileName%, %year%%month%%day%01, R
					}
				}
			}
			
			; for day 31
			if ((day == 31) && ((month == 4)||(month == 6)||(month == 9)||(month == 11)))
			{
				; set the day to 1
				day = 01
				
				; increase the month
				month++
				
				; rename the folder
				if (month < 10)
				{				
					FileMoveDir, %A_LoopFileName%, %year%0%month%%day%01, R
				}
				else
				{				
					FileMoveDir, %A_LoopFileName%, %year%%month%%day%01, R
				}
			}

			; for Feb 29
			if ((day == 29) && (month == 2))
			{
				if ((year == 2000)||(year == 2004)||(year == 2008)||(year == 2012))
					Continue
					
				else
				{
					; set the day to 1
					day = 01
					
					; set the month to 03
					month = 03
					
					; rename the folder
					FileMoveDir, %A_LoopFileName%, %year%%month%%day%01, R
				}
			}

			; for day 00
			if (day == 00)
			{
				if (month == 01)
				{
					; decrease the year
					year--
					
					; set the month to 12
					month = 12
					
					; set the day to 31
					day = 31
					
					; rename the folder
					FileMoveDir, %A_LoopFileName%, %year%%month%%day%01, R
				}
				
				else if ((month == 02)||(month == 04)||(month == 06)||(month == 08)||(month == 09)||(month == 11))
				{
					; set the day to 31
					day = 31
					
					; decrease the month
					month--
					
					; rename the folder
					if (month < 10)
					{				
						FileMoveDir, %A_LoopFileName%, %year%0%month%%day%01, R
					}
					else
					{				
						FileMoveDir, %A_LoopFileName%, %year%%month%%day%01, R
					}					
				}
				
				else if ((month == 05)||(month == 07)||(month == 10)||(month == 12))
				{
					; set the day to 30
					day = 30
					
					; decrease the month
					month--
					
					; rename the folder
					if (month < 10)
					{				
						FileMoveDir, %A_LoopFileName%, %year%0%month%%day%01, R
					}
					else
					{				
						FileMoveDir, %A_LoopFileName%, %year%%month%%day%01, R
					}									
				}
				
				else
				{
					; set the month to 02
					month = 02
					
					; check for leap year
					if ((year == 2000)||(year == 2004)||(year == 2008)||(year == 2012))
					{
						day = 29
					}
					else
					{
						day = 28
					}
					
					; rename the folder
					FileMoveDir, %A_LoopFileName%, %year%%month%%day%01, R					
				}
			}
			
			; pause
			Sleep, 100
		}
	}

	; reset the variables
	radiobutton = 1
	dayadjust = 0
	
	; notification
	MsgBox, Rename Complete:`n`n%displaypath%
Return

RenameSetup:
	WinGetPos, winX, winY, winWidth, winHeight, Greensheet
	winY+=%winHeight%

	; create the GUI
	Gui, 4:+AlwaysOnTop
	Gui, 4:+ToolWindow
	Gui, 4:Add, Text, x10 y10 w230 h20, %displaypath%
	Gui, 4:Add, Radio, vradiobutton Checked, No Adjustment
	Gui, 4:Add, Radio, , Add Day(s)
	Gui, 4:Add, Radio, , Subtract Day(s)
	Gui, 4:Add, Text, x10 y95 w40 h20, Days:
	Gui, 4:Add, Edit, x40 y95 w20 vdayadjust, 0
	Gui, 4:Add, Button, x10 y125 w40 Default gRenameSubmit, OK
	Gui, 4:Add, Button, x60 y125 gRenameCancel, Cancel
	Gui, 4:Show, x%winX% y%winY% h160 w220, Rename Setup
Return

RenameSubmit:
	; save values and destroy the Gui
	Gui, 4:Submit
	Gui, 4:Destroy
	
	; set the rename setup flag
	renameflag = 1
	
	MsgBox, Click "Rename Issues" to rename:  %displaypath%
Return

RenameCancel:
	; destroy the Gui
	Gui, 4:Destroy
	
	; reset the variables
	radiobutton = 1
	dayadjust = 0
	renameflag = 0
Return
; ===================RENAME FOLDERS

; ===================CREATE FOLDERS
Move:
	; initialize run counter
	runcount = 0

	; create GUI
	WinGetPos, winX, winY, winWidth, winHeight, Greensheet
	Gui, 5:+ToolWindow
	Gui, 5:Add, Text,, %displaypath%
	Gui, 5:Add, Text,, Processing: YYYYMMDDEE
	Gui, 5:Add, Text,, Please wait . . .
	Gui, 5:Show, x%winX% y%winY%, Move PDFs
			
	; loop to process all folders without recursion
	Loop, *, 2, 0
	{
		; update status window
		ControlSetText, Static2, Issue: %A_LoopFileName%, Move PDFs
	
		; create folders for images
		FileCreateDir, %A_LoopFileName%\01_jpg
		FileCreateDir, %A_LoopFileName%\02_pdf
			
		; move all PDFs into 02_pdf folder
		FileMove, %A_LoopFileName%\*.pdf, %A_LoopFileName%\02_pdf
			
		; update the counter
		runcount++
	}

	; destroy the GUI
	Gui, 5:Destroy
		
	; create a dialog to indicate that the script is done
	MsgBox, PDF Move Complete:`n`n%displaypath%`n`nIssues: %runcount%
Return
; ===================CREATE FOLDERS

; ===================EXPORT JPG
Process:
	; error message
	if (titlefolderpath = _)
	{
		MsgBox, Please select a folder and try again.
		Return
	}
	
	runcount = 0
	
	; input boxes for starting numbers
	WinGetPos, winX, winY, winWidth, winHeight, Greensheet
	InputBox, volume, Volume, Enter the starting volume number:,, 230, 125, %winX%, %winY%,
	if ErrorLevel
		Return
	else
	{
		InputBox, issue, Issue, Enter the starting issue number:,, 230, 125, %winX%, %winY%,
		if ErrorLevel
			Return
		else
		{
			InputBox, issueadjust, Issue Increment, Enter the issue increment value:,, 230, 125, %winX%, %winY%,
			if ErrorLevel
				Return
			else
			{
				; loop to process all folders without recursion
				Loop, *, 2, 0
				{
					; increment the run counter
					runcount++
						
					; initialize the counters
					PDFcount = 0
					JPGcount = 0

					; get the issue folder path
					issuefolderpath = %A_LoopFileFullPath%
						
					; count the PDFs in 02_pdf
					Loop, %A_LoopFileName%\02_pdf\*.pdf
					{
						PDFcount++
					}

					; count the JPGs in 01_jpg
					Loop, %A_LoopFileName%\01_jpg\*.jpg
					{
						JPGcount++
					}

					IfEqual, PDFcount, 0
					{
						MsgBox, There are no PDFs to process:`n`n%A_LoopFileName%\02_pdf
						Continue
					}
						
					IfGreater, JPGcount, 0
					{
						MsgBox, This issue may have been processed:`n`n`t%A_LoopFileName%`n`nPDF Count: %PDFcount%`n JPG Count: %JPGcount%
						Continue
					}
						
					; update the GUIs
					ControlSetText, Static1, %PDFcount%, PDF Count	
					ControlSetText, Static2, %A_LoopFileName%, PDF Count	
					ControlSetText, Static2, %A_LoopFileName%, Volume	
					ControlSetText, Static2, %A_LoopFileName%, Issue	
					
					; open the 02_pdf folder
					Run, %A_LoopFileName%\02_pdf
						
					; wait for the folder to open
					WinWaitActive, 02_pdf
					Sleep, 300
						
					; select all files
					Send, {CtrlDown}a{CtrlUp}
					Sleep, 300

					; select "Combine supported files in Acrobat"
					Send, {AppsKey}
					Sleep, 300
					Send, {Down 2}
					Sleep, 300
					Send, {Enter}
					Sleep, 200

					; wait for Combine Files window
					WinWaitActive, Combine Files
					Sleep, 200

					; select large file icon
					Send, {Tab}
					Sleep, 200
					Send, {Tab}
					Sleep, 200
					Send, {Tab}
					Sleep, 200
					Send, {Tab}
					Sleep, 200
					Send, {Enter}
					Sleep, 200

					; highlight the Combine Files button
					Send, {Tab}
					Sleep, 200
					Send, {Tab}
					Sleep, 200
					Send, {Tab}

					; create PDF count dialog if first run or does not exist
					IfEqual, runcount, 1
						GoSub, PDFupdate

					IfEqual, runcount, 1
						GoSub, VolumeUpdate

					IfEqual, runcount, 1
						GoSub, IssueUpdate

					IfWinNotExist, PDF Count
						GoSub, PDFupdate
						
					; waits for Save As dialog
					WinWaitActive, Save As
					Sleep, 300

					; selects Cancel
					Send, {Tab 3}
					Sleep, 200
					Send, {Enter}
					Sleep, 200

					; Redo label for not enough JPGs
					Redo:
						
					; waits for Binder
					WinActivate, Binder
					WinWaitActive, Binder
					Sleep, 300

					; select File > Export > Image > JPEG
					Send, {AltDown}f{AltUp}
					Sleep, 300
					Send, t
					Sleep, 300
					Send, i
					Sleep, 300
					Send, {Enter}
					Sleep, 200

					; wait for Save As dialog
					WinWaitActive, Save As
					Sleep, 300

					; enter the path to 01_jpg
					Send, {Home}
					Sleep, 300
					Send, %A_LoopFileLongPath%\01_jpg\
					Sleep, 300
					Send, {Enter}
					Sleep, 300
					
					if (runcount != 1)
					{
						; increment the issue number
						issue += %issueadjust%
					}
					
					; update the issue number display
					ControlSetText, Static1, %issue%, Issue
						
					; create volume number dialog if it does not exist
					IfWinNotExist, Volume
						GoSub, VolumeUpdate
					
					Sleep, 200
						
					; create issue number dialog if it does not exist
					IfWinNotExist, Issue
						GoSub, IssueUpdate

					; count the JPGs every 5 seconds
					Loop
					{
						getKeyState, state, RShift
						if state = D
						{
							MsgBox, 2, Export JPG, %displaypath%\%A_LoopFileName%`n`n`t`tPDF Count: %PDFcount%`n`t`t JPG Count: %JPGcount%
							IfMsgBox Abort
							{
								Gui, 1:Destroy
								Gui, 2:Destroy
								Gui, 3:Destroy	
								Return
							}
							IfMsgBox Retry
							{
								FileDelete, %A_LoopFileName%\01_jpg\*.jpg
								Goto, Redo
							}
							IfMsgBox Ignore
								Continue
						}
							
						Sleep, 5000
							
						; initialize the counter
						JPGcount = 0
							
						Loop, %A_LoopFileName%\01_jpg\*.jpg
						{
							JPGcount++
						}

						; break the loop if done processing images
						IfEqual, PDFcount, %JPGcount%
						{			
							; create metadata.txt file
							FileAppend, volume: %volume%`nissue: %issue%`nnote:, %issuefolderpath%\metadata.txt

							Break
						}
					}
						
					; close the 02_pdf window
					WinClose, 02_pdf
					Sleep, 100
					
					; close the Binder window
					WinActivate, Binder
					WinWaitActive, Binder
					Sleep, 100
					WinClose, Binder
					Sleep, 100

					; select No changes saved
					WinWaitActive, Adobe Acrobat
					Sleep, 100
					Send, n
					Sleep, 100
						
					; initialize the counter
					JPGcount = 1
					JPGname = %JPGcount%

					; loop to rename JPEGs
					Loop, %A_LoopFileName%\01_jpg\*.jpg
					{
						; prepend zeros if necessary
						IfLess, JPGcount, 1000
						{
							JPGname = 0%JPGcount%
							IfLess, JPGcount, 100
							{
								JPGname = 0%JPGname%
								IfLess, JPGcount, 10
								{
									JPGname = 0%JPGname%
								}
							}
						}

						; rename the file
						FileMove, %A_LoopFileFullPath%, %A_LoopFileDir%\%JPGname%.jpg

						; increment the counter
						JPGcount++
					}			
				}
				
				; destroy the GUIs
				Gui, 1:Destroy
				Gui, 2:Destroy
				Gui, 3:Destroy	
			}
		}
	}

	MsgBox, Image Processing Complete:`n`n%displaypath%`n`nIssues: %runcount%
Return
; ===================EXPORT JPG

; ===================PDF UPDATE
PDFupdate:
	; create the GUI
	WinGetPos, winX, winY, winWidth, winHeight, 02_pdf
;	winY+=%winHeight%
	Gui, 1:+AlwaysOnTop
	Gui, 1:+ToolWindow
	Gui, 1:Font, cGreen s25 bold, Arial
	Gui, 1:Add, Text, x30 y20 w40 h35, %PDFcount%
	Gui, 1:Font, cBlack s8 norm, Arial
	Gui, 1:Add, Text, x20 y8 w60 h12, %A_LoopFileName%
	Gui, 1:Add, Button, x20 y55 w18 gPDFminus, -
	Gui, 1:Add, Button, x60 y55 gPDFplus, +
	Gui, 1:Show, x%winX% y%winY% h85 w100, PDF Count
Return

PDFplus:
	; increment the PDF count
	PDFcount++
	
	; update the GUI
	ControlSetText, Static1, %PDFcount%, PDF Count
Return

PDFminus:
	; decrement the PDF count
	PDFcount--
	
	; update the GUI
	ControlSetText, Static1, %PDFcount%, PDF Count	
Return
; ===================PDF COUNT

; ===================VOLUME NUMBER
VolumeUpdate:
	; create the GUI
	WinGetPos, winX, winY, winWidth, winHeight, PDF Count
	winY+=%winHeight%
	Gui, 2:+AlwaysOnTop
	Gui, 2:+ToolWindow
	Gui, 2:Font, cGreen s25 bold, Arial
	Gui, 2:Add, Text, x20 y20 w60 h35, %volume%
	Gui, 2:Font, cBlack s8 norm, Arial
	Gui, 2:Add, Text, x20 y8 w60 h12, %A_LoopFileName%
	Gui, 2:Add, Button, x20 y55 w18 gVolumeMinus, -
	Gui, 2:Add, Button, x60 y55 gVolumePlus, +
	Gui, 2:Show, x%winX% y%winY% h85 w100, Volume
Return

VolumePlus:
	; increment the volume number
	volume++
	
	; update the GUI
	ControlSetText, Static1, %volume%, Volume
Return

VolumeMinus:
	; decrement the volume number
	volume--
	
	; update the GUI
	ControlSetText, Static1, %volume%, Volume	
Return
; ===================VOLUME NUMBER

; ===================ISSUE NUMBER
IssueUpdate:
	; create the GUI
	WinGetPos, winX, winY, winWidth, winHeight, PDF Count
	winX+=%winWidth%
	Gui, 3:+AlwaysOnTop
	Gui, 3:+ToolWindow
	Gui, 3:Font, cGreen s25 bold, Arial
	Gui, 3:Add, Text, x20 y20 w60 h35, %issue%
	Gui, 3:Font, cBlack s8 norm, Arial
	Gui, 3:Add, Text, x20 y8 w60 h12, %A_LoopFileName%
	Gui, 3:Add, Button, x20 y55 w18 gIssueMinus, -
	Gui, 3:Add, Button, x60 y55 gIssuePlus, +
	Gui, 3:Show, x%winX% y%winY% h85 w100, Issue
Return

IssuePlus:
	; increment the volume number
	issue++
	
	; update the GUI
	ControlSetText, Static1, %issue%, Issue
Return

IssueMinus:
	; decrement the volume number
	issue--
	
	; update the GUI
	ControlSetText, Static1, %issue%, Issue	
Return
; ===================ISSUE NUMBER

; ===========================================SUBROUTINES
; =================METADATA
IssueFolderPath:
	; if the titlefolderpath variable is empty
	if titlefolderpath = _
	{
		; create dialog to select the title folder
		FileSelectFolder, titlefolderpath, %scannedpath%, 0, Next Issue`n`nSelect the TITLE folder:
		if ErrorLevel
			Return

		; extract title folder name from path
		StringGetPos, foldernamepos, titlefolderpath, \, R1
		foldernamepos++
		StringTrimLeft, titlefoldername, titlefolderpath, foldernamepos
	}

	; store the clipboard contents
	temp = %clipboard%

	; activate the title folder
	SetTitleMatchMode 1
	IfWinNotActive, %titlefoldername%
	{
		WinActivate, %titlefoldername%, , , ,
		WinWaitActive, %titlefoldername%, , , ,
	}
	Sleep, 200

	; grab the folder name & path
	Loop {
		Send, {F2}
		Sleep, 200
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100
		
		if RegExMatch(clipboard, "\d\d\d\d\d\d\d\d\d\d")
		{
			Send, {Enter}
			Sleep, 100
			issuefoldername = %clipboard%
			issuefolderpath = %titlefolderpath%\%issuefoldername%
			Break
		}
	}

	; restore the clipboard contents
	clipboard = %temp%	
Return

MetaHarvest:
	; create display date variables
	month := SubStr(issuefoldername, 5, 2)
	if (month == 01) {
		monthname = Jan.
	}
	else if (month == 02) {
		monthname = Feb.
	}
	else if (month == 03) {
		monthname = Mar.
	}
	else if (month == 04) {
		monthname = Apr.
	}
	else if (month == 05) {
		monthname = May
	}
	else if (month == 06) {
		monthname = June
	}
	else if (month == 07) {
		monthname = July
	}
	else if (month == 08) {
		monthname = Aug.
	}
	else if (month == 09) {
		monthname = Sept.
	}
	else if (month == 10) {
		monthname = Oct.
	}
	else if (month == 11) {
		monthname = Nov.
	}
	else if (month == 12) {
		monthname = Dec.
	}
	day := SubStr(issuefoldername, 7, 2)
	if (SubStr(day, 1, 1) == 0)
	{
		day := SubStr(day, 2)
	}
	year := SubStr(issuefoldername, 1, 4)

	; initialize the loop counter
	metacount = 0
	; initialize the note variable
	note =
	; initialize the displaynote variable
	displaynote =
	
	; if there is a metadata.txt file in the folder
	IfExist, %issuefolderpath%\metadata.txt
	{
		; read in the metadata.txt document
		Loop, read , %issuefolderpath%\metadata.txt
		{
			; increment the counter
			metacount++
				
			; case for volume number
			if (metacount == 1)
			{
				StringTrimLeft, volumenum, A_LoopReadLine, 8
			}
				
			; case for issue number
			else if (metacount == 2)
			{
				StringTrimLeft, issuenum, A_LoopReadLine, 7
			}
				
			; case for the note
			else
			{
				StringTrimLeft, note, A_LoopReadLine, 6
			}
		}
		
		; if the note field is blank
		if (note == "")
			displaynote = `n`n`n`t`t`t       <<< BLANK >>>

		else
		{
			; loop to parse the note
			Loop, Parse, note, ., %A_Space%
			{
				displaynote .= A_LoopField
				
				; special handling variables for notes with periods
				StringGetPos, volumepos, A_LoopField, %volumespecial%, L
				StringGetPos, issuepos, A_LoopField, %issuespecial%, L
				StringGetPos, volumeissuepos, A_LoopField, %volumeissuespecial%, L
				StringLen, notelength, A_LoopField
				
				; if the parsed note is short
				if (notelength <= 8)
				{
					; end the note if it is a number
					if A_LoopField is integer
					{
						displaynote .= .`n`n
					}
					; otherwise continue the note
					else
					{
						displaynote .= "."
						displaynote .= A_Space
						continue
					}
				}
				
				; continue the note if it is the first part of a special case
				else if ((volumepos == 0) || (issuepos == 0) || (volumeissuepos == 0))
				{
					displaynote .= "."
					displaynote .= A_Space
					continue
				}
				
				; end normal notes with a period
				else displaynote .= .`n`n
			}
		}

		WinGetPos, winX, winY, winWidth, winHeight, Greensheet
		winY+=%winHeight%

		; create display windows if necessary
		IfWinNotExist, Note
		{		
			; create Note GUI
			Gui, 8:+AlwaysOnTop
			Gui, 8:+ToolWindow
			Gui, 8:Add, Text, x15 y15 w380 h115, %displaynote%
			Gui, 8:Show, x%winX% y%winY% h125 w400, Note
		}
		IfWinNotExist, Volume
		{
			; create VolumeNum GUI
			Gui, 6:+AlwaysOnTop
			Gui, 6:+ToolWindow
			Gui, 6:Font, cRed s15 bold, Arial
			Gui, 6:Font, cRed s25 bold, Arial
			Gui, 6:Add, Text, x15 y8 w70 h40, %volumenum%
			Gui, 6:Show, x%winX% y%winY% h55 w100, Volume
		}
		IfWinNotExist, Issue
		{		
			; create IssueNum GUI
			Gui, 7:+AlwaysOnTop
			Gui, 7:+ToolWindow
			Gui, 7:Font, cRed s25 bold, Arial
			Gui, 7:Add, Text, x15 y8 w75 h40, %issuenum%
			Gui, 7:Show, x%winX% y%winY% h55 w100, Issue
		}
		IfWinNotExist, Date
		{		
			; create Date GUI
			Gui, 9:+AlwaysOnTop
			Gui, 9:+ToolWindow
			Gui, 9:Font, cRed s15 bold, Arial
			Gui, 9:Add, Text, x35 y15 w160 h25, %monthname% %day%, %year%
			Gui, 9:Font, cRed s10 bold, Arial
			Gui, 9:Show, x%winX% y%winY% h55 w200, Date
		}

		; update the metadata display windows
		ControlSetText, Static1, %volumenum%, Volume
		ControlSetText, Static1, %issuenum%, Issue
		ControlSetText, Static1, %displaynote%, Note
		ControlSetText, Static1, %monthname% %day%`, %year%, Date

		; get ACDSee window ID
		SetTitleMatchMode 2
		WinGet, acdseeid, ID, ACDSee
		
		; activate ACDSee
		SetTitleMatchMode 1
		IfWinNotActive, ahk_id %acdseeid%
		WinActivate, ahk_id %acdseeid%
		WinWaitActive, ahk_id %acdseeid%
		Sleep, 200			
	}

	; if there is no metadata.txt file in the folder
	else
	{
		; create display variables for issue folder path
		StringGetPos, issuefolderpos, issuefolderpath, \, R3
		StringLeft, issuefolder1, issuefolderpath, issuefolderpos
		StringTrimLeft, issuefolder2, issuefolderpath, issuefolderpos		

		; print an error message
		MsgBox, 0, Error, There is no metadata.txt file in this folder.`n`n%issuefolder1%`n%issuefolder2%`n`nUse New Meta (Alt + n) to create a new file.
	}
Return
; =================METADATA
; ===========================================SUBROUTINES

; =================MENUS
GreensheetArchives:
	Run, %Qpath%
Return

GreensheetProcessing:
	Run, http://digitalprojects.library.unt.edu/projects/index.php/Greensheet_Processing
Return

IssueCheckPY:
	if (titlefolderpath = _)
	{
		MsgBox, Please select a folder and try again.
		Return
	}

	; run the issueCheck.py script in the standard Windows terminal
	Run, C:\Windows\System32\cmd.exe /k C:\Python27\python.exe C:\Python27\issueCheck.py "%titlefolderpath%" > "%titlefolderpath%\report-%A_YYYY%-%A_MM%-%A_DD%.txt"
	Sleep, 1000

	; activate terminal window
	SetTitleMatchMode 2
	WinWait, cmd.exe, , , ,
	IfWinNotActive, cmd.exe, , , ,
	WinActivate, cmd.exe, , , ,
	WinWaitActive, cmd.exe, , , ,
	Sleep, 100

	; close the terminal
	Send, exit
	Send, {Enter}
Return

Reload:
Reload

Exit:
ExitApp

Documentation:
	Run, http://digitalprojects.library.unt.edu/projects/index.php/Greensheet_Born_Digital
Return

About:
	MsgBox, Greensheet Born Digital`n`nAndrew.Weidner@unt.edu
Return
; =================MENUS
