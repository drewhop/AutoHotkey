/*
 * TDNP_Metadata.ahk
 *
 * Description: tool for digital newspaper creation and quality control
 */

#NoEnv
#persistent

; ===========================================VARIABLES
; file
scannedpath = Q:\scanned_for_tdnp
collationpath = Q:\TDNP\Collation Sheets

; help
docURL = https://github.com/drewhop/AutoHotkey/wiki/TDNP_Metadata

; edit
titlefolderpath = _
titlefoldername = _

; issue separation
issuefile =

; metadata display
metacount =
issuenum =
volumenum =
note =
displaynote =
volumespecial = Incorrect volume number
issuespecial = Incorrect issue number
volumeissuespecial = Incorrect volume and

; report
today =

; search
searchstring = Do not use quotes.

; scoreboard
issuescore = 0
imagescore = 0
metadatascore = 0
openscore = 0
nextscore = 0
prevscore = 0
gotoscore = 0
displaymetascore = 0
editmetascore = 0
backscore = 0
folderscore = 0
hotkeys = 0
; ===========================================VARIABLES

Gui, 1:Color, d0d0d0, 912206
Gui, 1:Show, h0 w0, TDNP_Metadata

; ===========================================MENUS
; see MENU FUNCTIONS at end of file

; File menu
Menu, FileMenu, Add, &Collation Sheets, CollationSheets
Menu, FileMenu, Add, &scanned_for_tdnp, OpenTDNP
Menu, FileMenu, Add, &Microfilm Scanning Queue, Microfilm
Menu, FileMenu, Add
Menu, FileMenu, Add, Create &Report, IssueCheckPY
Menu, FileMenu, Add
Menu, FileMenu, Add, Reloa&d, Reload
Menu, FileMenu, Add, E&xit, Exit

; Edit menu
Menu, EditMenu, Add, Title Folder &Path, TitleFolderPath
Menu, EditMenu, Add
Menu, EditMenu, Add, &Display Current Folder, DisplayValues

; Issue Separation menu
Menu, IssuesMenu, Add, &Enter Folder Names, EnterFolderNames
Menu, IssuesMenu, Add, &Create Issue Folders, CreateIssueFolders
Menu, IssuesMenu, Add
Menu, IssuesMenu, Add, &Paste Images, PasteImages

; Search menu
Menu, SearchMenu, Add, US Directory: &Texas, DirectoryTitleTexas
Menu, SearchMenu, Add, US Directory: &Oklahoma, DirectoryTitleOklahoma
Menu, SearchMenu, Add, US Directory: &LCCN, DirectoryLCCN
Menu, SearchMenu, Add
Menu, SearchMenu, Add, Portal (TX): Title, PortalTitle
Menu, SearchMenu, Add, Gateway (OK): Title, GatewayTitle

; Help menu
Menu, HelpMenu, Add, &Documentation, Documentation
Menu, HelpMenu, Add, &IssueCheck.py, Report
Menu, HelpMenu, Add, &About, About

; create menus
Menu, MenuBar, Add, &File, :FileMenu
Menu, MenuBar, Add, &Edit, :EditMenu
Menu, MenuBar, Add, &Issue Separation, :IssuesMenu
Menu, MenuBar, Add, &Search, :SearchMenu
Menu, MenuBar, Add, &Help, :HelpMenu

; create menu toolbar
Gui, Menu, MenuBar
; ===========================================MENUS

Gui, 1:Show, h190 w345, TDNP_Metadata
winactivate, TDNP_Metadata

; ===========================================LABELS
; row 1
; Issues mini-counter: static 2
; Images mini-counter: static 4
; Last Hotkey: static 7
Gui, 1:Add, Text, x15  y10 w40 h15, Issues:
Gui, 1:Add, Text, x50  y10 w15 h15, 0
Gui, 1:Add, Text, x95  y10 w40 h15, Images:
Gui, 1:Add, Text, x135 y10 w15 h15, 0
Gui, 1:Add, Text, x175 y10 w80 h15, New Meta ( n )
Gui, 1:Add, Text, x255 y10 w80 h15, Last HotKey
Gui, 1:Add, Text, x265 y35 w60 h15,

; row 2  Static 8-11
Gui, 1:Add, Text, x15  y70 w80 h15, Open ( i )
Gui, 1:Add, Text, x95  y70 w80 h15, Next ( o )
Gui, 1:Add, Text, x175 y70 w80 h15, Previous ( p )
Gui, 1:Add, Text, x255 y70 w80 h15, GoTo ( g )

;row 3  Static 12-15
Gui, 1:Add, Text, x15  y130 w70 h15, Disp Meta ( 0 )
Gui, 1:Add, Text, x95  y130 w70 h15, Edit Meta ( m )
Gui, 1:Add, Text, x175 y130 w70 h15, Back ( b )
Gui, 1:Add, Text, x255 y130 w70 h15, HotKeys
; ===========================================LABELS

; ===========================================DATA
Gui, 1:Font, s15,
; static 16-18
Gui, 1:Add, Text, x25  y30 w55 h25, 0
Gui, 1:Add, Text, x105 y30 w55 h25, 0
Gui, 1:Add, Text, x185 y30 w55 h25, 0
; static 19-22
Gui, 1:Add, Text, x25  y90 w55 h25, 0
Gui, 1:Add, Text, x105 y90 w55 h25, 0
Gui, 1:Add, Text, x185 y90 w55 h25, 0
Gui, 1:Add, Text, x265 y90 w55 h25, 0
; static 23-26
Gui, 1:Add, Text, x25  y150 w45 h25, 0
Gui, 1:Add, Text, x105 y150 w45 h25, 0
Gui, 1:Add, Text, x185 y150 w45 h25, 0
Gui, 1:Add, Text, x265 y150 w45 h25, 0
; ===========================================DATA

; ===========================================BOXES
; decorative box
Gui, 1:Add, GroupBox, x2 y-12 w340 h200,       

; row 1
Gui, 1:Add, GroupBox, x15  y15 w75 h45, 
Gui, 1:Add, GroupBox, x95  y15 w75 h45, 
Gui, 1:Add, GroupBox, x175 y15 w75 h45, 
Gui, 1:Add, GroupBox, x255 y15 w75 h45, 

; row 2
Gui, 1:Add, GroupBox, x15  y75 w75 h45, 
Gui, 1:Add, GroupBox, x95  y75 w75 h45, 
Gui, 1:Add, GroupBox, x175 y75 w75 h45, 
Gui, 1:Add, GroupBox, x255 y75 w75 h45,

;row 3
Gui, 1:Add, GroupBox, x15  y135 w75 h45,         
Gui, 1:Add, GroupBox, x95  y135 w75 h45,           
Gui, 1:Add, GroupBox, x175 y135 w75 h45,           
Gui, 1:Add, GroupBox, x255 y135 w75 h45,            
; ===========================================BOXES

; ===========================================SCRIPTS
; ==============================
; NEW META
; creates new metadata.txt file in currently active issue folder
; Hotkey: Alt + n
!n::
	; harvest the issue folder path
	Gosub, IssueFolderPath

	; if there is no metadata.txt file in the folder
	IfNotExist, %issuefolderpath%\metadata.txt
	{
		; create new text document
		FileAppend, volume:`nissue:`nnote:, %issuefolderpath%\metadata.txt

		; wait two seconds
		Sleep, 2000

		; loop to check for the new metadata.txt document
		Loop
		{
			; if the file does not yet exist
			IfNotExist, %issuefolderpath%\metadata.txt
				; wait another second
				Sleep, 1000
			else
				Break
		}

		; open the metadata.txt file
		Run, metadata.txt, %issuefolderpath%

		; prepare metadata.txt for data entry
		SetTitleMatchMode 1
		WinWait, metadata, , , ,
		IfWinNotActive, metadata, , , ,
		WinActivate, metadata, , , ,
		WinWaitActive, metadata, , , ,
		Sleep, 100
		Send, {Down}{Left}
  
		; update the scoreboard
		hotkeys++
		metadatascore++
		ControlSetText, Static7, NEWMETA, TDNP_Metadata
		ControlSetText, Static18, %metadatascore%, TDNP_Metadata
		ControlSetText, Static26, %hotkeys%, TDNP_Metadata
	Return
	}

	; if there is already a metadata.txt file in the folder
	else
	{
		; create display variables for issue folder path
		StringGetPos, issuefolderpos, issuefolderpath, \, R3
		StringLeft, issuefolder1, issuefolderpath, issuefolderpos
		StringTrimLeft, issuefolder2, issuefolderpath, issuefolderpos		

		; print a message
		MsgBox, 0, New Meta, There is already a metadata.txt file in this folder.`n`n%issuefolder1%`n%issuefolder2%`n`nOpening existing file.

		; open the existing metadata.txt document
		Run, metadata.txt, %issuefolderpath%

		; update the scoreboard
		hotkeys++
		editmetascore++
		ControlSetText, Static7, EDITMETA, TDNP_Metadata
		ControlSetText, Static24, %editmetascore%, TDNP_Metadata
		ControlSetText, Static26, %hotkeys%, TDNP_Metadata
	}
Return
; ==============================

; ==============================
; OPEN
; opens selected issue folder and first TIFF file
; Hotkey: Alt + i
!i::
	; open selected directory
	Send, {Enter}

	; wait for the TDNP issue folder to load
	SetTitleMatchMode RegEx
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100
  
	; open the first TIFF file
	Send, {Down}
	Sleep, 100
	Send, {Up}
	Sleep, 100
	Send, {Enter}

	; update the scoreboard
	hotkeys++
	openscore++
	ControlSetText, Static7, OPEN, TDNP_Metadata
	ControlSetText, Static19, %openscore%, TDNP_Metadata
	ControlSetText, Static26, %hotkeys%, TDNP_Metadata
Return
; ==============================

; ==============================
; OPEN+
; opens selected issue folder and first TIFF file
; and displays issue metadata in moveable windows
; Hotkey: Win + Alt + i
#!i::
	; open selected directory
	Send, {Enter}

	; wait for the TDNP issue folder to load
	SetTitleMatchMode RegEx
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100
  
	; open the first TIFF file
	Send, {Down}
	Sleep, 100
	Send, {Up}
	Sleep, 100
	Send, {Enter}
	Sleep, 500

	; update the scoreboard
	hotkeys++
	openscore++
	ControlSetText, Static7, OPEN+, TDNP_Metadata
	ControlSetText, Static19, %openscore%, TDNP_Metadata
	ControlSetText, Static26, %hotkeys%, TDNP_Metadata

	; metadata harvest subroutine
	Gosub, MetaHarvest
Return
; ==============================
  
; ==============================
; NEXT
; opens first TIFF file in next issue
; Hotkey: Alt + o
!o::
	; navigation subroutine
	Gosub, NextIssue

	; update the scoreboard
	hotkeys++
	nextscore++
	ControlSetText, Static7, NEXT, TDNP_Metadata
	ControlSetText, Static20, %nextscore%, TDNP_Metadata
	ControlSetText, Static26, %hotkeys%, TDNP_Metadata
Return
; ==============================

; ==============================
; NEXT+
; opens first TIFF file in next issue
; and displays issue metadata in moveable windows
; Hotkey: Win + Alt + o
#!o::
	; navigation subroutine
	Gosub, NextIssue
	Sleep, 500

	; update the scoreboard
	hotkeys++
	nextscore++
	ControlSetText, Static7, NEXT+, TDNP_Metadata
	ControlSetText, Static20, %nextscore%, TDNP_Metadata
	ControlSetText, Static26, %hotkeys%, TDNP_Metadata

	; metadata harvest subroutine
	Gosub, MetaHarvest
Return
; ==============================

; ==============================
; PREVIOUS
; opens first TIFF file in previous issue
; Hotkey: Alt + p
!p::
	; navigation subroutine
	Gosub, PreviousIssue

	; update the scoreboard
	hotkeys++
	prevscore++
	ControlSetText, Static7, PREVIOUS, TDNP_Metadata
	ControlSetText, Static21, %prevscore%, TDNP_Metadata
	ControlSetText, Static26, %hotkeys%, TDNP_Metadata
Return
; ==============================

; ==============================
; PREVIOUS+
; opens first TIFF file in previous issue
; and displays issue metadata in moveable windows
; Hotkey: Win + Alt + p
#!p::
	; navigation subroutine
	Gosub, PreviousIssue
	Sleep, 500
	
	; update the scoreboard
	hotkeys++
	prevscore++
	ControlSetText, Static7, PREV+, TDNP_Metadata
	ControlSetText, Static21, %prevscore%, TDNP_Metadata
	ControlSetText, Static26, %hotkeys%, TDNP_Metadata

	; metadata harvest subroutine
	Gosub, MetaHarvest	
Return
; ==============================

; ==============================
; GoTo
; opens a specific issue folder and first TIFF file
; Hotkey: Alt + g
!g::
	; navigation subroutine
	Gosub, GoToIssue

	; update scoreboard
	hotkeys++
	gotoscore++
	ControlSetText, Static7, GOTO, TDNP_Metadata									
	ControlSetText, Static22, %gotoscore%, TDNP_Metadata									
	ControlSetText, Static26, %hotkeys%, TDNP_Metadata									
Return
; ==============================

; ==============================
; GoTo+
; opens a specific issue folder and first TIFF file
; and displays issue metadata in moveable windows
; Hotkey: Win + Alt + g
#!g::
	; navigation subroutine
	Gosub, GoToIssue
			
	; update scoreboard
	hotkeys++
	gotoscore++
	ControlSetText, Static7, GOTO+, TDNP_Metadata									
	ControlSetText, Static22, %gotoscore%, TDNP_Metadata									
	ControlSetText, Static26, %hotkeys%, TDNP_Metadata									

	; metadata harvest subroutine
	Gosub, MetaHarvest
Return
; ==============================

; ==============================
; DISPLAY META
; display the issue metadata
; Hotkey: Alt + 0
!0::
	; metadata harvest subroutine
	Gosub, MetaHarvest

	; update the scoreboard
	ControlSetText, Static7, DISPMETA, TDNP_Metadata
Return
; ==============================

; ==============================
; EDIT META
; opens the metadata.txt file for editing
; Hotkey: Alt + m
!m::
	; harvest the issue folder path
	Gosub, IssueFolderPath
	
	; if there is a metadata.txt file in the folder
	IfExist, %issuefolderpath%\metadata.txt
	{
		; open the metadata.txt document
		Run, metadata.txt, %issuefolderpath%

		; update the scoreboard
		hotkeys++
		editmetascore++
		ControlSetText, Static7, EDITMETA, TDNP_Metadata
		ControlSetText, Static24, %editmetascore%, TDNP_Metadata
		ControlSetText, Static26, %hotkeys%, TDNP_Metadata
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

; ==============================
; BACK
; opens the title folder with titlefolderpath variable
; Hotkey: Alt + b
!b::
	; if the titlefolderpath variable is empty
	if titlefolderpath = _
	{
		; create dialog to select the title folder
		FileSelectFolder, titlefolderpath, %scannedpath%, 0, BACK`n`nSelect the TITLE folder:
		if ErrorLevel
			Return

		; extract title folder name from path
		StringGetPos, foldernamepos, titlefolderpath, \, R1
		foldernamepos++
		StringTrimLeft, titlefoldername, titlefolderpath, foldernamepos
	}

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

	; open the title folder
	Run, %titlefolderpath%

	; update scoreboard
	hotkeys++
	backscore++
	ControlSetText, Static7, BACK, TDNP_Metadata
	ControlSetText, Static25, %backscore%, TDNP_Metadata
	ControlSetText, Static26, %hotkeys%, TDNP_Metadata
Return
; ==============================

; ==============================
; NewFolder
; creates new folder
; Hotkey: Win + n
#n::
	; exit and re-enter folder to clear any selections
	; up one directory
	Send, {AltDown}v{AltUp}
	Sleep, 100
	Send, g
	Sleep, 100
	Send, u
	Sleep, 400
	Send, {Enter}
	Sleep, 400

	; create new folder with right click menu
	Send, {AppsKey}
	Sleep, 200
	Send, w
	Sleep, 200
	Send, f

	; update the scoreboard
	folderscore++
	ControlSetText, Static7, Folder %folderscore%, TDNP_Metadata
Return
; ==============================
; ===========================================SCRIPTS

; ===========================================SUBROUTINES
; =================METADATA
IssueFolderPath:
	; save clipboard contents
	temp = %clipboard%
	
	; activate the current TDNP issue folder
	SetTitleMatchMode RegEx
	WinWait, ^[1-2][0-9]{9}$, , , ,
	IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
	WinActivate, ^[1-2][0-9]{9}$, , , ,
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100
	Send, {Tab 5}

	; copy issue folder path to clipboard
	Send, {F4}
	Sleep, 100
	Send, {CtrlDown}a{CtrlUp}
	Sleep, 100
	Send, {CtrlDown}c{CtrlUp}
	Sleep, 100
	Send, {Enter}
	Sleep, 100

	; grab the issue date from the folder path
	StringRight, date, clipboard, 10
	
	; loop checks for correctly copied folder path
	; up to 5 times, aborts script if unsuccessful
	Loop 5
	{
		; continue the script if the path copied to clipboard
		if RegExMatch(date, "\d\d\d\d\d\d\d\d\d\d")
			Break
			
		; or reattempt to copy folder path
		else
		{
			; wait one second
			Sleep, 1000
			
			; reactivate the issue folder
			SetTitleMatchMode RegEx
			WinWait, ^[1-2][0-9]{9}$, , , ,
			IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
			WinActivate, ^[1-2][0-9]{9}$, , , ,
			WinWaitActive, ^[1-2][0-9]{9}$, , , ,
			Sleep, 100
			Send, {Tab 5}

			; copy issue folder path to clipboard
			Send, {F4}
			Sleep, 100
			Send, {CtrlDown}a{CtrlUp}
			Sleep, 100
			Send, {CtrlDown}c{CtrlUp}
			Sleep, 100
			Send, {Enter}
			Sleep, 100
			
			; grab the issue date from the folder path
			StringRight, date, clipboard, 10
			
			; next loop iteration
			Continue
		}
		
		; restore clipboard contents
		clipboard = %temp%		

		; abort script if unable to copy path after 5 attempts
		MsgBox, 0, Error, Unable to copy file path.`n`nScript aborted.
		Exit
	}
	
	; assign the folder path
	issuefolderpath = %clipboard%
	
	; restore clipboard contents
	clipboard = %temp%		
Return

MetaHarvest:
	; harvest the issue folder path
	Gosub, IssueFolderPath
	
	; create display date variables
	month := SubStr(date, 5, 2)
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
	day := SubStr(date, 7, 2)
	if (SubStr(day, 1, 1) == 0)
	{
		day := SubStr(day, 2)
	}
	year := SubStr(date, 1, 4)

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

		; create display windows if first metadata display run
		if (displaymetascore == 0)
		{
			WinGetPos, winX, winY, winWidth, winHeight, Metadata
			winY+=%winHeight%

			; create VolumeNum GUI
			Gui, 2:+AlwaysOnTop
			Gui, 2:+ToolWindow
			Gui, 2:Font, cRed s15 bold, Arial
			Gui, 2:Font, cRed s25 bold, Arial
			Gui, 2:Add, GroupBox, x0 y-18 w100 h74,       
			Gui, 2:Add, Text, x15 y8 w70 h40, %volumenum%
			Gui, 2:Show, x%winX% y%winY% h55 w100, Volume

			; create IssueNum GUI
			Gui, 3:+AlwaysOnTop
			Gui, 3:+ToolWindow
			Gui, 3:Font, cRed s25 bold, Arial
			Gui, 3:Add, GroupBox, x0 y-18 w100 h74,       
			Gui, 3:Add, Text, x15 y8 w70 h40, %issuenum%
			Gui, 3:Show, x%winX% y%winY% h55 w100, Issue

			; create Note GUI
			Gui, 4:+AlwaysOnTop
			Gui, 4:+ToolWindow
			Gui, 4:Add, Text, x15 y15 w380 h115, %displaynote%
			Gui, 4:Show, x%winX% y%winY% h125 w400, Note
			
			; create Date GUI
			Gui, 5:+AlwaysOnTop
			Gui, 5:+ToolWindow
			Gui, 5:Font, cRed s15 bold, Arial
			Gui, 5:Add, Text, x35 y15 w160 h25, %monthname% %day%, %year%
			Gui, 5:Font, cRed s10 bold, Arial
			Gui, 5:Add, GroupBox, x0 y-7 w200 h63,       
			Gui, 5:Show, x%winX% y%winY% h55 w200, Date
		}

		; update the metadata display windows
		else
		{
			ControlSetText, Static1, %volumenum%, Volume
			ControlSetText, Static1, %issuenum%, Issue
			ControlSetText, Static1, %displaynote%, Note
			ControlSetText, Static1, %monthname% %day%`, %year%, Date
		}

		; update the scoreboard
		hotkeys++
		displaymetascore++
		ControlSetText, Static23, %displaymetascore%, TDNP_Metadata									
		ControlSetText, Static26, %hotkeys%, TDNP_Metadata

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

; =================NAVIGATION
NextIssue:
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

	; activate the current TDNP issue folder
	SetTitleMatchMode RegEx
	WinWait, ^[1-2][0-9]{9}$, , , ,
	IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
	WinActivate, ^[1-2][0-9]{9}$, , , ,
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100
	Send, {Tab 5}

	; up one directory
	Send, {AltDown}v{AltUp}
	Sleep, 100
	Send, g
	Sleep, 100
	Send, u
	Sleep, 100

	; activate the title folder
	SetTitleMatchMode 1
	WinWait, %titlefoldername%, , , ,
	IfWinNotActive, %titlefoldername%, , , ,
	WinActivate, %titlefoldername%, , , ,
	WinWaitActive, %titlefoldername%, , , ,
	Sleep, 100

	; open the next folder
	Send, {Down}
	Sleep, 100
	Send, {Enter}

	; reactivate the issue folder
	SetTitleMatchMode RegEx
	WinWait, ^[1-2][0-9]{9}$, , , ,
	IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
	WinActivate, ^[1-2][0-9]{9}$, , , ,
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100
  
	; open the first TIFF file
	Send, {Down}
	Sleep, 100
	Send, {Up}
	Sleep, 100
	Send, {Enter}
	Sleep, 500
Return

PreviousIssue:
	; if the titlefolderpath variable is empty
	if titlefolderpath = _
	{
		; create dialog to select the title folder
		FileSelectFolder, titlefolderpath, %scannedpath%, 0, Previous Issue`n`nSelect the TITLE folder:
		if ErrorLevel
			Return

		; extract title folder name from path
		StringGetPos, foldernamepos, titlefolderpath, \, R1
		foldernamepos++
		StringTrimLeft, titlefoldername, titlefolderpath, foldernamepos
	}

	; activate the current TDNP issue folder
	SetTitleMatchMode RegEx
	WinWait, ^[1-2][0-9]{9}$, , , ,
	IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
	WinActivate, ^[1-2][0-9]{9}$, , , ,
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100

	; up one directory
	Send, {AltDown}v{AltUp}
	Sleep, 100
	Send, g
	Sleep, 100
	Send, u
	Sleep, 100

	; activate the title folder
	SetTitleMatchMode 1
	WinWait, %titlefoldername%, , , ,
	IfWinNotActive, %titlefoldername%, , , ,
	WinActivate, %titlefoldername%, , , ,
	WinWaitActive, %titlefoldername%, , , ,
	Sleep, 100
  
	; open the previous folder
	Send, {Up}
	Sleep, 100
	Send, {Enter}

	; activate the next issue folder
	SetTitleMatchMode RegEx
	WinWait, ^[1-2][0-9]{9}$, , , ,
	IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
	WinActivate, ^[1-2][0-9]{9}$, , , ,
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100

	; open the first TIFF file
	Send, {Down}
	Sleep, 100
	Send, {Up}
	Sleep, 100
	Send, {Enter}
Return

GoToIssue:
	; if the titlefolderpath variable is empty
	if titlefolderpath = _
	{
		; create dialog to select the title folder
		FileSelectFolder, titlefolderpath, %scannedpath%, 0, GoTo Issue`n`nSelect the TITLE folder:
		if ErrorLevel
			Return

		; extract title folder name from path
		StringGetPos, foldernamepos, titlefolderpath, \, R1
		foldernamepos++
		StringTrimLeft, titlefoldername, titlefolderpath, foldernamepos
	}

	; create input box to enter the folder name to go to
	InputBox, input, GoTo Issue, Issue Folder Name:,, 150, 125,,,,,%issue%
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
					issue = %input%

					; case for active issue folder
					SetTitleMatchMode RegEx
					IfWinExist, ^[1-2][0-9]{9}$
					{
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

						; open the issue folder
						Run, %titlefolderpath%\%issue%
						
						; wait for issue folder to load
						WinWaitActive, ^[1-2][0-9]{9}$, , , ,
						; if the folder does not exist
						IfWinActive, , Windows can't find, , , ,
						{
							; exit the script
							Return
						}
						Sleep, 100

						; open the first TIFF file
						Send, {Down}
						Sleep, 100
						Send, {Up}
						Sleep, 100
						Send, {Enter}

						; end the title folder name check loop
						Break
					}
					
					; case for active title folder
					IfWinNotExist, ^[1-2][0-9]{9}$
					SetTitleMatchMode 2
					IfWinExist, %titlefoldername%
					{
						; activate title folder
						WinWait, %titlefoldername%, , , ,
						IfWinNotActive, %titlefoldername%, , , ,
						WinActivate, %titlefoldername%, , , ,
						WinWaitActive, %titlefoldername%, , , ,
						Sleep, 100

						; open the issue
						SetKeyDelay, 150
						Send, %issue%
						SetKeyDelay, 10
						Sleep, 100
						Send, {Enter}
						
						; wait for issue folder to load
						SetTitleMatchMode RegEx
						WinWaitActive, ^[1-2][0-9]{9}$, , , ,
						Sleep, 100

						; open the first TIFF file
						Send, {Down}
						Sleep, 100
						Send, {Up}
						Sleep, 100
						Send, {Enter}

						; end the title folder name check loop
						Break
					}
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
; =================NAVIGATION
; ===========================================SUBROUTINES

; ===========================================MENU FUNCTIONS
; =================FILE
; open the scanned_for_tdnp folder
OpenTDNP:
	Run, %scannedpath%
Return

CollationSheets:
	Run, %collationpath%
Return

Microfilm:
	Run, http://digitalprojects.library.unt.edu/projects/index.php/Microfilm_Scanning_Queue
Return

; create a report
IssueCheckPY:
	; set scoreboard hotkey
	ControlSetText, Static7, REPORT, TDNP_Metadata
	
	; create dialog to select the title folder
	FileSelectFolder, titlefolderpath, %scannedpath%, 0, Create Report`n`nSelect the TITLE folder:
		if ErrorLevel
			Return
		else
		{
			; create input box for report date, auto-populate with today's date
			InputBox, input, Create Report, Enter report date:,, 150, 125,,,,,%A_YYYY%-%A_MM%-%A_DD%
			if ErrorLevel
				Return
			else
			{
				; loop to validate correct report name format: YYYY-MM-DD
				Loop
				{
					; case for valid report name
					if RegExMatch(input, "\d\d\d\d-\d\d-\d\d")
					{
						today = %input%
						
						; run the issueCheck.py script in the standard Windows terminal
						Run, C:\Windows\System32\cmd.exe /k C:\Python27\python.exe C:\Python27\issueCheck.py "%titlefolderpath%" > "%titlefolderpath%\report-%today%.txt"
						Sleep, 1000
						
						; end the loop
						Break
					}
					
					; case for invalid report name
					else
					{
						; print error message
						MsgBox, 0, Create Report, Please enter a date in the format: YYYY-MM-DD`n`nExample: 2013-01-31
						
						; create new report date input box and re-enter loop
						InputBox, input, Create Report, Enter today's date:,, 150, 125,,,,,%A_YYYY%-%A_MM%-%A_DD%
						if ErrorLevel
							Return
					}
				}
				
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
			}
		}
Return

; reload application
Reload:
Reload

; exit application
Exit:
ExitApp
; =================FILE

; =================EDIT
; titlefolderpath variable input
TitleFolderPath:
FileSelectFolder, titlefolderpath, %scannedpath%, 0, Edit Folder Path`n`nSelect the TITLE folder:
	if ErrorLevel
		Return

	; extract title folder name from path
	StringGetPos, foldernamepos, titlefolderpath, \, R1
	foldernamepos++
	StringTrimLeft, titlefoldername, titlefolderpath, foldernamepos
Return

DisplayValues:
if titlefolderpath = _
{
	MsgBox, 0, Current Folder, Title Folder Name: %titlefoldername%`n`nTitle Folder Path: %titlefolderpath%
}
else
{
	; create display variables for title folder path
	StringGetPos, titlefolderpos, titlefolderpath, \, R2
	StringLeft, titlefolder1, titlefolderpath, titlefolderpos
	StringTrimLeft, titlefolder2, titlefolderpath, titlefolderpos

	MsgBox, 0, Current Folder, Title Folder Name:`n`n`t%titlefoldername%`n`nTitle Folder Path:`n`n`t%titlefolder1%`n`t%titlefolder2%
}
Return
; =================EDIT

; =================ISSUE SEPARATION
EnterFolderNames:
	; initialize variables
	issuefile =
	previous =
	input =
	count = 1
	currentissuenum = 0
	ControlSetText, Static7, ISSUES, TDNP_Metadata
	ControlSetText, Static2, %currentissuenum%, TDNP_Metadata
	currentimagescore = 0
	ControlSetText, Static4, %currentimagescore%, TDNP_Metadata

	; loop to accept issue folder names
	; and store in the issuefile variable
	; Cancel button exits data entry
	Loop
	{
		; set the input box display variable
		previous := input
		InputBox, input, Folder Names, Enter Issue %count%,, 150, 125,,,,,%previous%
		if ErrorLevel
			Return
		else
		{
			; loop to validate issue folder names
			Loop
			{
				; case for valid issue folder name
				if RegExMatch(input, "^\d\d\d\d\d\d\d\d\d\d$")
				{
					; append folder name to issuefile
					issuefile .= input
					
					; append new line to issuefile
					issuefile .= "`n"
					
					; update the counters
					count++
					currentissuenum++

					; update the scoreboard
					ControlSetText, Static2, %currentissuenum%, TDNP_Metadata
					
					; end the loop
					Break
				}
				
				; case for invalid issue folder name
				else
				{
					; print an error message
					MsgBox, 0, Enter Folder Names, Please enter the folder name in the format: YYYYMMDDEE`n`nExample: 1885013101
					
					; create another issue folder name input box and re-enter loop
					InputBox, input, Folder Names, Enter issue number %count%,, 150, 125,,,,,%previous%
					if ErrorLevel
						Return
				}
			}
		}
	}
Return

CreateIssueFolders:
	; abort script if issuefile is empty
	if issuefile =
	{
		MsgBox, 0, Create Folders, There are no issues entered.`n`nSelect "File > Enter Folder Names" to enter issues. 
		Return
	}
	else
	{
		; select title folder
		FileSelectFolder, titlefolderpath, %scannedpath%, 1, CREATE FOLDERS`n`nSelect the TITLE folder:
		if ErrorLevel
			Return

		; make the title folder the working directory
		SetWorkingDir %titlefolderpath%
		
		; create display variables for title folder path
		StringGetPos, titlefolderpos, titlefolderpath, \, R2
		StringLeft, titlefolder1, titlefolderpath, titlefolderpos
		StringTrimLeft, titlefolder2, titlefolderpath, titlefolderpos		
		
		; parse the issue file for first and last folders
		createfoldercount = 0
		Loop, parse, issuefile, `n
		{
			createfoldercount++
			if (createfoldercount == 1)
				firstfolder := A_LoopField
			else if (createfoldercount == currentissuenum)
				lastfolder := A_LoopField				
		}
		
		; confirm directory and folders to create
		MsgBox, 4, Create Issue Folders, %titlefolder1%`n%titlefolder2%`n`n`tFirst:`t%firstfolder%`n`n`tLast:`t%lastfolder%`n`n`t`tCreate %currentissuenum% folders?`n`nYes to Continue`nNo to Exit
		IfMsgBox, No, Return
		
		; loop creates issue folders stored in issuefile
		Loop, parse, issuefile, `n
		{
			FileCreateDir, %A_LoopField%
		}
		
		; update the scoreboard
		issuescore += %currentissuenum%
		ControlSetText, Static16, %issuescore%, TDNP_Metadata
	}

	; update scoreboard last hotkey
	ControlSetText, Static7, CREATE, TDNP_Metadata
Return

PasteImages:
	; abort script if issuefile is empty
	if issuefile =
	{
		MsgBox, 0, Paste Images, There are no issues entered.`n`nSelect "File > Enter Folder Names" to enter issues. 
		Return
	}
	else
	{
		; set last hotkey value
		ControlSetText, Static7, IMAGES, TDNP_Metadata

		; initialize micro-counter
		currentimagescore = 0
		ControlSetText, Static4, %currentimagescore%, TDNP_Metadata

		; select title folder
		FileSelectFolder, titlefolderpath, %scannedpath%, 0, PASTE IMAGES`n`nSelect the TITLE folder:
		if ErrorLevel
			Return
		else
		{
			; create display variables for title folder path
			StringGetPos, titlefolderpos, titlefolderpath, \, R2
			StringLeft, titlefolder1, titlefolderpath, titlefolderpos
			StringTrimLeft, titlefolder2, titlefolderpath, titlefolderpos
		
			; loop pastes clipboard contents to folders stored in issuefile
			Loop, parse, issuefile, `n
			{
				; exit script if no more issues
				if A_LoopField =
				{
					Return
				}
				else
				{
					; create display date variables
					month := SubStr(A_LoopField, 5, 2)
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
					day := SubStr(A_LoopField, 7, 2)
					if (SubStr(day, 1, 1) == 0)
					{
						day := SubStr(day, 2)
					}
					year := SubStr(A_LoopField, 1, 4)

					; create a confirm dialog for each paste operation, No cancels the script
					MsgBox, 4, Paste Images, %titlefolder1%`n%titlefolder2%`n`n`t`t%A_LoopField%`n`n`t`t%monthname% %day%`, %year%`n`nYes to Continue`nNo to Exit
					IfMsgBox, No, Return
					else
					{
						; set the path to paste the clipboard contents
						pastepath = %titlefolderpath%\%A_LoopField%
				
						; loop to parse and paste the clipboard
						Loop, parse, clipboard, `n, `r
						{
							FileMove, %A_LoopField%, %pastepath%
						}

						; update scoreboard
						imagescore++
						currentimagescore++
						ControlSetText, Static4, %currentimagescore%, TDNP_Metadata
						ControlSetText, Static17, %imagescore%, TDNP_Metadata
					}
				}
			}
		}
	}
Return
; =================ISSUE SEPARATION

; =================SEARCH
; US Newspaper Directory Texas search
DirectoryTitleTexas:
	InputBox, searchstring, Texas Search,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, http://chroniclingamerica.loc.gov/search/titles/results/?state=Texas&county=&city=&year1=1690&year2=2013&terms=%searchstring%&frequency=&language=&ethnicity=&labor=&material_type=&lccn=&rows=20
Return

; US Newspaper Directory Oklahoma search
DirectoryTitleOklahoma:
	InputBox, searchstring, Oklahoma Search,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, http://chroniclingamerica.loc.gov/search/titles/results/?state=Oklahoma&county=&city=&year1=1690&year2=2013&terms=%searchstring%&frequency=&language=&ethnicity=&labor=&material_type=&lccn=&rows=20
Return

; US Newspaper Directory LCCN search
DirectoryLCCN:
	InputBox, searchstring, LCCN Search,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, http://chroniclingamerica.loc.gov/search/titles/results/?state=&county=&city=&year1=1690&year2=2013&terms=&frequency=&language=&ethnicity=&labor=&material_type=&lccn=%searchstring%&rows=20
Return

; Portal to TX History search: Newspaper & Title
PortalTitle:
	InputBox, searchstring, Portal Title Search,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, http://texashistory.unt.edu/search/?q=%searchstring%&t=dc_title&fq=dc_type`%3Atext_newspaper
Return

; Gateway to OK History search: Newspaper & Title
GatewayTitle:
	InputBox, searchstring, Gateway Title Search,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, http://gateway.okhistory.org/search/?q=%searchstring%&t=dc_title&fq=&fq=&fq=dc_type`%3Atext_newspaper
Return
; =================SEARCH

; =================HELP
; load software documentation
Documentation:
	Run, %docURL%
Return

; open IssueCheck.py wiki page
Report:
	Run, http://digitalprojects.library.unt.edu/projects/index.php/TDNP_IssueCheck.py
Return

; display version info
About:
MsgBox, 0, About, TDNP_Metadata.ahk`nVersion 2.6`nAndrew.Weidner@unt.edu
Return
; =================HELP
; ===========================================MENU FUNCTIONS

GuiClose:
ExitApp
