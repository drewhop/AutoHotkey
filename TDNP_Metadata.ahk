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
issue =
temp =
count =

; metadata display
metacount =
issuenum =
volumenum =
note =
displaynote =
volumespecial = Incorrect volume number
issuespecial = Incorrect issue number
volumeissuespecial = Incorrect volume and
metaopenscore = 0
metanextscore = 0
metaprevscore = 0
metagotoscore = 0
metadisplayscore = 0

; navigation
activegoto = 0

; report
today =

; search
searchstring = Do not use quotes.

; scoreboard
folderscore = 0
imagescore = 0
metadatascore = 0
openscore = 0
nextscore = 0
prevscore = 0
editmetascore = 0
backscore = 0
gotoscore = 0
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
Menu, EditMenu, Add, Title Folder &Name, TitleFolderName
Menu, EditMenu, Add, Title Folder &Path, TitleFolderPath
Menu, EditMenu, Add
Menu, EditMenu, Add, &Display Current Values, DisplayValues

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

Gui, 1:Show, h130 w345, TDNP_Metadata
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
Gui, 1:Add, Text, x255 y10 w50 h15, HotKey
Gui, 1:Add, Text, x265 y35 w60 h15,

; row 2
Gui, 1:Add, Text, x15  y70 w80 h15, Open ( i )
Gui, 1:Add, Text, x95  y70 w80 h15, Next ( o )
Gui, 1:Add, Text, x175 y70 w80 h15, Previous ( p )
Gui, 1:Add, Text, x255 y70 w80 h15, Edit Meta ( m )
; ===========================================LABELS

; ===========================================DATA
Gui, 1:Font, s15,
; static 12-14
Gui, 1:Add, Text, x25  y30 w55 h25, 0
Gui, 1:Add, Text, x105 y30 w55 h25, 0
Gui, 1:Add, Text, x185 y30 w55 h25, 0
; static 15-18
Gui, 1:Add, Text, x25  y90 w55 h25, 0
Gui, 1:Add, Text, x105 y90 w55 h25, 0
Gui, 1:Add, Text, x185 y90 w55 h25, 0
Gui, 1:Add, Text, x265 y90 w55 h25, 0
; ===========================================DATA

; ===========================================BOXES
; decorative box
Gui, 1:Add, GroupBox, x2 y-12 w340 h140,       

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
; ===========================================BOXES

; ===========================================SCRIPTS
; ==============================
; BACK
; opens the title folder with titlefolderpath variable
; Hotkey: Alt + b
!b::
	; if titlefolderpath variable is empty
	if titlefolderpath = _
	{
		; create dialog to select the title folder
		FileSelectFolder, titlefolderpath, %scannedpath%, 0, Back`n`nSelect the TITLE folder:
		if ErrorLevel
			Return
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
	backscore++
	ControlSetText, Static7, Back: %backscore%, TDNP_Metadata
Return
; ==============================

; ==============================
; GoTo
; opens a specific issue folder and first TIFF file
; Hotkey: Alt + g
!g::	
	; if titlefoldername variable is empty
	if titlefoldername = _
	{
		; create input box to enter the title folder name
		InputBox, input, GoTo Issue, Enter Title Folder Name:,, 225, 125,,,,,
			if ErrorLevel
				Return
			else
			{
				titlefoldername = %input%
			}
	}
	
	; if titlefolderpath variable is empty
	if titlefolderpath = _
	{
		; create dialog to select the title folder
		FileSelectFolder, titlefolderpath, %scannedpath%, 0, GoTo`n`nSelect the TITLE folder:
		if ErrorLevel
			Return
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

						; set the GoTo indicator to TRUE (1)
						activegoto = 1

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

						; set the GoTo indicator to TRUE (1)
						activegoto = 1
						
						; end the title folder name check loop
						Break
					}
				}
				
				; if the issue folder name entered is not valid
				; print error message and re-enter loop
				else
				{
					MsgBox, Please enter a folder name in the format: YYYYMMDDEE`n`nExample: 1942061901
					InputBox, input, GoTo Issue, Issue Folder Name:,, 150, 125,,,,,
						if ErrorLevel
							Return
				}
			}
			
			; update scoreboard
			gotoscore++
			ControlSetText, Static7, GoTo: %gotoscore%, TDNP_Metadata									
		}	
Return
; ==============================

; ==============================
; GoTo+
; opens a specific issue folder and first TIFF file
; and displays issue metadata in moveable windows
; Hotkey: Win + Alt + g
#!g::	
	; if titlefoldername variable is empty
	if titlefoldername = _
	{
		; create input box to enter the title folder name
		InputBox, input, GoTo Issue, Enter Title Folder Name:,, 225, 125,,,,,
			if ErrorLevel
				Return
			else
			{
				titlefoldername = %input%
			}
	}
	
	; if titlefolderpath variable is empty
	if titlefolderpath = _
	{
		; create dialog to select the title folder
		FileSelectFolder, titlefolderpath, %scannedpath%, 0, GoTo`n`nSelect the TITLE folder:
		if ErrorLevel
			Return
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
						Sleep, 500

						; set the GoTo indicator to TRUE (1)
						activegoto = 1

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
						Sleep, 500

						; set the GoTo indicator to TRUE (1)
						activegoto = 1
												
						; end the title folder name check loop
						Break
					}
				}
				
				; if the issue folder name entered is not valid
				; print error message and re-enter loop
				else
				{
					MsgBox, Please enter a folder name in the format: YYYYMMDDEE`n`nExample: 1942061901
					InputBox, input, GoTo Issue, Issue Folder Name:,, 150, 125,,,,,
						if ErrorLevel
							Return
				}
			}

			; metadata harvest subroutine
			Gosub, MetaHarvest
			
			; update scoreboard
			gotoscore++
			metagotoscore++
			ControlSetText, Static7, GoTo: %gotoscore%, TDNP_Metadata									
		}	
Return
; ==============================

; ==============================
; NewMetaTXT
; creates new metadata.txt file in currently active issue folder
; Hotkey: Alt + n
!n::
	; in case issue folder is already active
	; TAB five times so the script works
	SetTitleMatchMode RegEx
	IfWinActive, ^[1-2][0-9]{9}$, , , ,
	Send, {Tab 5}

	; activate issue folder
	SetTitleMatchMode RegEx
	WinWait, ^[1-2][0-9]{9}$, , , ,
	IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
	WinActivate, ^[1-2][0-9]{9}$, , , ,
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100

	; save clipboard contents
	temp = %clipboard%

	; copy issue folder path to clipboard
	Send, {F4}
	Sleep, 100
	Send, {CtrlDown}a{CtrlUp}
	Sleep, 100
	Send, {CtrlDown}c{CtrlUp}
	Sleep, 100
	Send, {Enter}
	Sleep, 100

	; if there is no metadata.txt file in the folder
	IfNotExist, %clipboard%\metadata.txt
	{
		; create new text document
		FileAppend, volume:`nissue:`nnote:, %clipboard%\metadata.txt

		; wait two seconds
		Sleep, 2000

		; loop to check for the new metadata.txt document
		Loop
		{
			; if the file does not yet exist
			IfNotExist, %clipboard%\metadata.txt
				; wait another second
				Sleep, 1000
			else
				Break
		}

		; open the metadata.txt file
		Run, metadata.txt, %clipboard%

		; prepare metadata.txt for data entry
		SetTitleMatchMode 1
		WinWait, metadata, , , ,
		IfWinNotActive, metadata, , , ,
		WinActivate, metadata, , , ,
		WinWaitActive, metadata, , , ,
		Sleep, 100
		Send, {Down}{Left}
  
		; restore clipboard contents
		clipboard = %temp%

		; update the scoreboard
		metadatascore++
		ControlSetText, Static7, NEWMETA, TDNP_Metadata
		ControlSetText, Static14, %metadatascore%, TDNP_Metadata
	Return
	}
  
	; if there is already a metadata.txt file in the folder
	else
	{
		; print a message
		MsgBox, There is already a metadata.txt file in this folder.`n`n%clipboard%`n`nOpening existing file.

		; open the existing metadata.txt document
		Run, metadata.txt, %clipboard%

		; restore clipboard contents
		clipboard = %temp%
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
	openscore++
	ControlSetText, Static7, OPEN, TDNP_Metadata
	ControlSetText, Static15, %openscore%, TDNP_Metadata
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

	; metadata harvest subroutine
	Gosub, MetaHarvest

	; update the scoreboard
	openscore++
	metaopenscore++
	ControlSetText, Static7, OPEN+, TDNP_Metadata
	ControlSetText, Static15, %openscore%, TDNP_Metadata
Return
; ==============================
  
; ==============================
; NEXT
; opens first TIFF file in next issue
; Hotkey: Alt + o
!o::
	; if the titlefoldername variable is empty
	if titlefoldername = _
	{
		; create an input box to enter the title folder name
		InputBox, input, Next Issue, Enter Title Folder Name:,, 225, 125,,,,,
			if ErrorLevel
				Return
			else
			{
				titlefoldername = %input%
			}  
	}

	; case for TRUE (1) GoTo indicator
	if activegoto = 1
	{
		; if the titlefolderpath variable is empty
		if titlefolderpath = _
		{
			; create dialog to select the title folder
			FileSelectFolder, titlefolderpath, %scannedpath%, 0, Next Issue`n`nSelect the TITLE folder:
			if ErrorLevel
				Return
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

		; activate title folder
		SetTitleMatchMode 2
		WinWaitActive, %titlefoldername%, , , ,
		Sleep, 100

		; open the next issue
		SetKeyDelay, 150
		Send, %issue%
		SetKeyDelay, 10
		Sleep, 100
		Send, {Down}
		Sleep, 100
		Send, {Enter}
	
		; activate issue folder
		SetTitleMatchMode RegEx
		WinWait, ^[1-2][0-9]{9}$, , , ,
		IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
		WinActivate, ^[1-2][0-9]{9}$, , , ,
		WinWaitActive, ^[1-2][0-9]{9}$, , , ,
		Sleep, 100
	
		; open first page
		Send, {Down}
		Sleep, 100
		Send, {Up}
		Sleep, 100
		Send, {Enter}

		; update the scoreboard
		nextscore++
		ControlSetText, Static7, NEXT, TDNP_Metadata
		ControlSetText, Static16, %nextscore%, TDNP_Metadata

		; set the GoTo indicator to FALSE (0)
		activegoto = 0
	Return
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
	Send, b
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

	; update the scoreboard
	nextscore++
	ControlSetText, Static7, NEXT, TDNP_Metadata
	ControlSetText, Static16, %nextscore%, TDNP_Metadata
Return
; ==============================

; ==============================
; NEXT+
; opens first TIFF file in next issue
; and displays issue metadata in moveable windows
; Hotkey: Win + Alt + o
#!o::
	; if the titlefoldername variable is empty
	if titlefoldername = _
	{
		; create an input box to enter the title folder name
		InputBox, input, Next Issue, Enter Title Folder Name:,, 225, 125,,,,,
			if ErrorLevel
				Return
			else
			{
				titlefoldername = %input%
			}  
	}

	; case for TRUE (1) GoTo indicator
	if activegoto = 1
	{
		; if the titlefolderpath variable is empty
		if titlefolderpath = _
		{
			; create dialog to select the title folder
			FileSelectFolder, titlefolderpath, %scannedpath%, 0, Next Issue`n`nSelect the TITLE folder:
			if ErrorLevel
				Return
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

		; activate title folder
		SetTitleMatchMode 2
		WinWaitActive, %titlefoldername%, , , ,
		Sleep, 100

		; open the next issue
		SetKeyDelay, 150
		Send, %issue%
		SetKeyDelay, 10
		Sleep, 100
		Send, {Down}
		Sleep, 100
		Send, {Enter}
	
		; activate issue folder
		SetTitleMatchMode RegEx
		WinWait, ^[1-2][0-9]{9}$, , , ,
		IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
		WinActivate, ^[1-2][0-9]{9}$, , , ,
		WinWaitActive, ^[1-2][0-9]{9}$, , , ,
		Sleep, 100
	
		; open first page
		Send, {Down}
		Sleep, 100
		Send, {Up}
		Sleep, 100
		Send, {Enter}
		Sleep, 100

		; set the GoTo indicator to FALSE (0)
		activegoto = 0
	}

	else
	{
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
		Send, b
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
	}

	; metadata harvest subroutine
	Gosub, MetaHarvest
	
	; update the scoreboard
	nextscore++
	metanextscore++
	ControlSetText, Static7, NEXT+, TDNP_Metadata
	ControlSetText, Static16, %nextscore%, TDNP_Metadata
Return
; ==============================

; ==============================
; PREVIOUS
; opens first TIFF file in previous issue
; Hotkey: Alt + p
!p::
	; if the titlefoldername variable is empty
	if titlefoldername = _
	{
		; create an input box to enter the title folder name
		InputBox, input, Previous Issue, Enter Title Folder Name:,, 225, 125,,,,,
			if ErrorLevel
				Return
			else
			{
				titlefoldername = %input%
			}  
	}

	; case for TRUE (1) GoTo indicator
	if activegoto = 1
	{
		; if the titlefolderpath variable is empty
		if titlefolderpath = _
		{
			; create dialog to select the title folder
			FileSelectFolder, titlefolderpath, %scannedpath%, 0, Previous Issue`n`nSelect the TITLE folder:
			if ErrorLevel
				Return
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
		
		; activate title folder
		SetTitleMatchMode 2
		WinWaitActive, %titlefoldername%, , , ,
		Sleep, 100

		; open the previous issue
		SetKeyDelay, 150
		Send, %issue%
		SetKeyDelay, 10
		Sleep, 100
		Send, {Up}
		Sleep, 100
		Send, {Enter}
	
		; activate issue folder
		SetTitleMatchMode RegEx
		WinWait, ^[1-2][0-9]{9}$, , , ,
		IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
		WinActivate, ^[1-2][0-9]{9}$, , , ,
		WinWaitActive, ^[1-2][0-9]{9}$, , , ,
		Sleep, 100
	
		; open first page
		Send, {Down}
		Sleep, 100
		Send, {Up}
		Sleep, 100
		Send, {Enter}

		; update the scoreboard
		prevscore+=1
		ControlSetText, Static7, PREVIOUS, TDNP_Metadata
		ControlSetText, Static17, %prevscore%, TDNP_Metadata

		; set the GoTo indicator to FALSE (0)
		activegoto = 0
	Return
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
	Send, b
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

	; update the scoreboard
	prevscore++
	ControlSetText, Static7, PREV, TDNP_Metadata
	ControlSetText, Static17, %prevscore%, TDNP_Metadata
Return
; ==============================

; ==============================
; PREVIOUS+
; opens first TIFF file in previous issue
; and displays issue metadata in moveable windows
; Hotkey: Win + Alt + p
#!p::
	; if the titlefoldername variable is empty
	if titlefoldername = _
	{
		; create an input box to enter the title folder name
		InputBox, input, Previous Issue, Enter Title Folder Name:,, 225, 125,,,,,
			if ErrorLevel
				Return
			else
			{
				titlefoldername = %input%
			}  
	}

	; case for TRUE (1) GoTo indicator
	if activegoto = 1
	{
		; if the titlefolderpath variable is empty
		if titlefolderpath = _
		{
			; create dialog to select the title folder
			FileSelectFolder, titlefolderpath, %scannedpath%, 0, Previous Issue`n`nSelect the TITLE folder:
			if ErrorLevel
				Return
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
		
		; activate title folder
		SetTitleMatchMode 2
		WinWaitActive, %titlefoldername%, , , ,
		Sleep, 100

		; open the previous issue
		SetKeyDelay, 150
		Send, %issue%
		SetKeyDelay, 10
		Sleep, 100
		Send, {Up}
		Sleep, 100
		Send, {Enter}
	
		; activate issue folder
		SetTitleMatchMode RegEx
		WinWait, ^[1-2][0-9]{9}$, , , ,
		IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
		WinActivate, ^[1-2][0-9]{9}$, , , ,
		WinWaitActive, ^[1-2][0-9]{9}$, , , ,
		Sleep, 100
	
		; open first page
		Send, {Down}
		Sleep, 100
		Send, {Up}
		Sleep, 100
		Send, {Enter}
		Sleep, 100

		; set the GoTo indicator to FALSE (0)
		activegoto = 0
	}

	else
	{
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
		Send, b
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
		Sleep, 500
	}
	
	; metadata harvest subroutine
	Gosub, MetaHarvest
	
	; update the scoreboard
	prevscore++
	metaprevscore++
	ControlSetText, Static7, PREV+, TDNP_Metadata
	ControlSetText, Static17, %prevscore%, TDNP_Metadata
Return
; ==============================

; ==============================
; Display Meta
; display the issue metadata
; Hotkey: Alt + 0
!0::
	; save clipboard contents
	temp = %clipboard%

	; reactivate the issue folder
	SetTitleMatchMode RegEx
	WinWait, ^[1-2][0-9]{9}$, , , ,
	IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
	WinActivate, ^[1-2][0-9]{9}$, , , ,
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100

	; copy issue folder path to clipboard
	Send, {F4}
	Sleep, 100
	Send, {CtrlDown}a{CtrlUp}
	Sleep, 100
	Send, {CtrlDown}c{CtrlUp}
	Sleep, 100
	Send, {Enter}
	Sleep, 100

	; grab the issue date from the file path
	StringRight, date, clipboard, 10
	
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
	IfExist, %clipboard%\metadata.txt
	{
		; read in the metadata.txt document
		Loop, read , %clipboard%\metadata.txt
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
		
		if (note == "")
			displaynote = `n`n`n`t`t`t       <<< BLANK >>>

		else
		{
			; loop to parse the note
			Loop, Parse, note, ., %A_Space%
			{
				displaynote .= A_LoopField
				
				StringGetPos, volumepos, A_LoopField, %volumespecial%, L
				StringGetPos, issuepos, A_LoopField, %issuespecial%, L
				StringGetPos, volumeissuepos, A_LoopField, %volumeissuespecial%, L
				StringLen, notelength, A_LoopField
				
				if (notelength <= 8)
				{
					if A_LoopField is integer
					{
						displaynote .= .`n`n
					}
					else
					{
						displaynote .= "."
						displaynote .= A_Space
						continue
					}
				}
				else if ((volumepos == 0) || (issuepos == 0) || (volumeissuepos == 0))
				{
					displaynote .= "."
					displaynote .= A_Space
					continue
				}
				else displaynote .= .`n`n
			}
		}

		; create display windows if first metadata display run
		if ((metaopenscore == 0) && (metanextscore == 0) && (metaprevscore == 0) && (metagotoscore == 0) && (metadisplayscore == 0))
		{
			WinGetPos, winX, winY, winWidth, winHeight, Metadata
			winY+=%winHeight%

			; create VolumeNum GUI
			Gui, 2:+AlwaysOnTop
			Gui, 2:+ToolWindow
			Gui, 2:Font, cRed s15 bold, Arial
			Gui, 2:Font, cRed s25 bold, Arial
			Gui, 2:Add, GroupBox, x0 y-18 w100 h74,       
			Gui, 2:Add, Text, x15 y8 w70 h35, %volumenum%
			Gui, 2:Show, x%winX% y%winY% h55 w100, Volume

			; create IssueNum GUI
			Gui, 3:+AlwaysOnTop
			Gui, 3:+ToolWindow
			Gui, 3:Font, cRed s25 bold, Arial
			Gui, 3:Add, GroupBox, x0 y-18 w100 h74,       
			Gui, 3:Add, Text, x15 y8 w70 h35, %issuenum%
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
		; print an error message
		MsgBox, There is no metadata.txt file in this folder.`n`n%clipboard%`n`nUse New Meta (Alt + n) to create a new file.
	}

	; restore clipboard contents
	clipboard = %temp%		

	; update the scoreboard
	metadisplayscore++
	ControlSetText, Static7, REFRESH, TDNP_Metadata
Return
; ==============================

; ==============================
; EditMetaTXT
; opens the metadata.txt file for editing
; Hotkey: Alt + m
!m::
	; if the issue folder is active
	; TAB 5 times so the script works
	SetTitleMatchMode RegEx
	IfWinActive, ^[1-2][0-9]{9}$, , , ,
		Send, {Tab 5}

	; activate the current TDNP issue folder
	SetTitleMatchMode RegEx
	WinWait, ^[1-2][0-9]{9}$, , , ,
	IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
	WinActivate, ^[1-2][0-9]{9}$, , , ,
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100
	Send, {Tab 5}

	; save clipboard contents
	temp = %clipboard%

	; copy issue folder path to clipboard
	Send, {F4}
	Sleep, 100
	Send, {CtrlDown}a{CtrlUp}
	Sleep, 100
	Send, {CtrlDown}c{CtrlUp}
	Sleep, 100
	Send, {Enter}
	Sleep, 100

	; if there is a metadata.txt file in the folder
	IfExist, %clipboard%\metadata.txt
	{
		; open the metadata.txt document
		Run, metadata.txt, %clipboard%

		; restore clipboard contents
		clipboard = %temp%
  
		; update the scoreboard
		editmetascore++
		ControlSetText, Static7, EDITMETA, TDNP_Metadata
		ControlSetText, Static18, %editmetascore%, TDNP_Metadata
	Return
	}

	; if there is no metadata.txt file in the folder
	else
	{
		; print an error message
		MsgBox, There is no metadata.txt file in this folder.`n`n%clipboard%`n`nUse New Meta (Alt + n) to create a new file.
	}

 	; restore clipboard contents
    clipboard = %temp%
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
	Send, b
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
MetaHarvest:
	; save clipboard contents
	temp = %clipboard%

	; reactivate the issue folder
	SetTitleMatchMode RegEx
	WinWait, ^[1-2][0-9]{9}$, , , ,
	IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
	WinActivate, ^[1-2][0-9]{9}$, , , ,
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100

	; copy issue folder path to clipboard
	Send, {F4}
	Sleep, 100
	Send, {CtrlDown}a{CtrlUp}
	Sleep, 100
	Send, {CtrlDown}c{CtrlUp}
	Sleep, 100
	Send, {Enter}
	Sleep, 100

	; grab the issue date from the file path
	StringRight, date, clipboard, 10
	
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
	IfExist, %clipboard%\metadata.txt
	{
		; read in the metadata.txt document
		Loop, read , %clipboard%\metadata.txt
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
		
		if (note == "")
			displaynote = `n`n`n`t`t`t       <<< BLANK >>>

		else
		{
			; loop to parse the note
			Loop, Parse, note, ., %A_Space%
			{
				displaynote .= A_LoopField
				
				StringGetPos, volumepos, A_LoopField, %volumespecial%, L
				StringGetPos, issuepos, A_LoopField, %issuespecial%, L
				StringGetPos, volumeissuepos, A_LoopField, %volumeissuespecial%, L
				StringLen, notelength, A_LoopField
				
				if (notelength <= 8)
				{
					if A_LoopField is integer
					{
						displaynote .= .`n`n
					}
					else
					{
						displaynote .= "."
						displaynote .= A_Space
						continue
					}
				}
				else if ((volumepos == 0) || (issuepos == 0) || (volumeissuepos == 0))
				{
					displaynote .= "."
					displaynote .= A_Space
					continue
				}
				else displaynote .= .`n`n
			}
		}

		; create display windows if first metadata display run
		if ((metaopenscore == 0) && (metanextscore == 0) && (metaprevscore == 0) && (metagotoscore == 0) && (metadisplayscore == 0))
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
		; print an error message
		MsgBox, There is no metadata.txt file in this folder.`n`n%clipboard%`n`nUse New Meta (Alt + n) to create a new file.
	}

	; restore clipboard contents
	clipboard = %temp%		
Return
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
						MsgBox Please enter a date in the format: YYYY-MM-DD`n`nExample: 2013-01-31
						
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
; titlefoldername variable input
TitleFolderName:
InputBox, input, Title Folder Name,,, 250, 100,,,,,%titlefoldername%
	if ErrorLevel
		Return
	else
		titlefoldername = %input%
Return

; titlefolderpath variable input
TitleFolderPath:
FileSelectFolder, titlefolderpath, %scannedpath%, 0, Edit Title Folder`n`nSelect the TITLE folder:
	if ErrorLevel
		Return
Return

DisplayValues:
MsgBox,
(
Title Folder Name:  %titlefoldername%

Title Folder Path
%titlefolderpath%
)
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
					MsgBox Please enter the issue date in the format: YYYYMMDDEE`n`nExample: 1885013101
					
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
	; update scoreboard last hotkey
	ControlSetText, Static7, CREATE, TDNP_Metadata
	
	; select title folder
	FileSelectFolder, titlefolderpath, %scannedpath%, 1, CREATE FOLDERS`n`nSelect the TITLE folder:
	if ErrorLevel
		Return		
	else
	{
		; make the title folder the working directory
		SetWorkingDir %titlefolderpath%
		
		; create display variables for title folder path
		StringGetPos, titlefolderpos, titlefolderpath, \, R2
		StringLeft, titlefolder1, titlefolderpath, titlefolderpos
		StringTrimLeft, titlefolder2, titlefolderpath, titlefolderpos		
		
		; confirm number of folders and directory
		MsgBox, 4, Create Issue Folders, %titlefolder1%`n%titlefolder2%`n`n`t`tCreate %currentissuenum% folders?`n`nYes to Continue`nNo to Exit
		IfMsgBox, No, Return
		
		; loop creates issue folders stored in issuefile
		Loop, parse, issuefile, `n
		{
			FileCreateDir, %A_LoopField%
		}
		
		; update the scoreboard
		issuenum += %currentissuenum%
		ControlSetText, Static12, %issuenum%, TDNP_Metadata
	}
Return

PasteImages:
	; abort script if issuefile is empty
	if issuefile =
	{
		MsgBox, There are no issues to paste.`n`nChoose "File > Enter Folder Names" to enter issues. 
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
					MsgBox, 4, Paste Images, %titlefolder1%`n%titlefolder2%`n`n`t`t%monthname% %day%`, %year%`n`n`t`t%A_LoopField%`n`nYes to Continue`nNo to Exit
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
						ControlSetText, Static13, %imagescore%, TDNP_Metadata
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
MsgBox TDNP_Metadata.ahk`nVersion 2.5`nAndrew.Weidner@unt.edu
Return
; =================HELP
; ===========================================MENU FUNCTIONS

GuiClose:
ExitApp
