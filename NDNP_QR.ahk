/*
 * NDNP_QR.ahk
 *
 * DESCRIPTION: tool for NDNP data Quality Review
 *
 * VERSION HISTORY ******************
 *
 * Version 1.0 2012-12-20
 * Version 1.1 2013-01-11: added Search menu, modified inputboxes
 * Version 1.2 2013-01-15: added Next+, Previous+,
 *                         DVV Pages Loop, & Validate/Verify
 * Version 1.3 2013-02-27: reworked interface; added DVV Thumbs loop
 *                         ; added reel report tool
 * Version 1.4 2013-03-07: added batch report tool; added JP2 check
 *                         for NextJP2 and PreviousJP2
 *                         ; added page count & questionable to report
 * Version 1.5 2013-03-11: modified for FirstImpression viewer
 *                         ; added edition label to report
 * Version 1.6 2013-03-15: reworked batch report to sort batch.xml
 *                         ; reworked folder navigation
 *                         ; added Open+ and Metadata Viewer
 *
 * TABLE OF CONTENTS ****************
 *
 * Line#
 *   283  Open
 *   351  Open+
 *   743  Next
 *   825  Next+
 *  1232  Previous
 *  1314  Previous+
 *  1722  Metadata
 *  2058  Zoom In
 *  2094  Zoom Out
 *  2130  ViewXML
 *  2170  EditXML
 *  2271  Reel Report
 *  3127  Batch Report
 *  3596  Metadata Viewer
 *  4343  DVV Pages
 *  4409  DVV Thumbs
 *
 * REQUIRED SOFTWARE ****************
 *
 * Notepad++ (metadata extraction)
 * http://notepad-plus-plus.org/
 *
 * First Impression Viewer (default TIFF viewer)
 * http://www.softpedia.com/get/Multimedia/Graphic/Graphic-Viewers/First-Impression.shtml
 *
 * Digital Viewer and Validator (Pages & Thumbs Loops)
 * http://www.loc.gov/ndnp/tools/
 *
 * GUI NOTE *************************
 * The GUI dimensions may need to be adjusted for your system.
 * See the AutoHotkey GUI documentation for more information:
 * http://www.autohotkey.com/docs/commands/Gui.htm
 *
 * CONTACT: Andrew.Weidner@unt.edu
 */

#NoEnv
#persistent

; =====================================================
; ===========================================VARIABLES
; EDIT THIS VARIABLE
; for the drive where you store your QR batches
batchdrive = E:\

; EDIT THIS VARIABLE
; for the path to the Notepad++ folder
notepadpath = C:\Program Files (x86)\Notepad++

; scoreboard
hotkeys = 0
keystrokes = 0
openscore = 0
nextscore = 0
prevscore = 0
Metadatascore = 0
viewissuexmlscore = 0
editissuexmlscore = 0
zoominscore = 0
zoomoutscore = 0
clipboard =

; search
searchstring =

; Metadata
volume =
issue =
date =
month =
monthname =
day =

; Tools
delay =
count = 0
pagecount = 0
thumbscount= 0
batchpath =
reelnumber =
folderpath =
reportpath =
reportcount = 0
; ===========================================VARIABLES
; =====================================================


Gui, 1:Color, d0d0d0, 912206
Gui, 1:Show, h0 w0, NDNP_QR


; =================================================
; ===========================================MENUS
; see MENU FUNCTIONS (line #)

; File menu (2229)
Menu, FileMenu, Add, Open &Batch Folder, OpenBatch
Menu, FileMenu, Add
Menu, FileMenu, Add, Va&lidate Batch, DVVvalidate
Menu, FileMenu, Add, Ve&rify Batch, DVVverify
Menu, FileMenu, Add
Menu, FileMenu, Add, &Reload, Reload
Menu, FileMenu, Add, E&xit, Exit

; Tools menu (2270)
Menu, ToolsMenu, Add, &Reel Report, ReelReport
Menu, ToolsMenu, Add, &Batch Report, BatchReport
Menu, ToolsMenu, Add
Menu, ToolsMenu, Add, &Metadata Viewer, MetaViewer
Menu, ToolsMenu, Add
Menu, ToolsMenu, Add, DVV &Pages Loop, DVVpages
Menu, ToolsMenu, Add, DVV &Thumbs Loop, DVVthumbs

; Search menu (4483)
Menu, SearchMenu, Add, US Directory: &Texas, DirectoryTitleTexas
Menu, SearchMenu, Add, US Directory: &Oklahoma, DirectoryTitleOklahoma
Menu, SearchMenu, Add, US Directory: &NewMexico, DirectoryTitleNewMexico
Menu, SearchMenu, Add
Menu, SearchMenu, Add, US Directory: &LCCN, DirectoryLCCN
Menu, SearchMenu, Add
Menu, SearchMenu, Add, ChronAm: Te&xas, ChronAmTexas
Menu, SearchMenu, Add, ChronAm: O&klahoma, ChronAmOklahoma
Menu, SearchMenu, Add, ChronAm: New &Mexico, ChronAmNewMexico

; Help menu (4548)
Menu, HelpMenu, Add, &Documentation, Documentation
Menu, HelpMenu, Add, &NDNP Awardee Wiki, NDNPwiki
Menu, HelpMenu, Add, &About, About

; create menus
Menu, MenuBar, Add, &File, :FileMenu
Menu, MenuBar, Add, &Tools, :ToolsMenu
Menu, MenuBar, Add, &Search, :SearchMenu
Menu, MenuBar, Add, &Help, :HelpMenu

; create menu toolbar
Gui, Menu, MenuBar
; ===========================================MENUS
; =================================================


; ==================================================
; ===========================================LABELS
; scoreboard - Last Hotkey name: Static 3
Gui, 1:Add, Text, x15  y10 w40  h15, HotKeys
Gui, 1:Add, Text, x95  y10 w70  h15, Last HotKey
Gui, 1:Add, Text, x105 y35 w60  h15, 
Gui, 1:Add, Text, x175 y10 w70 h15, Report
Gui, 1:Add, Text, x255 y10 w70  h15, DVV

; Metadata window
Gui, 2:Font,, Arial
Gui, 2:Add, Text, x40 y55  w100 h20, Volume:
Gui, 2:Add, Text, x49 y80  w100 h20, Issue:
Gui, 2:Add, Text, x44 y105 w100 h20, ? Date:
Gui, 2:Add, Text, x45 y130 w100 h20, Pages:


; keystroke counters
; row 1
Gui, 1:Add, Text, x15  y70 w100 h20, Open ( i )
Gui, 1:Add, Text, x95  y70 w100 h20, Next ( o )
Gui, 1:Add, Text, x175 y70 w100 h20, Previous ( p )
Gui, 1:Add, Text, x255 y70 w100 h20, Metadata ( m )
; row 2
Gui, 1:Add, Text, x15  y130 w70 h20, Zoom In ( k )
Gui, 1:Add, Text, x95  y130 w70 h20, Zoom Out ( l )
Gui, 1:Add, Text, x175 y130 w70 h20, ViewXML ( q )
Gui, 1:Add, Text, x255 y130 w70 h20, EditXML ( w )
; ===========================================LABELS
; ==================================================


; ================================================
; ===========================================DATA
Gui, 1:Font, s15,

; Hotkeys & loops: Static 14-16
Gui, 1:Add, Text, x25  y30 w55 h25, 0
Gui, 1:Add, Text, x185 y30 w55 h25, 0
Gui, 1:Add, Text, x265 y30 w55 h25, 0
; row 1: Static 17-20
Gui, 1:Add, Text, x25  y90 w55 h25, 0
Gui, 1:Add, Text, x105 y90 w55 h25, 0
Gui, 1:Add, Text, x185 y90 w55 h25, 0
Gui, 1:Add, Text, x265 y90 w55 h25, 0
; row 2: Static 21-24
Gui, 1:Add, Text, x25  y150 w45 h25, 0
Gui, 1:Add, Text, x105 y150 w45 h25, 0
Gui, 1:Add, Text, x185 y150 w45 h25, 0
Gui, 1:Add, Text, x265 y150 w45 h25, 0

; sets Metadata window font
Gui, 2:Font, cRed s12 bold, Arial

; Metadata data fields
; Static 5-9
Gui, 2:Add, Text, x55 y20  w90  h20,
Gui, 2:Add, Text, x90 y55  w100 h20,
Gui, 2:Add, Text, x90 y80  w100 h20,
Gui, 2:Add, Text, x90 y105 w100 h20,
Gui, 2:Add, Text, x90 y130 w100 h20,
; ===========================================DATA
; ================================================


; ===============================================
; =========================================BOXES
; totals section
Gui, 1:Add, GroupBox, x2 y-12 w340 h200,       

; HotKeys, Last Hotkey, and loops
Gui, 1:Add, GroupBox, x15  y15 w75 h45,         
Gui, 1:Add, GroupBox, x95  y15 w75 h45,           
Gui, 1:Add, GroupBox, x175 y15 w75 h45,           
Gui, 1:Add, GroupBox, x255 y15 w75 h45,           

; row 1
Gui, 1:Add, GroupBox, x15  y75 w75 h45,         
Gui, 1:Add, GroupBox, x95  y75 w75 h45,           
Gui, 1:Add, GroupBox, x175 y75 w75 h45,           
Gui, 1:Add, GroupBox, x255 y75 w75 h45,           
; row 2
Gui, 1:Add, GroupBox, x15  y135 w75 h45,         
Gui, 1:Add, GroupBox, x95  y135 w75 h45,           
Gui, 1:Add, GroupBox, x175 y135 w75 h45,           
Gui, 1:Add, GroupBox, x255 y135 w75 h45,           

; Metadata decorative
Gui, 2:Add, GroupBox, x0 y-8 w200 h169,       

; Metadata date
Gui, 2:Add, GroupBox, x40 y5 w120 h40,       
Gui, 2:Add, GroupBox, x38 y3 w124 h44,       
; =========================================BOXES
; ===============================================


; format main GUI
WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
winX+=%winWidth%
Gui, 1:Show, x%winX% y%winY% h190 w345, NDNP_QR
WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR

; format Metadata GUI
Gui, 2:+AlwaysOnTop
Gui, 2:+ToolWindow
winX+=%winWidth%
Gui, 2:Show, x%winX% y%winY% h160 w200, Metadata

winactivate, NDNP_QR

Pause::Pause


; =================================================
; =========================================SCRIPTS
; =============================OPEN
; opens selected issue folder and first TIFF file in First Impression
; HotKey = Alt + i
!i::
  ; close any First Impression windows
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

	; wait for the reel folder
	SetTitleMatchMode RegEx
	IfWinNotActive, ^[0-9]{10}[A0-9]$, , , ,
	WinActivate, ^[0-9]{10}[A0-9]$, , , ,
	WinWaitActive, ^[0-9]{10}[A0-9]$, , , ,
	Sleep, 100

	; open the selected issue folder
	Send, {Enter}

	; wait for the issue folder
	WinWaitActive, ^1[0-9]{9}$, , , ,
	Sleep, 100

	; open the first TIFF
	Send, {Down 2}
	Send, {Enter}
	Sleep, 100

	; get the unique window id#
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
	Send, {Enter}

	; update the scoreboard
	hotkeys++
	openscore++
	ControlSetText, Static3, OPEN, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static17, %openscore%, NDNP_QR
Return
; =============================OPEN

; =============================OPEN+
; opens selected issue folder, first TIFF file and displays metadata
; HotKey = Ctrl + Alt + i
^!i::
	; close the Edition Label window
	Gui, 9:Destroy

	; close any First Impression windows
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

	; activate the reel folder
	SetTitleMatchMode RegEx
	IfWinNotActive, ^[0-9]{10}[A0-9]$, , , ,
	WinActivate, ^[0-9]{10}[A0-9]$, , , ,
	WinWaitActive, ^[0-9]{10}[A0-9]$, , , ,
	Sleep, 100

	; open the selected issue folder
	Send, {Enter}

	; wait for the issue folder
	WinWaitActive, ^1[0-9]{9}$, , , ,
	Sleep, 100

	; open the first TIFF
	Send, {Down 2}
	Send, {Enter}
	Sleep, 100

	; get the unique window id#
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

	; EXTRACT THE METADATA
	; store clipboard contents
	temp = %clipboard%
	
	; clear clipboard contents
	clipboard =

	; activate the current NDNP issue folder
	SetTitleMatchMode RegEx
	WinWait, ^[1-2][0-9]{9}$, , , ,
	IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
	WinActivate, ^[1-2][0-9]{9}$, , , ,
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100

	; moves to the issue.xml file
	Send, {End}
	Send, {Up}
	Sleep, 100
  
	; opens the issue.xml file in Notepad++
	Send, {AppsKey}
	Sleep, 200
	Send, n
	Sleep, 100
	Send, {Enter}
	; resets window
	Send, {Home}

	; wait for Notepad++
	SetTitleMatchMode 2
	WinWait, Notepad++, , , ,
	IfWinNotActive, Notepad++, , , ,
	WinActivate, Notepad++, , , ,
	WinWaitActive, Notepad++, , , ,
	Sleep, 200

	; activate Find dialog
	Send, {CtrlDown}f{CtrlUp}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100

	; find the volume number
	SendInput, ="volume"
	Send, {Enter}
	Sleep, 100

	; if no volume number
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		volume = none
		
		; update the volume number
		ControlSetText, Static6, %volume%, Metadata
	}
				
	else
	{
		; move to the number
		SendInput, <mods:number>
		Send, {Enter}
		Sleep, 100
		
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the volume number
		Send, {Right}
		Send, {ShiftDown}{End}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100

		; remove the </mods:number> tag
		; and store the volume number
		StringTrimRight, volume, clipboard, 14

		; update the volume number
		ControlSetText, Static6, %volume%, Metadata

		; activate Find dialog
		Send, {CtrlDown}f{CtrlUp}
	}
				
	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100

	; find the issue number
	SendInput, ="issue"
	Send, {Enter}
	Sleep, 100
				
	; if no issue number
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		issue = none
		
		; update the issue number
		ControlSetText, Static7, %issue%, Metadata		
	}
				
	else
	{
		; move to the number
		SendInput, <mods:number>
		Send, {Enter}
		Sleep, 100
		
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the issue number
		Send, {Right}
		Send, {ShiftDown}{End}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100

		; remove the </mods:number> tag
		; and store the issue number
		StringTrimRight, issue, clipboard, 14

		; update the issue number
		ControlSetText, Static7, %issue%, Metadata

		; activate Find dialog
		Send, {CtrlDown}f{CtrlUp}
	}
				
	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100
			
	; find the edition
	SendInput, ="edition"
	Send, {Enter}
	Sleep, 100
				
	; move to the edition label
	SendInput, <mods:caption>
	Send, {Enter}
	Sleep, 100
		
	; if no edition label
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		editionlabel =
	}
				
	else
	{		
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the edition label
		Send, {Right}
		Send, {ShiftDown}{End}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100

		; remove the </mods:caption> tag
		; and store the issue number
		StringTrimRight, editionlabel, clipboard, 15

		WinGetPos, winX, winY, winWidth, winHeight, Metadata
		winX+=%winWidth%

		; create Edition Label GUI
		Gui, 9:+AlwaysOnTop
		Gui, 9:+ToolWindow
		Gui, 9:Font, cRed s15 bold, Arial
		Gui, 9:Add, Text, x15 y15 w380 h25, %editionlabel%
		Gui, 9:Font, cRed s10 bold, Arial
		Gui, 9:Add, GroupBox, x0 y-7 w400 h63,       
		Gui, 9:Show, x%winX% y%winY% h55 w400, Edition

		; wait for Notepad++
		SetTitleMatchMode 2
		WinWait, Notepad++, , , ,
		IfWinNotActive, Notepad++, , , ,
		WinActivate, Notepad++, , , ,
		WinWaitActive, Notepad++, , , ,
		Sleep, 200
			
		; activate Find dialog
		Send, {CtrlDown}f{CtrlUp}
	}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100
				
	; find the publication date
	SendInput, </mods:dateIssued>
	Send, {Enter}

	; close the Find Window
	Send, {Esc}
	Sleep, 100

	; copy the publication date
	Send, {Left}{ShiftDown}{Left 10}{ShiftUp}
	Send, {CtrlDown}c{CtrlUp}
  
	; store the publication date
	date = %clipboard%

	; update the date field
	ControlSetText, Static5, %date%, Metadata

	; clear the questionabledate variable
	questionabledate =

	; activate Find dialog
	Send, {CtrlDown}f{CtrlUp}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100

	; find a questionable date
	SendInput, "questionable">
	Send, {Enter}
	Sleep, 100

	; if no questionable date
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		Sleep, 100
					
		; update the questionable date
		ControlSetText, Static8, %questionabledate%, Metadata		
	}
				
	else
	{
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the questionable publication date
		Send, {Right}{ShiftDown}{Right 10}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
  
		; store the questionable publication date
		questionabledate = %clipboard%
		
		; update the questionable date
		ControlSetText, Static8, %questionabledate%, Metadata		
		
		; activate the Find window
		Send, {CtrlDown}f{CtrlUp}
	}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100
	
	; initialize the numpages variable
	numpages = 0
	
	; count the pages
	Send, pageFileGrp
	Sleep, 100
	Loop
	{
		Send, {Enter}
		Sleep, 100
		
		ifWinActive, Find, Can't find the text:, ,
		{
			; close the Find windows
			Send, {Enter}
			Sleep, 100
			Send, {Esc}
			
			; end the page count loop
			Break
		}
		
		; add 1 to numpages
		numpages++		
	}

	; update the number of pages
	ControlSetText, Static9, %numpages%, Metadata
	
	; close the issue.xml file
	Send, {AltDown}f{AltUp}
	Sleep, 100
	Send, c
	
	; activate First Impression
	WinActivate, ahk_id %firstid%
	
	; bring all metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Date
	WinSet, Top,, Questionable
	WinSet, Top,, Volume
	WinSet, Top,, Issue
	WinSet, Top,, Edition
	WinSet, Top,, Timer	

	; restore the clipboard
	clipboard = %temp%

	; update the scoreboard
	hotkeys+=2
	openscore++
	Metadatascore++
	ControlSetText, Static3, OPEN+, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static17, %openscore%, NDNP_QR
	ControlSetText, Static20, %Metadatascore%, NDNP_QR
Return
; =============================OPEN+

; =============================NEXT
; opens the first TIFF file in the next issue folder
; Hotkey: Alt + o
!o::
	; close any First Impression windows
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
	Send, b
  
	; wait for the reel folder
	WinWaitActive, ^[0-9]{10}[A0-9]$, , , ,
	Sleep, 100

	; open the next folder
	Send, {Down}
	Sleep, 100
	Send, {Enter}

	; wait for the issue folder
	WinWaitActive, ^1[0-9]{9}$, , , ,
	Sleep, 100

	; open the first TIFF
	Send, {Down 2}
	Send, {Enter}
	Sleep, 100

	; get the unique window id#
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
	Send, {Enter}

	; update the scoreboard
	hotkeys++
	nextscore++
	ControlSetText, Static3, NEXT, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static18, %nextscore%, NDNP_QR
Return
; =============================NEXT

; =============================NEXT+
; opens the first TIFF file in the next issue folder
; and extracts the issue metadata
; Hotkey: Ctrl + Alt + o
^!o::
	; close the Edition Label window
	Gui, 9:Destroy

	; close any First Impression windows
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
	Send, b
  
	; wait for the reel folder
	WinWaitActive, ^[0-9]{10}[A0-9]$, , , ,
	Sleep, 100
	
	; open the next folder
	Send, {Down}
	Sleep, 100
	Send, {Enter}

	; wait for the issue folder
	WinWaitActive, ^1[0-9]{9}$, , , ,
	Sleep, 100

	; open the first TIFF
	Send, {Down 2}
	Send, {Enter}
	Sleep, 100

	; get the unique window id#
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

	; EXTRACT THE METADATA
	; store clipboard contents
	temp = %clipboard%
	
	; clear clipboard contents
	clipboard =

	; activate the current NDNP issue folder
	SetTitleMatchMode RegEx
	WinWait, ^[1-2][0-9]{9}$, , , ,
	IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
	WinActivate, ^[1-2][0-9]{9}$, , , ,
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100

	; moves to the issue.xml file
	Send, {End}
	Send, {Up}
	Sleep, 100
  
	; opens the issue.xml file in Notepad++
	Send, {AppsKey}
	Sleep, 200
	Send, n
	Sleep, 100
	Send, {Enter}
	; resets window
	Send, {Home}

	; wait for Notepad++
	SetTitleMatchMode 2
	WinWait, Notepad++, , , ,
	IfWinNotActive, Notepad++, , , ,
	WinActivate, Notepad++, , , ,
	WinWaitActive, Notepad++, , , ,
	Sleep, 200

	; activate Find dialog
	Send, {CtrlDown}f{CtrlUp}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100

	; find the volume number
	SendInput, ="volume"
	Send, {Enter}
	Sleep, 100

	; if no volume number
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		volume = none
		
		; update the volume number
		ControlSetText, Static6, %volume%, Metadata
	}
				
	else
	{
		; move to the number
		SendInput, <mods:number>
		Send, {Enter}
		Sleep, 100
		
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the volume number
		Send, {Right}
		Send, {ShiftDown}{End}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100

		; remove the </mods:number> tag
		; and store the volume number
		StringTrimRight, volume, clipboard, 14

		; update the volume number
		ControlSetText, Static6, %volume%, Metadata

		; activate Find dialog
		Send, {CtrlDown}f{CtrlUp}
	}
				
	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100

	; find the issue number
	SendInput, ="issue"
	Send, {Enter}
	Sleep, 100
				
	; if no issue number
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		issue = none
		
		; update the issue number
		ControlSetText, Static7, %issue%, Metadata		
	}
				
	else
	{
		; move to the number
		SendInput, <mods:number>
		Send, {Enter}
		Sleep, 100
		
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the issue number
		Send, {Right}
		Send, {ShiftDown}{End}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100

		; remove the </mods:number> tag
		; and store the issue number
		StringTrimRight, issue, clipboard, 14

		; update the issue number
		ControlSetText, Static7, %issue%, Metadata

		; activate Find dialog
		Send, {CtrlDown}f{CtrlUp}
	}
				
	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100
			
	; find the edition
	SendInput, ="edition"
	Send, {Enter}
	Sleep, 100
				
	; move to the edition label
	SendInput, <mods:caption>
	Send, {Enter}
	Sleep, 100
		
	; if no edition label
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		editionlabel =
	}
				
	else
	{		
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the edition label
		Send, {Right}
		Send, {ShiftDown}{End}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100

		; remove the </mods:caption> tag
		; and store the issue number
		StringTrimRight, editionlabel, clipboard, 15

		WinGetPos, winX, winY, winWidth, winHeight, Metadata
		winX+=%winWidth%

		; create Edition Label GUI
		Gui, 9:+AlwaysOnTop
		Gui, 9:+ToolWindow
		Gui, 9:Font, cRed s15 bold, Arial
		Gui, 9:Add, Text, x15 y15 w380 h25, %editionlabel%
		Gui, 9:Font, cRed s10 bold, Arial
		Gui, 9:Add, GroupBox, x0 y-7 w400 h63,       
		Gui, 9:Show, x%winX% y%winY% h55 w400, Edition

		; wait for Notepad++
		SetTitleMatchMode 2
		WinWait, Notepad++, , , ,
		IfWinNotActive, Notepad++, , , ,
		WinActivate, Notepad++, , , ,
		WinWaitActive, Notepad++, , , ,
		Sleep, 200
			
		; activate Find dialog
		Send, {CtrlDown}f{CtrlUp}
	}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100
				
	; find the publication date
	SendInput, </mods:dateIssued>
	Send, {Enter}

	; close the Find Window
	Send, {Esc}
	Sleep, 100

	; copy the publication date
	Send, {Left}{ShiftDown}{Left 10}{ShiftUp}
	Send, {CtrlDown}c{CtrlUp}
  
	; store the publication date
	date = %clipboard%

	; update the date field
	ControlSetText, Static5, %date%, Metadata

	; clear the questionabledate variable
	questionabledate =

	; activate Find dialog
	Send, {CtrlDown}f{CtrlUp}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100

	; find a questionable date
	SendInput, "questionable">
	Send, {Enter}
	Sleep, 100

	; if no questionable date
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		Sleep, 100
					
		; update the questionable date
		ControlSetText, Static8, %questionabledate%, Metadata		
	}
				
	else
	{
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the questionable publication date
		Send, {Right}{ShiftDown}{Right 10}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
  
		; store the questionable publication date
		questionabledate = %clipboard%
		
		; update the questionable date
		ControlSetText, Static8, %questionabledate%, Metadata		
		
		; activate the Find window
		Send, {CtrlDown}f{CtrlUp}
	}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100
	
	; initialize the numpages variable
	numpages = 0
	
	; count the pages
	Send, pageFileGrp
	Sleep, 100
	Loop
	{
		Send, {Enter}
		Sleep, 100
		
		ifWinActive, Find, Can't find the text:, ,
		{
			; close the Find windows
			Send, {Enter}
			Sleep, 100
			Send, {Esc}
			
			; end the page count loop
			Break
		}
		
		; add 1 to numpages
		numpages++		
	}

	; update the number of pages
	ControlSetText, Static9, %numpages%, Metadata
	
	; close the issue.xml file
	Send, {AltDown}f{AltUp}
	Sleep, 100
	Send, c
	
	; activate First Impression
	WinActivate, ahk_id %firstid%
	
	; bring all metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Date
	WinSet, Top,, Questionable
	WinSet, Top,, Volume
	WinSet, Top,, Issue
	WinSet, Top,, Edition
	WinSet, Top,, Timer	

	; restore the clipboard
	clipboard = %temp%

	; update the scoreboard
	hotkeys+=2
	nextscore++
	Metadatascore++
	ControlSetText, Static3, NEXT+, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static18, %nextscore%, NDNP_QR
	ControlSetText, Static20, %Metadatascore%, NDNP_QR
Return
; =============================NEXT+

; =============================PREVIOUS
; opens the first TIFF file in the previous issue folder
; Hotkey: Alt + p
!p::
	; close any First Impression windows
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
	Send, b
  
	; wait for the reel folder
	WinWaitActive, ^[0-9]{10}[A0-9]$, , , ,
	Sleep, 100
	
	; open the  previous folder
	Send, {Up}
	Sleep, 100
	Send, {Enter}

	; wait for the issue folder
	WinWaitActive, ^1[0-9]{9}$, , , ,
	Sleep, 100

	; open the first TIFF
	Send, {Down 2}
	Send, {Enter}
	Sleep, 100

	; get the unique window id#
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
	Send, {Enter}

	; update the scoreboard
	hotkeys++
	revscore++
	ControlSetText, Static3, PREV, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static19, %revscore%, NDNP_QR
Return
; =============================PREVIOUS

; =============================PREVIOUS+
; opens the first TIFF file in the previous issue folder
; and extracts the issue metadata
; Hotkey: Ctrl + Alt + p
^!p::
	; close the Edition Label Window
	Gui, 9:Destroy

	; close any First Impression windows
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
	Send, b
  
	; wait for the reel folder
	WinWaitActive, ^[0-9]{10}[A0-9]$, , , ,
	Sleep, 100

	; open the  previous folder
	Send, {Up}
	Sleep, 100
	Send, {Enter}

	; wait for the issue folder
	WinWaitActive, ^1[0-9]{9}$, , , ,
	Sleep, 100

	; open the first TIFF
	Send, {Down 2}
	Send, {Enter}
	Sleep, 100

	; get the unique window id#
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

	; EXTRACT THE METADATA
	; store clipboard contents
	temp = %clipboard%
	
	; clear clipboard contents
	clipboard =

	; activate the current NDNP issue folder
	SetTitleMatchMode RegEx
	WinWait, ^[1-2][0-9]{9}$, , , ,
	IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
	WinActivate, ^[1-2][0-9]{9}$, , , ,
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100

	; move to the issue.xml file
	Send, {End}
	Send, {Up}
	Sleep, 100
  
	; open the issue.xml file in Notepad++
	Send, {AppsKey}
	Sleep, 200
	Send, n
	Sleep, 100
	Send, {Enter}
	; resets window
	Send, {Home}

	; wait for Notepad++
	SetTitleMatchMode 2
	WinWait, Notepad++, , , ,
	IfWinNotActive, Notepad++, , , ,
	WinActivate, Notepad++, , , ,
	WinWaitActive, Notepad++, , , ,
	Sleep, 200

	; activate Find dialog
	Send, {CtrlDown}f{CtrlUp}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100

	; find the volume number
	SendInput, ="volume"
	Send, {Enter}
	Sleep, 100

	; if no volume number
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		volume = none
		
		; update the volume number
		ControlSetText, Static6, %volume%, Metadata
	}
				
	else
	{
		; move to the number
		SendInput, <mods:number>
		Send, {Enter}
		Sleep, 100
		
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the volume number
		Send, {Right}
		Send, {ShiftDown}{End}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100

		; remove the </mods:number> tag
		; and store the volume number
		StringTrimRight, volume, clipboard, 14

		; update the volume number
		ControlSetText, Static6, %volume%, Metadata

		; activate Find dialog
		Send, {CtrlDown}f{CtrlUp}
	}
				
	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100

	; find the issue number
	SendInput, ="issue"
	Send, {Enter}
	Sleep, 100
				
	; if no issue number
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		issue = none
		
		; update the issue number
		ControlSetText, Static7, %issue%, Metadata		
	}
				
	else
	{
		; move to the number
		SendInput, <mods:number>
		Send, {Enter}
		Sleep, 100
		
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the issue number
		Send, {Right}
		Send, {ShiftDown}{End}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100

		; remove the </mods:number> tag
		; and store the issue number
		StringTrimRight, issue, clipboard, 14

		; update the issue number
		ControlSetText, Static7, %issue%, Metadata

		; activate Find dialog
		Send, {CtrlDown}f{CtrlUp}
	}
				
	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100
			
	; find the edition
	SendInput, ="edition"
	Send, {Enter}
	Sleep, 100
				
	; move to the edition label
	SendInput, <mods:caption>
	Send, {Enter}
	Sleep, 100
		
	; if no edition label
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		editionlabel =
	}
				
	else
	{		
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the edition label
		Send, {Right}
		Send, {ShiftDown}{End}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100

		; remove the </mods:caption> tag
		; and store the issue number
		StringTrimRight, editionlabel, clipboard, 15

		WinGetPos, winX, winY, winWidth, winHeight, Metadata
		winX+=%winWidth%

		; create Edition Label GUI
		Gui, 9:+AlwaysOnTop
		Gui, 9:+ToolWindow
		Gui, 9:Font, cRed s15 bold, Arial
		Gui, 9:Add, Text, x15 y15 w380 h25, %editionlabel%
		Gui, 9:Font, cRed s10 bold, Arial
		Gui, 9:Add, GroupBox, x0 y-7 w400 h63,       
		Gui, 9:Show, x%winX% y%winY% h55 w400, Edition

		; wait for Notepad++
		SetTitleMatchMode 2
		WinWait, Notepad++, , , ,
		IfWinNotActive, Notepad++, , , ,
		WinActivate, Notepad++, , , ,
		WinWaitActive, Notepad++, , , ,
		Sleep, 200
			
		; activate Find dialog
		Send, {CtrlDown}f{CtrlUp}
	}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100
				
	; find the publication date
	SendInput, </mods:dateIssued>
	Send, {Enter}

	; close the Find Window
	Send, {Esc}
	Sleep, 100

	; copy the publication date
	Send, {Left}{ShiftDown}{Left 10}{ShiftUp}
	Send, {CtrlDown}c{CtrlUp}
  
	; store the publication date
	date = %clipboard%

	; update the date field
	ControlSetText, Static5, %date%, Metadata

	; clear the questionabledate variable
	questionabledate =

	; activate Find dialog
	Send, {CtrlDown}f{CtrlUp}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100

	; find a questionable date
	SendInput, "questionable">
	Send, {Enter}
	Sleep, 100

	; if no questionable date
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		Sleep, 100
					
		; update the questionable date
		ControlSetText, Static8, %questionabledate%, Metadata		
	}
				
	else
	{
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the questionable publication date
		Send, {Right}{ShiftDown}{Right 10}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
  
		; store the questionable publication date
		questionabledate = %clipboard%
		
		; update the questionable date
		ControlSetText, Static8, %questionabledate%, Metadata		
		
		; activate the Find window
		Send, {CtrlDown}f{CtrlUp}
	}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100
	
	; initialize the numpages variable
	numpages = 0
	
	; count the pages
	Send, pageFileGrp
	Sleep, 100
	Loop
	{
		Send, {Enter}
		Sleep, 100
		
		ifWinActive, Find, Can't find the text:, ,
		{
			; close the Find windows
			Send, {Enter}
			Sleep, 100
			Send, {Esc}
			
			; end the page count loop
			Break
		}
		
		; add 1 to numpages
		numpages++		
	}

	; update the number of pages
	ControlSetText, Static9, %numpages%, Metadata

	; close the issue.xml file
	Send, {AltDown}f{AltUp}
	Sleep, 100
	Send, c

	; activate First Impression
	SetTitleMatchMode 1
	WinActivate, ahk_id %firstid%
	
	; bring all metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Date
	WinSet, Top,, Questionable
	WinSet, Top,, Volume
	WinSet, Top,, Issue
	WinSet, Top,, Edition
	WinSet, Top,, Timer	

	; restore the clipboard
	clipboard = %temp%

	; update the scoreboard
	hotkeys+=2
	revscore++
	Metadatascore++
	ControlSetText, Static3, PREV+, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static19, %revscore%, NDNP_QR
	ControlSetText, Static20, %Metadatascore%, NDNP_QR
Return
; =============================PREVIOUS+

; =============================METADATA
; displays issue metadata: date, volume no., issue no., questionable date,
;                          number of pages, edition label
; Hotkey: Alt + m
!m::
	; store clipboard contents
	temp = %clipboard%
	
	; clear clipboard contents
	clipboard =
	
	; close active Edition Label window
	Gui, 9:Destroy

	; activate the current NDNP issue folder
	SetTitleMatchMode RegEx
	WinWait, ^[1-2][0-9]{9}$, , , ,
	IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
	WinActivate, ^[1-2][0-9]{9}$, , , ,
	WinWaitActive, ^[1-2][0-9]{9}$, , , ,
	Sleep, 100

	; moves to the issue.xml file
	Send, {End}
	Send, {Up}
	Sleep, 100
  
	; opens the issue.xml file in Notepad++
	Send, {AppsKey}
	Sleep, 200
	Send, n
	Sleep, 100
	Send, {Enter}
	; resets window
	Send, {Home}

	; wait for Notepad++
	SetTitleMatchMode 2
	WinWait, Notepad++, , , ,
	IfWinNotActive, Notepad++, , , ,
	WinActivate, Notepad++, , , ,
	WinWaitActive, Notepad++, , , ,
	Sleep, 200

	; activate Find dialog
	Send, {CtrlDown}f{CtrlUp}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100

	; find the volume number
	SendInput, ="volume"
	Send, {Enter}
	Sleep, 100

	; if no volume number
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		volume = none
		
		; update the volume number
		ControlSetText, Static6, %volume%, Metadata
	}
				
	else
	{
		; move to the number
		SendInput, <mods:number>
		Send, {Enter}
		Sleep, 100
		
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the volume number
		Send, {Right}
		Send, {ShiftDown}{End}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100

		; remove the </mods:number> tag
		; and store the volume number
		StringTrimRight, volume, clipboard, 14

		; update the volume number
		ControlSetText, Static6, %volume%, Metadata

		; activate Find dialog
		Send, {CtrlDown}f{CtrlUp}
	}
				
	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100

	; find the issue number
	SendInput, ="issue"
	Send, {Enter}
	Sleep, 100
				
	; if no issue number
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		issue = none
		
		; update the issue number
		ControlSetText, Static7, %issue%, Metadata		
	}
				
	else
	{
		; move to the number
		SendInput, <mods:number>
		Send, {Enter}
		Sleep, 100
		
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the issue number
		Send, {Right}
		Send, {ShiftDown}{End}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100

		; remove the </mods:number> tag
		; and store the issue number
		StringTrimRight, issue, clipboard, 14

		; update the issue number
		ControlSetText, Static7, %issue%, Metadata

		; activate Find dialog
		Send, {CtrlDown}f{CtrlUp}
	}
				
	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100
			
	; find the edition
	SendInput, ="edition"
	Send, {Enter}
	Sleep, 100
				
	; move to the edition label
	SendInput, <mods:caption>
	Send, {Enter}
	Sleep, 100
		
	; if no edition label
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		editionlabel =
	}
				
	else
	{		
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the edition label
		Send, {Right}
		Send, {ShiftDown}{End}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100

		; remove the </mods:caption> tag
		; and store the issue number
		StringTrimRight, editionlabel, clipboard, 15

		WinGetPos, winX, winY, winWidth, winHeight, Metadata
		winX+=%winWidth%

		; create Edition Label GUI
		Gui, 9:+AlwaysOnTop
		Gui, 9:+ToolWindow
		Gui, 9:Font, cRed s15 bold, Arial
		Gui, 9:Add, Text, x15 y15 w380 h25, %editionlabel%
		Gui, 9:Font, cRed s10 bold, Arial
		Gui, 9:Add, GroupBox, x0 y-7 w400 h63,       
		Gui, 9:Show, x%winX% y%winY% h55 w400, Edition

		; wait for Notepad++
		SetTitleMatchMode 2
		WinWait, Notepad++, , , ,
		IfWinNotActive, Notepad++, , , ,
		WinActivate, Notepad++, , , ,
		WinWaitActive, Notepad++, , , ,
		Sleep, 200
			
		; activate Find dialog
		Send, {CtrlDown}f{CtrlUp}
	}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100
				
	; find the publication date
	SendInput, </mods:dateIssued>
	Send, {Enter}

	; close the Find Window
	Send, {Esc}
	Sleep, 100

	; copy the publication date
	Send, {Left}{ShiftDown}{Left 10}{ShiftUp}
	Send, {CtrlDown}c{CtrlUp}
  
	; store the publication date
	date = %clipboard%

	; update the date field
	ControlSetText, Static5, %date%, Metadata

	; clear the questionabledate variable
	questionabledate =

	; activate Find dialog
	Send, {CtrlDown}f{CtrlUp}

	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100

	; find a questionable date
	SendInput, "questionable">
	Send, {Enter}
	Sleep, 100

	; if no questionable date
	ifWinActive, Find, Can't find the text:, ,
	{
		Send, {Enter}
		Sleep, 100
					
		; update the questionable date
		ControlSetText, Static8, %questionabledate%, Metadata		
	}
				
	else
	{
		; close the Find Window
		Send, {Esc}
		Sleep, 100

		; copy the questionable publication date
		Send, {Right}{ShiftDown}{Right 10}{ShiftUp}
		Send, {CtrlDown}c{CtrlUp}
  
		; store the questionable publication date
		questionabledate = %clipboard%
		
		; update the questionable date
		ControlSetText, Static8, %questionabledate%, Metadata
		
		; activate the Find window
		Send, {CtrlDown}f{CtrlUp}
	}
	
	; wait for Find window to load
	WinWaitActive, Find, , , ,
	Sleep, 100
	
	; initialize the numpages variable
	numpages = 0
	
	; count the pages
	Send, pageFileGrp
	Sleep, 100
	Loop
	{
		Send, {Enter}
		Sleep, 100
		
		ifWinActive, Find, Can't find the text:, ,
		{
			; close the Find windows
			Send, {Enter}
			Sleep, 100
			Send, {Esc}
			
			; end the page count loop
			Break
		}
		
		; add 1 to numpages
		numpages++		
	}

	; update the number of pages
	ControlSetText, Static9, %numpages%, Metadata
	
	; close the issue.xml file
	Send, {AltDown}f{AltUp}
	Sleep, 100
	Send, c

	; get the First Impression unique window id#
	SetTitleMatchMode 2
	WinGet, firstid, ID, First Impression

	; activate First Impression
	SetTitleMatchMode 1
	WinActivate, ahk_id %firstid%
	
	; bring all metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Date
	WinSet, Top,, Questionable
	WinSet, Top,, Volume
	WinSet, Top,, Issue
	WinSet, Top,, Edition
	WinSet, Top,, Timer	

	; restore the clipboard
	clipboard = %temp%
	
	; update the scoreboard
	hotkeys++
	Metadatascore++
	ControlSetText, Static3, METADATA, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static20, %Metadatascore%, NDNP_QR
Return
; =============================METADATA

; =============================ZOOM IN
; masthead view for First Impression
; HotKey = Alt + k
!k::
	; get the First Impression unique window id#
	SetTitleMatchMode 2
	WinGet, firstid, ID, First Impression

	; activate First Impression
	SetTitleMatchMode 1
	WinActivate, ahk_id %firstid%
	Sleep, 100
	
	; zoom in
	Send, {Enter}
	Sleep, 100
	Send, {NumpadSub 20}
	
	; bring all metadata windows to front
	WinSet, Top,, Metadata	
	WinSet, Top,, Date
	WinSet, Top,, Questionable
	WinSet, Top,, Volume
	WinSet, Top,, Issue
	WinSet, Top,, Edition
	WinSet, Top,, Timer	
	
	; update the scoreboard
	hotkeys++
	zoominscore++
	ControlSetText, Static3, ZOOM IN, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static21, %zoominscore%, NDNP_QR
Return
; =============================ZOOM IN

; =============================ZOOM OUT
; full page view for First Impression
; HotKey = Alt + l
!l::
	; get the First Impression unique window id#
	SetTitleMatchMode 2
	WinGet, firstid, ID, First Impression

	; activate First Impression
	SetTitleMatchMode 1
	WinActivate, ahk_id %firstid%
	Sleep, 100
	
	; zoom out
	Send, {Enter}
	Sleep, 100
	Send, {Enter}
	
	; bring all metadata windows to front
	WinSet, Top,, Metadata	
	WinSet, Top,, Date
	WinSet, Top,, Questionable
	WinSet, Top,, Volume
	WinSet, Top,, Issue
	WinSet, Top,, Edition
	WinSet, Top,, Timer	
		
	; update the scoreboard
	hotkeys++
	zoomoutscore++
	ControlSetText, Static3, ZOOM OUT, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static22, %zoomoutscore%, NDNP_QR
Return
; =============================ZOOM OUT

; =============================VIEW ISSUE XML
; opens the issue.xml file in the default application
; Hotkey: Alt + q
!q::
	SetTitleMatchMode RegEx

	; if issue folder already active
	IfWinActive, ^[1-2][0-9]{9}$, , , ,
	{
		; open the issue.xml
		Send, {End}
		Sleep, 100
		Send, {Up}
		Send, {Enter}
	}
	
	else
	{
		; activate the current NDNP issue folder
		IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
		WinActivate, ^[1-2][0-9]{9}$, , , ,
		WinWaitActive, ^[1-2][0-9]{9}$, , , ,
		Sleep, 100

		; open the issue.xml
		Send, {End}
		Sleep, 100
		Send, {Up}
		Send, {Enter}
	}
		
	; update the scoreboard
	hotkeys++
	viewissuexmlscore++
	ControlSetText, Static3, ViewXML, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static23, %viewissuexmlscore%, NDNP_QR
Return  
; =============================VIEW ISSUE XML

; =============================EDIT ISSUE XML
; opens issue.xml with Notepad++
; Hotkey: Alt + w
!w::
	SetTitleMatchMode RegEx

	; if issue folder already active
	IfWinActive, ^[1-2][0-9]{9}$, , , ,
	{
		; select the issue.xml file
		Send, {End}
		Sleep, 100
		Send, {Up}
		Sleep, 100
  
		; open the issue.xml file in Notepad++
		Send, {AppsKey}
		Sleep, 200
		Send, n
		Sleep, 100
		Send, {Enter}	
	}
	
	else
	{
		; activate the current NDNP issue folder
		IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
		WinActivate, ^[1-2][0-9]{9}$, , , ,
		WinWaitActive, ^[1-2][0-9]{9}$, , , ,
		Sleep, 100
	
		; select the issue.xml file
		Send, {End}
		Sleep, 100
		Send, {Up}
		Sleep, 100
  
		; open the issue.xml file in Notepad++
		Send, {AppsKey}
		Sleep, 200
		Send, n
		Sleep, 100
		Send, {Enter}
	}
		
	; update the scoreboard
	hotkeys++
	editissuexmlscore++
	ControlSetText, Static3, EditXML, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static24, %editissuexmlscore%, NDNP_QR
Return
; =============================EDIT ISSUE XML
; =========================================SCRIPTS
; =================================================


; ========================================================
; =========================================MENU FUNCTIONS
; =================FILE
; open a batch folder
; edit the file path variable in line 26 for your system
OpenBatch:
	; create dialog to select the batch folder
	FileSelectFolder, batchpath, %batchdrive%, 0, Select a folder:
	If ErrorLevel
		Return
	Else
		Run, %batchpath%
Return

; validate a batch with the command line
; NDNP_QR.AHK or .EXE must reside in the DVV folder
DVVvalidate:
	FileSelectFolder, batchpath, %batchdrive%, 0, `nSelect the batch to VALIDATE:
	if ErrorLevel
		Return
	else
		Run, %WINDIR%\System32\cmd.exe /k validationprocessor.bat batch %batchpath%\batch.xml update
Return

; verify a batch with the command line
; NDNP_QR.AHK or .EXE must reside in the DVV folder
DVVverify:
	FileSelectFolder, batchpath, %batchdrive%, 0, `nSelect the batch to VERIFY:
	if ErrorLevel
		Return
	else
		Run, %WINDIR%\System32\cmd.exe /k validationprocessor.bat batch %batchpath%\batch_1.xml verify
Return

; reload application
Reload:
Reload

; exit application
Exit:
ExitApp
; =================FILE

; ======================TOOLS
; ======REEL REPORT
; tool to create a report of all the issues in a reel folder
ReelReport:
	; update last hotkey & loop label
	ControlSetText, Static3, REPORT, NDNP_QR
	ControlSetText, Static4, Reel Report, NDNP_QR

	; initialize the issue and total pages counters
	reportcount = 0
	totalpages = 0

	; create dialog to select a reel folder
	FileSelectFolder, folderpath, %batchdrive%, 0, REEL REPORT`n`nSelect a REEL folder:
	if ErrorLevel
		Return
	else
	{		
		; create input box to enter the delay time
		InputBox, delay, Delay,,, 40, 100,,,,,8
		if ErrorLevel
			Return
		else
		{
			; close any First Impression windows
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
			
			; store the start time
			start = %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% 

			; create a variable for the report path (in LCCN folder)
			StringTrimRight, reportpath, folderpath, 12
			
			; get the reel number from the reel folder path
			StringRight, reelnumber, folderpath, 11

			; create the report
			FileAppend, %folderpath%`n`nPages`tDate`t`tVolume`tIssue`n`n, %reportpath%\%reelnumber%-report.txt

			; open the reel folder
			Run, %folderpath%

			; activate the reel folder
			SetTitleMatchMode RegEx
			WinWait, ^[0-9]{10}[A0-9]$, , , ,
			IfWinNotActive, ^[0-9]{10}[A0-9]$, , , ,
			WinActivate, ^[0-9]{10}[A0-9]$, , , ,
			WinWaitActive, ^[0-9]{10}[A0-9]$, , , ,
			Sleep, 100

			; open the first issue folder
			Send, {Down}
			Sleep, 200		
			Send, {Up}
			Sleep, 200
			Send, {Enter}
		
			Loop
			{
				; *************************************************
				; ALTERED FILE FAILSAFE
				SetTitleMatchMode 2	
				ifWinExist, Save
				{
					WinActivate, Save, , , ,
					WinWaitActive, Save, , , ,
					Sleep, 100

					; close Save dialog without saving changes
					Send, n
					Sleep, 100
	
					; add an error message to the report
					FileAppend, `n`nReport for reel %reelnumber% aborted.`n`nIssues: %reportcount%`n Pages: %totalpages%`nCHECK FOR ERRORS: %date%, %reportpath%\%reelnumber%-report.txt

					; print the start and end times
					FileAppend, `n`nSTART: %start%, %reportpath%\%reelnumber%-report.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%, %reportpath%\%reelnumber%-report.txt

					; close the metadata windows
					Gui, 4:Destroy
					Gui, 5:Destroy
					Gui, 6:Destroy
					Gui, 7:Destroy
					Gui, 9:Destroy

					; create a message box to indicate the script was aborted
					MsgBox, The report for reel %reelnumber% was aborted.`n`nIssues: %reportcount%`nPages: %totalpages%`n`nCHECK FOR ERRORS: %date%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%
					
				; exit the script
				Return
				}
				; *************************************************
			
				; activate the issue folder
				SetTitleMatchMode RegEx
				WinWait, ^[1-2][0-9]{9}$, , 5, ,
				if ErrorLevel
				; *************************************************
				; AUTO EXIT FUNCTION
				; executes at end of issue folders
				{
					; add the issue count to the report
					FileAppend, `nIssues: %reportcount%`n Pages: %totalpages%, %reportpath%\%reelnumber%-report.txt

					; print the start and end times
					FileAppend, `n`nSTART: %start%, %reportpath%\%reelnumber%-report.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%, %reportpath%\%reelnumber%-report.txt

					; close the metadata windows
					Gui, 4:Destroy
					Gui, 5:Destroy
					Gui, 6:Destroy
					Gui, 7:Destroy
					Gui, 9:Destroy

					; create a message box to indicate that the script ended
					MsgBox, Reel: %reelnumber%`nIssues: %reportcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%

				;exit the script
				return
				}
				; *************************************************						
				IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
				WinActivate, ^[1-2][0-9]{9}$, , , ,
				WinWaitActive, ^[1-2][0-9]{9}$, , , ,
				Sleep, 100

				; open the first TIFF
				Send, {Down}
				Sleep, 100
				Send, {Down}
				Sleep, 100
				Send, {Enter}
				Sleep, 300
				
				; activate the current NDNP issue folder
				SetTitleMatchMode RegEx
				WinWait, ^[1-2][0-9]{9}$, , , ,
				IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
				WinActivate, ^[1-2][0-9]{9}$, , , ,
				WinWaitActive, ^[1-2][0-9]{9}$, , , ,
				Sleep, 100
				
				; open issue.xml in Notepad++
				Send, {End}
				Sleep, 100
				Send, {Up}
				Sleep, 100
				Send, {AppsKey}
				Sleep, 100
				Send, n
				Sleep, 100
				Send, {Enter}

				; wait for Notepad++
				SetTitleMatchMode 2
				WinWait, Notepad++, , , ,
				IfWinNotActive, Notepad++, , , ,
				WinActivate, Notepad++, , , ,
				WinWaitActive, Notepad++, , , ,
				Sleep, 200
  
				; activate Find dialog
				Send, {CtrlDown}f{CtrlUp}

				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100

				; find the volume number
				SendInput, ="volume"
				Send, {Enter}
				Sleep, 100

				; if no volume number
				ifWinActive, Find, Can't find the text:, ,
				{
					Send, {Enter}
					volume = none
		
					; update the volume number
					ControlSetText, Static6, %volume%, Metadata
					ControlSetText, Static1, %volume%, Volume
				}
				
				else
				{
					; move to the number
					SendInput, <mods:number>
					Send, {Enter}
					Sleep, 100
			
					; close the Find Window
					Send, {Esc}
					Sleep, 100

					; copy the volume number
					Send, {Right}
					Send, {ShiftDown}{End}{ShiftUp}
					Send, {CtrlDown}c{CtrlUp}
					Sleep, 100

					; remove the </mods:number> tag
					; and store the volume number
					StringTrimRight, volume, clipboard, 14

					; update the volume number
					ControlSetText, Static6, %volume%, Metadata
					ControlSetText, Static1, %volume%, Volume

					; activate Find dialog
					Send, {CtrlDown}f{CtrlUp}
				}
				
				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100

				; find the issue number
				SendInput, ="issue"
				Send, {Enter}
				Sleep, 100
				
				; if no issue number
				ifWinActive, Find, Can't find the text:, ,
				{
					Send, {Enter}
					issue = none
		
					; update the issue number
					ControlSetText, Static7, %issue%, Metadata		
					ControlSetText, Static1, %issue%, Issue
				}
				
				else
				{
					; move to the number
					SendInput, <mods:number>
					Send, {Enter}
					Sleep, 100
		
					; close the Find Window
					Send, {Esc}
					Sleep, 100

					; copy the issue number
					Send, {Right}
					Send, {ShiftDown}{End}{ShiftUp}
					Send, {CtrlDown}c{CtrlUp}
					Sleep, 100

					; remove the </mods:number> tag
					; and store the issue number
					StringTrimRight, issue, clipboard, 14

					; update the issue number
					ControlSetText, Static7, %issue%, Metadata
					ControlSetText, Static1, %issue%, Issue

					; activate Find dialog
					Send, {CtrlDown}f{CtrlUp}
				}
				
				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100
				
				; find the edition
				SendInput, ="edition"
				Send, {Enter}
				Sleep, 100
				
				; move to the edition label
				SendInput, <mods:caption>
				Send, {Enter}
				Sleep, 100
		
				; if no edition label
				ifWinActive, Find, Can't find the text:, ,
				{
					Send, {Enter}
					editionlabel =
				}
				
				else
				{		
					; close the Find Window
					Send, {Esc}
					Sleep, 100

					; copy the edition label
					Send, {Right}
					Send, {ShiftDown}{End}{ShiftUp}
					Send, {CtrlDown}c{CtrlUp}
					Sleep, 100

					; remove the </mods:caption> tag
					; and store the issue number
					StringTrimRight, editionlabel, clipboard, 15

					if reportcount > 1
					{
						WinGetPos, winX, winY, winWidth, winHeight, Date
						winX+=%winWidth%

						; create Edition Label GUI
						Gui, 9:+AlwaysOnTop
						Gui, 9:+ToolWindow
						Gui, 9:Font, cRed s15 bold, Arial
						Gui, 9:Add, Text, x15 y15 w380 h25, %editionlabel%
						Gui, 9:Font, cRed s10 bold, Arial
						Gui, 9:Add, GroupBox, x0 y-7 w400 h63,       
						Gui, 9:Show, x%winX% y%winY% h55 w400, Edition
					}

					; wait for Notepad++
					SetTitleMatchMode 2
					WinWait, Notepad++, , , ,
					IfWinNotActive, Notepad++, , , ,
					WinActivate, Notepad++, , , ,
					WinWaitActive, Notepad++, , , ,
					Sleep, 200
					
					; activate Find dialog
					Send, {CtrlDown}f{CtrlUp}
				}

				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100

				; find the publication date
				SendInput, </mods:dateIssued>
				Send, {Enter}

				; close the Find Window
				Send, {Esc}
				Sleep, 100

				; copy the publication date
				Send, {Left}{ShiftDown}{Left 10}{ShiftUp}
				Send, {CtrlDown}c{CtrlUp}
  
				; store the publication date
				date = %clipboard%

					; create display date variables
					month := SubStr(date, 6, 2)
						if (month = 01) {
							monthname = Jan.
						}
						else if (month = 02) {
							monthname = Feb.
						}
						else if (month = 03) {
							monthname = Mar.
						}
						else if (month = 04) {
							monthname = Apr.
						}
						else if (month = 05) {
							monthname = May
						}
						else if (month = 06) {
							monthname = June
						}
						else if (month = 07) {
							monthname = July
						}
						else if (month = 08) {
							monthname = Aug.
						}
						else if (month = 09) {
							monthname = Sept.
						}
						else if (month = 10) {
							monthname = Oct.
						}
						else if (month = 11) {
							monthname = Nov.
						}
						else if (month = 12) {
							monthname = Dec.
						}
					day := SubStr(date, 9)
					if SubStr(day, 1, 1) = 0
					{
						day := SubStr(day, 2)
					}
					year := SubStr(date, 1, 4)

				; update the date field
				ControlSetText, Static5, %date%, Metadata
				ControlSetText, Static1, %monthname% %day%`, %year%, Date

				; clear the questionabledate variable
				questionabledate =
				questionabledisplay =

				; activate Find dialog
				Send, {CtrlDown}f{CtrlUp}

				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100

				; find a questionable date
				SendInput, "questionable">
				Send, {Enter}
				Sleep, 100

				; if no questionable date
				ifWinActive, Find, Can't find the text:, ,
				{
					Send, {Enter}
					Sleep, 100
					
					; update the questionable date
					ControlSetText, Static8, %questionabledate%, Metadata		
				}
				
				else
				{
					; close the Find Window
					Send, {Esc}
					Sleep, 100

					; copy the questionable publication date
					Send, {Right}{ShiftDown}{Right 10}{ShiftUp}
					Send, {CtrlDown}c{CtrlUp}
  
					; store the questionable publication date
					questionabledate = %clipboard%
					questionabledisplay = Questionable: %questionabledate%
		
					; create display questionable date variables
					month := SubStr(questionabledate, 6, 2)
						if (month = 01) {
							monthnameQ = Jan.
						}
						else if (month = 02) {
							monthnameQ = Feb.
						}
						else if (month = 03) {
							monthnameQ = Mar.
						}
						else if (month = 04) {
							monthnameQ = Apr.
						}
						else if (month = 05) {
							monthnameQ = May
						}
						else if (month = 06) {
							monthnameQ = June
						}
						else if (month = 07) {
							monthnameQ = July
						}
						else if (month = 08) {
							monthnameQ = Aug.
						}
						else if (month = 09) {
							monthnameQ = Sept.
						}
						else if (month = 10) {
							monthnameQ = Oct.
						}
						else if (month = 11) {
							monthnameQ = Nov.
						}
						else if (month = 12) {
							monthnameQ = Dec.
						}
					dayQ := SubStr(questionabledate, 9)
					if SubStr(dayQ, 1, 1) = 0
					{
						dayQ := SubStr(dayQ, 2)
					}
					yearQ := SubStr(questionabledate, 1, 4)

					; update the questionable date
					ControlSetText, Static8, %questionabledate%, Metadata

					if reportcount > 1
					{
						WinGetPos, winX, winY, winWidth, winHeight, Date
						winY+=%winHeight%

						; create Questionable Date GUI
						Gui, 7:+AlwaysOnTop
						Gui, 7:+ToolWindow
						Gui, 7:Font, cRed s15 bold, Arial
						Gui, 7:Add, Text, x35 y15 w160 h25, %monthnameQ% %dayQ%, %yearQ%
						Gui, 7:Font, cRed s10 bold, Arial
						Gui, 7:Add, GroupBox, x0 y-7 w200 h63,       
						Gui, 7:Show, x%winX% y%winY% h55 w200, Questionable
					}
					
					; wait for Notepad++
					SetTitleMatchMode 2
					WinWait, Notepad++, , , ,
					IfWinNotActive, Notepad++, , , ,
					WinActivate, Notepad++, , , ,
					WinWaitActive, Notepad++, , , ,
					Sleep, 200

					; activate the Find window
					Send, {CtrlDown}f{CtrlUp}
				}
	
				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100
	
				; initialize the numpages variable
				numpages = 0
				
				; count the pages
				Send, pageFileGrp
				Sleep, 100
				Loop
				{
					Send, {Enter}
					Sleep, 100
		
					ifWinActive, Find, Can't find the text:, ,
					{
						; close the Find windows
						Send, {Enter}
						Sleep, 100
						Send, {Esc}
			
						; end the page count loop
						Break
					}
		
					; add 1 to numpages & totalpages
					numpages++
				}

				; update total pages count
				totalpages+=%numpages%

				; update the number of pages
				ControlSetText, Static9, %numpages%, Metadata
				
				; close issue.xml
				Send, {AltDown}f{AltUp}
				Sleep, 100
				Send, c
			
				; add the data to the report
				FileAppend, %numpages%`t%date%`t%volume%`t%issue%`t%editionlabel%`t%questionabledisplay%`n, %reportpath%\%reelnumber%-report.txt
				Sleep, 100
	
				; update the scoreboard
				reportcount++
				ControlSetText, Static15, %reportcount%, NDNP_QR	
				
				SetTitleMatchMode 2
				WinGet, firstid, ID, First Impression

				; activate First Impression
				SetTitleMatchMode 1
				IfWinNotActive, ahk_id %firstid%
				WinActivate, ahk_id %firstid%
				WinWaitActive, ahk_id %firstid%
				Sleep, 200			
  
				; zoom out
				Send, {NumpadSub 20}
				Sleep, 300
				
				; send First Impression to bottom of stack
				WinSet, Bottom,, ahk_id %firstid%
				
				; create GUIs for Date, Volume, Issue, Questionable, and Edition Label if first pass
				if reportcount = 1
				{
					WinGetPos, winX, winY, winWidth, winHeight, Metadata
					winX+=%winWidth%

					; create Date GUI
					; this window may be closed without exiting the loop
					Gui, 4:+AlwaysOnTop
					Gui, 4:+ToolWindow
					Gui, 4:Font, cRed s15 bold, Arial
					Gui, 4:Add, Text, x35 y15 w160 h25, %monthname% %day%, %year%
					Gui, 4:Font, cRed s10 bold, Arial
					Gui, 4:Add, GroupBox, x0 y-7 w200 h63,       
					Gui, 4:Show, x%winX% y%winY% h55 w200, Date
		
					; create Volume GUI
					; this window may be closed without exiting the loop
					Gui, 5:+AlwaysOnTop
					Gui, 5:+ToolWindow
					Gui, 5:Font, cRed s25 bold, Arial
					Gui, 5:Add, GroupBox, x0 y-18 w200 h74,       
					Gui, 5:Add, Text, x15 y8 w170 h35, %volume%
					Gui, 5:Show, x%winX% y%winY% h55 w200, Volume

					; create Issue GUI
					; this window may be closed without exiting the loop
					Gui, 6:+AlwaysOnTop
					Gui, 6:+ToolWindow
					Gui, 6:Font, cRed s25 bold, Arial
					Gui, 6:Add, GroupBox, x0 y-18 w200 h74,       
					Gui, 6:Add, Text, x15 y8 w170 h35, %issue%
					Gui, 6:Show, x%winX% y%winY% h55 w200, Issue

					if questionabledate !=
					{
						WinGetPos, winX, winY, winWidth, winHeight, Date
						winY+=%winHeight%

						; create Questionable Date GUI
						Gui, 7:+AlwaysOnTop
						Gui, 7:+ToolWindow
						Gui, 7:Font, cRed s15 bold, Arial
						Gui, 7:Add, Text, x35 y15 w160 h25, %monthnameQ% %dayQ%, %yearQ%
						Gui, 7:Font, cRed s10 bold, Arial
						Gui, 7:Add, GroupBox, x0 y-7 w200 h63,       
						Gui, 7:Show, x%winX% y%winY% h55 w200, Questionable
					}

					if editionlabel !=
					{
						WinGetPos, winX, winY, winWidth, winHeight, Date
						winX+=%winWidth%

						; create Edition Label GUI
						Gui, 9:+AlwaysOnTop
						Gui, 9:+ToolWindow
						Gui, 9:Font, cRed s15 bold, Arial
						Gui, 9:Add, Text, x15 y15 w380 h25, %editionlabel%
						Gui, 9:Font, cRed s10 bold, Arial
						Gui, 9:Add, GroupBox, x0 y-7 w400 h63,       
						Gui, 9:Show, x%winX% y%winY% h55 w400, Edition
					}				

					MsgBox, THE REPORT LOOP IS PAUSED`n`nPosition the metadata windows as desired.`n`nClick OK and the loop will continue in %delay% seconds.
				}
				
				; bring all metadata windows to front
				WinSet, Top,, Metadata
				WinSet, Top,, Date
				WinSet, Top,, Questionable
				WinSet, Top,, Volume
				WinSet, Top,, Issue
				WinSet, Top,, Edition
				
				; create Delay Timer
				seconds = %delay%
				WinGetPos, winX, winY, winWidth, winHeight, Metadata
				winX+=%winWidth%
				Gui, 8:+AlwaysOnTop
				Gui, 8:+ToolWindow
				Gui, 8:Font, cGreen s25 bold, Arial
				Gui, 8:Add, GroupBox, x0 y-18 w50 h74,       
				Gui, 8:Add, Text, x15 y8 w30 h35, %seconds%
				Gui, 8:Show, x%winX% y%winY% h55 w50, Timer

				; start delay timer
				Loop
				{
					Sleep, 1000
					seconds--
					ControlSetText, Static1, %seconds%, Timer
					if (seconds = 0)
					{
						Gui, 8:Destroy
						Break
					}
				}
							
				; open Note dialog if Right Shift is held down
				getKeyState, state, RShift
				if state = D
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
  
						; activate the issue folder
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

						; activate the reel folder
						SetTitleMatchMode RegEx
						WinWait, ^[0-9]{10}[0-9A]$, , , ,
						IfWinNotActive, ^[0-9]{10}[0-9A]$, , , ,
						WinActivate, ^[0-9]{10}[0-9A]$, , , ,
						WinWaitActive, ^[0-9]{10}[0-9A]$, , , ,
						Sleep, 100
	
						; open the next issue folder
						Send, {Down}
						Sleep, 100
						Send, {Enter}
					
					Continue
					}

					else
					{
						; add the note to the report
						FileAppend, Note: %note%`n, %reportpath%\%reelnumber%-report.txt

						; close the Questionable Date & Edition Label windows
						Gui, 7:Destroy
						Gui, 9:Destroy
					}
				}

				; *************************************************
				; MANUAL EXIT FUNCTION
				; executes if Left Shift is held down during page display delay
				getKeyState, state, LShift
				if state = D
				{
					; add the issue count to the report
					FileAppend, `nReport for reel %reelnumber% was cancelled.`n`nIssues: %reportcount%`n Pages: %totalpages%, %reportpath%\%reelnumber%-report.txt

					; print the start and end times
					FileAppend, `n`nSTART: %start%, %reportpath%\%reelnumber%-report.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%, %reportpath%\%reelnumber%-report.txt

					; close the metadata windows
					Gui, 4:Destroy
					Gui, 5:Destroy
					Gui, 6:Destroy
					Gui, 7:Destroy
					Gui, 9:Destroy

					; create a message box to indicate the script was cancelled
					MsgBox, The report for reel %reelnumber% was cancelled.`n`nIssues: %reportcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%
				;exit the script
				return
				}
				; *************************************************						

				; close the Questionable Date & Edition Label windows
				Gui, 7:Destroy
				Gui, 9:Destroy

				; close any First Impression windows
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
  
				; activate the issue folder
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

				; activate the reel folder
				SetTitleMatchMode RegEx
				WinWait, ^[0-9]{10}[0-9A]$, , , ,
				IfWinNotActive, ^[0-9]{10}[0-9A]$, , , ,
				WinActivate, ^[0-9]{10}[0-9A]$, , , ,
				WinWaitActive, ^[0-9]{10}[0-9A]$, , , ,
				Sleep, 100
	
				; open the next issue folder
				Send, {Down}
				Sleep, 100
				Send, {Enter}
			}
		}
	}
Return
; ======REEL REPORT

; ======BATCH REPORT
BatchReport:
	; delay time (milliseconds) for loading issue.xml files
	delay = 1000
	
	; update last hotkey & report label
	ControlSetText, Static3, REPORT, NDNP_QR
	ControlSetText, Static4, Batch Report, NDNP_QR

	; initialize the issue and total pages counters
	reportcount = 0
	totalpages = 0

	; create dialog to select a batch folder
	FileSelectFolder, folderpath, %batchdrive%, 0, BATCH REPORT`n`nSelect a BATCH folder:
	if ErrorLevel
		Return
	else
	{		
		; enter the batch name and press OK (Cancel ends the script)
		InputBox, input, Batch Name,,, 150, 100,,,,,%batchname%
		if ErrorLevel
			Return
		else
		{
			; store the batch name
			batchname = %input%

			; store the start time
			start = %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% 
			
			; create the report file in the batch folder
			FileAppend, %folderpath%`n`nIdentifier`t`tPages`tDate`t`tVolume`tIssue`n`n, %folderpath%\%batchname%-report.txt

			; open the batch.xml file in Notepad++
			Run, "%notepadpath%\notepad++.exe" "%folderpath%\batch.xml"

			; initialize the issue counter
			issuecount = 0
			
			; initialize the issuefile
			issuefile =
	
			; *******************************************************
			; loop through every issue listed in the batch.xml file *
			; and store file paths in the issuefile variable        *
			;                                                       *
			; Notepad++ Find dialog:                                *
			;   "Wrap Around" must be UNCHECKED                     *
			;   or the loop will continue infinitely                *
			; *******************************************************
			Loop
			{
				; *************************************************
				; ALTERED FILE FAILSAFE
				SetTitleMatchMode 2	
				ifWinExist, Save
				{
					; close Save dialog without saving changes
					Send, n
					Sleep, 100

					; add an error message to the report including the last file opened
					FileAppend, `nReport for batch %batchname% aborted.`n`nIssues: %issuecount%`n Pages: %totalpages%, %folderpath%\%batchname%-report.txt
					FileAppend, `nCHECK FOR ERRORS: %issuepath%, %folderpath%\%batchname%-report.txt

					; print the start and end times
					FileAppend, `n`nSTART: %start%, %folderpath%\%batchname%-report.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%, %folderpath%\%batchname%-report.txt
		
					; create a message box to indicate the script was aborted
					MsgBox, The report for batch %batchname% was aborted.`n`nIssues: %issuecount%`nPages: %totalpages%`n`nCHECK FOR ERRORS: %issuepath%`n`nSTART:`t%start%`n`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%

				; exit the script
				return
				}
				; *************************************************

				; activate batch.xml window
				SetTitleMatchMode 2
				WinWait, batch.xml, , 10, ,
				if ErrorLevel
				; *************************************************
				; TIMEOUT EXIT FUNCTION
				; executes if batch.xml window cannot be found
				{
					; add the issue count to the report
					FileAppend, `nBatch: %batchname%`nIssues: %issuecount%`n Pages: %totalpages%, %folderpath%\%batchname%-report.txt

					; print the start and end times
					FileAppend, `n`nSTART: %start%, %folderpath%\%batchname%-report.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%, %folderpath%\%batchname%-report.txt

					; create a message box to indicate that the script ended
					MsgBox, The report for batch %batchname% timed out.`n`nIssues: %issuecount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%

				;exit the script
				return
				}
				; *************************************************
				IfWinNotActive, batch.xml, , , ,
				WinActivate, batch.xml, , , ,
				WinWaitActive, batch.xml, , , ,
				Sleep, 100

				; activate the Find dialog and locate the next issue.xml file
				Send, {CtrlDown}f{CtrlUp}
				Sleep, 50
				Send, editionOrder
				Sleep, 50
				Send, {Enter}
				Sleep, 50
				
				; if no more issues in batch.xml
				ifWinActive, Find, Can't find the text:, ,
				{
					Send, {Enter}
					Sleep, 50
					
					; close the Find dialog
					Send, {Esc}
					Sleep, 50
										
					; close the batch.xml
					Send, {AltDown}f{AltUp}
					Sleep, 50
					Send, c

					; end the loop
					Break
				}
				
				; close the Find dialog
				Send, {Esc}
				Sleep, 50
				
				; copy the issue.xml file path
				Send, {Right 7}
				Sleep, 50
				Send, {ShiftDown}{End}{ShiftUp}
				Sleep, 50	
				Send, {CtrlDown}c{CtrlUp}
				Sleep, 50
	
				; remove the </issue> tag and store the issue.xml path
				StringTrimRight, issuepath, clipboard, 8

				; append issue path to issuefile
				issuefile .= issuepath
					
				; append new line to issuefile
				issuefile .= "`n"
				
				; update the issue count
				issuecount++
			}
			
			; sort the issuefile variable
			Sort, issuefile

			; *****************************************
			; loop through sorted issuefile           *
			; extract metadata and add to report      *
			;                                         *
			; Notepad++ Find dialog:                  *
			;   "Wrap Around" must be UNCHECKED       *
			;   or the loop will continue infinitely  *
			; *****************************************
			Loop, parse, issuefile, `n
			{
				; *************************************************
				; AUTO EXIT FUNCTION
				; exit script if issuefile is empty
				if A_LoopField =
				{
					; add the issue count to the report
					FileAppend, `nBatch: %batchname%`nIssues: %issuecount%`n Pages: %totalpages%, %folderpath%\%batchname%-report.txt

					; print the start and end times
					FileAppend, `n`nSTART: %start%, %folderpath%\%batchname%-report.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%, %folderpath%\%batchname%-report.txt

					; create a message box to indicate that the script ended
					MsgBox, Batch: %batchname%`nIssues: %issuecount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%

				;exit the script
				return
				}
				; *************************************************
				
				; *************************************************
				; ALTERED FILE FAILSAFE
				SetTitleMatchMode 2	
				ifWinExist, Save
				{
					; close Save dialog without saving changes
					Send, n
					Sleep, 100

					; add an error message to the report including the last file opened
					FileAppend, `nReport for batch %batchname% aborted.`n`nIssues: %issuecount%`n Pages: %totalpages%, %folderpath%\%batchname%-report.txt
					FileAppend, `nCHECK FOR ERRORS: %issuepath%, %folderpath%\%batchname%-report.txt
	
					; print the start and end times
					FileAppend, `n`nSTART: %start%, %folderpath%\%batchname%-report.txt
					FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%, %folderpath%\%batchname%-report.txt
		
					; create a message box to indicate the script was aborted
					MsgBox, The report for batch %batchname% was aborted.`n`nIssues: %issuecount%`nPages: %totalpages%`n`nCHECK FOR ERRORS: %issuepath%`n`nSTART:`t%start%`n`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%

				; exit the script
				return
				}
				; *************************************************
			
				; remove the issue.xml file name and store the issue folder path
				StringTrimRight, issuefolderpath, A_LoopField, 15
	
				; remove the dates and store the LCCN and Reel #
				StringTrimRight, identifier, A_LoopField, 26
				Sleep, 100	
	
				; open the issue.xml file in Notepad++
				Run, "%notepadpath%\notepad++.exe" "%folderpath%\%A_LoopField%"
				Sleep, %delay%
	
				; HARVEST ISSUE DATA: number of pages, volume #, issue #, publication date, questionable date
	
				; initialize the number of pages variable
				numpages = 0
				
				; count the number of .TIF files
				Loop, %folderpath%\%issuefolderpath%\*.tif
				{
					numpages++
				}
				
				; add the page count to the total
				totalpages += numpages
				
				; update the number of pages
				ControlSetText, Static9, %numpages%, Metadata

				; activate Find dialog
				Send, {CtrlDown}f{CtrlUp}

				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100

				; find the volume number
				SendInput, ="volume"
				Send, {Enter}
				Sleep, 100

				; if no volume number
				ifWinActive, Find, Can't find the text:, ,
				{
					Send, {Enter}
					volume = none
					
					; update the volume number
					ControlSetText, Static6, %volume%, Metadata
				}
				
				else
				{
					; move to the number
					SendInput, <mods:number>
					Send, {Enter}
					Sleep, 100
			
					; close the Find Window
					Send, {Esc}
					Sleep, 100

					; copy the volume number
					Send, {Right}
					Send, {ShiftDown}{End}{ShiftUp}
					Send, {CtrlDown}c{CtrlUp}
					Sleep, 100

					; remove the </mods:number> tag
					; and store the volume number
					StringTrimRight, volume, clipboard, 14

					; update the volume number
					ControlSetText, Static6, %volume%, Metadata

					; activate Find dialog
					Send, {CtrlDown}f{CtrlUp}
				}
				
				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100

				; find the issue number
				SendInput, ="issue"
				Send, {Enter}
				Sleep, 100
				
				; if no issue number
				ifWinActive, Find, Can't find the text:, ,
				{
					Send, {Enter}
					issue = none

					; update the issue number
					ControlSetText, Static7, %issue%, Metadata		
				}
				
				else
				{
					; move to the number
					SendInput, <mods:number>
					Send, {Enter}
					Sleep, 100
			
					; close the Find Window
					Send, {Esc}
					Sleep, 100

					; copy the issue number
					Send, {Right}
					Send, {ShiftDown}{End}{ShiftUp}
					Send, {CtrlDown}c{CtrlUp}
					Sleep, 100
  
					; remove the </mods:number> tag
					; and store the issue number
					StringTrimRight, issue, clipboard, 14

					; update the issue number
					ControlSetText, Static7, %issue%, Metadata		

					; activate Find dialog
					Send, {CtrlDown}f{CtrlUp}
				}
				
				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100
				
				; find the edition
				SendInput, ="edition"
				Send, {Enter}
				Sleep, 100
				
				; move to the edition label
				SendInput, <mods:caption>
				Send, {Enter}
				Sleep, 100
		
				; if no edition label
				ifWinActive, Find, Can't find the text:, ,
				{
					Send, {Enter}
					editionlabel =
				}
				
				else
				{		
					; close the Find Window
					Send, {Esc}
					Sleep, 100

					; copy the edition label
					Send, {Right}
					Send, {ShiftDown}{End}{ShiftUp}
					Send, {CtrlDown}c{CtrlUp}
					Sleep, 100

					; remove the </mods:caption> tag
					; and store the issue number
					StringTrimRight, editionlabel, clipboard, 15
					
					; activate Find dialog
					Send, {CtrlDown}f{CtrlUp}
				}

				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100
				
				; find the publication date
				SendInput, </mods:dateIssued>
				Send, {Enter}

				; close the Find Window
				Send, {Esc}
				Sleep, 100

				; copy the publication date
				Send, {Left}{ShiftDown}{Left 10}{ShiftUp}
				Send, {CtrlDown}c{CtrlUp}
  
				; store the publication date
				date = %clipboard%

				; update the date field
				ControlSetText, Static5, %date%, Metadata

				; clear the questionabledate variables
				questionabledate =
				questionabledisplay =

				; activate Find dialog
				Send, {CtrlDown}f{CtrlUp}

				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100

				; find a questionable date
				SendInput, "questionable">
				Send, {Enter}
				Sleep, 100

				; if no questionable date
				ifWinActive, Find, Can't find the text:, ,
				{
					Send, {Enter}
					Sleep, 100
					
					; update the questionable date
					ControlSetText, Static8, %questionabledate%, Metadata
		
					; close the Find Window
					Send, {Esc}
					Sleep, 100
				}
				
				else
				{
					; close the Find Window
					Send, {Esc}
					Sleep, 100

					; copy the questionable publication date
					Send, {Right}{ShiftDown}{Right 10}{ShiftUp}
					Send, {CtrlDown}c{CtrlUp}
  
					; store the questionable publication date
					questionabledate = %clipboard%
					questionabledisplay = Questionable: %questionabledate%
					
					; update the questionable date
					ControlSetText, Static8, %questionabledate%, Metadata
				}

				; close the issue.xml file
				Send, {AltDown}f{AltUp}
				Sleep, 100
				Send, c

				; add the issue data to the report
				FileAppend, %identifier%`t%numpages%`t%date%`t%volume%`t%issue%`t%editionlabel%`t%questionabledisplay%`n, %folderpath%\%batchname%-report.txt
				
				; update the scoreboard
				reportcount++
				ControlSetText, Static15, %reportcount%, NDNP_QR
				Sleep, 100
			}
		}
	}
Return
; ======BATCH REPORT

; ======METADATA VIEWER
MetaViewer:
	; update last hotkey & loop label
	ControlSetText, Static3, VIEWER, NDNP_QR
	ControlSetText, Static4, Meta Viewer, NDNP_QR

	; initialize the issue and total pages counters
	reportcount = 0
	totalpages = 0

	; create dialog to select a reel folder
	FileSelectFolder, folderpath, %batchdrive%, 0, METADATA VIEWER`n`nSelect a REEL folder:
	if ErrorLevel
		Return
	else
	{		
		; create input box to enter the delay time
		InputBox, delay, Delay,,, 40, 100,,,,,5
		if ErrorLevel
			Return
		else
		{
			; close any First Impression windows
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
			
			; store the start time
			start = %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% 
			
			; get the reel number from the reel folder path
			StringRight, reelnumber, folderpath, 11

			; open the reel folder
			Run, %folderpath%

			; activate the reel folder
			SetTitleMatchMode RegEx
			WinWait, ^[0-9]{10}[A0-9]$, , , ,
			IfWinNotActive, ^[0-9]{10}[A0-9]$, , , ,
			WinActivate, ^[0-9]{10}[A0-9]$, , , ,
			WinWaitActive, ^[0-9]{10}[A0-9]$, , , ,
			Sleep, 100

			; open the first issue folder
			Send, {Down}
			Sleep, 200		
			Send, {Up}
			Sleep, 200
			Send, {Enter}
		
			Loop
			{
				; *************************************************
				; ALTERED FILE FAILSAFE
				SetTitleMatchMode 2	
				ifWinExist, Save
				{
					WinActivate, Save, , , ,
					WinWaitActive, Save, , , ,
					Sleep, 100

					; close Save dialog without saving changes
					Send, n
					Sleep, 100
	
					; close the metadata windows
					Gui, 4:Destroy
					Gui, 5:Destroy
					Gui, 6:Destroy
					Gui, 7:Destroy
					Gui, 9:Destroy

					; create a message box to indicate the script was aborted
					MsgBox, The metadata viewer for reel %reelnumber% was aborted.`n`nIssues: %reportcount%`nPages: %totalpages%`n`nCHECK FOR ERRORS: %date%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%
					
				; exit the script
				Return
				}
				; *************************************************
			
				; activate the issue folder
				SetTitleMatchMode RegEx
				WinWait, ^[1-2][0-9]{9}$, , 5, ,
				if ErrorLevel
				; *************************************************
				; AUTO EXIT FUNCTION
				; executes at end of issue folders
				{
					; close the metadata windows
					Gui, 4:Destroy
					Gui, 5:Destroy
					Gui, 6:Destroy
					Gui, 7:Destroy
					Gui, 9:Destroy

					; create a message box to indicate that the script ended
					MsgBox, Reel: %reelnumber%`nIssues: %reportcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%

				;exit the script
				return
				}
				; *************************************************						
				IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
				WinActivate, ^[1-2][0-9]{9}$, , , ,
				WinWaitActive, ^[1-2][0-9]{9}$, , , ,
				Sleep, 100

				; open the first TIFF
				Send, {Down}
				Sleep, 100
				Send, {Down}
				Sleep, 100
				Send, {Enter}
				Sleep, 300
				
				; activate the current NDNP issue folder
				SetTitleMatchMode RegEx
				WinWait, ^[1-2][0-9]{9}$, , , ,
				IfWinNotActive, ^[1-2][0-9]{9}$, , , ,
				WinActivate, ^[1-2][0-9]{9}$, , , ,
				WinWaitActive, ^[1-2][0-9]{9}$, , , ,
				Sleep, 100
				
				; open issue.xml in Notepad++
				Send, {End}
				Sleep, 100
				Send, {Up}
				Sleep, 100
				Send, {AppsKey}
				Sleep, 100
				Send, n
				Sleep, 100
				Send, {Enter}

				; wait for Notepad++
				SetTitleMatchMode 2
				WinWait, Notepad++, , , ,
				IfWinNotActive, Notepad++, , , ,
				WinActivate, Notepad++, , , ,
				WinWaitActive, Notepad++, , , ,
				Sleep, 200
  
				; activate Find dialog
				Send, {CtrlDown}f{CtrlUp}

				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100

				; find the volume number
				SendInput, ="volume"
				Send, {Enter}
				Sleep, 100

				; if no volume number
				ifWinActive, Find, Can't find the text:, ,
				{
					Send, {Enter}
					volume = none
		
					; update the volume number
					ControlSetText, Static6, %volume%, Metadata
					ControlSetText, Static1, %volume%, Volume
				}
				
				else
				{
					; move to the number
					SendInput, <mods:number>
					Send, {Enter}
					Sleep, 100
			
					; close the Find Window
					Send, {Esc}
					Sleep, 100

					; copy the volume number
					Send, {Right}
					Send, {ShiftDown}{End}{ShiftUp}
					Send, {CtrlDown}c{CtrlUp}
					Sleep, 100

					; remove the </mods:number> tag
					; and store the volume number
					StringTrimRight, volume, clipboard, 14

					; update the volume number
					ControlSetText, Static6, %volume%, Metadata
					ControlSetText, Static1, %volume%, Volume

					; activate Find dialog
					Send, {CtrlDown}f{CtrlUp}
				}
				
				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100

				; find the issue number
				SendInput, ="issue"
				Send, {Enter}
				Sleep, 100
				
				; if no issue number
				ifWinActive, Find, Can't find the text:, ,
				{
					Send, {Enter}
					issue = none
		
					; update the issue number
					ControlSetText, Static7, %issue%, Metadata		
					ControlSetText, Static1, %issue%, Issue
				}
				
				else
				{
					; move to the number
					SendInput, <mods:number>
					Send, {Enter}
					Sleep, 100
		
					; close the Find Window
					Send, {Esc}
					Sleep, 100

					; copy the issue number
					Send, {Right}
					Send, {ShiftDown}{End}{ShiftUp}
					Send, {CtrlDown}c{CtrlUp}
					Sleep, 100

					; remove the </mods:number> tag
					; and store the issue number
					StringTrimRight, issue, clipboard, 14

					; update the issue number
					ControlSetText, Static7, %issue%, Metadata
					ControlSetText, Static1, %issue%, Issue

					; activate Find dialog
					Send, {CtrlDown}f{CtrlUp}
				}
				
				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100
				
				; find the edition
				SendInput, ="edition"
				Send, {Enter}
				Sleep, 100
				
				; move to the edition label
				SendInput, <mods:caption>
				Send, {Enter}
				Sleep, 100
		
				; if no edition label
				ifWinActive, Find, Can't find the text:, ,
				{
					Send, {Enter}
					editionlabel =
				}
				
				else
				{		
					; close the Find Window
					Send, {Esc}
					Sleep, 100

					; copy the edition label
					Send, {Right}
					Send, {ShiftDown}{End}{ShiftUp}
					Send, {CtrlDown}c{CtrlUp}
					Sleep, 100

					; remove the </mods:caption> tag
					; and store the issue number
					StringTrimRight, editionlabel, clipboard, 15

					if reportcount > 1
					{
						WinGetPos, winX, winY, winWidth, winHeight, Date
						winX+=%winWidth%

						; create Edition Label GUI
						Gui, 9:+AlwaysOnTop
						Gui, 9:+ToolWindow
						Gui, 9:Font, cRed s15 bold, Arial
						Gui, 9:Add, Text, x15 y15 w380 h25, %editionlabel%
						Gui, 9:Font, cRed s10 bold, Arial
						Gui, 9:Add, GroupBox, x0 y-7 w400 h63,       
						Gui, 9:Show, x%winX% y%winY% h55 w400, Edition
					}

					; wait for Notepad++
					SetTitleMatchMode 2
					WinWait, Notepad++, , , ,
					IfWinNotActive, Notepad++, , , ,
					WinActivate, Notepad++, , , ,
					WinWaitActive, Notepad++, , , ,
					Sleep, 200
					
					; activate Find dialog
					Send, {CtrlDown}f{CtrlUp}
				}

				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100

				; find the publication date
				SendInput, </mods:dateIssued>
				Send, {Enter}

				; close the Find Window
				Send, {Esc}
				Sleep, 100

				; copy the publication date
				Send, {Left}{ShiftDown}{Left 10}{ShiftUp}
				Send, {CtrlDown}c{CtrlUp}
  
				; store the publication date
				date = %clipboard%

					; create display date variables
					month := SubStr(date, 6, 2)
						if (month = 01) {
							monthname = Jan.
						}
						else if (month = 02) {
							monthname = Feb.
						}
						else if (month = 03) {
							monthname = Mar.
						}
						else if (month = 04) {
							monthname = Apr.
						}
						else if (month = 05) {
							monthname = May
						}
						else if (month = 06) {
							monthname = June
						}
						else if (month = 07) {
							monthname = July
						}
						else if (month = 08) {
							monthname = Aug.
						}
						else if (month = 09) {
							monthname = Sept.
						}
						else if (month = 10) {
							monthname = Oct.
						}
						else if (month = 11) {
							monthname = Nov.
						}
						else if (month = 12) {
							monthname = Dec.
						}
					day := SubStr(date, 9)
					if SubStr(day, 1, 1) = 0
					{
						day := SubStr(day, 2)
					}
					year := SubStr(date, 1, 4)

				; update the date field
				ControlSetText, Static5, %date%, Metadata
				ControlSetText, Static1, %monthname% %day%`, %year%, Date

				; clear the questionabledate variable
				questionabledate =
				questionabledisplay =

				; activate Find dialog
				Send, {CtrlDown}f{CtrlUp}

				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100

				; find a questionable date
				SendInput, "questionable">
				Send, {Enter}
				Sleep, 100

				; if no questionable date
				ifWinActive, Find, Can't find the text:, ,
				{
					Send, {Enter}
					Sleep, 100
					
					; update the questionable date
					ControlSetText, Static8, %questionabledate%, Metadata		
				}
				
				else
				{
					; close the Find Window
					Send, {Esc}
					Sleep, 100

					; copy the questionable publication date
					Send, {Right}{ShiftDown}{Right 10}{ShiftUp}
					Send, {CtrlDown}c{CtrlUp}
  
					; store the questionable publication date
					questionabledate = %clipboard%
					questionabledisplay = Questionable: %questionabledate%
		
					; create display questionable date variables
					month := SubStr(questionabledate, 6, 2)
						if (month = 01) {
							monthnameQ = Jan.
						}
						else if (month = 02) {
							monthnameQ = Feb.
						}
						else if (month = 03) {
							monthnameQ = Mar.
						}
						else if (month = 04) {
							monthnameQ = Apr.
						}
						else if (month = 05) {
							monthnameQ = May
						}
						else if (month = 06) {
							monthnameQ = June
						}
						else if (month = 07) {
							monthnameQ = July
						}
						else if (month = 08) {
							monthnameQ = Aug.
						}
						else if (month = 09) {
							monthnameQ = Sept.
						}
						else if (month = 10) {
							monthnameQ = Oct.
						}
						else if (month = 11) {
							monthnameQ = Nov.
						}
						else if (month = 12) {
							monthnameQ = Dec.
						}
					dayQ := SubStr(questionabledate, 9)
					if SubStr(dayQ, 1, 1) = 0
					{
						dayQ := SubStr(dayQ, 2)
					}
					yearQ := SubStr(questionabledate, 1, 4)

					; update the questionable date
					ControlSetText, Static8, %questionabledate%, Metadata

					if reportcount > 1
					{
						WinGetPos, winX, winY, winWidth, winHeight, Date
						winY+=%winHeight%

						; create Questionable Date GUI
						Gui, 7:+AlwaysOnTop
						Gui, 7:+ToolWindow
						Gui, 7:Font, cRed s15 bold, Arial
						Gui, 7:Add, Text, x35 y15 w160 h25, %monthnameQ% %dayQ%, %yearQ%
						Gui, 7:Font, cRed s10 bold, Arial
						Gui, 7:Add, GroupBox, x0 y-7 w200 h63,       
						Gui, 7:Show, x%winX% y%winY% h55 w200, Questionable
					}
					
					; wait for Notepad++
					SetTitleMatchMode 2
					WinWait, Notepad++, , , ,
					IfWinNotActive, Notepad++, , , ,
					WinActivate, Notepad++, , , ,
					WinWaitActive, Notepad++, , , ,
					Sleep, 200

					; activate the Find window
					Send, {CtrlDown}f{CtrlUp}
				}
	
				; wait for Find window to load
				WinWaitActive, Find, , , ,
				Sleep, 100
	
				; initialize the numpages variable
				numpages = 0
				
				; count the pages
				Send, pageFileGrp
				Sleep, 100
				Loop
				{
					Send, {Enter}
					Sleep, 100
		
					ifWinActive, Find, Can't find the text:, ,
					{
						; close the Find windows
						Send, {Enter}
						Sleep, 100
						Send, {Esc}
			
						; end the page count loop
						Break
					}
		
					; add 1 to numpages & totalpages
					numpages++
				}

				; update total pages count
				totalpages+=%numpages%

				; update the number of pages
				ControlSetText, Static9, %numpages%, Metadata
				
				; close issue.xml
				Send, {AltDown}f{AltUp}
				Sleep, 100
				Send, c
			
				; update the scoreboard
				reportcount++
				ControlSetText, Static15, %reportcount%, NDNP_QR	
				
				SetTitleMatchMode 2
				WinGet, firstid, ID, First Impression

				; activate First Impression
				SetTitleMatchMode 1
				IfWinNotActive, ahk_id %firstid%
				WinActivate, ahk_id %firstid%
				WinWaitActive, ahk_id %firstid%
				Sleep, 200			
  
				; zoom out
				Send, {NumpadSub 20}
				Sleep, 300
				
				; send First Impression to bottom of stack
				WinSet, Bottom,, ahk_id %firstid%
				
				; create GUIs for Date, Volume, Issue, Questionable, and Edition Label if first pass
				if reportcount = 1
				{
					WinGetPos, winX, winY, winWidth, winHeight, Metadata
					winX+=%winWidth%

					; create Date GUI
					; this window may be closed without exiting the loop
					Gui, 4:+AlwaysOnTop
					Gui, 4:+ToolWindow
					Gui, 4:Font, cRed s15 bold, Arial
					Gui, 4:Add, Text, x35 y15 w160 h25, %monthname% %day%, %year%
					Gui, 4:Font, cRed s10 bold, Arial
					Gui, 4:Add, GroupBox, x0 y-7 w200 h63,       
					Gui, 4:Show, x%winX% y%winY% h55 w200, Date
		
					; create Volume GUI
					; this window may be closed without exiting the loop
					Gui, 5:+AlwaysOnTop
					Gui, 5:+ToolWindow
					Gui, 5:Font, cRed s25 bold, Arial
					Gui, 5:Add, GroupBox, x0 y-18 w200 h74,       
					Gui, 5:Add, Text, x15 y8 w170 h35, %volume%
					Gui, 5:Show, x%winX% y%winY% h55 w200, Volume

					; create Issue GUI
					; this window may be closed without exiting the loop
					Gui, 6:+AlwaysOnTop
					Gui, 6:+ToolWindow
					Gui, 6:Font, cRed s25 bold, Arial
					Gui, 6:Add, GroupBox, x0 y-18 w200 h74,       
					Gui, 6:Add, Text, x15 y8 w170 h35, %issue%
					Gui, 6:Show, x%winX% y%winY% h55 w200, Issue

					if questionabledate !=
					{
						WinGetPos, winX, winY, winWidth, winHeight, Date
						winY+=%winHeight%

						; create Questionable Date GUI
						Gui, 7:+AlwaysOnTop
						Gui, 7:+ToolWindow
						Gui, 7:Font, cRed s15 bold, Arial
						Gui, 7:Add, Text, x35 y15 w160 h25, %monthnameQ% %dayQ%, %yearQ%
						Gui, 7:Font, cRed s10 bold, Arial
						Gui, 7:Add, GroupBox, x0 y-7 w200 h63,       
						Gui, 7:Show, x%winX% y%winY% h55 w200, Questionable
					}

					if editionlabel !=
					{
						WinGetPos, winX, winY, winWidth, winHeight, Date
						winX+=%winWidth%

						; create Edition Label GUI
						Gui, 9:+AlwaysOnTop
						Gui, 9:+ToolWindow
						Gui, 9:Font, cRed s15 bold, Arial
						Gui, 9:Add, Text, x15 y15 w380 h25, %editionlabel%
						Gui, 9:Font, cRed s10 bold, Arial
						Gui, 9:Add, GroupBox, x0 y-7 w400 h63,       
						Gui, 9:Show, x%winX% y%winY% h55 w400, Edition
					}				

					MsgBox, THE METADATA VIEWER IS PAUSED`n`nPosition the metadata windows as desired.`n`nClick OK and the loop will continue in %delay% seconds.
				}
				
				; bring all metadata windows to front
				WinSet, Top,, Metadata
				WinSet, Top,, Date
				WinSet, Top,, Questionable
				WinSet, Top,, Volume
				WinSet, Top,, Issue
				WinSet, Top,, Edition
				
				; create Delay Timer
				seconds = %delay%
				WinGetPos, winX, winY, winWidth, winHeight, Metadata
				winX+=%winWidth%
				Gui, 8:+AlwaysOnTop
				Gui, 8:+ToolWindow
				Gui, 8:Font, cGreen s25 bold, Arial
				Gui, 8:Add, GroupBox, x0 y-18 w50 h74,       
				Gui, 8:Add, Text, x15 y8 w30 h35, %seconds%
				Gui, 8:Show, x%winX% y%winY% h55 w50, Timer

				; start delay timer
				Loop
				{
					Sleep, 1000
					seconds--
					ControlSetText, Static1, %seconds%, Timer
					if (seconds = 0)
					{
						Gui, 8:Destroy
						Break
					}
				}
							
				; *************************************************
				; MANUAL EXIT FUNCTION
				; executes if Left Shift is held down during page display delay
				getKeyState, state, LShift
				if state = D
				{
					; close the metadata windows
					Gui, 4:Destroy
					Gui, 5:Destroy
					Gui, 6:Destroy
					Gui, 7:Destroy
					Gui, 9:Destroy

					; create a message box to indicate the script was cancelled
					MsgBox, The metadata viewer for reel %reelnumber% was cancelled.`n`nIssues: %reportcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%
				;exit the script
				return
				}
				; *************************************************						

				; close the Questionable Date & Edition Label windows
				Gui, 7:Destroy
				Gui, 9:Destroy

				; close any First Impression windows
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
  
				; activate the issue folder
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

				; activate the reel folder
				SetTitleMatchMode RegEx
				WinWait, ^[0-9]{10}[0-9A]$, , , ,
				IfWinNotActive, ^[0-9]{10}[0-9A]$, , , ,
				WinActivate, ^[0-9]{10}[0-9A]$, , , ,
				WinWaitActive, ^[0-9]{10}[0-9A]$, , , ,
				Sleep, 100
	
				; open the next issue folder
				Send, {Down}
				Sleep, 100
				Send, {Enter}
			}
		}
	}
Return
; ======METADATA VIEWER

; ======DVV PAGES LOOP
; loop to automatically view pages in the DVV
; script by Ann.Howington@unt.edu
; modified for NDNP_QR app
; Pause to pause/restart
; hold down F1 to cancel
DVVpages:
	; create input box to enter the delay time
	InputBox, delay, DVV Pages Loop, Enter the delay in milliseconds:`n1000 milliseconds = 1 second,, 250, 140,,,,,3000
	if ErrorLevel
		Return
	else
	{
		; update last hotkey
		ControlSetText, Static3, PAGES, NDNP_QR
		ControlSetText, Static5, DVV Pages, NDNP_QR
		
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
			ControlSetText, Static16, %count%, NDNP_QR
			
			; delay for the specified time
			Sleep, %delay%
			
			; end the loop if F1 is held down
			if getKeyState("F1")
			break
		}
		
		; close the counter window
		Gui, 3:Destroy
	}
Return
; ======DVV PAGES LOOP

; ======DVV THUMBS LOOP
; loop to view the issue thumbs in the DVV
; script by Ann.Howington@unt.edu
; modified for NDNP_QR app
; Pause to pause/restart
; hold down F1 to cancel
DVVthumbs:
	; create input box to enter the delay time
	InputBox, delay, DVV Thumbs Loop, Enter the delay in milliseconds:`n1000 milliseconds = 1 second,, 250, 140,,,,,7000
	if ErrorLevel
		Return
	else
	{
		; update last hotkey
		ControlSetText, Static3, THUMBS, NDNP_QR
		ControlSetText, Static5, DVV Thumbs, NDNP_QR
		
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
			ControlSetText, Static16, %count%, NDNP_QR

			; delay for the specified time
			Sleep, %delay%
			
			; end the loop if F1 is held down
			if getKeyState("F1")
			break
		}
		
		; close the counter window
		Gui, 3:Destroy
	}
Return
; ======DVV THUMBS LOOP
; ======================TOOLS

; =================SEARCH
; US Newspaper Directory Texas search
DirectoryTitleTexas:
	InputBox, searchstring, Texas Search,,, 250, 100,,,,,Do not use quotes.
	if ErrorLevel
		Return
	else
		Run, http://chroniclingamerica.loc.gov/search/titles/results/?state=Texas&county=&city=&year1=1690&year2=2013&terms=%searchstring%&frequency=&language=&ethnicity=&labor=&material_type=&lccn=&rows=20
Return

; US Newspaper Directory Oklahoma search
DirectoryTitleOklahoma:
	InputBox, searchstring, Oklahoma Search,,, 250, 100,,,,,Do not use quotes.
	if ErrorLevel
		Return
	else
		Run, http://chroniclingamerica.loc.gov/search/titles/results/?state=Oklahoma&county=&city=&year1=1690&year2=2013&terms=%searchstring%&frequency=&language=&ethnicity=&labor=&material_type=&lccn=&rows=20
Return

; US Newspaper Directory NewMexico search
DirectoryTitleNewMexico:
	InputBox, searchstring, New Mexico Search,,, 250, 100,,,,,Do not use quotes.
	if ErrorLevel
		Return
	else
		Run, http://chroniclingamerica.loc.gov/search/titles/results/?state=New+Mexico&county=&city=&year1=1690&year2=2013&terms=%searchstring%&frequency=&language=&ethnicity=&labor=&material_type=&lccn=&rows=20
Return

; US Newspaper Directory LCCN search
DirectoryLCCN:
	InputBox, searchstring, LCCN Search,,, 250, 100,,,,,Do not use quotes.
	if ErrorLevel
		Return
	else
		Run, http://chroniclingamerica.loc.gov/search/titles/results/?state=&county=&city=&year1=1690&year2=2013&terms=&frequency=&language=&ethnicity=&labor=&material_type=&lccn=%searchstring%&rows=20
Return

; Chronicling America Texas search
ChronAmTexas:
	InputBox, searchstring, ChronAm Texas Search,,, 250, 100,,,,,Do not use quotes.
	if ErrorLevel
		Return
	else
		Run, http://chroniclingamerica.loc.gov/search/pages/results/?state=Texas&date1=1836&date2=1922&proxtext=%searchstring%&x=0&y=0&dateFilterType=yearRange&rows=20&searchType=basic
Return

; Chronicling America Oklahoma search
ChronAmOklahoma:
	InputBox, searchstring, ChronAm Oklahoma Search,,, 250, 100,,,,,Do not use quotes.
	if ErrorLevel
		Return
	else
		Run, http://chroniclingamerica.loc.gov/search/pages/results/?state=Oklahoma&date1=1836&date2=1922&proxtext=%searchstring%&x=0&y=0&dateFilterType=yearRange&rows=20&searchType=basic
Return

; Chronicling America New Mexico search
ChronAmNewMexico:
	InputBox, searchstring, ChronAm New Mexico Search,,, 250, 100,,,,,Do not use quotes.
	if ErrorLevel
		Return
	else
		Run, http://chroniclingamerica.loc.gov/search/pages/results/?state=New+Mexico&date1=1836&date2=1922&proxtext=%searchstring%&x=0&y=0&dateFilterType=yearRange&rows=20&searchType=basic
Return
; =================SEARCH

; =================HELP
; display version info
About:
MsgBox NDNP_QR.ahk`nVersion 1.6`nAndrew.Weidner@unt.edu
Return

; open NDNP Awardee wiki
NDNPwiki:
  Run, http://www.loc.gov/extranet/wiki/library_services/ndnp/index.php/AwardeePage
Return

; open NDNP_QR wiki page
Documentation:
  Run, https://github.com/drewhop/AutoHotkey/wiki/NDNP_QR
Return
; =================HELP
; =========================================MENU FUNCTIONS
; ========================================================

GuiClose:
ExitApp
