/*
 * NDNP_QR.ahk
 *
 * DESCRIPTION: application for NDNP data Quality Review
 *
 *
 * VERSION HISTORY ******************
 *
 * Version 1.0 2012-12-20
 * Version 1.1 2013-01-11: added Search menu; modified inputboxes
 * Version 1.2 2013-01-15: added Next+, Previous+,
 *                         DVV Pages Loop, & Validate/Verify
 * Version 1.3 2013-02-27: reworked interface; added DVV Thumbs loop
 *                         ; added reel report tool
 * Version 1.4 2013-03-07: added batch report tool
 *                         ; added page count & questionable to report
 * Version 1.5 2013-03-11: customized for FirstImpression viewer
 *                         ; added edition label to reports
 * Version 1.6 2013-03-15: reworked batch report to sort batch.xml
 *                         ; reworked folder navigation
 *                         ; added Open+ and Metadata Viewer
 * Version 1.7 2013-05-10: cleaned up code with subroutines; added Edit menu
 *                         ; new scripts: GoTo, GoTo+, Metadata Windows, Back
 *                         ; reworked search scripts
 * Version 1.8 2013-05-17: cleaned up Reports, ExtractMeta, and OpenFirstTiff
 *                         ; added skip issues option for navigation
 * Version 1.9 2013-00-00: code review & production batch testing
 * Version 2.0 2013-00-00: NDNP Listserv Release
 *
 *
 * REQUIRED SOFTWARE ****************
 *
 * First Impression Viewer (default TIFF viewer)
 * http://www.softpedia.com/get/Multimedia/Graphic/Graphic-Viewers/First-Impression.shtml
 *
 * Notepad++ (metadata extraction)
 * http://notepad-plus-plus.org/
 *
 * Digital Viewer and Validator (Pages & Thumbs Loops)
 * http://www.loc.gov/ndnp/tools/
 *
 *
 * GUI NOTE *************************
 * The GUI dimensions may need to be adjusted for your system.
 * See the AutoHotkey GUI documentation for more information:
 * http://www.autohotkey.com/docs/commands/Gui.htm
 *
 *
 * CONTACT: Andrew.Weidner@unt.edu
 */

#NoEnv
#persistent

; =========================================VARIABLES
; EDIT THIS VARIABLE
; for the drive where you store your QR batches
batchdrivedefault = E:\

; EDIT THIS VARIABLE
; for the path to the Notepad++ folder
notepadpathdefault = C:\Program Files (x86)\Notepad++

; EDIT THIS VARIABLE
; for the path to the folder containing cmd.exe
CMDpathdefault = C:\Windows\System32

; EDIT THIS VARIABLE
; for the path to the DVV folder
DVVpathdefault = C:\dvv

; documentation
docURL = https://github.com/drewhop/AutoHotkey/wiki/NDNP_QR

; scoreboard
hotkeys = 0
openscore = 0
nextscore = 0
prevscore = 0
gotoscore = 0
metadatascore = 0
viewissuexmlscore = 0
editissuexmlscore = 0
zoominscore = 0
zoomoutscore = 0

; navigation
batchdrive = %batchdrivedefault%
notepadpath = %notepadpathdefault%
CMDpath = %CMDpathdefault%
DVVpath = %DVVpathdefault%
reelfoldername = _
reelfolderpath = _
navskip = 1
navchoice = 1

; tools
count = 0
pagecount = 0
thumbscount = 0
loopcount = 0

; search
LCCNstring = Enter an LCCN
directorysearchstring = Do not use quotes.
chronamsearchstring = Do not use quotes.
date1 = 1836
date2 = 1922
; =========================================VARIABLES

Gui, 1:Color, d0d0d0, 912206
Gui, 1:Show, h0 w0, NDNP_QR

; =========================================MENUS
; see MENU FUNCTIONS

; File
Menu, FileMenu, Add, &Open Batch Folder, OpenBatch
Menu, FileMenu, Add
Menu, FileMenu, Add, V&alidate Batch, DVVvalidate
Menu, FileMenu, Add, V&erify Batch, DVVverify
Menu, FileMenu, Add
Menu, FileMenu, Add, &Reload, Reload
Menu, FileMenu, Add, E&xit, Exit

; Edit
Menu, EditMenu, Add, &Select Reel Folder, EditReelFolder
Menu, EditMenu, Add, &Display Current Reel Folder, DisplayReelFolder
Menu, EditMenu, Add
Menu, EditMenu, Add, &Folder Navigation, NavSkip
Menu, EditMenu, Add
Menu, EditMenu, Add, &Batch Drive, EditBatchDrive
Menu, EditMenu, Add, &Notepad++ Folder, EditNotepadFolder
Menu, EditMenu, Add, &CMD Folder, EditCMDFolder
Menu, EditMenu, Add, D&VV Folder, EditDVVFolder

; Tools
Menu, ToolsMenu, Add, &Batch Report, BatchReport
Menu, ToolsMenu, Add
Menu, ToolsMenu, Add, &Reel Report, ReelReport
Menu, ToolsMenu, Add, &Metadata Viewer, MetaViewer
Menu, ToolsMenu, Add
Menu, ToolsMenu, Add, DVV &Pages Loop, DVVpages
Menu, ToolsMenu, Add, DVV &Thumbs Loop, DVVthumbs

; Search
Menu, SearchMenu, Add, US &Directory Search, DirectorySearch
Menu, SearchMenu, Add
Menu, SearchMenu, Add, US Directory: &LCCN, DirectoryLCCN
Menu, SearchMenu, Add
Menu, SearchMenu, Add, &ChronAm Search, ChronAmSearch

; Help
Menu, HelpMenu, Add, &Documentation, Documentation
Menu, HelpMenu, Add, &NDNP Awardee Wiki, NDNPwiki
Menu, HelpMenu, Add, &About, About

; create menus
Menu, MenuBar, Add, &File, :FileMenu
Menu, MenuBar, Add, &Edit, :EditMenu
Menu, MenuBar, Add, &Tools, :ToolsMenu
Menu, MenuBar, Add, &Search, :SearchMenu
Menu, MenuBar, Add, &Help, :HelpMenu

; create menu toolbar
Gui, Menu, MenuBar
; =========================================MENUS

; =========================================LABELS
; row 1
; Last Hotkey: Static 3
Gui, 1:Add, Text, x15  y10 w40  h15, HotKeys
Gui, 1:Add, Text, x95  y10 w70  h15, Last HotKey
Gui, 1:Add, Text, x105 y35 w60  h15, 
Gui, 1:Add, Text, x175 y10 w70  h15, Report / DVV
Gui, 1:Add, Text, x255 y10 w70  h15, Metadata ( m )

; row 2
Gui, 1:Add, Text, x15  y70 w100 h20, Open ( i )
Gui, 1:Add, Text, x95  y70 w100 h20, Next ( o )
Gui, 1:Add, Text, x175 y70 w100 h20, Previous ( p )
Gui, 1:Add, Text, x255 y70 w100 h20, GoTo ( g )

; row 3
Gui, 1:Add, Text, x15  y130 w70 h20, Zoom In ( k )
Gui, 1:Add, Text, x95  y130 w70 h20, Zoom Out ( l )
Gui, 1:Add, Text, x175 y130 w70 h20, ViewXML ( q )
Gui, 1:Add, Text, x255 y130 w70 h20, EditXML ( w )

; Metadata window
Gui, 2:Font,, Arial
Gui, 2:Add, Text, x40 y55  w100 h20, Volume:
Gui, 2:Add, Text, x49 y80  w100 h20, Issue:
Gui, 2:Add, Text, x44 y105 w100 h20, ? Date:
Gui, 2:Add, Text, x45 y130 w100 h20, Pages:
; =========================================LABELS

; =========================================DATA
; set larger data font
Gui, 1:Font, s15,

; row 1: Static 14-16
Gui, 1:Add, Text, x25  y30 w55 h25, 0
Gui, 1:Add, Text, x185 y30 w55 h25, 0
Gui, 1:Add, Text, x265 y30 w55 h25, 0
; row 2: Static 17-20
Gui, 1:Add, Text, x25  y90 w55 h25, 0
Gui, 1:Add, Text, x105 y90 w55 h25, 0
Gui, 1:Add, Text, x185 y90 w55 h25, 0
Gui, 1:Add, Text, x265 y90 w55 h25, 0
; row 3: Static 21-24
Gui, 1:Add, Text, x25  y150 w45 h25, 0
Gui, 1:Add, Text, x105 y150 w45 h25, 0
Gui, 1:Add, Text, x185 y150 w45 h25, 0
Gui, 1:Add, Text, x265 y150 w45 h25, 0

; Metadata window font
Gui, 2:Font, cRed s12 bold, Arial

; Metadata: Static 5-9
Gui, 2:Add, Text, x55 y20  w90  h20,
Gui, 2:Add, Text, x90 y55  w100 h20,
Gui, 2:Add, Text, x90 y80  w100 h20,
Gui, 2:Add, Text, x90 y105 w100 h20,
Gui, 2:Add, Text, x90 y130 w100 h20,
; =========================================DATA

; =========================================BOXES
; main window decorative
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
; row 3
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

; activate the NDNP_QR window
winactivate, NDNP_QR

; pause key pauses any script or function
Pause::Pause

; =========================================SCRIPTS
; =============================OPEN
; opens selected issue folder and first TIFF file in First Impression
; HotKey = Alt + i
!i::
	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows

	; open first TIFF in selected folder
	Gosub, OpenFirstTIFF

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

	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows

	; open first TIFF in selected folder
	Gosub, OpenFirstTIFF
	
	; zoom out
	Send, {NumpadSub 20}

	; send First Impression to bottom of stack
	WinSet, Bottom,, ahk_id %firstid%
	Sleep, 300

	; extract the metadata
	GoSub, ExtractMeta
		
	; activate First Impression
	WinActivate, ahk_id %firstid%
	
	; bring metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Edition

	; update the scoreboard
	hotkeys+=2
	openscore++
	metadatascore++
	ControlSetText, Static3, OPEN+, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static16, %metadatascore%, NDNP_QR
	ControlSetText, Static17, %openscore%, NDNP_QR
Return
; =============================OPEN+

; =============================NEXT
; opens the first TIFF file in the next issue folder
; Hotkey: Alt + o
!o::
	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows
	
	; up to reel folder
	Gosub, IssueToReel

	; move to the next folder
	SetKeyDelay, 50
	Send, {Down %navskip%}
	SetKeyDelay, 10
	Sleep, 100

	; open first TIFF in selected folder
	Gosub, OpenFirstTIFF

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

	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows

	; up to the reel directory
	Gosub, IssueToReel
	
	; move to the next folder
	SetKeyDelay, 50
	Send, {Down %navskip%}
	SetKeyDelay, 10
	Sleep, 100

	; open first TIFF in selected folder
	Gosub, OpenFirstTIFF

	; zoom out
	Send, {NumpadSub 20}

	; send First Impression to bottom of stack
	WinSet, Bottom,, ahk_id %firstid%
	Sleep, 300

	; extract the metadata
	GoSub, ExtractMeta

	; activate First Impression
	WinActivate, ahk_id %firstid%
	
	; bring metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Edition

	; update the scoreboard
	hotkeys+=2
	nextscore++
	metadatascore++
	ControlSetText, Static3, NEXT+, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static16, %metadatascore%, NDNP_QR
	ControlSetText, Static18, %nextscore%, NDNP_QR
Return
; =============================NEXT+

; =============================PREVIOUS
; opens the first TIFF file in the previous issue folder
; Hotkey: Alt + p
!p::
	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows
	
	; up to reel folder
	Gosub, IssueToReel

	; move to the previous folder
	SetKeyDelay, 50
	Send, {Up %navskip%}
	SetKeyDelay, 10
	Sleep, 100

	; open first TIFF in selected folder
	Gosub, OpenFirstTIFF

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

	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows
	
	; up to reel folder
	Gosub, IssueToReel

	; move to the previous folder
	SetKeyDelay, 50
	Send, {Up %navskip%}
	SetKeyDelay, 10
	Sleep, 100

	; open first TIFF in selected folder
	Gosub, OpenFirstTIFF

	; zoom out
	Send, {NumpadSub 20}

	; send First Impression to bottom of stack
	WinSet, Bottom,, ahk_id %firstid%
	Sleep, 300

	; extract the metadata
	GoSub, ExtractMeta

	; activate First Impression
	WinActivate, ahk_id %firstid%
	
	; bring metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Edition

	; update the scoreboard
	hotkeys+=2
	revscore++
	metadatascore++
	ControlSetText, Static3, PREV+, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static16, %metadatascore%, NDNP_QR
	ControlSetText, Static19, %revscore%, NDNP_QR
Return
; =============================PREVIOUS+

; =============================GOTO
; opens a specific issue folder and first TIFF file
; Hotkey: Alt + g
!g::
	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows
	
	; navigation subroutine
	Gosub, GoToIssue

	; zoom out
	Send, {Enter}

	; update the scoreboard
	hotkeys++
	gotoscore++
	ControlSetText, Static3, GOTO, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static20, %gotoscore%, NDNP_QR
Return
; =============================GOTO

; =============================GOTO+
; opens a specific issue folder and first TIFF file
; and extracts the issue metadata
; Hotkey: Ctrl + Alt + g
^!g::
	; close the Edition Label Window
	Gui, 9:Destroy

	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows
	
	; navigation subroutine
	Gosub, GoToIssue
	
	; zoom out
	Send, {NumpadSub 20}

	; send First Impression to bottom of stack
	WinSet, Bottom,, ahk_id %firstid%
	Sleep, 300

	; get the issue folder path
	Gosub, IssueFolderPath
	
	; extract the metadata
	Gosub, ExtractMeta
	
	; activate First Impression
	WinActivate, ahk_id %firstid%
	
	; bring metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Edition

	; update the scoreboard
	hotkeys+=2
	gotoscore++
	metadatascore++
	ControlSetText, Static3, GOTO+, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static16, %metadatascore%, NDNP_QR
	ControlSetText, Static20, %gotoscore%, NDNP_QR
Return
; =============================GOTO+

; =============================METADATA
; displays issue metadata in Metadata window
; Hotkey: Alt + m
!m::
	; close the Edition Label Window
	Gui, 9:Destroy

	; get the unique FI window id#
	SetTitleMatchMode 2
	WinGet, firstid, ID, First Impression

	; harvest the issuefolder path
	Gosub, IssueFolderPath

	; extract the metadata
	GoSub, ExtractMeta

	; activate First Impression
	WinActivate, ahk_id %firstid%
	
	; bring metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Edition

	; update the scoreboard
	hotkeys++
	metadatascore++
	ControlSetText, Static3, METADATA, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static16, %metadatascore%, NDNP_QR
Return
; =============================METADATA

; =============================METADATA WINDOWS
; creates separate windows for the issue metadata
; Hotkey: Ctrl + Alt + m
^!m::
	Gosub, CreateMetaWindows
Return
; =============================CREATE METADATA WINDOWS

; =============================ZOOM IN
; First Impression masthead view
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
; First Impression full page view
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
; Web browsers work well for this
; Hotkey: Alt + q
!q::
	; harvest the issuefolder path
	Gosub, IssueFolderPath

	; open the issue.xml file in default application
	Run, "%issuefolderpath%\%issuefoldername%.xml"
		
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
	; harvest the issuefolder path
	Gosub, IssueFolderPath

	; open the issue.xml file in Notepad++
	Run, "%notepadpath%\notepad++.exe" "%issuefolderpath%\%issuefoldername%.xml"		
		
	; update the scoreboard
	hotkeys++
	editissuexmlscore++
	ControlSetText, Static3, EditXML, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static24, %editissuexmlscore%, NDNP_QR
Return
; =============================EDIT ISSUE XML

; =============================BACK
; opens the reel folder with reelfolderpath variable
; Hotkey: Alt + b
!b::
	; check for reel folder variable
	Gosub, ReelFolderCheck
	
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

	; open the reel folder
	Run, %reelfolderpath%

	; update scoreboard
	hotkeys++
	backscore++
	ControlSetText, Static3, BACK %backscore%, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
Return
; =============================BACK
; =========================================SCRIPTS

; =========================================SUBROUTINES
; ======================NAVIGATION SUBS
ReelFolderCheck:
	if (reelfolderpath == "_")
	{
		; create dialog to select the reel folder
		FileSelectFolder, reelfolderpath, %batchdrive%, 0, `nSelect the REEL folder:
		if ErrorLevel
		{
			reelfolderpath = _
			Exit
		}

		; extract reel folder name from path
		StringGetPos, foldernamepos, reelfolderpath, \, R1
		foldernamepos++
		StringTrimLeft, reelfoldername, reelfolderpath, foldernamepos
	}
Return

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

OpenFirstTIFF:
	; wait for the reel folder
	SetTitleMatchMode 1
	IfWinNotActive, %reelfoldername%, , , ,
	WinActivate, %reelfoldername%, , , ,
	WinWaitActive, %reelfoldername%, , , ,
	Sleep, 100

	; open the selected issue folder
	Send, {AltDown}f{AltUp}
	Sleep, 100
	Send, o

	; wait for the issue folder
	SetTitleMatchMode RegEx
	WinWaitActive, ^1[0-9]{9}$, , , ,
	Sleep, 100

	; get the issue folder path
	Gosub, IssueFolderPath
	
	; find and open the first TIF file
	Loop, %issuefolderpath%\*.tif
	{
		Run, %issuefolderpath%\%A_LoopFileName%
		Break
	}
	
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
Return

GoToIssue:
	; create input box to enter the folder name to go to
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
						MsgBox, %issuefolder% does not exist in this directory:`n`n%reelfolderpath%
						
						; exit the script
						Return
					}
					
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
						Run, %reelfolderpath%\%issuefolder%
						
						; wait for issue folder to load
						WinWaitActive, ^[1-2][0-9]{9}$, , 5, ,
						; if the folder does not exist
						if ErrorLevel
						{
							; open the reel folder
							Run, %reelfolderpath%
							
							; exit the script
							Return
						}

						; assign the issuefolderpath variable
						issuefolderpath = %reelfolderpath%\%issuefolder%
												
						; find and open the first TIF file
						Loop, %issuefolderpath%\*.tif
						{
							Run, %issuefolderpath%\%A_LoopFileName%
							Break
						}

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

						; break title check loop
						Break
					}
					
					; case for active reel folder
					IfWinNotExist, ^[1-2][0-9]{9}$
					SetTitleMatchMode 2
					IfWinExist, %reelfoldername%
					{
						; activate reel folder
						WinWait, %reelfoldername%, , , ,
						IfWinNotActive, %reelfoldername%, , , ,
						WinActivate, %reelfoldername%, , , ,
						WinWaitActive, %reelfoldername%, , , ,
						Sleep, 100

						; open the issue
						SetKeyDelay, 150
						Send, %issuefolder%
						SetKeyDelay, 10
						Sleep, 100
						Send, {AltDown}f{AltUp}
						Sleep, 100
						Send, o
						
						; wait for the issue folder
						SetTitleMatchMode RegEx
						WinWaitActive, ^1[0-9]{9}$, , , ,
						Sleep, 100

						; get the issue folder path
						Gosub, IssueFolderPath
						
						; find and open the first TIF file
						Loop, %issuefolderpath%\*.tif
						{
							Run, %issuefolderpath%\%A_LoopFileName%
							Break
						}

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

						; break title check loop
						Break
					}
					
					; case for no relevant folders
					IfWinNotExist, ^[1-2][0-9]{9}$
					SetTitleMatchMode 2
					IfWinNotExist, %reelfoldername%
					{
						;open the reel folder
						Run, %reelfolderpath%
						
						; activate reel folder
						WinWait, %reelfoldername%, , , ,
						IfWinNotActive, %reelfoldername%, , , ,
						WinActivate, %reelfoldername%, , , ,
						WinWaitActive, %reelfoldername%, , , ,
						Sleep, 100

						; open the issue
						SetKeyDelay, 150
						Send, %issuefolder%
						SetKeyDelay, 10
						Sleep, 100
						Send, {AltDown}f{AltUp}
						Sleep, 100
						Send, o
						
						; wait for the issue folder
						SetTitleMatchMode RegEx
						WinWaitActive, ^1[0-9]{9}$, , , ,
						Sleep, 100

						; get the issue folder path
						Gosub, IssueFolderPath
						
						; find and open the first TIF file
						Loop, %issuefolderpath%\*.tif
						{
							Run, %issuefolderpath%\%A_LoopFileName%
							Break
						}

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

						; break title check loop
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
; ======================NAVIGATION SUBS

; ======================METADATA SUBS
IssueFolderPath:
	; save clipboard contents
	temp = %clipboard%
	
	; activate the current NDNP issue folder
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
	
	; extract issue folder name from path
	StringGetPos, foldernamepos, issuefolderpath, \, R1
	foldernamepos++
	StringTrimLeft, issuefoldername, issuefolderpath, foldernamepos

	; restore clipboard contents
	clipboard = %temp%		
Return

ExtractMeta:
	; store clipboard contents
	temp = %clipboard%
	
	; initialize the issuexml variable
	issuexml =
			
	; initialize the metadata values
	volume =
	issue =
	date =
	editionlabel =
	questionable =
	questionabledisplay =
	monthname =
	day =
	year =
	monthnameQ =
	dayQ =
	yearQ =
			
	; initialize the metadata flags
	volumeflag = 0
	issueflag = 0

	; read in the issue.xml file to the issuexml variable
	FileRead, issuexml, %issuefolderpath%\%issuefoldername%.xml
				
	; loop to parse the issuexml variable
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

			; create Edition Label GUI
			; except for first pass of Reel Report or Metadata Viewer
			if ((loopcount == 0) || (loopcount > 1))
			{
				WinGetPos, winX, winY, winWidth, winHeight, Metadata
				winY+=%winHeight%

				Gui, 9:+AlwaysOnTop
				Gui, 9:+ToolWindow
				Gui, 9:Font, cRed s15 bold, Arial
				Gui, 9:Add, Text, x15 y15 w380 h25, %editionlabel%
				Gui, 9:Font, cRed s10 bold, Arial
				Gui, 9:Add, GroupBox, x0 y-7 w400 h63,       
				Gui, 9:Show, x%winX% y%winY% h55 w400, Edition
			}
		}

		; if the line contains an issue date
		IfInString, A_LoopField, <mods:dateIssued encoding="iso8601">
		{
			; assign the issue date
			StringMid, date, A_LoopField, 37, 10

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
		}

		; if the line contains a questionable issue date
		IfInString, A_LoopField, qualifier="questionable"
		{
			; assign the questionable issue date
			StringMid, questionable, A_LoopField, 62, 10

			; assign the report display variable
			questionabledisplay = Questionable: %questionable%
		
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
			dayQ := SubStr(questionable, 9)
			if SubStr(dayQ, 1, 1) = 0
			{
				dayQ := SubStr(dayQ, 2)
			}
			yearQ := SubStr(questionable, 1, 4)

			; if reel report or metadata viewer
			if loopcount > 1
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
		}

		; initialize the number of pages variable
		numpages = 0
					
		; count the number of .TIF files
		Loop, %issuefolderpath%\*.tif
		{
			numpages++
		}
	}

	; update the date field
	ControlSetText, Static5, %date%, Metadata
	ControlSetText, Static1, %monthname% %day%`, %year%, Date

	; update the volume number
	ControlSetText, Static6, %volume%, Metadata
	ControlSetText, Static1, %volume%, Volume

	; update the issue number
	ControlSetText, Static7, %issue%, Metadata		
	ControlSetText, Static1, %issue%, Issue

	; update the questionable date
	ControlSetText, Static8, %questionable%, Metadata		

	; update the number of pages
	ControlSetText, Static9, %numpages%, Metadata
				
	; create GUIs for Date, Volume, Issue, Questionable, and Edition Label
	; if first pass of Reel Report or Metadata Viewer
	if (loopcount == 1)
	{
		Gosub, CreateMetaWindows
	}
	
	; restore the clipboard
	clipboard = %temp%
Return

CreateMetaWindows:
	; position to right of Metadata window
	WinGetPos, winX, winY, winWidth, winHeight, Metadata
	winX+=%winWidth%

	; Issue number GUI
	Gui, 6:+AlwaysOnTop
	Gui, 6:+ToolWindow
	Gui, 6:Font, cRed s25 bold, Arial
	Gui, 6:Add, GroupBox, x0 y-18 w200 h74,       
	Gui, 6:Add, Text, x15 y8 w170 h35, %issue%
	Gui, 6:Show, x%winX% y%winY% h55 w200, Issue

	; Volume number GUI
	Gui, 5:+AlwaysOnTop
	Gui, 5:+ToolWindow
	Gui, 5:Font, cRed s25 bold, Arial
	Gui, 5:Add, GroupBox, x0 y-18 w200 h74,       
	Gui, 5:Add, Text, x15 y8 w170 h35, %volume%
	Gui, 5:Show, x%winX% y%winY% h55 w200, Volume

	; Date GUI
	Gui, 4:+AlwaysOnTop
	Gui, 4:+ToolWindow
	Gui, 4:Font, cRed s15 bold, Arial
	Gui, 4:Add, Text, x35 y15 w160 h25, %monthname% %day%, %year%
	Gui, 4:Font, cRed s10 bold, Arial
	Gui, 4:Add, GroupBox, x0 y-7 w200 h63,       
	Gui, 4:Show, x%winX% y%winY% h55 w200, Date
	
	; Questionable Date GUI
	if (questionabledate != "")
	{
		; position below Date window
		WinGetPos, winX, winY, winWidth, winHeight, Date
		winY+=%winHeight%

		Gui, 7:+AlwaysOnTop
		Gui, 7:+ToolWindow
		Gui, 7:Font, cRed s15 bold, Arial
		Gui, 7:Add, Text, x35 y15 w160 h25, %monthnameQ% %dayQ%, %yearQ%
		Gui, 7:Font, cRed s10 bold, Arial
		Gui, 7:Add, GroupBox, x0 y-7 w200 h63,       
		Gui, 7:Show, x%winX% y%winY% h55 w200, Questionable
	}

	; Edition Label GUI
	if (editionlabel != "")
	{
		; position below Metadata window
		WinGetPos, winX, winY, winWidth, winHeight, Metadata
		winY+=%winHeight%

		Gui, 9:+AlwaysOnTop
		Gui, 9:+ToolWindow
		Gui, 9:Font, cRed s15 bold, Arial
		Gui, 9:Add, Text, x15 y15 w380 h25, %editionlabel%
		Gui, 9:Font, cRed s10 bold, Arial
		Gui, 9:Add, GroupBox, x0 y-7 w400 h63,       
		Gui, 9:Show, x%winX% y%winY% h55 w400, Edition
	}				

	; pause first instance of Reel Report and Metadata Viewer
	if (loopcount == 1)
	{
		MsgBox, 0, Loop Paused, Position the metadata windows as desired.`n`nClick OK and the loop will continue in %delay% seconds.
	}
Return

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
		if (seconds = 0)
		{
			Gui, 8:Destroy
			Break
		}
	}
Return
; ======================METADATA SUBS
; =========================================SUBROUTINES

; =========================================MENU FUNCTIONS
; ======================FILE MENU
; open a batch folder
; edit the file path variable in line 26 for your system
OpenBatch:
	; create dialog to select the batch folder
	FileSelectFolder, batchpath, E:\, 0, Select a folder:
	If ErrorLevel
		Return
	Else
		Run, %batchpath%
Return

; validate a batch with the command line
; modify CMDpath & DVVpath in VARIABLES for your system 
DVVvalidate:
	FileSelectFolder, batchpath, E:\, 0, `nSelect the batch to VALIDATE:
	if ErrorLevel
		Return
	else
	{
		SetWorkingDir, %DVVpath%
		Run, "%CMDpath%\cmd.exe" /k validationprocessor.bat batch %batchpath%\batch.xml update
	}
Return

; verify a batch with the command line
; modify CMDpath & DVVpath in VARIABLES for your system 
DVVverify:
	FileSelectFolder, batchpath, E:\, 0, `nSelect the batch to VERIFY:
	if ErrorLevel
		Return
	else
	{
		SetWorkingDir, %DVVpath%
		Run, "%CMDpath%\cmd.exe" /k validationprocessor.bat batch %batchpath%\batch_1.xml verify
	}
Return

; reload application
Reload:
Reload

; exit application
Exit:
ExitApp
; ======================FILE MENU

; ======================EDIT MENU
EditReelFolder:
	; create dialog to select the reel folder
	FileSelectFolder, reelfolderpath, %batchdrive%, 0, `nSelect the REEL folder:
	if ErrorLevel
	{
		reelfolderpath = _
		Return
	}

	; extract reel folder name from path
	StringGetPos, foldernamepos, reelfolderpath, \, R1
	foldernamepos++
	StringTrimLeft, reelfoldername, reelfolderpath, foldernamepos
Return

DisplayReelFolder:
	if (reelfolderpath == "_")
	{
		MsgBox, 0, Current Reel, Reel Folder Name: %reelfoldername%`n`nReel Folder Path: %reelfolderpath%
	}
	else
	{
		; create display variables for title folder path
		StringGetPos, reelfolderpos, reelfolderpath, \, R2
		StringLeft, reelfolder1, reelfolderpath, reelfolderpos
		StringTrimLeft, reelfolder2, reelfolderpath, reelfolderpos

		MsgBox, 0, Current Reel, Reel Folder Name:`n`n`t%reelfoldername%`n`nReel Folder Path:`n`n`t%reelfolder1%`n`t%reelfolder2%
	}
Return

NavSkip:
	; create the Folder Navigation GUI
	Gui, 12:Add, Text,, FOLDER NAVIGATION`n`nNumber of folders:
	Gui, 12:Add, DropDownList, Choose%navchoice% R7 vnavskip, 1|2|3|5|10|15|20
	Gui, 12:Add, Button, x10 y80 gNavSkipGo default, OK
	Gui, 12:Add, Button, x45 y80 gNavSkipCancel, Cancel
	Gui, 12:Show,, Folder Navigation
Return

NavSkipGo:
	; assign the navskip variable value
	Gui, 12:Submit
	
	; assign the navchoice variable value
	if (navskip == 1)
		navchoice = %navskip%
	else if (navskip == 2)
		navchoice = %navskip%
	else if (navskip == 3)
		navchoice = %navskip%
	else if (navskip == 5)
		navchoice = 4
	else if (navskip == 10)
		navchoice = 5
	else if (navskip == 15)
		navchoice = 6
	else navchoice = 7
		
	; close the Folder Navigation GUI
	Gui, 12:Destroy
Return

NavSkipCancel:
	Gui, 12:Destroy
Return

EditBatchDrive:
	; create dialog to select the batch storage drive
	FileSelectFolder, batchdrive,, 0, `nSelect the drive where your batches are stored:
	if ErrorLevel
	{
		batchdrive = %batchdrivedefault%
		Return
	}
Return

EditNotepadFolder:
	; create dialog to select the Notepad++ folder
	FileSelectFolder, notepadpath, C:\, 0, `nSelect the folder where Notepad++ is installed:
	if ErrorLevel
	{
		notepadpath = %notepadpathdefault%
		Return
	}
Return

EditCMDFolder:
	; create dialog to select the CMD.exe folder
	FileSelectFolder, CMDpath, C:\, 0, `nSelect the folder where CMD.exe is installed:
	if ErrorLevel
	{
		CMDpath = %CMDpathdefault%
		Return
	}
Return

EditDVVFolder:
	; create dialog to select the DVV folder
	FileSelectFolder, DVVpath, C:\, 0, `nSelect the folder where the DVV is installed:
	if ErrorLevel
	{
		DVVpath = %DVVpathdefault%
		Return
	}
Return
; ======================EDIT MENU

; ======================TOOLS MENU
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
	FileSelectFolder, batchfolderpath, %batchdrive%, 0, BATCH REPORT`n`nSelect a BATCH folder:
	if ErrorLevel
		Return
	else
	{	
		; extract batch name from path
		StringGetPos, foldernamepos, batchfolderpath, \, R1
		foldernamepos++
		StringTrimLeft, batchname, batchfolderpath, foldernamepos

		; store the start time
		start = %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% 
			
		; create the report file in the batch folder
		FileAppend, %batchfolderpath%`n`nIdentifier`t`tPages`tDate`t`tVolume`tIssue`n`n, %batchfolderpath%\%batchname%-report.txt

		; read in the batch.xml file to the batchxml variable
		FileRead, batchxml, %batchfolderpath%\batch.xml

		; initialize the issue counter
		issuecount = 0

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

		; initialize the total page counter
		totalpages = 0

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
				FileAppend, `n  END: %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%, %batchfolderpath%\%batchname%-report.txt

				; create a message box to indicate that the script ended
				MsgBox, 0, Batch Report, Batch: %batchname%`nIssues: %issuecount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%

				;exit the report tool
				return
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

	; initialize the issue and total pages counters
	loopcount = 0
	totalpages = 0

	; create dialog to select a reel folder
	FileSelectFolder, reelfolderpath, %batchdrive%, 0, REEL REPORT`n`nSelect a REEL folder:
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
			Gosub, CloseFirstImpressionWindows			
			
			; store the start time
			start = %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% 

			; create a variable for the report path (LCCN folder)
			StringGetPos, foldernamepos, reelfolderpath, \, R1
			StringLeft, reportpath, reelfolderpath, foldernamepos

			; extract reel number from path
			foldernamepos++
			StringTrimLeft, reelnumber, reelfolderpath, foldernamepos

			; create the report file
			FileAppend, %reelfolderpath%`n`nPages`tDate`t`tVolume`tIssue`n`n, %reportpath%\%reelnumber%-report.txt

			; open the reel folder
			Run, %reelfolderpath%

			; activate the reel folder
			SetTitleMatchMode 1
			WinWait, %reelnumber%, , , ,
			IfWinNotActive, %reelnumber%, , , ,
			WinActivate, %reelnumber%, , , ,
			WinWaitActive, %reelnumber%, , , ,
			Sleep, 100

			; open the first issue folder
			Send, {Down}
			Sleep, 200		
			Send, {Up}
			Sleep, 200
			Send, {AltDown}f{AltUp}
			Sleep, 100
			Send, o
		
			Loop
			{
				; activate the issue folder
				SetTitleMatchMode RegEx
				WinWait, ^[1-2][0-9]{9}$, , 5, ,
				if ErrorLevel
				; *************************************************
				; AUTO EXIT at end of issue folders
				{
					; add the issue count to the report
					FileAppend, `nIssues: %loopcount%`n Pages: %totalpages%, %reportpath%\%reelnumber%-report.txt

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
					MsgBox, 0, Reel Report, Reel: %reelnumber%`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%

					; reset variables
					loopcount = 0
					reelfoldername = _
					reelfolderpath = _

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

				; get the First Impression window ID
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

				; harvest the issue folder path
				Gosub, IssueFolderPath
				
				; increment the counter
				loopcount++
				
				; extract the metadata
				Gosub, ExtractMeta
				
				; update total pages count
				totalpages+=%numpages%

				; add the data to the report
				FileAppend, %numpages%`t%date%`t%volume%`t%issue%`t%editionlabel%`t%questionabledisplay%`n, %reportpath%\%reelnumber%-report.txt
				Sleep, 100
	
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
						Gosub, CloseFirstImpressionWindows
  
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
						Send, u
						Sleep, 100

						; activate the reel folder
						SetTitleMatchMode RegEx
						WinWait, ^[0-9]{10}[0-9A]$, , , ,
						IfWinNotActive, ^[0-9]{10}[0-9A]$, , , ,
						WinActivate, ^[0-9]{10}[0-9A]$, , , ,
						WinWaitActive, ^[0-9]{10}[0-9A]$, , , ,
						Sleep, 100
	
						; open the next issue folder
						SetKeyDelay, 50
						Send, {Down %navskip%}
						SetKeyDelay, 50
						Sleep, 100
						Send, {AltDown}f{AltUp}
						Sleep, 100
						Send, o
					
						; next loop iteration
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
				; MANUAL EXIT
				; hold down Left Shift during page display delay
				getKeyState, state, LShift
				if state = D
				{
					; add the issue count to the report
					FileAppend, `nReport for reel %reelnumber% was cancelled.`n`nIssues: %loopcount%`n Pages: %totalpages%, %reportpath%\%reelnumber%-report.txt

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
					MsgBox, 0, Reel Report, The report for reel %reelnumber% was cancelled.`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%
					
					; reset variables
					loopcount = 0
					reelfoldername = _
					reelfolderpath = _

					;exit the report tool
					return
				}
				; *************************************************						

				; close the Questionable Date & Edition Label windows
				Gui, 7:Destroy
				Gui, 9:Destroy

				; close any First Impression windows
				Gosub, CloseFirstImpressionWindows
  
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
				Send, u
				Sleep, 100

				; activate the reel folder
				SetTitleMatchMode 1
				WinWait, %reelnumber%, , , ,
				IfWinNotActive, %reelnumber%, , , ,
				WinActivate, %reelnumber%, , , ,
				WinWaitActive, %reelnumber%, , , ,
				Sleep, 100
	
				; open the next issue folder
				SetKeyDelay, 50
				Send, {Down %navskip%}
				SetKeyDelay, 10
				Sleep, 100
				Send, {AltDown}f{AltUp}
				Sleep, 100
				Send, o
			}
		}
	}
Return
; ======REEL REPORT

; ======METADATA VIEWER
; opens the first TIFF and displays the issue metadata
; for each issue in a reel folder
MetaViewer:
	; update last hotkey & loop label
	ControlSetText, Static3, VIEWER, NDNP_QR
	ControlSetText, Static4, Meta Viewer, NDNP_QR

	; initialize the issue and total pages counters
	loopcount = 0
	totalpages = 0

	; create dialog to select a reel folder
	FileSelectFolder, reelfolderpath, %batchdrive%, 0, METADATA VIEWER`n`nSelect a REEL folder:
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
			Gosub, CloseFirstImpressionWindows			
			
			; store the start time
			start = %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% 
			
			; get the reel number from the reel folder path
			StringGetPos, foldernamepos, reelfolderpath, \, R1
			foldernamepos++
			StringTrimLeft, reelnumber, reelfolderpath, foldernamepos

			; open the reel folder
			Run, %reelfolderpath%

			; activate the reel folder
			SetTitleMatchMode 1
			WinWait, %reelnumber%, , , ,
			IfWinNotActive, %reelnumber%, , , ,
			WinActivate, %reelnumber%, , , ,
			WinWaitActive, %reelnumber%, , , ,
			Sleep, 100

			; open the first issue folder
			Send, {Down}
			Sleep, 200		
			Send, {Up}
			Sleep, 200
			Send, {AltDown}f{AltUp}
			Sleep, 100
			Send, o
		
			Loop
			{
				; activate the issue folder
				SetTitleMatchMode RegEx
				WinWait, ^[1-2][0-9]{9}$, , 5, ,
				if ErrorLevel
				; *************************************************
				; AUTO EXIT at end of issue folders
				{
					; close the metadata windows
					Gui, 4:Destroy
					Gui, 5:Destroy
					Gui, 6:Destroy
					Gui, 7:Destroy
					Gui, 9:Destroy

					; create a message box to indicate that the script ended
					MsgBox, 0, Metadata, Viewer, Reel: %reelnumber%`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%

					; reset variables
					loopcount = 0
					reelfoldername = _
					reelfolderpath = _
					
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

				; get the First Impression window ID
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

				; harvest the issue folder path
				Gosub, IssueFolderPath

				; increment the counter
				loopcount++
				
				; extract the metadata
				Gosub, ExtractMeta
				
				; update total pages count
				totalpages+=%numpages%

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

					; create a message box to indicate the script was cancelled
					MsgBox, 0, Metadata Viewer, The metadata viewer for reel %reelnumber% was cancelled.`n`nIssues: %loopcount%`nPages: %totalpages%`n`nSTART:`t%start%`nEND:`t%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%

					; reset variables
					loopcount = 0
					reelfoldername = _
					reelfolderpath = _
					
					;exit the metadata viewer
					return
				}
				; *************************************************						

				; close the Questionable Date & Edition Label windows
				Gui, 7:Destroy
				Gui, 9:Destroy

				; close any First Impression windows
				Gosub, CloseFirstImpressionWindows
  
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
				Send, u
				Sleep, 100

				; activate the reel folder
				SetTitleMatchMode 1
				WinWait, %reelnumber%, , , ,
				IfWinNotActive, %reelnumber%, , , ,
				WinActivate, %reelnumber%, , , ,
				WinWaitActive, %reelnumber%, , , ,
				Sleep, 100
	
				; open the next issue folder
				SetKeyDelay, 50
				Send, {Down %navskip%}
				SetKeyDelay, 10
				Sleep, 100
				Send, {AltDown}f{AltUp}
				Sleep, 100
				Send, o
			}
		}
	}
Return
; ======METADATA VIEWER

; ======DVV PAGES LOOP
; loop to view individual pages in the DVV
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
; ======================TOOLS MENU

; ======================SEARCH MENU
DirectoryLCCN:
	InputBox, LCCNstring, US Directory: LCCN,,, 200, 100,,,,,%LCCNstring%
	if ErrorLevel
		Return
	else
		Run, http://chroniclingamerica.loc.gov/search/titles/results/?state=&county=&city=&year1=1690&year2=2013&terms=&frequency=&language=&ethnicity=&labor=&material_type=&lccn=%LCCNstring%&rows=20
Return

DirectorySearch:
	Gui, 10:Add, Text,, Enter a search term:
	Gui, 10:Add, Edit, w170 vdirectorysearchstring, %directorysearchstring%
	Gui, 10:Add, Text,, Choose a state (optional):
	Gui, 10:Add, DropDownList, r10 vstate, Alabama|Arizona|California|Colorado|District of Columbia|Florida|Georgia|Hawaii|Idaho|Illinois|Indiana|Kansas|Kentucky|Louisiana|Minnesota|Mississippi|Missouri|Montana|Nebraska|New Mexico|New York|North Dakota|Ohio|Oklahoma|Oregon|Pennsylvania|South Carolina|Tennessee|Texas|Utah|Vermont|Virginia|Washington
	Gui, 10:Add, Button, x10 y100 gDirectoryGo default, Search
	Gui, 10:Add, Button, x65 y100 gDirSearchCancel, Cancel
	Gui, 10:Show,, US Directory
Return

ChronAmSearch:
	Gui, 11:Add, Text,, Enter a search term:
	Gui, 11:Add, Edit, w170 vchronamsearchstring, %chronamsearchstring%
	Gui, 11:Add, Text,, Choose a state (optional):
	Gui, 11:Add, DropDownList, r10 vstate, Alabama|Arizona|California|Colorado|District of Columbia|Florida|Georgia|Hawaii|Idaho|Illinois|Indiana|Kansas|Kentucky|Louisiana|Minnesota|Mississippi|Missouri|Montana|Nebraska|New Mexico|New York|North Dakota|Ohio|Oklahoma|Oregon|Pennsylvania|South Carolina|Tennessee|Texas|Utah|Vermont|Virginia|Washington
	Gui, 11:Add, Text,, Start:%A_Tab%End:
	Gui, 11:Add, Edit, x10 y115 vdate1, %date1%
	Gui, 11:Add, Edit, x50 y115 vdate2, %date2%
	Gui, 11:Add, Button, x10 y150 gChronAmGo default, Search
	Gui, 11:Add, Button, x65 y150 gCASearchCancel, Cancel
	Gui, 11:Show,, ChronAm
Return

DirectoryGo:
	; assign the search string
	Gui, 10:Submit
	
	; close the search GUI
	Gui, 10:Destroy
	
	; load the results in the default Web browser
	Run, http://chroniclingamerica.loc.gov/search/titles/results/?state=%state%&county=&city=&year1=1690&year2=2013&terms=%directorysearchstring%&frequency=&language=&ethnicity=&labor=&material_type=&lccn=&rows=20
Return

ChronAmGo:
	; assign the search string
	Gui, 11:Submit
	
	; close the search GUI
	Gui, 11:Destroy
	
	; load the results in the default Web browser
	Run, http://chroniclingamerica.loc.gov/search/pages/results/?state=%state%&date1=%date1%&date2=%date2%&proxtext=%chronamsearchstring%&x=0&y=0&dateFilterType=yearRange&rows=20&searchType=basic
Return

DirSearchCancel:
	Gui, 10:Destroy
Return

CASearchCancel:
	Gui, 11:Destroy
Return
; ======================SEARCH MENU

; ======================HELP MENU
; display version info
About:
	MsgBox, 0, About, NDNP_QR.ahk`nVersion 1.8`nAndrew.Weidner@unt.edu
Return

; open NDNP Awardee wiki
NDNPwiki:
  Run, http://www.loc.gov/extranet/wiki/library_services/ndnp/index.php/AwardeePage
Return

; open NDNP_QR wiki page
Documentation:
  Run, %docURL%
Return
; ======================HELP MENU
; =========================================MENU FUNCTIONS

GuiClose:
ExitApp
