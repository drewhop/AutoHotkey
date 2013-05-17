; ********************************
; NDNP_QR_hotkeys.ahk
; required file for NDNP_QR.ahk
; ********************************

; =============================OPEN
; opens first TIFF file for selected issue folder
; HotKey = Alt + i
!i::
  ; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
	; set the script flag
	openflag = 1

	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows

	; open first TIFF in selected folder
	Gosub, OpenFirstTIFF

	; zoom out
	Send, {Enter}

	; reset the script flag
	openflag = 0

	; update the scoreboard
	hotkeys++
	openscore++
	ControlSetText, Static3, OPEN, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static17, %openscore%, NDNP_QR
Return
; =============================OPEN

; =============================OPEN+
; opens first TIFF file for selected issue folder
; and displays the issue metadata
; HotKey = Ctrl + Alt + i
^!i::
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
	; set the script flag
	openplusflag = 1

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

	; reset the script flag
	openplusflag = 0

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
; opens the first TIFF file for the next issue folder
; Hotkey: Alt + o
!o::
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
	; set the script flag
	nextflag = 1

	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows

	; activate reel folder
	SetTitleMatchMode 1
	WinActivate, %reelfoldername%, , , ,
	WinWaitActive, %reelfoldername%, , 5, ,
	if ErrorLevel
	{
		; print error message after 5 seconds
		MsgBox, 0, Next, NDNP_QR_hotkeys.ahk`n`nCannot find folder %reelfoldername%`n`nOptions:`n`t"File > Open Reel Folder"`n`t"Edit > Reel Folder > Set Path"
		
		; exit the script
		Return
	}
	Sleep, 100
	
	; move to the next folder
	SetKeyDelay, 50
	Send, {Down %navskip%}
	SetKeyDelay, 10
	Sleep, 100

	; open first TIFF in selected folder
	Gosub, OpenFirstTIFF

	; zoom out
	Send, {Enter}

	; reset the script flag
	nextflag = 0

	; update the scoreboard
	hotkeys++
	nextscore++
	ControlSetText, Static3, NEXT, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static18, %nextscore%, NDNP_QR
Return
; =============================NEXT

; =============================NEXT+
; opens the first TIFF file for the next issue folder
; and displays the issue metadata
; Hotkey: Ctrl + Alt + o
^!o::
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
	; set the script flag
	nextplusflag = 1

	; close the Edition Label window
	Gui, 9:Destroy

	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows

	; activate reel folder
	SetTitleMatchMode 1
	WinActivate, %reelfoldername%, , , ,
	WinWaitActive, %reelfoldername%, , 5, ,
	if ErrorLevel
	{
		; print error message after 5 seconds
		MsgBox, 0, Next+, NDNP_QR_hotkeys.ahk`n`nCannot find folder %reelfoldername%`n`nOptions:`n`t"File > Open Reel Folder"`n`t"Edit > Reel Folder > Set Path"
		
		; exit the script
		Return
	}
	Sleep, 100
	
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

	; reset the script flag
	nextplusflag = 0

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
; opens the first TIFF file for the previous issue folder
; Hotkey: Alt + p
!p::
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
	; set the script flag
	previousflag = 1

	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows
	
	; activate reel folder
	SetTitleMatchMode 1
	WinActivate, %reelfoldername%, , , ,
	WinWaitActive, %reelfoldername%, , 5, ,
	if ErrorLevel
	{
		; print error message after 5 seconds
		MsgBox, 0, Previous, NDNP_QR_hotkeys.ahk`n`nCannot find folder %reelfoldername%`n`nOptions:`n`t"File > Open Reel Folder"`n`t"Edit > Reel Folder > Set Path"
		
		; exit the script
		Return
	}
	Sleep, 100
	
	; move to the previous folder
	SetKeyDelay, 50
	Send, {Up %navskip%}
	SetKeyDelay, 10
	Sleep, 100

	; open first TIFF in selected folder
	Gosub, OpenFirstTIFF

	; zoom out
	Send, {Enter}

	; reset the script flag
	previousflag = 0

	; update the scoreboard
	hotkeys++
	revscore++
	ControlSetText, Static3, PREV, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static19, %revscore%, NDNP_QR
Return
; =============================PREVIOUS

; =============================PREVIOUS+
; opens the first TIFF file for the previous issue folder
; and displays the issue metadata
; Hotkey: Ctrl + Alt + p
^!p::
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
	; set the script flag
	previousplusflag = 1

	; close the Edition Label Window
	Gui, 9:Destroy

	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows
	
	; activate reel folder
	SetTitleMatchMode 1
	WinActivate, %reelfoldername%, , , ,
	WinWaitActive, %reelfoldername%, , 5, ,
	if ErrorLevel
	{
		; print error message after 5 seconds
		MsgBox, 0, Previous+, NDNP_QR_hotkeys.ahk`n`nCannot find folder %reelfoldername%`n`nOptions:`n`t"File > Open Reel Folder"`n`t"Edit > Reel Folder > Set Path"
		
		; exit the script
		Return
	}
	Sleep, 100
	
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

	; reset the script flag
	previousplusflag = 0

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
; opens the first TIFF file for a specific issue folder
; Hotkey: Alt + g
!g::
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
	; set the script flag
	gotoflag = 1

	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows
	
	; navigation subroutine
	Gosub, GoToIssue

	; open first TIFF in selected folder
	Gosub, OpenFirstTIFF

	; zoom out
	Send, {Enter}

	; reset the script flag
	gotoflag = 0

	; update the scoreboard
	hotkeys++
	gotoscore++
	ControlSetText, Static3, GOTO, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static20, %gotoscore%, NDNP_QR
Return
; =============================GOTO

; =============================GOTO+
; opens the first TIFF file for a specific issue folder
; and displays the issue metadata
; Hotkey: Ctrl + Alt + g
^!g::
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
	; set the script flag
	gotoplusflag = 1

	; close the Edition Label Window
	Gui, 9:Destroy

	; check for reel folder variable
	Gosub, ReelFolderCheck
	
	; close all First Impression windows
	Gosub, CloseFirstImpressionWindows
	
	; navigation subroutine
	Gosub, GoToIssue

	; open first TIFF in selected folder
	Gosub, OpenFirstTIFF
	
	; zoom out
	Send, {NumpadSub 20}

	; send First Impression to bottom of stack
	WinSet, Bottom,, ahk_id %firstid%
	Sleep, 300

	; extract the metadata
	Gosub, ExtractMeta
	
	; activate First Impression
	WinActivate, ahk_id %firstid%
	
	; bring metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Edition

	; set the script flag
	gotoplusflag = 0

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
; displays the metadata for the selected issue folder
; Hotkey: Alt + m
!m::
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
	; set the script flag
	metadataflag = 1

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

	; set the script flag
	metadataflag = 0

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
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
	Gosub, CreateMetaWindows
Return
; =============================METADATA WINDOWS

; =============================ZOOM IN
; First Impression masthead view
; HotKey = Alt + k
!k::
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
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
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
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
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
	; set the script flag
	viewissuexmlflag = 1

	; harvest the issuefolder path
	Gosub, IssueFolderPath

	; open the issue.xml file in default application
	Run, "%issuefolderpath%\%issuefoldername%.xml"
		
	; reset the script flag
	viewissuexmlflag = 0

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
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
	; set the script flag
	editissuexmlflag = 1

	; harvest the issuefolder path
	Gosub, IssueFolderPath

	; open the issue.xml file in Notepad++
	Run, "%notepadpath%\notepad++.exe" "%issuefolderpath%\%issuefoldername%.xml"		
		
	; reset the script flag
	editissuexmlflag = 0

	; update the scoreboard
	hotkeys++
	editissuexmlscore++
	ControlSetText, Static3, EditXML, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static24, %editissuexmlscore%, NDNP_QR
Return
; =============================EDIT ISSUE XML

; =============================BACK
; resets desktop to the current reel folder
; Hotkey: Alt + b
!b::
	; deactivate the NDNP_QR window
	SetTitleMatchMode 1
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	
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

	; activate current reel folder if it exists
	SetTitleMatchMode 1
	IfWinExist, %reelfoldername%
		WinActivate, %reelfoldername%

	; otherwise open the reel folder
	else
	{
		Run, %reelfolderpath%
		
		; wait for the folder to load
		WinWaitActive, %reelfoldername%
		
		; select first issue
		Send, {Down}
		Sleep, 50
		Send, {Up}
	}
	
	; update scoreboard
	hotkeys++
	backscore++
	ControlSetText, Static3, BACK %backscore%, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
Return
; =============================BACK