; ********************************
; NDNP_QR_hotkeys.ahk
; required file for NDNP_QR.ahk
; ********************************

; =============================DEACTIVATE
DeactivateWindows:
	SetTitleMatchMode 1
	IfWinActive, %reelfoldername%
		WinActivate, ahk_class Shell_TrayWnd
	IfWinActive, NDNP_QR
		WinActivate, ahk_class Shell_TrayWnd
	IfWinActive, Metadata
		WinActivate, ahk_class Shell_TrayWnd
	IfWinActive, Date
		WinActivate, ahk_class Shell_TrayWnd
	IfWinActive, Volume
		WinActivate, ahk_class Shell_TrayWnd
	IfWinActive, Issue
		WinActivate, ahk_class Shell_TrayWnd
	IfWinActive, Questionable
		WinActivate, ahk_class Shell_TrayWnd
	IfWinActive, Edition
		WinActivate, ahk_class Shell_TrayWnd
Return
; =============================DEACTIVATE

; =============================OPEN
; opens first TIFF file for selected issue folder
; HotKey = Alt + i
!i::
	Gosub, DeactivateWindows
	Sleep, 200
	
	openflag = 1
	
	; NDNP_QR_navigation.ahk
	Gosub, ReelFolderCheck
	Gosub, CloseFirstImpressionWindows
	Gosub, ResetReel
	Gosub, OpenFirstTIFF

	; zoom out
	Send, {Enter}

	openflag = 0

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
	Gosub, DeactivateWindows
	Sleep, 200
	
	openplusflag = 1

	IfWinNotExist, Metadata
	{
		Gosub, RedrawMetaWindow ; NDNP_QR_metadata.ahk
	}
	
	Gui, 7:Destroy ; Questionable
	Gui, 9:Destroy ; Edition Label
	
	; NDNP_QR_navigation.ahk
	Gosub, ReelFolderCheck
	Gosub, CloseFirstImpressionWindows
	Gosub, ResetReel
	Gosub, OpenFirstTIFF
	
	; zoom out
	Send, {NumpadSub 20}

	; send First Impression backward
	WinSet, Bottom,, ahk_id %firstid%
	Sleep, 300

	GoSub, ExtractMeta ; NDNP_QR_metadata.ahk
		
	; activate First Impression
	WinActivate, ahk_id %firstid%
	
	; bring metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Edition
	WinSet, Top,, Date
	WinSet, Top,, Volume
	WinSet, Top,, Issue
	WinSet, Top,, Questionable

	openplusflag = 0

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
	Gosub, DeactivateWindows
	Sleep, 200
	
	nextflag = 1

	; NDNP_QR_navigation.ahk
	Gosub, ReelFolderCheck
	Gosub, CloseFirstImpressionWindows

	; activate reel folder
	SetTitleMatchMode 1
	WinActivate, %reelfoldername%, , , ,
	WinWaitActive, %reelfoldername%, , 5, ,
	if ErrorLevel
	{
		; print error message after 5 seconds
		MsgBox, 0, Next, NDNP_QR_hotkeys.ahk`n`nCannot find folder %reelfoldername%`n`nOptions:`n`t"File > Open Reel Folder"`n`t"Edit > Reel Folder > Set Path"
		Return
	}
	Sleep, 100
	
	; move to the next folder
	SetKeyDelay, 50
	Send, {Down %navskip%}
	SetKeyDelay, 10
	Sleep, 100

	Gosub, OpenFirstTIFF ; NDNP_QR_navigation.ahk

	; zoom out
	Send, {Enter}

	nextflag = 0

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
	Gosub, DeactivateWindows
	Sleep, 200
	
	nextplusflag = 1

	IfWinNotExist, Metadata
	{
		Gosub, RedrawMetaWindow ; NDNP_QR_metadata.ahk
	}
	
	Gui, 7:Destroy ; Questionable
	Gui, 9:Destroy ; Edition Label
	
	; NDNP_QR_navigation.ahk
	Gosub, ReelFolderCheck
	Gosub, CloseFirstImpressionWindows

	; activate reel folder
	SetTitleMatchMode 1
	WinActivate, %reelfoldername%, , , ,
	WinWaitActive, %reelfoldername%, , 5, ,
	if ErrorLevel
	{
		; print error message after 5 seconds
		MsgBox, 0, Next+, NDNP_QR_hotkeys.ahk`n`nCannot find folder %reelfoldername%`n`nOptions:`n`t"File > Open Reel Folder"`n`t"Edit > Reel Folder > Set Path"
		Return
	}
	Sleep, 100
	
	; move to the next folder
	SetKeyDelay, 50
	Send, {Down %navskip%}
	SetKeyDelay, 10
	Sleep, 100

	Gosub, OpenFirstTIFF ; NDNP_QR_navigation.ahk

	; zoom out
	Send, {NumpadSub 20}

	; send First Impression backward
	WinSet, Bottom,, ahk_id %firstid%
	Sleep, 300

	GoSub, ExtractMeta ; NDNP_QR_metadata.ahk

	; activate First Impression
	WinActivate, ahk_id %firstid%
	
	; bring metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Edition
	WinSet, Top,, Date
	WinSet, Top,, Volume
	WinSet, Top,, Issue
	WinSet, Top,, Questionable

	nextplusflag = 0

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
	Gosub, DeactivateWindows
	Sleep, 200
	
	previousflag = 1

	; NDNP_QR_navigation.ahk
	Gosub, ReelFolderCheck
	Gosub, CloseFirstImpressionWindows
	
	; activate reel folder
	SetTitleMatchMode 1
	WinActivate, %reelfoldername%, , , ,
	WinWaitActive, %reelfoldername%, , 5, ,
	if ErrorLevel
	{
		; print error message after 5 seconds
		MsgBox, 0, Previous, NDNP_QR_hotkeys.ahk`n`nCannot find folder %reelfoldername%`n`nOptions:`n`t"File > Open Reel Folder"`n`t"Edit > Reel Folder > Set Path"
		Return
	}
	Sleep, 100
	
	; move to the previous folder
	SetKeyDelay, 50
	Send, {Up %navskip%}
	SetKeyDelay, 10
	Sleep, 100

	Gosub, OpenFirstTIFF ; NDNP_QR_navigation.ahk

	; zoom out
	Send, {Enter}

	previousflag = 0

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
	Gosub, DeactivateWindows
	Sleep, 200
	
	previousplusflag = 1

	IfWinNotExist, Metadata
	{
		Gosub, RedrawMetaWindow ; NDNP_QR_metadata.ahk
	}
	
	Gui, 7:Destroy ; Questionable
	Gui, 9:Destroy ; Edition Label
	
	; NDNP_QR_navigation.ahk
	Gosub, ReelFolderCheck
	Gosub, CloseFirstImpressionWindows
	
	; activate reel folder
	SetTitleMatchMode 1
	WinActivate, %reelfoldername%, , , ,
	WinWaitActive, %reelfoldername%, , 5, ,
	if ErrorLevel
	{
		; print error message after 5 seconds
		MsgBox, 0, Previous+, NDNP_QR_hotkeys.ahk`n`nCannot find folder %reelfoldername%`n`nOptions:`n`t"File > Open Reel Folder"`n`t"Edit > Reel Folder > Set Path"
		Return
	}
	Sleep, 100
	
	; move to the previous folder
	SetKeyDelay, 50
	Send, {Up %navskip%}
	SetKeyDelay, 10
	Sleep, 100

	Gosub, OpenFirstTIFF ; NDNP_QR_navigation.ahk

	; zoom out
	Send, {NumpadSub 20}

	; send First Impression backward
	WinSet, Bottom,, ahk_id %firstid%
	Sleep, 300

	GoSub, ExtractMeta ; NDNP_QR_metadata.ahk

	; activate First Impression
	WinActivate, ahk_id %firstid%
	
	; bring metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Edition
	WinSet, Top,, Date
	WinSet, Top,, Volume
	WinSet, Top,, Issue
	WinSet, Top,, Questionable

	previousplusflag = 0

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
	Gosub, DeactivateWindows
	Sleep, 200
	
	gotoflag = 1

	; NDNP_QR_navigation.ahk
	Gosub, ReelFolderCheck
	Gosub, CloseFirstImpressionWindows
	Gosub, GoToIssue
	Gosub, OpenFirstTIFF

	; zoom out
	Send, {Enter}

	gotoflag = 0

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
	Gosub, DeactivateWindows
	Sleep, 200
	
	gotoplusflag = 1

	IfWinNotExist, Metadata
	{
		Gosub, RedrawMetaWindow ; NDNP_QR_metadata.ahk
	}
	
	Gui, 7:Destroy ; Questionable
	Gui, 9:Destroy ; Edition Label
	
	; NDNP_QR_navigation.ahk
	Gosub, ReelFolderCheck
	Gosub, CloseFirstImpressionWindows
	Gosub, GoToIssue
	Gosub, OpenFirstTIFF
	
	; zoom out
	Send, {NumpadSub 20}

	; send First Impression backward
	WinSet, Bottom,, ahk_id %firstid%
	Sleep, 300

	Gosub, ExtractMeta ; NDNP_QR_metadata.ahk
	
	; activate First Impression
	WinActivate, ahk_id %firstid%
	
	; bring metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Edition
	WinSet, Top,, Date
	WinSet, Top,, Volume
	WinSet, Top,, Issue
	WinSet, Top,, Questionable

	gotoplusflag = 0

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
	Gosub, DeactivateWindows
	Sleep, 200
	
	metadataflag = 1

	IfWinNotExist, Metadata
	{
		Gosub, RedrawMetaWindow ; NDNP_QR_metadata.ahk
	}
	
	Gui, 7:Destroy ; Questionable
	Gui, 9:Destroy ; Edition Label
	SetTitleMatchMode 2
	WinGet, firstid, ID, First Impression
	
	Gosub, IssueFolderPath ; NDNP_QR_navigation.ahk
	GoSub, ExtractMeta ; NDNP_QR_metadata.ahk

	; activate First Impression
	WinActivate, ahk_id %firstid%
	
	; bring metadata windows to front
	WinSet, Top,, Metadata
	WinSet, Top,, Edition
	WinSet, Top,, Date
	WinSet, Top,, Volume
	WinSet, Top,, Issue
	WinSet, Top,, Questionable

	metadataflag = 0

	hotkeys++
	metadatascore++
	ControlSetText, Static3, METADATA, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
	ControlSetText, Static16, %metadatascore%, NDNP_QR
Return
; =============================METADATA

; =============================RESTORE METADATA WINDOW
; restores/redraws the Metadata window
; Hotkey: Win + Alt + m
#!m::
	Gosub, DeactivateWindows
	Sleep, 200
	
	Gosub, RedrawMetaWindow ; NDNP_QR_metadata.ahk
Return
; =============================RESTORE METADATA WINDOW

; =============================METADATA WINDOWS
; creates separate windows for the issue metadata
; Hotkey: Ctrl + Alt + m
^!m::
	Gosub, DeactivateWindows
	Sleep, 200
	
	Gosub, CreateMetaWindows ; NDNP_QR_metadata.ahk
Return
; =============================METADATA WINDOWS

; =============================ZOOM IN
; First Impression masthead view
; HotKey = Alt + k
!k::
	Gosub, DeactivateWindows
	Sleep, 200
	
	; get the First Impression window id#
	SetTitleMatchMode 2
	WinGet, firstid, ID, First Impression

	; activate First Impression
	SetTitleMatchMode 1
	WinActivate, ahk_id %firstid%
	Sleep, 100
	
	; zoom in
	Send, {Enter}
	Sleep, 100
	Send, {NumpadSub %imagewidth%}
	
	; bring all metadata windows to front
	WinSet, Top,, Metadata	
	WinSet, Top,, Date
	WinSet, Top,, Questionable
	WinSet, Top,, Volume
	WinSet, Top,, Issue
	WinSet, Top,, Edition
	WinSet, Top,, Timer	
	
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
	Gosub, DeactivateWindows
	Sleep, 200
	
	; get the First Impression window id#
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
	Gosub, DeactivateWindows
	Sleep, 200
	
	viewissuexmlflag = 1

	Gosub, IssueFolderPath ; NDNP_QR_navigation.ahk

	; open the issue.xml file in default application
	Run, "%issuefolderpath%\%issuefoldername%.xml"
		
	viewissuexmlflag = 0

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
	Gosub, DeactivateWindows
	Sleep, 200
	
	editissuexmlflag = 1

	Gosub, IssueFolderPath ; NDNP_QR_navigation.ahk

	; open the issue.xml file in Notepad++
	Run, "%notepadpath%\notepad++.exe" "%issuefolderpath%\%issuefoldername%.xml"		
		
	editissuexmlflag = 0

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
	Gosub, DeactivateWindows
	Sleep, 200
	
	; NDNP_QR_navigation.ahk
	Gosub, ReelFolderCheck
	Gosub, ResetReel
	
	hotkeys++
	backscore++
	ControlSetText, Static3, BACK %backscore%, NDNP_QR
	ControlSetText, Static14, %hotkeys%, NDNP_QR
Return
; =============================BACK

; =============================REEL NOTES
; date (Ctrl + 1)
^1::
	Send, %Ctrl1%
	Send, {Left}
Return

; volume number (Ctrl + 2)
^2::
	Send, %Ctrl2%
	Send, {Left}
Return

; issue number (Ctrl + 3)
^3::
	Send, %Ctrl3%
	Send, {Left}
Return

; odd number of pages (Ctrl + 4)
^4::
	Send, %Ctrl4%
Return

; front page lists X pages (Ctrl + 5)
^5::
	Send, %Ctrl5%
	Send, {Left 7}
Return

; add section label (Ctrl + 6)
^6::
	Send, %Ctrl6%
	Send, {Left}
Return

; add edition label (Ctrl + 7)
^7::
	Send, %Ctrl7%
	Send, {Left}
Return

; user-defined (Ctrl + 8)
^8::
	Send, %Ctrl8%
Return

; user-defined (Ctrl + 9)
^9::
	Send, %Ctrl9%
Return

; user-defined (Ctrl + 0)
^0::
	Send, %Ctrl0%
Return
; =============================REEL NOTES
