/*
 * Dashboard Tools.ahk
 *
 * Description: tool for searching and editing records published
 *              in the Portal to Texas History & UNT Digital Library
 */

#NoEnv
#persistent

; ===========================================VARIABLES
; EDIT THIS VARIABLE (if necessary)
; for the path to the Mozilla Firefox folder
firefoxpath = C:\Program Files (x86)\Mozilla Firefox

; search
input =
searchstring = Do not use quotes.

; scoreboard
hotkeys = 0
portalscore = 0
digilibscore = 0
editscore = 0
displayscore = 0
linkyscore = 0
seleniumscore = 0
; ===========================================VARIABLES

Gui, Color, d0d0d0, 912206
Gui, Show, h0 w0, DashboardTools

; =====================================MENUS
; see MENU FUNCTIONS at end of file

; File menu
Menu, FileMenu, Add, &Portal to TX History, OpenPortal
Menu, FileMenu, Add, UNT &Digital Library, OpenDigiLibrary
Menu, FileMenu, Add
Menu, FileMenu, Add, &Reload, Reload
Menu, FileMenu, Add, E&xit, Exit

; Edit menu
; Dashboard submenu
Menu, DashMenu, Add, &Portal, DashPortal
Menu, DashMenu, Add, &Digital Library, DashDigiLibrary
Menu, EditMenu, Add, &Dashboard, :DashMenu
; New Record Creator submenu
Menu, NRCMenu, Add, &Portal, NRCPortal
Menu, NRCMenu, Add, &Digital Library, NRCDigiLibrary
Menu, EditMenu, Add, &New Record Creator, :NRCMenu
Menu, EditMenu, Add
Menu, EditMenu, Add, &Edit Firefox Path, EditFirefoxPath
Menu, EditMenu, Add, &Display Firefox Path, DisplayFirefoxPath

; Search menu
; Portal submenu
Menu, PortalMenu, Add, &Full Text, PortalFullText
Menu, PortalMenu, Add, &Metadata, PortalMetadata
Menu, PortalMenu, Add, &Title, PortalTitle
Menu, PortalMenu, Add, &Subject, PortalSubject
Menu, PortalMenu, Add, &Creator, PortalCreator
Menu, SearchMenu, Add, &Portal Search, :PortalMenu
Menu, SearchMenu, Add
; Digital Library submenu
Menu, DigiLibMenu, Add, &Full Text, DigiLibFullText
Menu, DigiLibMenu, Add, &Metadata, DigiLibMetadata
Menu, DigiLibMenu, Add, &Title, DigiLibTitle
Menu, DigiLibMenu, Add, &Subject, DigiLibSubject
Menu, DigiLibMenu, Add, &Creator, DigiLibCreator
Menu, SearchMenu, Add, &Digital Library Search, :DigiLibMenu
Menu, SearchMenu, Add
Menu, SearchMenu, Add, Display &Search Hotkeys, DisplaySearch

; Help menu
Menu, HelpMenu, Add, &Documentation, Documentation
Menu, HelpMenu, Add, &About, About

; create menus
Menu, MenuBar, Add, &File, :FileMenu
Menu, MenuBar, Add, &Edit, :EditMenu
Menu, MenuBar, Add, &Search, :SearchMenu
Menu, MenuBar, Add, &Help, :HelpMenu

; create menu toolbar
Gui, Menu, MenuBar
; =====================================MENUS

; =====================================LABELS
; scoreboard labels - Last Hotkey: Static 3
Gui, Font,, Arial
Gui, Add, Text, x15  y10 w40 h15, HotKeys
Gui, Add, Text, x95  y10 w70 h15, Last HotKey
Gui, Add, Text, x105 y35 w60 h15,
Gui, Add, Text, x175 y10 w70 h15, Portal Search
Gui, Add, Text, x255 y10 w70 h15, DigiLib Search


Gui, Add, Text, x15  y70 w80 h15, Edit ( z )
Gui, Add, Text, x95  y70 w80 h15, Display ( q )
Gui, Add, Text, x175 y70 w80 h15, Linky ( k )
Gui, Add, Text, x255 y70 w80 h15, Selenium ( , )
; =====================================LABELS

; =====================================DATA
; scoreboard counters
Gui, Font, s15,

; Static 10-12
Gui, Add, Text, x25  y30 w55 h25, 0
Gui, Add, Text, x185 y30 w55 h25, 0
Gui, Add, Text, x265 y30 w55 h25, 0

; Static 13-16
Gui, Add, Text, x25  y90 w55 h25, 0
Gui, Add, Text, x105 y90 w55 h25, 0
Gui, Add, Text, x185 y90 w55 h25, 0
Gui, Add, Text, x265 y90 w55 h25, 0
; =====================================DATA

; =====================================BOXES
; decorative box
Gui, Add, GroupBox, x2 y-12 w340 h140,       

; row 1
Gui, Add, GroupBox, x15  y15 w75 h45, 
Gui, Add, GroupBox, x95  y15 w75 h45, 
Gui, Add, GroupBox, x175 y15 w75 h45, 
Gui, Add, GroupBox, x255 y15 w75 h45, 

; row 2
Gui, Add, GroupBox, x15  y75 w75 h45, 
Gui, Add, GroupBox, x95  y75 w75 h45, 
Gui, Add, GroupBox, x175 y75 w75 h45, 
Gui, Add, GroupBox, x255 y75 w75 h45, 
; =====================================BOXES

WinGetPos, winX, winY, winWidth, winHeight, DashboardTools
winX+=%winWidth%
Gui, Show, x%winX% y%winY% h130 w345, DashboardTools
winactivate, DashboardTools

; =====================================SCRIPTS
; ==============================PORTAL SEARCH
; PortalFullText:
!1::
  InputBox, searchstring, Portal FULL TEXT,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://texashistory.unt.edu/search/?q="%searchstring%"&t=fulltext
		hotkeys++
		portalscore++
		ControlSetText, Static3, P text, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static11, %portalscore%, DashboardTools
Return

; PortalMetadata:
!2::
	InputBox, searchstring, Portal METADATA,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://texashistory.unt.edu/search/?q="%searchstring%"&t=metadata
		hotkeys++
		portalscore++
		ControlSetText, Static3, P meta, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static11, %portalscore%, DashboardTools
Return

; PortalTitle:
!3::
	InputBox, searchstring, Portal TITLE,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://texashistory.unt.edu/search/?q="%searchstring%"&t=dc_title
		hotkeys++
		portalscore++
		ControlSetText, Static3, P title, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static11, %portalscore%, DashboardTools
Return

; PortalSubject:
!4::
	InputBox, searchstring, Portal SUBJECT,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://texashistory.unt.edu/search/?q="%searchstring%"&t=dc_subject
		hotkeys++
		portalscore++
		ControlSetText, Static3, P subject, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static11, %portalscore%, DashboardTools
Return

; PortalCreator:
!5::
	InputBox, searchstring, Portal CREATOR,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://texashistory.unt.edu/search/?q="%searchstring%"&t=dc_creator
		hotkeys++
		portalscore++
		ControlSetText, Static3, P creator, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static11, %portalscore%, DashboardTools
Return
; ==============================PORTAL SEARCH


; ==============================DIGITAL LIBRARY SEARCH
; DigiLibFullText:
!6::
	InputBox, searchstring, DigiLib FULL TEXT,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://digital.library.unt.edu/search/?q="%searchstring%"&t=fulltext
		hotkeys++
		digilibscore++
		ControlSetText, Static3, DL text, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static12, %digilibscore%, DashboardTools
Return

; DigiLibMetadata:
!7::
	InputBox, searchstring, DigiLib METADATA,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://digital.library.unt.edu/search/?q="%searchstring%"&t=metadata
		hotkeys++
		digilibscore++
		ControlSetText, Static3, DL meta, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static12, %digilibscore%, DashboardTools
Return

; DigiLibTitle:
!8::
	InputBox, searchstring, DigiLib TITLE,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://digital.library.unt.edu/search/?q="%searchstring%"&t=dc_title
		hotkeys++
		digilibscore++
		ControlSetText, Static3, DL title, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static12, %digilibscore%, DashboardTools
Return

; DigiLibSubject:
!9::
	InputBox, searchstring, DigiLib SUBJECT,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://digital.library.unt.edu/search/?q="%searchstring%"&t=dc_subject
		hotkeys++
		digilibscore++
		ControlSetText, Static3, DL subject, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static12, %digilibscore%, DashboardTools
Return

; DigiLibCreator:
!0::
	InputBox, searchstring, DigiLib CREATOR,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://digital.library.unt.edu/search/?q="%searchstring%"&t=dc_creator
		hotkeys++
		digilibscore++
		ControlSetText, Static3, DL creator, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static12, %digilibscore%, DashboardTools
Return
; ==============================DIGITAL LIBRARY SEARCH


; ==============================PORTAL DASHBOARD SEARCH
; PortalDashDescription:
^!1::
	InputBox, searchstring, Portal Dash DESCRIPTION,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://edit.texashistory.unt.edu/?q="%searchstring%"&t=dc_description
		hotkeys++
		portalscore++
		ControlSetText, Static3, PD desc, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static11, %portalscore%, DashboardTools
Return

; PortalDashMetadata:
^!2::
	InputBox, searchstring, Portal Dash METADATA,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://edit.texashistory.unt.edu/?q="%searchstring%"&t=metadata
		hotkeys++
		portalscore++
		ControlSetText, Static3, PD meta, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static11, %portalscore%, DashboardTools
Return

; PortalDashTitle:
^!3::
	InputBox, searchstring, Portal Dash TITLE,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://edit.texashistory.unt.edu/?q="%searchstring%"&t=dc_title
		hotkeys++
		portalscore++
		ControlSetText, Static3, PD title, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static11, %portalscore%, DashboardTools
Return

; PortalDashSubject:
^!4::
	InputBox, searchstring, Portal Dash SUBJECT,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://edit.texashistory.unt.edu/?q="%searchstring%"&t=dc_subject
		hotkeys++
		portalscore++
		ControlSetText, Static3, PD subject, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static11, %portalscore%, DashboardTools
Return

; PortalDashCreator:
^!5::
	InputBox, searchstring, Portal Dash CREATOR,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://edit.texashistory.unt.edu/?q="%searchstring%"&t=dc_creator
		hotkeys++
		portalscore++
		ControlSetText, Static3, PD creator, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static11, %portalscore%, DashboardTools
Return
; ==============================PORTAL DASHBOARD SEARCH


; ==============================DIGITAL LIBRARY DASHBOARD SEARCH
; DigiLibDashDescription:
^!6::
	InputBox, searchstring, DigiLib Dash DESCRIPTION,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://edit.digital.library.unt.edu/?q="%searchstring%"&t=dc_description
		hotkeys++
		digilibscore++
		ControlSetText, Static3, DLD desc, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static12, %digilibscore%, DashboardTools
Return

; DigiLibDashMetadata:
^!7::
	InputBox, searchstring, DigiLib Dash METADATA,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://edit.digital.library.unt.edu/?q="%searchstring%"&t=metadata
		hotkeys++
		digilibscore++
		ControlSetText, Static3, DLD meta, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static12, %digilibscore%, DashboardTools
Return

; DigiLibDashTitle:
^!8::
	InputBox, searchstring, DigiLib Dash TITLE,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://edit.digital.library.unt.edu/?q="%searchstring%"&t=dc_title
		hotkeys++
		digilibscore++
		ControlSetText, Static3, DLD title, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static12, %digilibscore%, DashboardTools
Return

; DigiLibDashSubject:
^!9::
	InputBox, searchstring, DigiLib Dash SUBJECT,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://edit.digital.library.unt.edu/?q="%searchstring%"&t=dc_subject
		hotkeys++
		digilibscore++
		ControlSetText, Static3, DLD subject, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static12, %digilibscore%, DashboardTools
Return

; DigiLibDashCreator:
^!0::
	InputBox, searchstring, DigiLib Dash CREATOR,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://edit.digital.library.unt.edu/?q="%searchstring%"&t=dc_creator
		hotkeys++
		digilibscore++
		ControlSetText, Static3, DLD creator, DashboardTools
		ControlSetText, Static10, %hotkeys%, DashboardTools
		ControlSetText, Static12, %digilibscore%, DashboardTools
Return
; ==============================DIGITAL LIBRARY DASHBOARD SEARCH


; ==============================
; Edit
; opens Portal / Digital Library Dashboard editor
; from Brief Record or Full Record views

; Hotkey = Alt + z
!z::

	; bring focus to the URL box
	Send, {CtrlDown}l{CtrlUp}
	Sleep, 5

	; move cursor to start of URL
	Send, {Home}
	Sleep, 5

	; add "edit." to beginning of URL
	Send, edit.
	Sleep, 5

	; open the record in Dashboard
	Send, {Enter}
  
	; update the scoreboard
	hotkeys++
	editscore++
	ControlSetText, Static3, EDIT, DashboardTools
	ControlSetText, Static10, %hotkeys%, DashboardTools
	ControlSetText, Static13, %editscore%, DashboardTools
  
return
; ==============================


; ==============================
; Display
; opens Portal / Digital library public view
; from Dashboard editor

; Hotkey = Alt + q
!q::

	; bring focus to the Firefox URL box
	Send, {CtrlDown}l{CtrlUp}
	Sleep, 5

	; move cursor to start of URL
	Send, {Home}
	Sleep, 5

	; remove "edit." from the beginning of URL
	Send, {Delete 5}
	Sleep, 5

	; open the record in Dashboard
	Send, {Enter}

	; update the scoreboard
	hotkeys++
	displayscore++
	ControlSetText, Static3, DISPLAY, DashboardTools
	ControlSetText, Static10, %hotkeys%, DashboardTools
	ControlSetText, Static14, %displayscore%, DashboardTools
  
return
; ==============================


; ==============================
; Linky
; opens selected links with Linky

; HotKey = Alt + k
!k::

	; activate Linky
	Send, {AppsKey}
	Sleep, 400
	Send, l
	Sleep, 400
	Send, {Down}
	Sleep, 400
	Send, {Enter}

	; focus on the Linky popup window
	SetTitleMatchMode 1
	WinWait, Select Links, , , ,
	IfWinNotActive, Select Links, , , ,
	WinActivate, Select Links, , , ,
	WinWaitActive, Select Links, , , ,
	Sleep, 200

	; open selected links
	Send, {AltDown}s{AltUp}
	Sleep, 200

	; open next results page
	Send, {ShiftDown}{Tab 2}{ShiftUp}
	Sleep, 100
	Send, {Enter}

	; update the scoreboard
	hotkeys++
	linkyscore++
	ControlSetText, Static3, LINKY, DashboardTools
	ControlSetText, Static10, %hotkeys%, DashboardTools
	ControlSetText, Static15, %linkyscore%, DashboardTools
  
return
; ==============================


; ==============================
; Selenium
; activates test suite loaded in Selenium IDE sidebar

; HotKey = Alt + j
^,::
	Sleep, 500
	
	; bring focus to Selenium sidebar
	Send, {ShiftDown}{Tab}
	Sleep, 200
	Send, {ShiftUp}
	Sleep, 200
	Send, {Tab}
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	Send, {Tab}
	Sleep, 100

	; play entire test suite
	Send, {AltDown}a{AltUp}
	Sleep, 500
	Send, p

	; update the scoreboard
	hotkeys++
	seleniumscore++
	ControlSetText, Static3, SELENIUM, DashboardTools
	ControlSetText, Static10, %hotkeys%, DashboardTools
	ControlSetText, Static16, %seleniumscore%, DashboardTools

return
; ==============================
; =====================================SCRIPTS


; =====================================MENU FUNCTIONS
; =================FILE
; Portal to Texas History home page
OpenPortal:
	Run, "%firefoxpath%\firefox.exe" http://texashistory.unt.edu/
Return

; UNT Digital Library home page
OpenDigiLibrary:
	Run, "%firefoxpath%\firefox.exe" http://digital.library.unt.edu
Return

; Portal Dashboard
DashPortal:
	Run, "%firefoxpath%\firefox.exe" http://edit.texashistory.unt.edu
Return

; Digital Library Dashboard
DashDigiLibrary:
	Run, "%firefoxpath%\firefox.exe" http://edit.digital.library.unt.edu
Return

; Portal New Record Creator
NRCPortal:
	Run, "%firefoxpath%\firefox.exe" http://edit.texashistory.unt.edu/nrc
Return

; Digital Library New Record Creator
NRCDigiLibrary:
	Run, "%firefoxpath%\firefox.exe" http://edit.digital.library.unt.edu/nrc
Return

; reload application
Reload:
Reload

; exit application
Exit:
ExitApp
; =================FILE

; =================EDIT
DisplayFirefoxPath:
MsgBox %firefoxpath%
Return

EditFirefoxPath:
	FileSelectFolder, firefoxpath, *C:\, 0, `nSelect the Mozilla Firefox folder:
	if ErrorLevel
		firefoxpath = C:\Program Files (x86)\Mozilla Firefox
Return
; =================EDIT

; =================SEARCH
PortalFullText:
	InputBox, searchstring, Portal FULL TEXT,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://texashistory.unt.edu/search/?q="%searchstring%"&t=fulltext
Return

PortalMetadata:
	InputBox, searchstring, Portal METADATA,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://texashistory.unt.edu/search/?q="%searchstring%"&t=metadata
Return

PortalTitle:
	InputBox, searchstring, Portal TITLE,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://texashistory.unt.edu/search/?q="%searchstring%"&t=dc_title
Return

PortalSubject:
	InputBox, searchstring, Portal SUBJECT,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://texashistory.unt.edu/search/?q="%searchstring%"&t=dc_subject
Return

PortalCreator:
	InputBox, searchstring, DigiLib CREATOR,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://texashistory.unt.edu/search/?q="%searchstring%"&t=dc_creator
Return

DigiLibFullText:
	InputBox, searchstring, DigiLib FULL TEXT,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://digital.library.unt.edu/search/?q="%searchstring%"&t=fulltext
Return

DigiLibMetadata:
	InputBox, searchstring, DigiLib METADATA,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://digital.library.unt.edu/search/?q="%searchstring%"&t=metadata
Return

DigiLibTitle:
	InputBox, searchstring, DigiLib TITLE,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://digital.library.unt.edu/search/?q="%searchstring%"&t=dc_title
Return

DigiLibSubject:
	InputBox, searchstring, DigiLib SUBJECT,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://digital.library.unt.edu/search/?q="%searchstring%"&t=dc_subject
Return

DigiLibCreator:
	InputBox, searchstring, DigiLib CREATOR,,, 250, 100,,,,,%searchstring%
	if ErrorLevel
		Return
	else
		Run, "%firefoxpath%\firefox.exe" http://digital.library.unt.edu/search/?q="%searchstring%"&t=dc_creator
Return

DisplaySearch:
MsgBox,
(
Alt + 1  Portal FULL TEXT
Alt + 2  Portal METADATA
Alt + 3  Portal TITLE
Alt + 4  Portal SUBJECT
Alt + 5  Portal CREATOR

Alt + 6  DigiLib FULL TEXT
Alt + 7  DigiLib METADATA
Alt + 8  DigiLib TITLE
Alt + 9  DigiLib SUBJECT
Alt + 0  DigiLib CREATOR

Ctrl + Alt + 1  Portal Dash DESCRIPTION
Ctrl + Alt + 2  Portal Dash METADATA
Ctrl + Alt + 3  Portal Dash TITLE
Ctrl + Alt + 4  Portal Dash SUBJECT
Ctrl + Alt + 5  Portal Dash CREATOR

Ctrl + Alt + 6  DigiLib Dash DESCRIPTION
Ctrl + Alt + 7  DigiLib Dash METADATA
Ctrl + Alt + 8  DigiLib Dash TITLE
Ctrl + Alt + 9  DigiLib Dash SUBJECT
Ctrl + Alt + 0  DigiLib Dash CREATOR
)
Return
; =================SEARCH

; =================HELP
; display version info
About:
MsgBox DashboardTools.ahk`nVersion 1.3`nAndrew.Weidner@unt.edu
Return

; open DashboardTools.ahk wiki page
Documentation:
Run, "%firefoxpath%\firefox.exe" https://github.com/drewhop/AutoHotkey/wiki/DashboardTools
Return
; =================HELP
; =====================================MENU FUNCTIONS

GuiClose:
ExitApp
