/*
 * NewspaperNotes.ahk
 *
 * Description: tool for digital newspaper metadata data entry
 */

#NoEnv
#persistent

; =====================================VARIABLES
; STANDARD notes
Ctrl1 = Incorrect date printed on front page: .
Ctrl2 = Issue number out of sequence.
Ctrl3 = Volume number out of sequence.
Ctrl3a = Volume and issue numbers out of sequence.
Ctrl4 = Incorrect issue number printed on front page: No. .
Ctrl5 = Incorrect volume number printed on front page: Vol. .
Ctrl5a = Incorrect volume and issue numbers printed on front page: Vol. , No. .
Ctrl6 = Missing page .
Ctrl6a = Missing pages .
Ctrl7 = Incorrect page number on page : printed as "".

; USER-DEFINED notes
input =
Ctrl8 =
Ctrl9 =
Ctrl0 =
value8 =
value9 =
value0 =

; scoreboard
hotkeys = 0
keystrokes = 0
scoresave = 0
score1 = 0
score2 = 0
score3 = 0
score4 = 0
score5 = 0
score6 = 0
score7 = 0
score8 = 0
score9 = 0
score10 = 0
; =====================================VARIABLES


Gui, Color, d0d0d0, 912206
Gui, Show, h0 w0, Newspaper Notes


; ===========================================MENUS
; see MENU FUNCTIONS at end of file

; File menu
Menu, FileMenu, Add, Display &Standard Values, DisplayStandardValues
Menu, FileMenu, Add, Display &User-Defined Values, DisplayUserValues
Menu, FileMenu, Add
Menu, FileMenu, Add, &Reload, Reload
Menu, FileMenu, Add, E&xit, Exit

; Edit menu
Menu, EditMenu, Add, Ctrl + &8, Control8
Menu, EditMenu, Add, Ctrl + &9, Control9
Menu, EditMenu, Add, Ctrl + &0, Control0

; Help menu
Menu, HelpMenu, Add, &Documentation, Documentation
Menu, HelpMenu, Add, &Newspaper Notes, NewspaperNotes
Menu, HelpMenu, Add, &About, About

; create menus
Menu, MyMenuBar, Add, &File, :FileMenu
Menu, MyMenuBar, Add, &Edit, :EditMenu
Menu, MyMenuBar, Add, &Help, :HelpMenu

; create menu bar
Gui, Menu, MyMenuBar
; ===========================================MENUS


; ===========================================LABELS
Gui, Font,, Arial

; scoreboard labels - Last Hotkey: Static 4
Gui, Add, Text, x18  y10 w100 h20, HotKeys
Gui, Add, Text, x90  y10 w100 h20, Keystrokes
Gui, Add, Text, x190 y10 w100 h20, Last HotKey
Gui, Add, Text, x205 y35 w65  h20, 

; keystroke counter labels
Gui, Add, Text, x295 y10  w100 h20, SAVE
Gui, Add, Text, x295 y22  w100 h15, Ctrl + .
Gui, Add, Text, x15  y75  w100 h20, Date (1)
Gui, Add, Text, x85  y75  w100 h20, Iss Seq (2)
Gui, Add, Text, x155 y75  w100 h20, Vol Seq (3)
Gui, Add, Text, x225 y75  w100 h20, Iss # (4)
Gui, Add, Text, x295 y75  w100 h20, Vol # (5)
Gui, Add, Text, x15  y130 w100 h20, Missing (6)
Gui, Add, Text, x85  y130 w100 h20, Page # (7)
Gui, Add, Text, x155 y130 w100 h20, Ctrl + 8
Gui, Add, Text, x225 y130 w100 h20, Ctrl + 9
Gui, Add, Text, x295 y130 w100 h20, Ctrl + 0
; ===========================================LABELS


; ===========================================DATA
Gui, Font, s15,

; scoreboard counters
; Hotkeys & Keystrokes: Static 17 & 18
Gui, Add, Text, x30  y30 w35 h25, 0
Gui, Add, Text, x105 y30 w50 h25, 0

; script keystroke counters
; SAVE Static 19 
Gui, Add, Text, x300 y35 w35 h20, 0
; row 1: Static 20-24
Gui, Add, Text, x20  y90 w50 h25, 0
Gui, Add, Text, x90  y90 w50 h25, 0
Gui, Add, Text, x160 y90 w50 h25, 0
Gui, Add, Text, x230 y90 w50 h25, 0
Gui, Add, Text, x300 y90 w50 h25, 0
; row 2: Static 25-29
Gui, Add, Text, x20  y145 w50 h25, 0
Gui, Add, Text, x90  y145 w50 h25, 0
Gui, Add, Text, x160 y145 w50 h25, 0
Gui, Add, Text, x230 y145 w50 h25, 0
Gui, Add, Text, x300 y145 w50 h25, 0
; ===========================================DATA


; ===========================================BOXES
; totals section
Gui, Add, GroupBox, x10 y-5 w275 h70,       

; HotKeys, Keystrokes, & Last Hotkey
Gui, Add, GroupBox, x18  y15 w65 h45,         
Gui, Add, GroupBox, x90  y15 w90 h45,           
Gui, Add, GroupBox, x190 y15 w85 h45,           

; Save counter
Gui, Add, GroupBox, x290 y-5 w65 h70,       

; script counters
Gui, Add, GroupBox, x10  y60  w65 h60, 
Gui, Add, GroupBox, x80  y60  w65 h60, 
Gui, Add, GroupBox, x150 y60  w65 h60, 
Gui, Add, GroupBox, x220 y60  w65 h60, 
Gui, Add, GroupBox, x290 y60  w65 h60, 
Gui, Add, GroupBox, x10  y115 w65 h60, 
Gui, Add, GroupBox, x80  y115 w65 h60, 
Gui, Add, GroupBox, x150 y115 w65 h60, 
Gui, Add, GroupBox, x220 y115 w65 h60, 
Gui, Add, GroupBox, x290 y115 w65 h60, 
; ===========================================BOXES


WinGetPos, winX, winY, winWidth, winHeight, Newspaper Notes
winX+=%winWidth%
Gui, Show, x%winX% y%winY% h180 w365, Newspaper Notes
winactivate, Newspaper Notes


; ========================================SCRIPTS
; ==================SAVE
; Save and Close Notepad document (Ctrl + . )
^.::
  SetTitleMatchMode 2
	WinMenuSelectItem, Notepad,, File, Save,,,,,, Notepad++,
	WinMenuSelectItem, Notepad,, File, Exit,,,,,, Notepad++,
	; update the scoreboard
	hotkeys+=1
	scoresave+=1
	ControlSetText, Static4, SAVE, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static19, %scoresave%, Newspaper Notes
Return
; ==================SAVE

; ==================STANDARD
; incorrect date (Ctrl + 1)
^1::
	Send, %Ctrl1%
	Send, {Left}
	; update the scoreboard
	hotkeys+=1
	keystrokes+=40
	score1+=40
	ControlSetText, Static4, Date, Newspaper Notes
	ControlSetText, Static17, %hotkeys%, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static20, %score1%, Newspaper Notes
Return

; issue sequence (Ctrl + 2)
^2::
	Send, %Ctrl2%
	; update the scoreboard
	hotkeys+=1
	keystrokes+=29
	score2+=29
	ControlSetText, Static4, Issue Seq., Newspaper Notes
	ControlSetText, Static17, %hotkeys%, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static21, %score2%, Newspaper Notes
Return

; volume sequence (Ctrl + 3)
^3::
	Send, %Ctrl3%
	; update the scoreboard
	hotkeys+=1
	keystrokes+=30
	score3+=30
	ControlSetText, Static4, Volume Seq., Newspaper Notes
	ControlSetText, Static17, %hotkeys%, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static22, %score3%, Newspaper Notes
Return

; volume and issue sequence (Ctrl+Alt+3)
^!3::
	Send, %Ctrl3a%
	; update the scoreboard
	hotkeys+=1
	keystrokes+=41
	score3+=41
	ControlSetText, Static4, Vol Iss Seq., Newspaper Notes
	ControlSetText, Static17, %hotkeys%, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static22, %score3%, Newspaper Notes
Return

; issue number (Ctrl + 4)
^4::
	Send, %Ctrl4%
	Send, {Left}
	; update the scoreboard
	hotkeys+=1
	keystrokes+=52
	score4+=52
	ControlSetText, Static4, Issue #, Newspaper Notes
	ControlSetText, Static17, %hotkeys%, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static23, %score4%, Newspaper Notes
Return

; volume number (Ctrl + 5)
^5::
	Send, %Ctrl5%
	Send, {Left}
	; update the scoreboard
	hotkeys+=1
	keystrokes+=54
	score5+=54
	ControlSetText, Static4, Volume #, Newspaper Notes
	ControlSetText, Static17, %hotkeys%, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static24, %score5%, Newspaper Notes
Return

; volume and issue number (Ctrl + Alt + 5)
^!5::
	Send, %Ctrl5a%
	Send, {Left 7}
	; update the scoreboard
	hotkeys+=1
	keystrokes+=77
	score5+=77
	ControlSetText, Static4, Vol Iss #, Newspaper Notes
	ControlSetText, Static17, %hotkeys%, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static24, %score5%, Newspaper Notes
Return

; missing page (Ctrl + 6)
^6::
	Send, %Ctrl6%
	Send, {Left}
	; update the scoreboard
	hotkeys+=1
	keystrokes+=15
	score6+=15
	ControlSetText, Static4, Missing Pg., Newspaper Notes
	ControlSetText, Static17, %hotkeys%, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static25, %score6%, Newspaper Notes
Return

; missing pages (Ctrl + Alt + 6)
^!6::
	Send, %Ctrl6a%
	Send, {Left}
	; update the scoreboard
	hotkeys+=1
	keystrokes+=16
	score6+=16
	ControlSetText, Static4, Missing Pgs., Newspaper Notes
	ControlSetText, Static17, %hotkeys%, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static25, %score6%, Newspaper Notes
Return

; page number (Ctrl + 7)
^7::
	Send, %Ctrl7%
	Send, {Left 16}
	; update the scoreboard
	hotkeys+=1
	keystrokes+=62
	score7+=62
	ControlSetText, Static4, Page No., Newspaper Notes
	ControlSetText, Static17, %hotkeys%, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static26, %score7%, Newspaper Notes
Return
; ==================STANDARD

; ==================USER-DEFINED
; Ctrl + 8
^8::
	Send, %Ctrl8%
	; update the scoreboard
	hotkeys+=1
	keystrokes+=%value8%
	score8+=%value8%
	ControlSetText, Static4, Ctrl + 8, Newspaper Notes
	ControlSetText, Static17, %hotkeys%, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static27, %score8%, Newspaper Notes
Return

; Ctrl + 9
^9::
	Send, %Ctrl9%
	; update the scoreboard
	hotkeys+=1
	keystrokes+=%value9%
	score9+=%value9%
	ControlSetText, Static4, Ctrl + 9, Newspaper Notes
	ControlSetText, Static17, %hotkeys%, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static28, %score9%, Newspaper Notes
Return

; Ctrl + 0
^0::
	Send, %Ctrl0%
	; update the scoreboard
	hotkeys+=1
	keystrokes+=%value0%
	score10+=%value0%
	ControlSetText, Static4, Ctrl + 0, Newspaper Notes
	ControlSetText, Static17, %hotkeys%, Newspaper Notes
	ControlSetText, Static18, %keystrokes%, Newspaper Notes
	ControlSetText, Static29, %score10%, Newspaper Notes
Return
; ==================USER-DEFINED
; ========================================SCRIPTS


; ====================================MENU FUNCTIONS
; =================FILE
; display user-defined variable values
DisplayUserValues:
MsgBox,
(
USER-DEFINED VALUES

Ctrl + 8 = %Ctrl8%

Ctrl + 9 = %Ctrl9%

Ctrl + 0 = %Ctrl0%
)
Return

; display standard variable values
DisplayStandardValues:
MsgBox,
(
STANDARD VALUES

Ctrl + 1 = %Ctrl1%

Ctrl + 2 = %Ctrl2%

Ctrl + 3 = %Ctrl3%

Ctrl + Alt + 3 = %Ctrl3a%

Ctrl + 4 = %Ctrl4%

Ctrl + 5 = %Ctrl5%

Ctrl + Alt + 5 = `n`t%Ctrl5a%

Ctrl + 6 = %Ctrl6%

Ctrl + Alt + 6 = %Ctrl6a%

Ctrl + 7 = %Ctrl7%
)
Return

; reload application
Reload:
Reload

; exit application
Exit:
ExitApp
; =================FILE

; =================EDIT
; USER-DEFINED input functions
/*
***********************begin template***********************
; define function name
ControlX:
; create input box, variable, window label, window text,, width, height,,,,,input field display
InputBox, input, Ctrl X, Enter the value for Ctrl + X,, 250, 125,,,,,%CtrlX%
	if ErrorLevel ; Cancel ends function
		Return
	else
	{
		; OK assigns input to CtrlX
		CtrlX = %input%
		; and length of string to valueX for scoreboard
		StringLen, valueX, CtrlX
	}
Return
************************end template************************
*/

; Ctrl8 input
Control8:
InputBox, input, Ctrl 8, Enter the value for Ctrl + 8,, 250, 125,,,,,%Ctrl8%
	if ErrorLevel
		Return
	else
	{
		Ctrl8 = %input%
		StringLen, value8, Ctrl8
	}
Return

; Ctrl9 input
Control9:
InputBox, input, Ctrl 9, Enter the value for Ctrl + 9,, 250, 125,,,,,%Ctrl9%
	if ErrorLevel
		Return
	else
	{
		Ctrl9 = %input%
		StringLen, value9, Ctrl9
	}
Return

; Ctrl0 input
Control0:
InputBox, input, Ctrl 0, Enter the value for Ctrl + 0,, 250, 125,,,,,%Ctrl0%
	if ErrorLevel
		Return
	else
	{
		Ctrl0 = %input%
		StringLen, value0, Ctrl0
	}
Return
; =================EDIT

; =================HELP
; display version info
About:
MsgBox NewspaperNotes.ahk`nVersion 1.5`nAndrew.Weidner@unt.edu
Return

; open NewspaperNotes.ahk wiki page
Documentation:
	Run, https://github.com/drewhop/AutoHotkey/wiki/NewspaperNotes
Return

; open Newspaper Notes input standard wiki page
NewspaperNotes:
	Run, http://digitalprojects.library.unt.edu/projects/index.php/Newspaper_Notes
Return
; =================HELP
; ====================================MENU FUNCTIONS

GuiClose:
ExitApp
