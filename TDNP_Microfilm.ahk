/*
 * TDNP_Microfilm
 *
 * Description: tool for microfilmed newspaper evalation
 */

#NoEnv
#persistent

; =====================================VARIABLES
; file
scannedpath = Q:\scanned_for_tdnp
collationpath = Q:\TDNP\Collation Sheets

; search
searchstring = Do not use quotes.

; tools
input =
inputAS =
inputRR =
mm =
in =
redrat =

; scoreboard
hotkeys = 0
keystrokes = 0
topmonthscore = 0
nextrowscore = 0
highlightscore = 0
; =====================================VARIABLES

Gui, Color, d0d0d0, 912206
Gui, Show, h0 w0, TDNP_Microfilm

; ==========================MENUS
; see MENU FUNCTIONS at end of file

; File menu
Menu, FileMenu, Add, &Collation Sheets, CollationSheets
Menu, FileMenu, Add, &scanned_for_tdnp, OpenTDNP
Menu, FileMenu, Add, &Microfilm Scanning Queue, Microfilm
Menu, FileMenu, Add
Menu, FileMenu, Add, Reloa&d, Reload
Menu, FileMenu, Add, E&xit, Exit

; Search menu
Menu, SearchMenu, Add, US Directory: &Texas, DirectoryTitleTexas
Menu, SearchMenu, Add, US Directory: &Oklahoma, DirectoryTitleOklahoma
Menu, SearchMenu, Add, US Directory: &LCCN, DirectoryLCCN
Menu, SearchMenu, Add
Menu, SearchMenu, Add, Portal (TX): Title, PortalTitle
Menu, SearchMenu, Add, Gateway (OK): Title, GatewayTitle

; Tools menu, Convert & Calculate submenus
Menu, ConvertMenu, Add, &Millimeters to Inches, mm2in
Menu, ConvertMenu, Add, &Inches to Millimeters, in2mm
Menu, ToolsMenu, Add, Con&vert, :ConvertMenu
Menu, ToolsMenu, Add
Menu, CalculateMenu, Add, &Actual Size, ActualSize
Menu, CalculateMenu, Add, &Reduction Ratio, ReductionRatio
Menu, ToolsMenu, Add, &Calculate, :CalculateMenu
Menu, ToolsMenu, Add
Menu, ToolsMenu, Add, &Frequency Definitions, Frequency

; Help menu
Menu, HelpMenu, Add, &Documentation, Documentation
Menu, HelpMenu, Add, &About, About

; create menus
Menu, MenuBar, Add, &File, :FileMenu
Menu, MenuBar, Add, &Search, :SearchMenu
Menu, MenuBar, Add, &Tools, :ToolsMenu
Menu, MenuBar, Add, &Help, :HelpMenu

; create menu toolbar
Gui, Menu, MenuBar
; ==========================MENUS

; ===========================================LABELS
Gui, Font,, Arial

; scoreboard - last hotkey: Static 4
Gui, Add, Text, x18  y10 w100 h20, HotKeys
Gui, Add, Text, x85  y10 w100 h20, Keystrokes
Gui, Add, Text, x170 y10 w100 h20, Last HotKey
Gui, Add, Text, x185 y35 w45  h20, 

; keystroke counters
Gui, Add, Text, x10  y70 w100 h20, TopMonth
Gui, Add, Text, x95  y70 w100 h20, NextRow
Gui, Add, Text, x180 y70 w100 h20, Highlight
; ===========================================LABELS

; ===========================================DATA
Gui, Font, s15,

; Hotkeys & Keystrokes: Static 8-9
Gui, Add, Text, x30  y30 w35 h25, 0
Gui, Add, Text, x100 y30 w50 h25, 0

; TopMonth, NextRow, Highlight: Static 10-12
Gui, Add, Text, x20  y90 w45 h25, 0
Gui, Add, Text, x105 y90 w45 h25, 0
Gui, Add, Text, x190 y90 w45 h25, 0
; ===========================================DATA

; =========================================BOXES
; totals section
Gui, Add, GroupBox, x10 y-5 w245 h70,       

; HotKeys, Keystrokes, & Last Hotkey
Gui, Add, GroupBox, x18  y15 w60  h45,         
Gui, Add, GroupBox, x85  y15 w80 h45,           
Gui, Add, GroupBox, x170 y15 w80 h45,           

; TopMonth, NextRow, Highlight
Gui, Add, GroupBox, x10  y75 w75 h45, 
Gui, Add, GroupBox, x95  y75 w75 h45, 
Gui, Add, GroupBox, x180 y75 w75 h45, 
; =========================================BOXES

WinGetPos, winX, winY, winWidth, winHeight, TDNP_Microfilm
winX+=%winWidth%
Gui, Show, x%winX% y%winY% h130 w265, TDNP_Microfilm
winactivate, TDNP_Microfilm

; =========================================SCRIPTS
; ===================
; TopMonth - TDNP Calendar Template
; moves from bottom row in April/August
; ----------- to top row in May/September
; Hotkey: Alt + '
!'::
  ; move to top row
	Send, {Up 38}

	; display month row
	Send, {Up}
	Send, {Down}

	; move right one month
	Send, {Right 8}

	; update the scoreboard
	hotkeys++
	keystrokes+=48
	topmonthscore+=48
	ControlSetText, Static4, TopMonth, TDNP_Microfilm
	ControlSetText, Static8, %hotkeys%, TDNP_Microfilm
	ControlSetText, Static9, %keystrokes%, TDNP_Microfilm
	ControlSetText, Static10, %topmonthscore%, TDNP_Microfilm
Return
; ===================

; ===================
; NextRow - TDNP Calendar Template
; a typewriter style carriage return
; Hotkey: Ctrl + '
^'::
	; move down two rows
	Send, {Down 2}

	; move left six cells
	Send, {Left 6}

	; update the scoreboard
	hotkeys++
	keystrokes+=8
	nextrowscore+=8
	ControlSetText, Static4, NextRow, TDNP_Microfilm
	ControlSetText, Static8, %hotkeys%, TDNP_Microfilm
	ControlSetText, Static9, %keystrokes%, TDNP_Microfilm
	ControlSetText, Static11, %nextrowscore%, TDNP_Microfilm
Return
; ===================

; ===================
; Highlight - Excel Spreadsheet
; changes the background color of a cell
; Hotkey: Ctrl + ,
^,::
	; activate the Home menu
	Send, {AltDown}h{AltUp}
	Sleep, 100

	; select background color
	Send, h
	Sleep, 100

	; select yellow
	Send, {Down 6}
	Sleep, 50
	Send, {Right 3}
	Sleep, 50
	Send, {Enter}

	; update the scoreboard
	hotkeys++
	keystrokes+=2
	highlightscore+=2
	ControlSetText, Static4, Highlight, TDNP_Microfilm
	ControlSetText, Static8, %hotkeys%, TDNP_Microfilm
	ControlSetText, Static9, %keystrokes%, TDNP_Microfilm
	ControlSetText, Static12, %highlightscore%, TDNP_Microfilm
Return
; ===================
; =========================================SCRIPTS

; ===================================MENU FUNCTIONS
; =================FILE
; open the scanned_for_tdnp folder
OpenTDNP:
	Run, %scannedpath%
Return

; open the Collation Sheets folder
CollationSheets:
	Run, %collationpath%
Return

Microfilm:
	Run, http://digitalprojects.library.unt.edu/projects/index.php/Microfilm_Scanning_Queue
Return

; reload application
Reload:
Reload

; close application
Exit:
ExitApp
; =================FILE

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

; =================TOOLS
; convert millimeters to inches
mm2in:
	; create input box to enter millimeters
	InputBox, input, MM to IN,,, 150, 100,,,,,enter millimeters
	if ErrorLevel
		Return
	else
	{
		; floating point number to 2 decimal places
		SetFormat, Float, 0.2
		
		; store millimeters input
		mm = %input%
		
		; make the calculation
		input*=0.0393701
		
		; print the result
		MsgBox %mm% millimeters`n`n%input% inches
	}
Return

; convert inches to millimeters
in2mm:
	; create input box to enter millimeters
	InputBox, input, IN to MM,,, 150, 100,,,,,enter inches
	if ErrorLevel
		Return
	else
	{
		; floating point number to 2 decimal places
		SetFormat, Float, 0.2

		; store inches input
		in = %input%

		; make the calculation
		input*=25.4
		
		; print the result
		MsgBox %in% inches`n`n%input% millimeters
	}
Return

; calculate the actual size of the original newspaper
; given the size on film and reduction ratio
ActualSize:
	; create input box to enter size of page on film in millimeters
	InputBox, inputAS, Actual Size, Enter size on film:,, 150, 125,,,,,millimeters
	if ErrorLevel
		Return
	else
	{
		; store the size on film
		mm = %inputAS%
		
		; create input box to enter the reduction ratio		
		InputBox, inputAS, Actual Size, Enter reduction ratio:,, 150, 125,,,,,integer
		if ErrorLevel
			Return
		else
		{
			; store the reduction ratio
			redrat = %inputAS%
			
			; floating point number to 2 decimal places			
			SetFormat, Float, 0.2
			
			; make the calculation
			inputAS*=%mm%
			inputAS*=0.0393701
			
			; print the result
			MsgBox Size on Film: `t%mm% millimeters`n`nReduction Ratio: `t%redrat% : 1`n`nActual Size: `t%inputAS% inches		  
		}
	}
Return

; calculate the reduction ratio of the filmed newspaper
; given the size on film and the size of the original newspaper
ReductionRatio:
	; create input box to enter size of page on film in millimeters
	InputBox, inputRR, Red. Ratio, Enter size on film:,, 150, 125,,,,,millimeters
	if ErrorLevel
		Return
	else
	{
		; store size of page on film in millimeters
		mm = %inputRR%
		
		; create input box to enter the reduction ratio
		InputBox, inputRR, Red. Ratio, Enter actual size:,, 150, 125,,,,,inches
		if ErrorLevel
			Return
		else
		{
			; store size of original page in inches
			in = %inputRR%

			; floating point number to 2 decimal places			
			SetFormat, Float, 0.2

			; make the calculation
			inputRR*=25.4
			inputRR/=%mm%

			; print the result
			MsgBox Size on Film: `t%mm% millimeters`n`nActual Size: `t%in% inches`n`nReduction Ratio: `t%inputRR% : 1
		}
	}
Return

; display newspaper publication frequency definitions
Frequency:
MsgBox,
(
Daily`t`t5-7 / week

Weekly`t`t1 / week
Semi-Weekly`t2 / week
Tri-Weekly`t3 / week

Bi-Weekly`t`t1 / two weeks
Monthly`t`t1 / month
)
Return
; =================TOOLS

; =================HELP
; open the TDNP_Microfilm wiki page
Documentation:
	Run, https://github.com/drewhop/AutoHotkey/wiki/TDNP_Microfilm
Return

; open US Newspaper Directory
Directory:
	Run, http://chroniclingamerica.loc.gov/search/titles/
Return

; display version info
About:
MsgBox TDNP_Microfilm.ahk`nVersion 1.0`nAndrew.Weidner@unt.edu
Return
; =================HELP
; ===================================MENU FUNCTIONS

GuiClose:
ExitApp
