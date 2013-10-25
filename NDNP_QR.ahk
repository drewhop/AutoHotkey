/*
 * NDNP_QR.ahk
 *
 * DESCRIPTION: application for NDNP data Quality Review
 *
 * REQUIRED FILES ****************************************
 * All files must reside in the same directory
 * when running the NDNP_QR.ahk script or at compile time
 *
 *    NDNP_QR.ahk (compile for NDNP_QR.exe)
 *    NDNP_QR_dvvloops.ahk
 *    NDNP_QR_hotkeys.ahk
 *    NDNP_QR_menus.ahk
 *    NDNP_QR_metadata.ahk
 *    NDNP_QR_navigation.ahk
 *    NDNP_QR_reelloops.ahk
 *    NDNP_QR_reports.ahk
 *
 * *******************************************************
 *
 * REQUIRED SOFTWARE ******************************
 * First Impression Viewer (TIFF viewer)
 * http://www.softpedia.com/get/Multimedia/Graphic/Graphic-Viewers/First-Impression.shtml
 *
 * DVV - Digital Viewer and Validator
 * (Validate & Verify Batch, Pages & Thumbs Loops)
 * http://www.loc.gov/ndnp/tools/
 *
 * Notepad++ (EditXML hotkey)
 * http://notepad-plus-plus.org/
 * ************************************************
 *
 * GUI NOTE ***************************************************
 * The GUI dimensions may need to be adjusted for your system.
 * See the AutoHotkey GUI documentation for more information:
 * http://www.autohotkey.com/docs/commands/Gui.htm
 * ************************************************************
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

; EDIT THIS VARIABLE
; for the tools delay time in seconds
; (Reel Report & Metadata Viewer)
; must be an integer in the range 3-9
delaydefault = 6

; EDIT THIS VARIABLE
; for the language code report default
languagecodedefault = spa

; default image width values
wide = 20
medium = 25
narrow = 30
; EDIT THIS VARIABLE
; for the default reel loop image display width (%wide%, %medium%, %narrow%)
imagewidth = %wide%

; documentation
docURL = https://github.com/drewhop/AutoHotkey/wiki/NDNP_QR

; STANDARD notes
Ctrl1 = Date printed on front page: .
Ctrl2 = Volume number printed on front page: .
Ctrl3 = Issue number printed on front page: .
Ctrl4 = Odd number of pages.
Ctrl5 = Front page lists  pages.
Ctrl6 = Add section label: .
Ctrl7 = Add edition label: .

; USER-DEFINED notes
Ctrl8 =
Ctrl9 =
Ctrl0 =
value8 =
value9 =
value0 =

; system
notepadpath = %notepadpathdefault%
CMDpath = %CMDpathdefault%
DVVpath = %DVVpathdefault%

; navigation
batchdrive = %batchdrivedefault%
reelfoldername = _
reelfolderpath = _
navskip = 1
navchoice = 1

; tools
lccnurl = http://chroniclingamerica.loc.gov/lccn/
lccndivider := "============================================================================="
reeldivider := "-----------------------------------------------------------------------------"
reportheader = Identifier`t`t`tVolume`tIssue`tDate (Questionable)`tPages (Missing)
delay = %delaydefault%
delaychoice := delay - 2
dvvdelay =
dvvdelayms =
dvvdelaychoice = 6
cancelbutton = 0
ispaused = 0
metaloopname =
dvvloopname =
count = 0
pagecount = 0
thumbscount = 0
loopcount = 0
notecount = 0
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
timerX =
timerY =
issuefolderpath =
issuefoldername =
languagecode = %languagecodedefault%
ocrterm =

; search
LCCNstring = Enter an LCCN
directorysearchstring = Do not use quotes.
chronamsearchstring = Do not use quotes.
date1 = 1836
date2 = 1922

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

; error handling
errortitle =
errorsubtitle =
tiffopenflag = 0
openflag = 0
openplusflag = 0
nextflag = 0
nextplusflag = 0
previousflag = 0
previousplusflag = 0
gotoflag = 0
gotoplusflag = 0
metadataflag = 0
viewissuexmlflag = 0
editissuexmlflag = 0
; =========================================VARIABLES

; BEGIN GUI FORMATTING *******************************************
; this section formats the NDNP_QR(1) and Metadata(2) windows
Gui, 1:Color, d0d0d0, 912206
Gui, 1:Show, h0 w0, NDNP_QR

; =========================================MENUS
; File
Menu, FileMenu, Add, &Open Reel Folder, OpenReel
Menu, FileMenu, Add
Menu, FileMenu, Add, V&alidate Batch, DVVvalidate
Menu, FileMenu, Add, V&erify Batch, DVVverify
Menu, FileMenu, Add
Menu, FileMenu, Add, &Reload, Reload
Menu, FileMenu, Add, E&xit, Exit

; Edit
Menu, NoteMenu, Add, Ctrl + &8, Note8
Menu, NoteMenu, Add, Ctrl + &9, Note9
Menu, NoteMenu, Add, Ctrl + &0, Note0
Menu, NoteMenu, Add
Menu, NoteMenu, Add, Display &Standard Notes, DisplayStandardNotes
Menu, NoteMenu, Add, Display &User-Defined Notes, DisplayUserNotes
Menu, EditMenu, Add, &Note Hotkeys, :NoteMenu
Menu, EditMenu, Add
Menu, EditMenu, Add, &Image Display Width, ImageWidth
Menu, EditMenu, Add
Menu, EditMenu, Add, &Folder Navigation, NavSkip
Menu, EditMenu, Add
Menu, ReelMenu, Add, &Set Path, EditReelFolder
Menu, ReelMenu, Add, &Display Current Reel, DisplayReelFolder
Menu, EditMenu, Add, &Reel Folder, :ReelMenu
Menu, EditMenu, Add
Menu, DirMenu, Add, &Batch Drive, EditBatchDrive
Menu, DirMenu, Add, &CMD Folder, EditCMDFolder
Menu, DirMenu, Add, &DVV Folder, EditDVVFolder
Menu, DirMenu, Add, &Notepad++ Folder, EditNotepadFolder
Menu, DirMenu, Add
Menu, DirMenu, Add, Display Current &Paths, DisplayCurrentPaths
Menu, EditMenu, Add, Directory &Paths, :DirMenu

; Tools
Menu, ToolsMenu, Add, &Batch Report, BatchReport
Menu, ToolsMenu, Add
Menu, ToolsMenu, Add, &Reel Report, ReelReport
Menu, ToolsMenu, Add, &Metadata Viewer, MetaViewer
Menu, ToolsMenu, Add
Menu, DVVMenu, Add, DVV &Pages, DVVpages
Menu, DVVMenu, Add, DVV &Thumbs, DVVthumbs
Menu, DVVMenu, Add
Menu, DVVMenu, Add, Set &Delay, DVVDelay
Menu, ToolsMenu, Add, DVV &Loops, :DVVMenu
Menu, ToolsMenu, Add
Menu, OCRMenu, Add, &Language Codes, LanguageCodeReport
Menu, OCRMenu, Add, OCR &Search, OCRSearch
Menu, ToolsMenu, Add, &OCR Reports, :OCRMenu

; Search
Menu, SearchMenu, Add, &US Directory Search, DirectorySearch
Menu, SearchMenu, Add, US Directory &LCCN, DirectoryLCCN
Menu, SearchMenu, Add
Menu, SearchMenu, Add, &ChronAm Search, ChronAmSearch
Menu, SearchMenu, Add, ChronAm &Browse, ChronAmBrowse
Menu, SearchMenu, Add, ChronAm &Data, ChronAmData

; Help
Menu, HelpMenu, Add, Online &Documentation, Documentation
Menu, HelpMenu, Add, &NDNP Awardee Wiki, NDNPwiki
Menu, HelpMenu, Add, &About NDNP_QR, About

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
; set main window font and size
; Gui, 1:Font, s10, Arial

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

; Metadata window font and size
Gui, 2:Font, s8, Arial
Gui, 2:Add, Text, x26 y20  w100 h20, Date:
Gui, 2:Add, Text, x13 y45  w100 h20, Volume:
Gui, 2:Add, Text, x22 y70  w100 h20, Issue:
Gui, 2:Add, Text, x17 y95  w100 h20, ? Date:
Gui, 2:Add, Text, x18 y120 w100 h20, Pages:
; =========================================LABELS

; =========================================DATA
; set larger main window data font
Gui, 1:Font, s15, Arial

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

; Metadata window data font
Gui, 2:Font, cRed s12 bold, Arial

; Metadata: Static 6-10
Gui, 2:Add, Text, x65 y20  w100 h20,
Gui, 2:Add, Text, x65 y45  w130 h20,
Gui, 2:Add, Text, x65 y70  w130 h20,
Gui, 2:Add, Text, x65 y95  w100 h20,
Gui, 2:Add, Text, x65 y120 w100 h20,
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
; =========================================BOXES

; main GUI
WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
winX+=%winWidth%
Gui, 1:Show, x%winX% y%winY% h190 w345, NDNP_QR
WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR

; Metadata GUI
; Gui, 2:+AlwaysOnTop
; Gui, 2:+ToolWindow
; winX+=%winWidth%
; Gui, 2:Show, x%winX% y%winY% h160 w200, Metadata
; END GUI FORMATTING *********************************************

WinActivate, NDNP_QR

; pause key pauses/restarts any active function
Pause::Pause

; =========================================HOTKEYS
#Include NDNP_QR_hotkeys.ahk
;  DeactiveWindows: deactivates an active NDNP_QR window
;  OPEN: opens first TIFF for selected issue folder
;  OPEN+: also displays issue metadata
;  NEXT: opens first TIFF for next issue folder
;  NEXT+: also displays issue metadata
;  PREVIOUS: opens first TIFF for previous issue folder
;  PREVIOUS+: also displays issue metadata
;  GOTO: opens first TIFF for specific issue folder
;  GOTO+: also displays issue metadata
;  METADATA: displays metadata for selected issue folder
;  REDRAW METADATA WINDOW: restores/redraws Metadata window
;  METADATA WINDOWS: loads separate windows for issue metadata
;  ZOOM IN: First Impression masthead view
;  ZOOM OUT: First Impression full page view
;  VIEW ISSUE XML: opens issue.xml file in default application
;  EDIT ISSUE XML: opens issue.xml in Notepad++
; =========================================HOTKEYS

; =========================================SUBROUTINES
#Include NDNP_QR_navigation.ahk
;  ReelFolderCheck: tests for stored "reelfolderpath" variable
;  ResetReel: opens or activates the current reel folder
;  CloseFirstImpressionWindows: loop closes all FI windows
;  OpenFirstTIFF: opens first .TIF for selected folder
;  GoToIssue: input dialog for GOTO hotkeys
;  IssueToReel: moves up one directory from any issue folder

#Include NDNP_QR_metadata.ahk
;  IssueFolderPath: assigns the "issuefolderpath" variable
;  ExtractMeta: parses issue.xml file & displays metadata
;  RedrawMetaWindow: restores/redraws Metadata window
;  CreateMetaWindows: loads separate metadata windows
; =========================================SUBROUTINES

; =========================================MENU FUNCTIONS
; ======================TOOLS MENU
#Include NDNP_QR_reports.ahk
;  BatchReport: creates a .TXT report of all issues in a batch
;  BatchParseOCR: helper subroutine for parsing batch.xml in OCR reports
;  LanguageCodeReport: creates a .TXT report of language codes in a reel's OCR
;  OCRSearch: creates a .TXT report of search term hits in a reel's OCR

#Include NDNP_QR_reelloops.ahk
;  ReelReport: creates a .TXT report of all issues in a reel
;              while displaying images and metadata
;  MetaViewer: displays images and metadata for issues in a reel
;              provides option to record notes
;  GetReportInfo: variable assignment subroutine
;  ReelLoopFunction: primary function for ReelReport & MetaViewer
;  DELAY TIMER: timer window functions

#Include NDNP_QR_dvvloops.ahk
;  DVVpages & DVVthumbs: viewing loops for the DVV
;  DVVnotes: notes function for the viewing loops
;  DVV COUNTER BUTTONS: counter button functions
;  DVV DELAY DIALOG: delay dialog functions
; ======================TOOLS MENU

; ******************************************************
#Include NDNP_QR_menus.ahk
; FILE MENU============================
;  OpenReel: opens reel folder & assigns variables
;  DVVvalidate: validates a batch with the command line
;  DVVverify: verifies a batch with the command line
;  Reload: reloads the application
;  Exit: exits the application
; EDIT MENU============================
;  NavSkip: dialog for Folder Navigation
;  NavSkipGo & NavSkipCancel: button functions
;  EditReelFolder: assigns the "reelfolderpath" variable
;  DisplayReelFolder: displays current reel folder variables
;  EditBatchDrive: assigns the "batchdrive" variable
;  EditNotepadFolder: assigns the "notepadpath" variable
;  EditCMDFolder: assigns the "CMDpath" variable
;  EditDVVFolder: assigns the "DVVpath" variable
; SEARCH MENU============================
;  DirectorySearch: dialog for US Directory Search
;  DirectoryGo & DirSearchCancel: button functions
;  DirectoryLCCN: dialog for US Directory: LCCN
;  ChronAmSearch: dialog for ChronAm Search
;  ChronAmGo & CASearchCancel: button functions
;  ChronAmBrowse: dialog for ChronAm Browse
;  ChronAmBrowseGo & CABrowseCancel: button functions
;  ChronAmData: dialog for ChronAm Data
;  ChronAmDataGo & CADataCancel: button functions
; ******************************************************

; ======================HELP MENU
Documentation:
  Run, %docURL%
Return

NDNPwiki:
  Run, http://www.loc.gov/extranet/wiki/library_services/ndnp/index.php/AwardeePage
Return

About:
	MsgBox, 0, About NDNP_QR, Quality Review utility for NDNP data.`n_____________________________`nVersion 2.1   (October 2013)`nAndrew.Weidner@unt.edu
Return
; ======================HELP MENU
; =========================================MENU FUNCTIONS

; exit application on main window close
GuiClose:
ExitApp
