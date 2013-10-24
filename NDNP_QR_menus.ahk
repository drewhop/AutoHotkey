; ********************************
; NDNP_QR_menus.ahk
; required file for NDNP_QR.ahk
; ********************************

; ======================REEL CHECK
ReelFolderCheckLoop:
	MsgBox, 0, Reel Folder, %reelfoldername% does not appear to be a REEL folder.`n`n`tPlease try again.

	reelfolderpath = _
	reelfoldername = _

	FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect the REEL folder:
	if ErrorLevel
	{
		reelfolderpath = _
		reelfoldername = _
		Exit
	}

	Gosub, GetReelInfo
Return
; ======================REEL CHECK

; ======================REEL INFO
GetReelInfo:
	; display variables for reel folder path
	StringGetPos, reelfolderpos, reelfolderpath, \, R2
	StringLeft, reelfolder1, reelfolderpath, reelfolderpos
	StringTrimLeft, reelfolder2, reelfolderpath, reelfolderpos

	; reel folder name
	StringGetPos, foldernamepos, reelfolderpath, \, R1
	foldernamepos++
	StringTrimLeft, reelfoldername, reelfolderpath, foldernamepos
Return
; ======================REEL INFO

; ======================FILE MENU
OpenReel:
	FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
	if ErrorLevel
	{
		reelfolderpath = _
		reelfoldername = _
		Return
	}
	
	; close the current reel folder
	SetTitleMatchMode 1
	IfWinExist, %reelfoldername%,, %reelfoldername%-
		WinClose, %reelfoldername%,,, %reelfoldername%-
		
	Gosub, GetReelInfo
		
	; loop checks to see if it is a reel folder
	Loop
	{
		StringLen, length, reelfoldername
		if (length != 11)
		{		
			Gosub, ReelFolderCheckLoop
		}
		else Break
	}
	
	; open the reel folder		
	Run, %reelfolderpath%
		
	; wait for the window to load
	SetTitleMatchMode 1
	WinWaitActive, %reelfoldername%,,, %reelfoldername%-
	Sleep, 100
		
	; select the first issue
	Send, {Down}
	Sleep, 100
	Send, {Up}
Return

; validate a batch with the command line
; modify CMDpath & DVVpath in VARIABLES for your system 
DVVvalidate:
	FileSelectFolder, batchpath, %batchdrive%, 2, `nSelect the batch to VALIDATE:
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
	FileSelectFolder, batchpath, %batchdrive%, 2, `nSelect the batch to VERIFY:
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
; =======FOLDER NAVIGATION
NavSkip:
	Gui, 12:+ToolWindow
	Gui, 12:Add, Text,, Number of folders:
	; navchoice value (1-7) pre-selected, assigns value (1,2,3,5,10,15,20) to navskip
	Gui, 12:Add, DropDownList, Choose%navchoice% R7 vnavskip, 1|2|3|5|10|15|20
	; run NavSkipGo if OK
	Gui, 12:Add, Button, w40 x10 y55 gNavSkipGo default, OK
	; run NavSkipCancel if Cancel
	Gui, 12:Add, Button, x65 y55 gNavSkipCancel, Cancel
	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	Gui, 12:Show, x%winX% y%winY%, Folder Navigation
Return

; OK button
NavSkipGo:
	Gui, 12:Submit
	
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
		
	Gui, 12:Destroy
Return

; Cancel button
NavSkipCancel:
	Gui, 12:Destroy
Return
; =======FOLDER NAVIGATION

; =======IMAGE WIDTH
ImageWidth:
	if (imagewidth == wide)
		imageselect = 1
	else if (imagewidth == medium)
		imageselect = 2
	else imageselect = 3
	
	Gui, 21:+ToolWindow
	Gui, 21:Add, Text,, Select an image width:
	Gui, 21:Add, DropDownList, AltSubmit Choose%imageselect% R3 vimagechoice, Wide|Medium|Narrow
	; run ScreenWidthGo if OK
	Gui, 21:Add, Button, w40 x10 y55 gImageWidthGo default, OK
	; run ScreenWidthCancel if Cancel
	Gui, 21:Add, Button, x65 y55 gImageWidthCancel, Cancel
	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	Gui, 21:Show, x%winX% y%winY%, Image Width
Return

; OK button
ImageWidthGo:
	Gui, 21:Submit
	
	if (imagechoice == 1)
		imagewidth = %wide%
	else if (imagechoice == 2)
		imagewidth = %medium%
	else imagewidth = %narrow%
		
	Gui, 21:Destroy
Return

; Cancel button
ImageWidthCancel:
	Gui, 21:Destroy
Return
; =======SCREEN WIDTH

; =======NOTE HOTKEYS
; Ctrl8 input
Note8:
InputBox, input, Ctrl 8, Enter the value for Ctrl + 8,, 250, 125,,,,,%Ctrl8%
	if ErrorLevel
		Return
	else
		Ctrl8 = %input%
Return

; Ctrl9 input
Note9:
InputBox, input, Ctrl 9, Enter the value for Ctrl + 9,, 250, 125,,,,,%Ctrl9%
	if ErrorLevel
		Return
	else
		Ctrl9 = %input%
Return

; Ctrl0 input
Note0:
InputBox, input, Ctrl 0, Enter the value for Ctrl + 0,, 250, 125,,,,,%Ctrl0%
	if ErrorLevel
		Return
	else
		Ctrl0 = %input%
Return

DisplayStandardNotes:
MsgBox,
(
STANDARD NOTES

Ctrl + 1 = %Ctrl1%

Ctrl + 2 = %Ctrl2%

Ctrl + 3 = %Ctrl3%

Ctrl + 4 = %Ctrl4%

Ctrl + 5 = %Ctrl5%

Ctrl + 6 = %Ctrl6%

Ctrl + 7 = %Ctrl7%
)
Return

DisplayUserNotes:
MsgBox,
(
USER-DEFINED NOTES

Ctrl + 8 = %Ctrl8%

Ctrl + 9 = %Ctrl9%

Ctrl + 0 = %Ctrl0%
)
Return
; =======NOTE HOTKEYS

; =======REEL FOLDER
EditReelFolder:
	FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
	if ErrorLevel
	{
		reelfolderpath = _
		reelfoldername = _
		Return
	}

	Gosub, GetReelInfo

	Loop ; valid reel folder check
	{
		StringLen, length, reelfoldername
		if (length != 11)
		{		
			Gosub, ReelFolderCheckLoop
		}
				
		else Break
	}
	
	Gosub, ResetReel
Return

DisplayReelFolder:
	if (reelfolderpath == "_")
	{
		MsgBox, 0, Current Reel, Reel Folder Name: %reelfoldername%`n`nReel Folder Path: %reelfolderpath%
	}
	else
	{
		MsgBox, 0, Current Reel, Reel Folder Name:`n`n`t%reelfoldername%`n`nReel Folder Path:`n`n`t%reelfolder1%`n`t%reelfolder2%
	}
Return
; =======REEL FOLDER

; =======DIRECTORY PATHS
EditBatchDrive:
	FileSelectFolder, batchdrive,, 0, `nSelect the drive where your batches are stored:
	if ErrorLevel
	{
		batchdrive = %batchdrivedefault%
		Return
	}
Return

EditNotepadFolder:
	FileSelectFolder, notepadpath, C:\, 0, `nSelect the folder where Notepad++ is installed:
	if ErrorLevel
	{
		notepadpath = %notepadpathdefault%
		Return
	}
Return

EditCMDFolder:
	FileSelectFolder, CMDpath, C:\, 0, `nSelect the folder where CMD.exe is installed:
	if ErrorLevel
	{
		CMDpath = %CMDpathdefault%
		Return
	}
Return

EditDVVFolder:
	FileSelectFolder, DVVpath, C:\, 0, `nSelect the folder where the DVV is installed:
	if ErrorLevel
	{
		DVVpath = %DVVpathdefault%
		Return
	}
Return

DisplayCurrentPaths:
	MsgBox, 0, Current Paths, Batch Drive:`n`n     %batchdrive%`n`nCMD Folder:`n`n     %CMDpath%`n`nDVV Folder:`n`n     %DVVpath%`n`nNotepad++ Folder:`n`n     %notepadpath%
Return
; =======DIRECTORY PATHS
; ======================EDIT MENU

; ======================SEARCH MENU
; =======DIRECTORY SEARCH
DirectorySearch:
	Gui, 10:+ToolWindow
	; search box input field
	Gui, 10:Add, Text,, Enter a search term:
	Gui, 10:Add, Edit, w130 vdirectorysearchstring, %directorysearchstring%
	; state dropdown menu
	Gui, 10:Add, Text,, Choose a state (optional):
	Gui, 10:Add, DropDownList, r10 vstate, Alabama|Arizona|California|Colorado|District of Columbia|Florida|Georgia|Hawaii|Idaho|Illinois|Indiana|Kansas|Kentucky|Louisiana|Minnesota|Mississippi|Missouri|Montana|Nebraska|New Mexico|New York|North Dakota|Ohio|Oklahoma|Oregon|Pennsylvania|South Carolina|Tennessee|Texas|Utah|Vermont|Virginia|Washington
	; run DirectoryGo on Search button
	Gui, 10:Add, Button, x10 y100 gDirectoryGo default, Search
	; run DirSearchCancel on Cancel button
	Gui, 10:Add, Button, x65 y100 gDirSearchCancel, Cancel

	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	winY+=%winHeight%	
	Gui, 10:Show, x%winX% y%winY%, US Directory Search
Return

; Search button
DirectoryGo:
	Gui, 10:Submit
	Gui, 10:Destroy
	Run, http://chroniclingamerica.loc.gov/search/titles/results/?state=%state%&county=&city=&year1=1690&year2=2013&terms=%directorysearchstring%&frequency=&language=&ethnicity=&labor=&material_type=&lccn=&rows=20
Return

; Cancel button
DirSearchCancel:
	Gui, 10:Destroy
Return
; =======DIRECTORY SEARCH

; =======DIRECTORY LCCN
DirectoryLCCN:
	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	winY+=%winHeight%
	InputBox, LCCNstring, US Directory: LCCN,,, 200, 100, %winX%, %winY%,,,%LCCNstring%
	if ErrorLevel
		Return
	else
		Run, http://chroniclingamerica.loc.gov/lccn/%LCCNstring%
Return
; =======DIRECTORY LCCN

; =======CHRONAM SEARCH
ChronAmSearch:
	Gui, 11:+ToolWindow
	; search box input field
	Gui, 11:Add, Text,, Enter a search term:
	Gui, 11:Add, Edit, w130 vchronamsearchstring, %chronamsearchstring%
	; state dropdown menu
	Gui, 11:Add, Text,, Choose a state (optional):
	Gui, 11:Add, DropDownList, r10 vstate, Alabama|Arizona|California|Colorado|District of Columbia|Florida|Georgia|Hawaii|Idaho|Illinois|Indiana|Kansas|Kentucky|Louisiana|Minnesota|Mississippi|Missouri|Montana|Nebraska|New Mexico|New York|North Dakota|Ohio|Oklahoma|Oregon|Pennsylvania|South Carolina|Tennessee|Texas|Utah|Vermont|Virginia|Washington
	; date input fields
	Gui, 11:Add, Text,, Start:%A_Tab%End:
	Gui, 11:Add, Edit, x10 y115 vdate1, %date1%
	Gui, 11:Add, Edit, x50 y115 vdate2, %date2%
	; run ChronAmGo on Search button
	Gui, 11:Add, Button, x10 y145 gChronAmGo default, Search
	; run CASearchCancel on Cancel button
	Gui, 11:Add, Button, x65 y145 gCASearchCancel, Cancel

	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	winY+=%winHeight%	
	Gui, 11:Show, x%winX% y%winY%, ChronAm Search
Return

; Search button
ChronAmGo:
	Gui, 11:Submit
	Gui, 11:Destroy
	Run, http://chroniclingamerica.loc.gov/search/pages/results/?state=%state%&date1=%date1%&date2=%date2%&proxtext=%chronamsearchstring%&x=0&y=0&dateFilterType=yearRange&rows=20&searchType=basic
Return

; Cancel button
CASearchCancel:
	Gui, 11:Destroy
Return
; =======CHRONAM SEARCH

; =======CHRONAM BROWSE
ChronAmBrowse:
	Gui, 13:+ToolWindow
	; state dropdown menu
	Gui, 13:Add, Text,, Choose a state (optional):
	Gui, 13:Add, DropDownList, r10 vstate, Alabama|Arizona|California|Colorado|District of Columbia|Florida|Georgia|Hawaii|Idaho|Illinois|Indiana|Kansas|Kentucky|Louisiana|Minnesota|Mississippi|Missouri|Montana|Nebraska|New Mexico|New York|North Dakota|Ohio|Oklahoma|Oregon|Pennsylvania|South Carolina|Tennessee|Texas|Utah|Vermont|Virginia|Washington
	; run ChronAmBrowseGo on Browse button
	Gui, 13:Add, Button, x10 y55 gChronAmBrowseGo default, Browse
	; run CABrowseCancel on Cancel button
	Gui, 13:Add, Button, x65 y55 gCABrowseCancel, Cancel

	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	winY+=%winHeight%	
	Gui, 13:Show, x%winX% y%winY%, ChronAm Browse
Return

; Search button
ChronAmBrowseGo:
	Gui, 13:Submit
	Gui, 13:Destroy
	Run, http://chroniclingamerica.loc.gov/newspapers/%state%
Return

; Cancel button
CABrowseCancel:
	Gui, 13:Destroy
Return
; =======CHRONAM BROWSE

; =======CHRONAM DATA
ChronAmData:
	statecode =
	databatch =

	Gui, 14:+ToolWindow
	; state code dropdown menu
	Gui, 14:Add, Text,, Choose a code (optional):
	Gui, 14:Add, DropDownList, r10 vstatecode, az|curiv|dlc|fu|hihouml|in|iune|khi|kyu|lu|mnhi|mohi|mthi|nbu|ndhi|nmu|nn|ohi|okhi|oru|pst|scu|tu|txdn|uuml|vi|vtu|wa
	; search box input field
	Gui, 14:Add, Text,, Enter a batch name (optional):
	Gui, 14:Add, Edit, w140 vdatabatch,
	; run ChronAmDataGo on Browse button
	Gui, 14:Add, Button, w40 x10 y100 gChronAmDataGo default, GO
	; run CADataCancel on Cancel button
	Gui, 14:Add, Button, x60 y100 gCADataCancel, Cancel

	SetTitleMatchMode 1
	WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
	winY+=%winHeight%	
	Gui, 14:Show, x%winX% y%winY%, ChronAm Data
Return

; GO button
ChronAmDataGo:
	Gui, 14:Submit
	Gui, 14:Destroy
	
	Run, http://chroniclingamerica.loc.gov/data/batches
	
	; if statecode or databatch are not empty
	if ((statecode != "") || (databatch != ""))
	{
		SetTitleMatchMode 1
		WinWaitActive, Index of /data/batches, , 10
		Sleep, 200
		
		; activate the Find dialog
		Send, {CtrlDown}f{CtrlUp}
		Sleep, 100
		
		if (statecode != "")
		{
			; input the state code
			Send, _%statecode%
		}
		
		if (databatch != "")
		{
			; input the batch name
			Send, _%databatch%
		}
	}
Return

; Cancel button
CADataCancel:
	Gui, 14:Destroy
Return
; =======CHRONAM DATA
; ======================SEARCH MENU
