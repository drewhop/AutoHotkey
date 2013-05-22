; ********************************
; NDNP_QR_menus.ahk
; required file for NDNP_QR.ahk
; ********************************

; ======================FILE MENU
; open a reel folder
OpenReel:
	; create dialog to select the reel folder
	FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
	if ErrorLevel
	{
		; reset variables on Cancel
		reelfolderpath = _
		reelfoldername = _
		Return
	}
	
	; close the current reel folder
	SetTitleMatchMode 1
	IfWinExist, %reelfoldername%
		WinClose, %reelfoldername%
		
	; create display variables for reel folder path
	StringGetPos, reelfolderpos, reelfolderpath, \, R2
	StringLeft, reelfolder1, reelfolderpath, reelfolderpos
	StringTrimLeft, reelfolder2, reelfolderpath, reelfolderpos

	; assign the new reel folder name
	StringGetPos, foldernamepos, reelfolderpath, \, R1
	foldernamepos++
	StringTrimLeft, reelfoldername, reelfolderpath, foldernamepos
	
/*
	; check to see if it is a reel folder
	StringLen, length, reelfoldername
	if (length != 11)
	{		
		; print error message
		MsgBox, 0, Open Reel Folder, %reelfoldername% does not appear to be a REEL folder.`n`n`tPlease try again.

		; reset the variables
		reelfolderpath = _
		reelfoldername = _
	}
*/
	; loop checks to see if it is a reel folder
	Loop
	{
		StringLen, length, reelfoldername
		if (length != 11)
		{		
			; print error message
			MsgBox, 0, Reel Folder, %reelfoldername% does not appear to be a REEL folder.`n`n`tPlease try again.

			; reset the variables
			reelfolderpath = _
			reelfoldername = _

			; create dialog to select the reel folder
			FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect the REEL folder:
			if ErrorLevel
			{
				; reset the variables on Cancel
				reelfolderpath = _
				reelfoldername = _
				Exit
			}

			; create display variables for reel folder path
			StringGetPos, reelfolderpos, reelfolderpath, \, R2
			StringLeft, reelfolder1, reelfolderpath, reelfolderpos
			StringTrimLeft, reelfolder2, reelfolderpath, reelfolderpos

			; extract reel folder name from path
			StringGetPos, foldernamepos, reelfolderpath, \, R1
			foldernamepos++
			StringTrimLeft, reelfoldername, reelfolderpath, foldernamepos
		}
				
		else Break
	}
	
;	else
;	{
		; open the reel folder		
		Run, %reelfolderpath%
		
		; wait for the window to load
		SetTitleMatchMode 1
		WinWaitActive, %reelfoldername%
		Sleep, 100
		
		; select the first issue
		Send, {Down}
		Sleep, 100
		Send, {Up}
;	}
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
; create the GUI
NavSkip:
	Gui, 12:+ToolWindow
	Gui, 12:Add, Text,, Number of folders:
	; navchoice value (1-7) pre-selected, assigns value (1,2,3,5,10,15,20) to navskip
	Gui, 12:Add, DropDownList, Choose%navchoice% R7 vnavskip, 1|2|3|5|10|15|20
	; run NavSkipGo if OK
	Gui, 12:Add, Button, w40 x10 y55 gNavSkipGo default, OK
	; run NavSkipCancel if Cancel
	Gui, 12:Add, Button, x65 y55 gNavSkipCancel, Cancel
	Gui, 12:Show,, Folder Navigation
Return

; OK button function
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

; Cancel button function
NavSkipCancel:
	; close the Folder Navigation GUI
	Gui, 12:Destroy
Return
; =======FOLDER NAVIGATION

; =======REEL FOLDER
EditReelFolder:
	; create dialog to select the reel folder
	FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect a REEL folder:
	if ErrorLevel
	{
		; reset variables on Cancel
		reelfolderpath = _
		reelfoldername = _
		Return
	}

	; create display variables for reel folder path
	StringGetPos, reelfolderpos, reelfolderpath, \, R2
	StringLeft, reelfolder1, reelfolderpath, reelfolderpos
	StringTrimLeft, reelfolder2, reelfolderpath, reelfolderpos

	; extract reel folder name from path
	StringGetPos, foldernamepos, reelfolderpath, \, R1
	foldernamepos++
	StringTrimLeft, reelfoldername, reelfolderpath, foldernamepos

	; loop checks to see if it is a reel folder
	Loop
	{
		StringLen, length, reelfoldername
		if (length != 11)
		{		
			; print error message
			MsgBox, 0, Reel Folder, %reelfoldername% does not appear to be a REEL folder.`n`n`tPlease try again.

			; reset the variables
			reelfolderpath = _
			reelfoldername = _

			; create dialog to select the reel folder
			FileSelectFolder, reelfolderpath, %batchdrive%, 2, `nSelect the REEL folder:
			if ErrorLevel
			{
				; reset the variables on Cancel
				reelfolderpath = _
				reelfoldername = _
				Exit
			}

			; create display variables for reel folder path
			StringGetPos, reelfolderpos, reelfolderpath, \, R2
			StringLeft, reelfolder1, reelfolderpath, reelfolderpos
			StringTrimLeft, reelfolder2, reelfolderpath, reelfolderpos

			; extract reel folder name from path
			StringGetPos, foldernamepos, reelfolderpath, \, R1
			foldernamepos++
			StringTrimLeft, reelfoldername, reelfolderpath, foldernamepos
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
		; reset to default on Cancel
		batchdrive = %batchdrivedefault%
		Return
	}
Return

EditNotepadFolder:
	FileSelectFolder, notepadpath, C:\, 0, `nSelect the folder where Notepad++ is installed:
	if ErrorLevel
	{
		; reset to default on Cancel
		notepadpath = %notepadpathdefault%
		Return
	}
Return

EditCMDFolder:
	FileSelectFolder, CMDpath, C:\, 0, `nSelect the folder where CMD.exe is installed:
	if ErrorLevel
	{
		; reset to default on Cancel
		CMDpath = %CMDpathdefault%
		Return
	}
Return

EditDVVFolder:
	FileSelectFolder, DVVpath, C:\, 0, `nSelect the folder where the DVV is installed:
	if ErrorLevel
	{
		; reset to default on Cancel
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
; create the GUI
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
	Gui, 10:Show,, US Directory Search
Return

; Search button function
DirectoryGo:
	; assign the search string
	Gui, 10:Submit
	
	; close the search GUI
	Gui, 10:Destroy
	
	; load the results in the default Web browser
	Run, http://chroniclingamerica.loc.gov/search/titles/results/?state=%state%&county=&city=&year1=1690&year2=2013&terms=%directorysearchstring%&frequency=&language=&ethnicity=&labor=&material_type=&lccn=&rows=20
Return

; Cancel button function
DirSearchCancel:
	Gui, 10:Destroy
Return
; =======DIRECTORY SEARCH

; =======DIRECTORY LCCN
DirectoryLCCN:
	InputBox, LCCNstring, US Directory: LCCN,,, 200, 100,,,,,%LCCNstring%
	if ErrorLevel
		Return
	else
		Run, http://chroniclingamerica.loc.gov/search/titles/results/?state=&county=&city=&year1=1690&year2=2013&terms=&frequency=&language=&ethnicity=&labor=&material_type=&lccn=%LCCNstring%&rows=20
Return
; =======DIRECTORY LCCN

; =======CHRONAM SEARCH
; create the GUI
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
	Gui, 11:Show,, ChronAm Search
Return

; Search button function
ChronAmGo:
	; assign the search string
	Gui, 11:Submit
	
	; close the search GUI
	Gui, 11:Destroy
	
	; load the results in the default Web browser
	Run, http://chroniclingamerica.loc.gov/search/pages/results/?state=%state%&date1=%date1%&date2=%date2%&proxtext=%chronamsearchstring%&x=0&y=0&dateFilterType=yearRange&rows=20&searchType=basic
Return

; Cancel button function
CASearchCancel:
	Gui, 11:Destroy
Return
; =======CHRONAM SEARCH

; =======CHRONAM BROWSE
; create the GUI
ChronAmBrowse:
	Gui, 13:+ToolWindow
	; state dropdown menu
	Gui, 13:Add, Text,, Choose a state (optional):
	Gui, 13:Add, DropDownList, r10 vstate, Alabama|Arizona|California|Colorado|District of Columbia|Florida|Georgia|Hawaii|Idaho|Illinois|Indiana|Kansas|Kentucky|Louisiana|Minnesota|Mississippi|Missouri|Montana|Nebraska|New Mexico|New York|North Dakota|Ohio|Oklahoma|Oregon|Pennsylvania|South Carolina|Tennessee|Texas|Utah|Vermont|Virginia|Washington
	; run ChronAmBrowseGo on Browse button
	Gui, 13:Add, Button, x10 y55 gChronAmBrowseGo default, Browse
	; run CABrowseCancel on Cancel button
	Gui, 13:Add, Button, x65 y55 gCABrowseCancel, Cancel
	Gui, 13:Show,, ChronAm Browse
Return

; Search button function
ChronAmBrowseGo:
	; assign the state
	Gui, 13:Submit
	
	; close the GUI
	Gui, 13:Destroy
	
	; load the results in the default Web browser
	Run, http://chroniclingamerica.loc.gov/newspapers/%state%
Return

; Cancel button function
CABrowseCancel:
	Gui, 13:Destroy
Return
; =======CHRONAM BROWSE

; =======CHRONAM DATA
; create the GUI
ChronAmData:
	; initialize the search variables
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
	Gui, 14:Show,, ChronAm Data
Return

; GO button function
ChronAmDataGo:
	; assign the variables (if any)
	Gui, 14:Submit
	
	; close the GUI
	Gui, 14:Destroy
	
	; load data/batches in the default Web browser
	Run, http://chroniclingamerica.loc.gov/data/batches
	
	; if statecode or databatch are not empty
	if ((statecode != "") || (databatch != ""))
	{
		; wait for page to load
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

; Cancel button function
CADataCancel:
	Gui, 14:Destroy
Return
; =======CHRONAM DATA
; ======================SEARCH MENU
