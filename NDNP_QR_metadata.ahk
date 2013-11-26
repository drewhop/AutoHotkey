; ********************************
; NDNP_QR_metadata.ahk
; required file for NDNP_QR.ahk
; ********************************

; ======================ISSUE FOLDER PATH
IssueFolderPath:
	temp = %clipboard%
	clipboard =

	SetTitleMatchMode 1
	WinActivate, %reelfoldername%, , , report,
	WinWaitActive, %reelfoldername%, , 3, ,
	if ErrorLevel
	{
		errortitle =
		errorsubtitle =
		
		; error handling
		if (openflag == 1)
		{
			errortitle = Error: Open
			openflag = 0
		}
		else if (openplusflag == 1)
		{
			errortitle = Error: Open+
			openplusflag = 0
		}
		else if (nextflag == 1)
		{
			errortitle = Error: Next
			nextflag = 0
		}
		else if (nextplusflag == 1)
		{
			errortitle = Error: Next+
			nextplusflag = 0
		}
		else if (previousflag == 1)
		{
			errortitle = Error: Previous
			previousflag = 0
		}
		else if (previousplusflag == 1)
		{
			errortitle = Error: Previous+
			previousplusflag = 0
		}
		else if (gotoflag == 1)
		{
			errortitle = Error: GoTo
			gotoflag = 0
		}
		else if (gotoplusflag == 1)
		{
			errortitle = Error: GoTo+
			gotoplusflag = 0
		}
		else if (metadataflag == 1)
		{
			errortitle = Error: Display Metadata
			metadataflag = 0
		}
		else if (viewissuexmlflag == 1)
		{
			errortitle = Error: ViewIssueXML
			viewissuexmlflag = 0
		}
		else if (editissuexmlflag == 1)
		{
			errortitle = Error: EditIssueXML
			editissuexmlflag = 0
		}
		else errortitle = Error
		
		if (tiffopenflag == 1)
		{
			errorsubtitle = OpenFirstTIFF`nNDNP_QR_navigation.ahk
			tiffopenflag = 0
		}
		
		; error message after 3 seconds
		MsgBox, 0, %errortitle%, %errorsubtitle%`n===IssueFolderPath===`nNDNP_QR_metadata.ahk`n`nCannot find folder %reelfoldername%`n`nOptions:`n`t"File > Open Reel Folder"`n`t"Edit > Reel Folder > Set Path"		
		Exit
	}
	Sleep, 200
	
	; copy name of selected folder
	Send, {F2}
	Sleep, 200
	Send, {CtrlDown}c{CtrlUp}
	Sleep, 200
	issuefoldername = %clipboard%
	Send, {Enter}
	Sleep, 100

	; case for opened issue folder
	SetTitleMatchMode RegEx
	IfWinActive, ^1[0-9]{9}$, , 5
	{
		Gosub, IssueToReel
	}	
	
	; exit script if issue name copied to clipboard
	if RegExMatch(issuefoldername, "\d\d\d\d\d\d\d\d\d\d")
	{
		issuefolderpath = %reelfolderpath%\%issuefoldername%
		clipboard = %temp%		
		Return
	}
			
	; loop checks twice for correctly copied issue folder name
	; aborts script if unsuccessful
	Loop 2
	{
		Sleep, 500

		SetTitleMatchMode 1
		WinActivate, %reelfoldername%, , , ,
		WinWaitActive, %reelfoldername%, , , ,
		Sleep, 100

		; copy name of selected folder
		Send, {F2}
		Sleep, 100
		Send, {CtrlDown}c{CtrlUp}
		Sleep, 100
		Send, {Enter}
		ClipWait
		issuefoldername = %clipboard%

		; return if issue name copied to clipboard
		if RegExMatch(issuefoldername, "\d\d\d\d\d\d\d\d\d\d")
		{
			issuefolderpath = %reelfolderpath%\%issuefoldername%
			clipboard = %temp%		
			Return
		}
	}

	; abort script if unable to copy path after 2 attempts
	; reset the variables
	errortitle =
	errorsubtitle =
		
	; error handling
	if (openflag == 1)
	{
		errortitle = Error: Open
		openflag = 0
	}
	else if (global openplusflag == 1)
	{
		errortitle = Error: Open+
		openplusflag = 0
	}
	else if (nextflag == 1)
	{
		errortitle = Error: Next
		nextflag = 0
	}
	else if (nextplusflag == 1)
	{
		errortitle = Error: Next+
		nextplusflag = 0
	}
	else if (previousflag == 1)
	{
		errortitle = Error: Previous
		previousflag = 0
	}
	else if (previousplusflag == 1)
	{
		errortitle = Error: Previous+
		previousplusflag = 0
	}
	else if (gotoflag == 1)
	{
		errortitle = Error: GoTo
		gotoflag = 0
	}
	else if (gotoplusflag == 1)
	{
		errortitle = Error: GoTo+
		gotoplusflag = 0
	}
	else if (metadataflag == 1)
	{
		errortitle = Error: Display Metadata
		metadataflag = 0
	}
	else if (viewissuexmlflag == 1)
	{
		errortitle = Error: ViewIssueXML
		viewissuexmlflag = 0
	}
	else if (editissuexmlflag == 1)
	{
		errortitle = Error: EditIssueXML
		editissuexmlflag = 0
	}
	else errortitle = Error
	
	if (tiffopenflag == 1)
	{
		errorsubtitle = OpenFirstTIFF`nNDNP_QR_navigation.ahk
		tiffopenflag = 0
	}

	MsgBox, 0, %errortitle%, ===Issue Folder Path===`nNDNP_QR_metadata.ahk`n`nUnable to copy folder name.`n`nScript will end now.
	clipboard = %temp%		
Exit
; ======================ISSUE FOLDER PATH

; ======================ISSUE METADATA
IssueMetadata:
	; if the line contains a volume number tag
	IfInString, A_LoopField, type="volume"
	{
		volumeflag = 1
		Continue
	}

	; if the volumeflag has been set
	if (volumeflag == 1)
	{
		StringTrimLeft, volume, A_LoopField, 13
		StringTrimRight, volume, volume, 14
		volumeflag = 0
	}

	; if the line contains an issue number tag
	IfInString, A_LoopField, type="issue"
	{
		issueflag = 1
		Continue
	}

	; if the issueflag has been set
	if (issueflag == 1)
	{
		StringTrimLeft, issue, A_LoopField, 13
		StringTrimRight, issue, issue, 14
		issueflag = 0
	}

	; if the line contains an edition label
	IfInString, A_LoopField, mods:caption
	{
		StringTrimLeft, editionlabel, A_LoopField, 14
		StringTrimRight, editionlabel, editionlabel, 15

		; create Edition Label GUI
		; except for Batch Report or first pass of Reel Report / Metadata Viewer
		if (batchreportflag != 1)
		{
			if ((loopcount == 0) || (loopcount > 1))
			{
				WinGetPos, winX, winY, winWidth, winHeight, Metadata
				winY+=%winHeight%
				Gui, 9:+AlwaysOnTop
				Gui, 9:+ToolWindow
				Gui, 9:Font, cRed s15 bold, Arial
				Gui, 9:Add, Text, x15 y8 w380 h25, %editionlabel%
				Gui, 9:Show, x%winX% y%winY% h40 w400, Edition
			}
		}
	}

	; if the line contains an issue date
	IfInString, A_LoopField, <mods:dateIssued encoding="iso8601">
	{
		StringMid, date, A_LoopField, 37, 10
		
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
		StringMid, questionable, A_LoopField, 62, 10

		month := SubStr(questionable, 6, 2)
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
		; create Questionable window if Date window exists
		SetTitleMatchMode 1
		IfWinExist, Date
		{
			WinGetPos, winX, winY, winWidth, winHeight, Date
			winY+=%winHeight%
			Gui, 7:+AlwaysOnTop
			Gui, 7:+ToolWindow
			Gui, 7:Font, cRed s15 bold, Arial
			Gui, 7:Add, Text, x35 y8 w160 h25, %monthnameQ% %dayQ%, %yearQ%
			Gui, 7:Show, x%winX% y%winY% h40 w200, Questionable
		}
	}
	
	; if the line contains a section label tag
	IfInString, A_LoopField, section label
	{
		sectionflag = 1
		sectioncount++
		Continue
	}

	; if the sectionflag has been set
	if (sectionflag == 1)
	{
		StringTrimLeft, section, A_LoopField, 13
		StringTrimRight, section, section, 14
		if (sectioncount > 1)
		{
			sectionlabels := sectionlabels . " | " . section
		}
		else
		{
			sectionlabels := section		
		}
		sectionflag = 0
		section := ""
	}

	; if the line indicates a missing page
	IfInString, A_LoopField, Not digitized`, published
	{
		missing++
	}
Return
; ======================ISSUE METADATA

; ======================EXTRACT METADATA
ExtractMeta:
	temp = %clipboard%
	issuexml =
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
	volumeflag = 0
	issueflag = 0
	missing = 0
	missingdisplay =
	sectionlabels =
	sectioncount = 0
	sectionflag = 0
	
	; read in issue.xml file to issuexml variable
	FileRead, issuexml, %issuefolderpath%\%issuefoldername%.xml
				
	Loop, parse, issuexml, `n, `r%A_Space%%A_Tab% ; parse issuexml
	{
		Gosub, IssueMetadata
	}

	numpages = 0					
	Loop, %issuefolderpath%\*.tif ; count tiffs
	{
		numpages++
	}

	; update date
	ControlSetText, Static6, %date%, Metadata
	ControlSetText, Static1, %monthname% %day%`, %year%, Date

	; update volume number
	ControlSetText, Static7, %volume%, Metadata
	SetTitleMatchMode 1
	IfWinExist, Volume
	{
		WinGetPos, winX, winY,,, Volume
		Gui, 5:Destroy
		Gui, 5:+AlwaysOnTop
		Gui, 5:+ToolWindow
		Gui, 5:Font, cRed s15 bold, Arial
		StringLen, lengthvol, volume
		if (lengthvol > 5)
		{
			Gui, 5:Add, Text, x15 y8 w170 h25, %volume%
			Gui, 5:Show, x%winX% y%winY% h40 w190, Volume
		}
		else
		{
			Gui, 5:Add, Text, x15 y8 w70 h35, %volume%
			Gui, 5:Show, x%winX% y%winY% h40 w100, Volume
		}
	}

	; update issue number
	ControlSetText, Static8, %issue%, Metadata
	SetTitleMatchMode 1
	IfWinExist, Issue
	{
		WinGetPos, winX, winY,,, Issue
		Gui, 6:Destroy
		Gui, 6:+AlwaysOnTop
		Gui, 6:+ToolWindow
		Gui, 6:Font, cRed s15 bold, Arial
		StringLen, lengthiss, issue
		if (lengthiss > 5)
		{
			Gui, 6:Add, Text, x15 y8 w170 h25, %issue%
			Gui, 6:Show, x%winX% y%winY% h40 w190, Issue
		}
		else
		{
			Gui, 6:Add, Text, x15 y8 w70 h25, %issue%
			Gui, 6:Show, x%winX% y%winY% h40 w100, Issue
		}
	}

	; update questionable date
	ControlSetText, Static9, %questionable%, Metadata		

	; update number of pages
	ControlSetText, Static10, %numpages%, Metadata
	
	; create GUIs for Date, Volume, Issue, Questionable, and Edition Label
	; if first pass of Reel Report or Metadata Viewer
	if (loopcount == 1)
	{
		Gosub, CreateMetaWindows
	}
	
	clipboard = %temp%
Return
; ======================EXTRACT METADATA

; ======================REDRAW METADATA WINDOW
RedrawMetaWindow:
	; get window coordinates
	SetTitleMatchMode 1
	IfWinExist, Metadata
	{
		WinGetPos, winX, winY,,, Metadata
	}
	else
	{
		WinGetPos, winX, winY, winWidth, winHeight, NDNP_QR
		winX+=%winWidth%
	}
	
	Gui, 2:Destroy ; Metadata window
	
	; create the Metadata window
	Gui, 2:Font, s8, Arial
	Gui, 2:Add, Text, x26 y20  w100 h20, Date:
	Gui, 2:Add, Text, x13 y45  w100 h20, Volume:
	Gui, 2:Add, Text, x22 y70  w100 h20, Issue:
	Gui, 2:Add, Text, x17 y95  w100 h20, ? Date:
	Gui, 2:Add, Text, x18 y120 w100 h20, Pages:
	Gui, 2:Font, cRed s12 bold, Arial
	Gui, 2:Add, Text, x65 y20  w100 h20,
	Gui, 2:Add, Text, x65 y45  w130 h20,
	Gui, 2:Add, Text, x65 y70  w130 h20,
	Gui, 2:Add, Text, x65 y95  w100 h20,
	Gui, 2:Add, Text, x65 y120 w100 h20,
	Gui, 2:+AlwaysOnTop
	Gui, 2:+ToolWindow
	Gui, 2:Show, x%winX% y%winY% h160 w200, Metadata
	ControlSetText, Static6, %date%, Metadata
	ControlSetText, Static7, %volume%, Metadata
	ControlSetText, Static8, %issue%, Metadata		
	ControlSetText, Static9, %questionable%, Metadata		
	ControlSetText, Static10, %numpages%, Metadata
Return
; ======================REDRAW METADATA WINDOW

; ======================CREATE METADATA WINDOWS
CreateMetaWindows:
	dateX = 0
	dateY = 0
	volumeX = 0
	volumeY = 0
	issueX = 0
	issueY = 0
	
	; get screen coordinates of existing windows
	SetTitleMatchMode 1
	IfWinExist, Date
	{
		WinGetPos, dateX, dateY,,, Date
	}
	IfWinExist, Volume
	{
		WinGetPos, volumeX, volumeY,,, Volume
	}
	IfWinExist, Issue
	{
		WinGetPos, issueX, issueY,,, Issue
	}
	
	Gui, 4:Destroy
	Gui, 5:Destroy
	Gui, 6:Destroy
	Gui, 7:Destroy
	Gui, 9:Destroy
	
	Gosub, RedrawMetaWindow	
	WinWaitActive, Metadata
	WinGetPos, winX, winY, winWidth, winHeight, Metadata
	winX+=%winWidth%

	; Issue number
	Gui, 6:+AlwaysOnTop
	Gui, 6:+ToolWindow
	Gui, 6:Font, cRed s15 bold, Arial
	StringLen, lengthiss, issue
	if (lengthiss > 5)
	{
		Gui, 6:Add, Text, x15 y8 w170 h35, %issue%
		if ((issueX == 0) && (issueY == 0))
			Gui, 6:Show, x%winX% y%winY% h40 w190, Issue
		else
			Gui, 6:Show, x%issueX% y%issueY% h40 w190, Issue
	}
	else
	{
		Gui, 6:Add, Text, x15 y8 w70 h25, %issue%
		if ((issueX == 0) && (issueY == 0))
			Gui, 6:Show, x%winX% y%winY% h40 w100, Issue
		else
			Gui, 6:Show, x%issueX% y%issueY% h40 w100, Issue
	}

	; Volume number
	Gui, 5:+AlwaysOnTop
	Gui, 5:+ToolWindow
	Gui, 5:Font, cRed s15 bold, Arial
	StringLen, lengthvol, volume
	if (lengthvol > 5)
	{
		Gui, 5:Add, Text, x15 y8 w170 h35, %volume%
		if ((volumeX == 0) && (volumeY == 0))
			Gui, 5:Show, x%winX% y%winY% h40 w190, Volume
		else
			Gui, 5:Show, x%volumeX% y%volumeY% h40 w190, Volume
	}
	else
	{
		Gui, 5:Add, Text, x15 y8 w70 h35, %volume%
		if ((volumeX == 0) && (volumeY == 0))
			Gui, 5:Show, x%winX% y%winY% h40 w100, Volume
		else
			Gui, 5:Show, x%volumeX% y%volumeY% h40 w100, Volume
	}
	
	; Date
	Gui, 4:+AlwaysOnTop
	Gui, 4:+ToolWindow
	Gui, 4:Font, cRed s15 bold, Arial
	Gui, 4:Add, Text, x35 y8 w160 h25, %monthname% %day%, %year%
	if ((dateX == 0) && (dateY == 0))
		Gui, 4:Show, x%winX% y%winY% h40 w200, Date
	else
		Gui, 4:Show, x%dateX% y%dateY% h40 w200, Date
	
	if (questionable != "") ; Questionable Date
	{
		WinGetPos, winX, winY, winWidth, winHeight, Date
		winY+=%winHeight%
		Gui, 7:+AlwaysOnTop
		Gui, 7:+ToolWindow
		Gui, 7:Font, cRed s15 bold, Arial
		Gui, 7:Add, Text, x35 y8 w160 h25, %monthnameQ% %dayQ%, %yearQ%
		Gui, 7:Show, x%winX% y%winY% h40 w200, Questionable
	}

	if (editionlabel != "") ; Edition Label
	{
		WinGetPos, winX, winY, winWidth, winHeight, Metadata
		winY+=%winHeight%
		Gui, 9:+AlwaysOnTop
		Gui, 9:+ToolWindow
		Gui, 9:Font, cRed s15 bold, Arial
		Gui, 9:Add, Text, x15 y8 w380 h25, %editionlabel%
		Gui, 9:Show, x%winX% y%winY% h40 w400, Edition
	}				
		
	if (loopcount == 1) ; pause first pass
	{
		MsgBox, 0, %metaloopname% Paused, Position the metadata windows as desired.`n`nClick OK and the loop will continue in %delay% seconds.
	}
Return
; ======================CREATE METADATA WINDOWS
