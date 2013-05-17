; ********************************
; NDNP_QR_metadata.ahk
; required file for NDNP_QR.ahk
; ********************************

; ======================ISSUE FOLDER PATH
IssueFolderPath:
  ; save clipboard contents
	temp = %clipboard%
	
	; empty the clipboard
	clipboard = _

	; activate reel folder
	SetTitleMatchMode 1
	WinActivate, %reelfoldername%, , , ,
	WinWaitActive, %reelfoldername%, , 5, ,
	if ErrorLevel
	{
		; reset the variables
		errortitle =
		errorsubtitle =
		
		; error handling variables
		if (openflag == 1)
		{
			errortitle = Open
			openflag = 0
		}
		else if (openplusflag == 1)
		{
			errortitle = Open+
			openplusflag = 0
		}
		else if (nextflag == 1)
		{
			errortitle = Next
			nextflag = 0
		}
		else if (nextplusflag == 1)
		{
			errortitle = Next+
			nextplusflag = 0
		}
		else if (previousflag == 1)
		{
			errortitle = Previous
			previousflag = 0
		}
		else if (previousplusflag == 1)
		{
			errortitle = Previous+
			previousplusflag = 0
		}
		else if (gotoflag == 1)
		{
			errortitle = GoTo
			gotoflag = 0
		}
		else if (gotoplusflag == 1)
		{
			errortitle = GoTo
			gotoplusflag = 0
		}
		else if (metadataflag == 1)
		{
			errortitle = Metadata
			metadataflag = 0
		}
		else if (viewissuexmlflag == 1)
		{
			errortitle = ViewIssueXML
			viewissuexmlflag = 0
		}
		else if (editissuexmlflag == 1)
		{
			errortitle = EditIssueXML
			editissuexmlflag = 0
		}
		else errortitle = Error
		
		if (tiffopenflag == 1)
		{
			errorsubtitle = OpenFirstTIFF`nNDNP_QR_navigation.ahk
			tiffopenflag = 0
		}
		
		; print error message after 5 seconds
		MsgBox, 0, %errortitle%, %errorsubtitle%`nIssueFolderPath`nNDNP_QR_metadata.ahk`n`nCannot find folder %reelfoldername%`n`nOptions:`n`t"File > Open Reel Folder"`n`t"Edit > Reel Folder > Set Path"
		
		; exit the script
		Exit
	}
	Sleep, 200
	
	; copy name of selected folder
	Send, {F2}
	Sleep, 100
	Send, {CtrlDown}c{CtrlUp}
	Sleep, 100
	Send, {Enter}
	Sleep, 100
	issuefoldername = %clipboard%

	; case for opened issue folder
	SetTitleMatchMode RegEx
	IfWinActive, ^1[0-9]{9}$, , 5
	{
		Gosub, IssueToReel
	}	
	
	; loop checks for correctly copied issue folder name
	; up to 5 times, aborts script if unsuccessful
	Loop 5
	{
		; continue the script if issue name copied to clipboard
		if RegExMatch(issuefoldername, "\d\d\d\d\d\d\d\d\d\d")
			Break
			
		; or reattempt to copy folder path
		else
		{
			; wait one second
			Sleep, 1000

			; re-activate the reel folder
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

			; next loop iteration
			Continue
		}
		
		; restore clipboard contents
		clipboard = %temp%		

		; abort script if unable to copy path after 5 attempts
		MsgBox, 0, %errortitle%, Issue Folder Path`nNDNP_QR_metadata.ahk`n`nUnable to copy folder name.`n`nScript aborted.
		Exit
	}

	; assign the issue folder path
	issuefolderpath = %reelfolderpath%\%issuefoldername%
	
	; restore clipboard contents
	clipboard = %temp%		
Return
; ======================ISSUE FOLDER PATH

; ======================EXTRACT METADATA
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
; ======================EXTRACT METADATA

; ======================CREATE METADATA WINDOWS
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
; ======================CREATE METADATA WINDOWS