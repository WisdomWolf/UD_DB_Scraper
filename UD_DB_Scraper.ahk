;todo: add some help options to tray menu
;todo: integrate login details
SetTitleMatchMode,2
SendMode Input
SetKeyDelay, 0, 10
FormatTime, newDate, %A_Now%, yyMMdd

csvFile := A_ScriptDir . "\UD_DB_New_" . newDate . ".csv"
manuallyStarted := false
system :=
newUnitsFound := 0

IfNotExist, %A_AppData%/CompassAHKTweaks
{
	FileCreateDir, %A_AppData%/CompassAHKTweaks
}
SetWorkingDir, %A_AppData%/CompassAHKTweaks

IfNotExist, %A_WorkingDir%\CompassTweaks.ini
{
	FileSelectFile, UDEXE, 3,, Universal Desktop Location
	If ErrorLevel {
		msgBox, 48, Error, You must specify Universal Desktop Location to continue.
		ExitApp
	}
	IniWrite, %UDEXE%, %A_WorkingDir%\CompassTweaks.ini, Universal Desktop, Location
} else {
	IniRead, UDEXE, %A_WorkingDir%\CompassTweaks.ini, Universal Desktop, Location
}

file := A_WorkingDir . "\UD_DB.csv"

;Bullet-Proof UD start-up
Process, Exist, UniversalDesktop.exe
if (errorlevel) {
	Process, Close, %errorlevel%
	Process, WaitClose, %errorlevel%, 5
	if (errorlevel) {
		msgBox Unable to kill UD
		return
	}
}
Run, %UDEXE%
TrayTip, UD Scraper, UD Launched!, 10, 1
WinWaitActive, Login ahk_class WindowsForms10.Window.8.app.0.33c0d9d, 5
if (errorlevel) {
	msgBox Script Timed Out While Waiting For Window.
	ExitApp
} else {
	TrayTip, UD Scraper, Ready! `nEnter login details and then press Right Ctrl to continue., 10, 1
}
KeyWait, RCTRL, D T60
if (errorlevel) {
	if (!manuallyStarted) {
		msgBox Script Timed Out Waiting for Keypress.
		GoSub,Exiter
	}
} else {
	TrayTip, UD Scraper, Beginning Scrape..., 10, 1
	ControlGetText, Field1, WindowsForms10.EDIT.app.0.33c0d9d2, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
	ControlGetText, Field2, WindowsForms10.EDIT.app.0.33c0d9d3, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
	;msgBox, % "Field1: " . Field1 . "`nField2: " . Field2
	Gosub,Scrape
	Exit
}

;enumeration
Scrape:
	Array := Object()
	Gosub,GenerateSystemArray
	LoopCount := ArrayCount
	Loop %LoopCount%
	{
		system := Array[1]
		FileAppend, Attempting scrape of %system%.  %ArrayCount% systems remaining.  Index %A_Index%`n, UD_Scraper_Log.txt
		Sleep 500
		errorlevel := 0
		Control, ChooseString, %system%, WindowsForms10.COMBOBOX.app.0.33c0d9d1, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
		if (errorlevel) {
			Sleep 200
			Control, ChooseString, %system%, WindowsForms10.COMBOBOX.app.0.33c0d9d5, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
			if (errorlevel) {
				FileAppend, System select fail at: %system%.`n, UD_Scraper_Log.txt
				MsgBox, System select fail at: %system%.
				GoSub,PauseScrape
			}
		}
		;ControlFocus, Enter
		ControlSend, Enter, {Enter}, Login ahk_class WindowsForms10.Window.8.app.0.33c0d9d
		if (errorlevel) {
			FileAppend, Control send enter failed. Exiting.`n, UD_Scraper_Log.txt
			msgBox, Control send enter failed. Exiting.
			GoSub,PauseScrape
		}
		WinWait, Login ahk_class WindowsForms10.Window.8.app.0.33c0d9d,Authenticating, 5
		WinWaitClose, Login ahk_class WindowsForms10.Window.8.app.0.33c0d9d,Authenticating, 20
		Sleep 500
		IfWinExist, Authentication Failure
		{
			TrayTip, CompassAHKTweaks, UD Scraper, Authentication Failure Detected
			ControlSend, OK, {ENTER}, ahk_class #32770
			WinWaitActive, Login ahk_class WindowsForms10.Window.8.app.0.33c0d9d,, 2
			Sleep, 100
			Control, EditPaste, %Field1%, WindowsForms10.EDIT.app.0.33c0d9d2, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
			Sleep, 100
			ControlSetText, WindowsForms10.EDIT.app.0.33c0d9d3, %Field2%, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
			TrayTip, UD Scraper, Continuing...
			FileAppend, Unabled to login to %system%. Continuing to next system.`n, UD_Scraper_Log.txt
			Array.Remove(1)
			ArrayCount -= 1
			Continue
		}
		IfWinExist, Universal Desktop,,Login
		{
			TrayTip, UD Scraper, Generating info for single login system.
			InputBox, NewUnitNumber, %system%, Please enter this unit number:,,220,150
			InputBox, NewUnitLabel, %system%, Please enter the name for this unit:,,220,150
			AddUnit(NewUnitLabel, system, "x", "x", "x", NewUnitNumber)
			FileAppend, Manually added %NewUnitNumber% %NewUnitLable% on %system%`n, UD_Scraper_Log.txt
			ControlSend, ahk_parent, {F4}, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
			WinWaitActive, Login ahk_class WindowsForms10.Window.8.app.0.33c0d9d, 5
			if (errorlevel) {
				ControlSend, ahk_parent, {F4}, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
				WinWaitActive, Login ahk_class WindowsForms10.Window.8.app.0.33c0d9d, 5
				if (errorlevel) {
					msgBox, UD Exit routine timed out.  Please manually close Universal Desktop and wait for login window to proceed.
					WinWaitActive, Login ahk_class WindowsForms10.Window.8.app.0.33c0d9d, 60
					if (errorlevel) {
						FileAppend, Attempts to exit UD Timed-Out. Exiting.`n, UD_Scraper_Log.txt
						msgBox, Timed Out.  Please try again later.
						GoSub,Exiter
					}
				}
			}
			Sleep 500
			ControlSend, WindowsForms10.EDIT.app.0.33c0d9d2, %Field1%, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
			ControlSend, WindowsForms10.EDIT.app.0.33c0d9d3, %Field2%, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
		} else {
			WinWait, Login ahk_class WindowsForms10.Window.8.app.0.33c0d9d, Getting Login, 3
			WinWaitClose, Login ahk_class WindowsForms10.Window.8.app.0.33c0d9d, Getting Login, 30
			GoSub,ResetTree
			GoSub,TraverseTree
			Sleep 300
			ControlSend, Cancel, {Space}, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
			if (errorlevel) {
				FileAppend, Control send cancel failed.  Exiting.`n, UD_Scraper_Log.txt
				msgBox, Control send cancel failed. Exiting.
				break
			}
			Array.Remove(1)
			ArrayCount -= 1
			FileAppend, Completed scrape of %system%.  %ArrayCount% systems remaining.`n, UD_Scraper_Log.txt
			WinWaitActive, Login ahk_class WindowsForms10.Window.8.app.0.33c0d9d,,3
		}
	}
	if (newUnitsFound) {
		msgBox, % "Data Scrape complete!`n" . newUnitsFound . " new units found."
		try
			Run, % csvFile
		catch
			MsgBox, % "Unable to find or launch " . csvFile
	} else {
		msgBox, % "Data Scrape complete!`nNo new units discovered."
	}
	GoSub,PauseScrape
return

#IfWinActive ahk_class WindowsForms10.Window.8.app.0.33c0d9d

NumpadEnter::
Enter::
Tab::
	manuallyStarted := true
	TrayTip, UD Scraper, Beginning Scrape..., 10, 1
	ControlGetText, Field1, WindowsForms10.EDIT.app.0.33c0d9d2, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
	ControlGetText, Field2, WindowsForms10.EDIT.app.0.33c0d9d3, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
	;msgBox, % "Field1: " . Field1 . "`nField2: " . Field2
	Gosub,Scrape
return

#IfWinActive Authentication Failure ahk_class #32770
LButton::
RButton::
MButton::
Enter::
NumpadEnter::
Space::
	TrayTip, UD Scraper, Please wait for the scrape to complete.
Return

#IfWinActive
^!x::
	FileAppend, Exit forced by user.`n`n`n, UD_Scraper_Log.txt
	GoSub,PauseScrape
return
	
GenerateSystemArray:
	IniRead, ArrayCount, %A_WorkingDir%\CompassTweaks.ini, Last Run, SystemCount, -1
	if (ArrayCount > 0) {
		TrayTip, UD Scraper, Getting Systems from ini file
		FileAppend, System information obtained from ini file.`n, %A_WorkingDir%\UD_Scraper_Log.txt
		Loop %ArrayCount%
		{
			IniRead, system, %A_WorkingDir%\CompassTweaks.ini, Last Run, SystemArray%A_Index%, x
			Array.Insert(system)
		}
		return
	}
	IfNotExist, %A_WorkingDir%\Scraper_List.txt
	{
		MsgBox, 4, UD Scraper, Couldn't locate system list. Would you like to generate one?
		IfMsgBox, Yes
		{
			ControlGet, SystemList, List,, WindowsForms10.COMBOBOX.app.0.33c0d9d1, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
			Loop, Parse, SystemList, `
			{
				FileAppend, %A_LoopField%, %A_WorkingDir%\Scraper_List.txt
			}
			MsgBox, System List Generation complete.
			GoSub, %A_ThisLabel%
		} else {
			exit
		}
	}
	else
	{
		TrayTip, UD Scraper, Gettings Systems from Scraper_List.txt
		FileAppend, System information obtained from Scraper List.`n, UD_Scraper_Log.txt
		Loop, Read, %A_WorkingDir%\Scraper_List.txt
		{
			ArrayCount += 1
			Array.Insert(A_LoopReadLine)
			FileAppend, Index %A_Index% is %A_LoopReadLine%`n, %A_WorkingDir%\UD_Scraper_Log.txt
		}
		FileAppend, `n, %A_WorkingDir%\UD_Scraper_Log.txt
	}
return

GrabText:
	;Get CustomerChoice Position
	ControlGetText, Customer, Edit4, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
	ControlGetText, Enterprise, Edit3, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
	ControlGetText, Division, Edit2, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
	ControlGet, StoreList, List,, WindowsForms10.COMBOBOX.app.0.33c0d9d1, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
	
	;Unit,UnitName,System,Customer,Enterprise,Division
	Loop,Parse,StoreList,`n 
	{	
		AddUnit(A_LoopField, system, Customer, Enterprise, Division)
	}
	TrayTip, AutoHotkey, Parsing data to csv..., 10, 1
return

ParseUnitNumber(unit)
{
	if (RegExMatch(unit, "^Z\s*-")) or (RegExMatch(unit, "[Cc]losed"))
		return "z"
	RegExMatch(unit, "\d{3,5}(?=\W)", result)
	if (result = "") {
		result = 5555
	}
	return result
}

AddUnit(Store, System, Customer, Enterprise, Division, UnitNumber="new")
{	
	global newUnitsFound
	global file
	UnitFound := false
	
	IfExist, %file%
	{
		Loop, Read, %file%
		{
			if (ErrorLevel) {
				TrayTip, UD Scraper, File Read Fail!
				break
			}
			unitDetails := StrSplit(A_LoopReadLine, ",")
			{
				If (unitDetails[2] = Store){
					UnitFound := true
					break
				}
			}
			If UnitFound
				break
		}
		If !UnitFound and FileExist(csvFile)
		{
			Loop, Read, %csvFile%
			{
				if (ErrorLevel) {
					TrayTip, UD Scraper, File Read Fail!
					break
				}
				unitDetails := StrSplit(A_LoopReadLine, ",")
				{
					If (unitDetails[2] = Store){
						UnitFound := true
						break
					}
				}
				If UnitFound
					break
			}
			If !UnitFound
			{
				AddToCSV(Store, System, Customer, Enterprise, Division, UnitNumber)
			}
		} else {
			If !UnitFound
			{
				AddToCSV(Store, System, Customer, Enterprise, Division, UnitNumber)
			}
		}
	} else {
		AddToCSV(Store, System, Customer, Enterprise, Division, UnitNumber)
		FileAppend, %UnitNumber%`,%Store%`,%system%`,%Customer%`,%Enterprise%`,%Division%`n, %A_ScriptDir%\UD_DB_BrandNew.csv
	}
Return
}

AddToCSV(Store, System, Customer, Enterprise, Division, UnitNumber)
{
	global file
	global csvFile
	global newUnitsFound
	
	If (UnitNumber = "new") {
		UnitNumber := ParseUnitNumber(Store)
		If (UnitNumber = "new") {
			msgBox, % "WTF Mate, Unit Number Parse failed"
			return
		}
		If (UnitNumber = "acketz" or UnitNumber = "z")
			return
	}
	newUnitsFound++
	FileAppend, %UnitNumber%`,%Store%`,%system%`,%Customer%`,%Enterprise%`,%Division%`n, %csvFile%
	FileAppend, %UnitNumber%`,%Store%`,%system%`,%Customer%`,%Enterprise%`,%Division%`n, %file%
}

ResetTree:
	Control, Choose, 1, WindowsForms10.COMBOBOX.app.0.33c0d9d4
	Control, Choose, 1, WindowsForms10.COMBOBOX.app.0.33c0d9d3
	Control, Choose, 1, WindowsForms10.COMBOBOX.app.0.33c0d9d2
	Control, Choose, 1, WindowsForms10.COMBOBOX.app.0.33c0d9d1
return

TraverseTree:
TraverseCustomer:
{
	TrayTip, AutoHotkey, Traversing Customer!, 10, 1
	Loop {
		Control, Choose, %A_Index%, WindowsForms10.COMBOBOX.app.0.33c0d9d4, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
		if (errorlevel != 1) 
			Gosub,TraverseEnterprise
		else
			break
	}
	return
}

TraverseEnterprise:
	Loop {
		Control, Choose, %A_Index%, WindowsForms10.COMBOBOX.app.0.33c0d9d3, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
		if (errorlevel != 1)
			Gosub, TraverseDivision
		else
			break
	}
return

TraverseDivision:
	Loop {
		Control, Choose, %A_Index%, WindowsForms10.COMBOBOX.app.0.33c0d9d2, ahk_class WindowsForms10.Window.8.app.0.33c0d9d
		if (errorlevel != 1)
			Gosub,GrabText
		else
			break
	}
return

;Bullet-Proof UD start-up
BulletProofUD:
	Process, Exist, UniversalDesktop.exe
	if (errorlevel) {
		Process, Close, %errorlevel%
		Process, WaitClose, %errorlevel%, 5
		if (errorlevel) {
			msgBox Unable to kill UD
			return
		}
	}
	Run, %UDEXE%
return

KillUDLogin:
	WinWait, Login
	WinClose, Login
return

PauseScrape:
	IniWrite, %ArrayCount%, %A_WorkingDir%\CompassTweaks.ini, Last Run, SystemCount
	Loop %ArrayCount%
	{
		system := Array[A_Index]
		IniWrite, %system%, %A_WorkingDir%\CompassTweaks.ini, Last Run, SystemArray%A_Index%
	}
	GoSub,Exiter
return

Exiter:
	GoSub,KillUDLogin
ExitApp