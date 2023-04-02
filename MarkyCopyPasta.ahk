﻿#Requires AutoHotkey v2.0
SetTitleMatchMode 2

WelcomeMessage()

!\:: Reload

^[:: CopyExcelColumnToCaMSys()
^]:: CheckStudentIDOrder()
!x:: CopyExcelColumnToCaMSys()
!c:: CheckStudentIDOrder()
!q:: ExitApp

; Test settings
; !w:: MsgBox A_Clipboard
; Test settings end

WelcomeMessage() {
	welcomemsg := "
	(
		While this script is running, a green 'H' icon will be in your system tray.`r`n
		This program copies a vertical column of marks from Excel into the coursework marks entry page in CaMSys.`r`n`r`n
		Three steps to use this program:`r`n
		(1) Open the Excel file containing your marks and place the Excel cell cursor at the top of the column of marks (on the first student's mark) you wish to copy. Make sure there is nothing in the cell below the last mark. Any students with no marks should be given a zero.`r`n
		(2) Open Chrome, log into CaMSys and navigate to the Coursework Marks Entry page for your subject. Click on the component required, and place your cursor inside the box for the first student's mark.`r`n
		(3) Press Ctrl+[. Do not touch the keyboard while the script runs.`r`n
		Repeat this with as many columns of marks as you need, selecting the correct start of columns in Excel and CaMSys respectively.`r`n`r`n
		This program can also check if the order of Student IDs in Excel and CaMSys matches:`r`n
		(1) Open the Excel file containing your marks and place the Excel cell cursor at the top of the column of Student IDs (on the first student's ID).`r`n
		(2) Open Chrome, log into CaMSys and navigate to the Coursework Marks Entry page for your subject. Make sure the cursor is not in the input box. (If you just opened the page, you don't have to do anything. Or you can click randomly somewhere on the text in the page.)`r`n
		(3) Press Ctrl+]. Do not touch the keyboard while the script runs.`r`n`r`n
		Press Alt+Q to Quit the script, or right-click the 'H' icon in your system tray and click Exit.`r`n`r`n
		This program was built by Willie Poh at Hackerspace MMU's Hackathon No. 23. Version 0.1 (Beta Test)
	)"

	MsgBox welcomemsg, "Welcome to MarkyCopyPasta!"
}

CopyExcelColumnToCaMSys() {
	WaitNoAltKey()
	Marks := CopyExcelColumn()

	if !Marks {
		MsgBox "Failed to copy only numbers from Excel."
		return false
	}

	PasteColumnInCaMSys(Marks)
}

CheckStudentIDOrder() {
	WaitNoAltKey()
	ExcelStudentIDs := GetStudentIDExcel()

	if !ExcelStudentIDs {
		MsgBox "Failed to copy only numbers from Excel."
		return false
	}

	CaMSysStudentIDs := GetStudentIDCaMSys()

	if CaMSysStudentIDs.Length < 1 {
		MsgBox "Failed to copy Student IDs from CaMSys."
		return false
	}

	; MsgBox StrJoin(",", ExcelStudentIds)
	; MsgBox StrJoin(",", CaMSysStudentIDs)

	if ExcelStudentIds.Length == CaMSysStudentIDs.Length
		longer := false
	else if ExcelStudentIds.Length > CaMSysStudentIDs.Length
		longer := "There are more Student IDs in Excel than in CaMSys."
	else if ExcelStudentIds.Length < CaMSysStudentIDs.Length
		longer := "There are more Student IDs in CaMSys than in Excel."

	For index, esid in ExcelStudentIds {
		if CaMSysStudentIDs.Has(index) {
			if esid != CaMSysStudentIDs[index] {
				if !longer {
					MsgBox "Student ID lists do not match! Mismatch found at position number " index
					Exit
				}
				else {
					MsgBox "Student ID lists do not match! Mismatch found at position number " index ". " longer
					Exit
				}
			}
		}
		else {
			MsgBox "Student ID lists do not match! Mismatch found at position number " index ". " longer
			Exit
		}
	}

	if !longer
		MsgBox "Student ID lists match!"
	else
		MsgBox "Student ID lists do not match! " longer
}

CopyExcelColumn() {
	if !SwitchToExcelWindow()
		Exit

	A_Clipboard := "xyzblah"

	; Copy contents of a cell to clipboard
	Send "^+{Down}"
	Sleep 60
	Send "^c"
	Sleep 200

	if A_Clipboard == "xyzblah"
		return false
	else if A_Clipboard == ""
		return false

	Marks := []
	Marks := StrSplit(A_Clipboard, "`r`n") ; Last item in this array is a blank

	Send "{Esc}"
	Sleep 30
	Send "{Right}"
	Sleep 30
	Send "{Left}"
	Sleep 30

	; Remove the last blank element of the array
	Marks.pop

	for mark in Marks
		if !IsNumber(mark)
			return false

	return Marks
}

PasteColumnInCaMSys(Marks) {
	WaitNoAltKey()
	if !SwitchToCaMSysWindow()
		Exit

	; Perform check - make sure it's in data entry mode
	A_Clipboard := "xyzblah"
	Sleep 60
	Send "^a"
	Sleep 100  
	Send "^c"
	Sleep 30
	if !IsNumber(A_Clipboard) {
		MsgBox "Cursor is not in input field. Please click / place the cursor into the first marks entry field."
		Exit
	}

	For index, Mark in Marks {
		if (index == 1) and (Mark == 0) {
			Send "{Tab}"
			if Marks[2] == 0 {
				Send "{Tab}"
				if Marks[3] == 0 {
					Send "{Tab}"
					if Marks[4] ==0 {
						MsgBox "Detected that your first four marks are ZEROES (0). When there are more than three ZEROES (O) at the top of your marks table, please enter them manually and start your cursor in Excel and CaMSys from the first non-zero mark."
						Exit
					}
					Send Marks[4]
					Send "+{Tab}"
				}
				Send Marks[3]
				Send "+{Tab}"
			}
			Send Marks[2]
			Send "+{Tab}"
			Send Mark
			Send "{Tab}"
		}
		else {
			Send Mark
			Send "{Tab}"
		}
	}
	Sleep 8000
	MsgBox "Finished copying marks from Excel to CaMSys! Please wait for the CaMSys page to finish 'spinning.' Remember to check marks entered and click 'Save' once confirmed."
}

GetStudentIDExcel() {
	if !SwitchToExcelWindow()
		Exit

	A_Clipboard := "xyzblah"

	; Copy contents of a cell to clipboard
	Send "^+{Down}"
	Sleep 60
	Send "^c"
	Sleep 200

	if A_Clipboard == "xyzblah"
		return false
	else if A_Clipboard == ""
		return false

	StudentIDs := []
	StudentIDs := StrSplit(A_Clipboard, "`r`n") ; Last item in this array is a blank

	Send "{Esc}"
	Sleep 30
	Send "{Right}"
	Sleep 30
	Send "{Left}"
	Sleep 30

	; Remove the last blank element of the array
	StudentIDs.pop

	for StudentID in StudentIDs
		if !IsNumber(StudentID)
			return false

	return StudentIDs
}

GetStudentIDCaMSys() {
	WaitNoAltKey()
	if !SwitchToCaMSysWindow()
		Exit

	; Make sure something new enters the clipboard
	A_Clipboard := "xyzblah"
	Send "^a"
	Sleep 60
	Send "^c"
	Sleep 1000
	if A_Clipboard == "xyzblah" {
		MsgBox "Failed to get Student IDs from page."
		Exit
	}

	Data := []
	Data := StrSplit(A_Clipboard, "`r`n")
	StudentIDs := []

	for index, datum in Data {
		if StrLen(datum) == 10 && IsInteger(datum) && datum > 1000000000
			StudentIDs.push(datum)
	}

	return StudentIDs
}

SwitchToExcelWindow() {
	if !WinActive("ahk_class XLMAIN") {
		if WinExist("ahk_class XLMAIN") {
			if CountExcelWindows() == 1 {
				WinActivate("ahk_class XLMAIN")
			}
			else {
				MsgBox("Too many Excel files open. Please open only the Excel file containing your student marks.")
				; ECHO please make sure only one Excel window is open.
				return false
			}
		}
		else {
			MsgBox("No Excel file currently open. Please open an Excel file containing your student marks.")
			return false
		}
	}
	return true
}

CountExcelWindows() {
	WinHandles := []
	WinList := WinGetList("ahk_class XLMAIN")
	For Each, Win in WinList {
		WinHandles.Push(Win)
	}
	return WinHandles.Length

}

SwitchToCaMSysWindow() {
	if (!WinExist("Course Work Marks - Google Chrome ahk_exe chrome.exe") and !WinExist("CW Marks Entry - Google Chrome ahk_exe chrome.exe")) {
		MsgBox("Your Google Chrome is not opened to the Coursework Marks entry page. Please open Google Chrome to the correct page and place the cursor on the first value of the column of marks you wish to copy to.")
		return false
	}
	else {
		if WinExist("CW Marks Entry - Google Chrome ahk_exe chrome.exe")
			WinActivate("CW Marks Entry - Google Chrome ahk_exe chrome.exe")
		else if WinExist("Course Work Marks - Google Chrome ahk_exe chrome.exe")
			WinActivate("Course Work Marks - Google Chrome ahk_exe chrome.exe")
	}
	Sleep 60
	return true
}

StrJoin(sep, params) {
	for param in params
		str .= param . sep
	return SubStr(str, 1, -StrLen(sep))
}

WaitNoAltKey() {
	Loop {
		if GetKeyState("LAlt", "P")
			Continue
		else if GetKeyState("RAlt", "P")
			Continue
		else if GetKeyState("c", "P")
			Continue
		else if GetKeyState("x", "P")
			Continue
		else Break
	}
}