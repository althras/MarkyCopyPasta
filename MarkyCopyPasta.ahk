#Requires AutoHotkey v2.0
SetTitleMatchMode 2

if FileExist("Hackerspace.ico") {
	TraySetIcon("Hackerspace.ico")
}

A_TrayMenu.Add()  ; Creates a separator line.
A_TrayMenu.Add("Instructions/Help", MenuHandler)  ; Creates a new menu item.
MenuHandler(ItemName, ItemPos, MyMenu) {
	WelcomeMessage()
}

; Fire the welcome message on load.
; WelcomeMessage()

; ########## SETUP HOTKEYS ##########
!NumpadDiv:: Reload
^[:: CopyExcelColumnToChrome("Others")
^]:: CopyExcelColumnToChrome("CWSub")
^\:: CopyExcelColumnToChrome("FinalOBE")
^;:: CheckStudentIDOrder()

; Test settings
; !w:: MsgBox "Student ID lists do not match! Mismatch found at position number " 1, "Mismatch", "iconx"
; Test settings end

; ########## HOTKEYS AFTER THIS LINE ONLY WORK IN GUI ##########
#HotIf WinActive("MarkyCopyPasta ahk_class AutoHotkeyGUI")
Esc::ExitApp

; ########## SETTING UP GUI ##########
MarkyCopyPasta := Gui(, "MarkyCopyPasta v1.1")
MarkyCopyPasta.SetFont(, "Calibri")
MarkyCopyPasta.SetFont("Bold s13")
MarkyCopyPasta.Add("Text", "w300 Center", "Welcome to MarkyCopyPasta")
MarkyCopyPasta.SetFont("Norm s10")
MCPCheckID := MarkyCopyPasta.Add("Text", "w300", "Please click 'Help' for instructions.")
ExcelReadyColor := MarkyCopyPasta.Add("Text", "BackgroundGreen", "     ")
ExcelReadyText := MarkyCopyPasta.Add("Text", "vExcelReadyText x+10 w200", "Excel Status")
ChromeReadyColor := MarkyCopyPasta.Add("Text", "vChromeReadyColor BackgroundGreen xM", "     ")
ChromeReadyText := MarkyCopyPasta.Add("Text", "vChromeReadyText x+10 w200", "CLiC Status")
MCPCheckReady := MarkyCopyPasta.Add("Button", "w300 xM", "&Recheck Readiness")
MCPCheckReady.OnEvent("Click", (*) => UpdateReadiness())
MarkyCopyPasta.Add("Text", "xM", "")
MCPCheckID := MarkyCopyPasta.Add("Button", "w300", "Compare Student ID Order (Ctrl + `;)")
MCPCheckID.OnEvent("Click", (*) => CheckStudentIDOrder())
MarkyCopyPasta.Add("Text", "xM", "Copy Marks from Excel to:")
MCPEntryOthers := MarkyCopyPasta.Add("Button", "w300", "CW or Exam Marks Entry pages (Ctrl + [)")
MCPEntryOthers.OnEvent("Click", (*) => CopyExcelColumnToChrome("Others"))
MCPEntryCWSub := MarkyCopyPasta.Add("Button", "w300", "CW Marks Sub Component Entry page (Ctrl + ])")
MCPEntryCWSub.OnEvent("Click", (*) => CopyExcelColumnToChrome("CWSub"))
MCPEntryFinalOBE := MarkyCopyPasta.Add("Button", "w300", "OBE Exam Marks Entry (w/Breakup) page (Ctrl + \)")
MCPEntryFinalOBE.OnEvent("Click", (*) => CopyExcelColumnToChrome("FinalOBE"))
MarkyCopyPasta.Add("Text", "w300 Center", "")
MCPHelp := MarkyCopyPasta.Add("Button", "w75 x90", "&Help")
MCPHelp.OnEvent("Click", (*) => WelcomeMessage())
MCPQuit := MarkyCopyPasta.Add("Button", "w75 x+0", "&Quit")
MCPQuit.OnEvent("Click", (*) => ExitApp())
MarkyCopyPasta.Show()
MarkyCopyPasta.OnEvent("Close", (*) => ExitApp())
UpdateReadiness()
Return


; ########## FUNCTIONS ##########
WelcomeMessage() {
	welcomemsg := "
	(
		This program copies a vertical column of marks from Excel into the coursework marks entry page in CLiC. How to use:`r`n`r`n
		(1) Open the Excel file with your marks and place the Excel cell cursor at the top of the column of marks (on the first student's mark) you wish to copy. For Exam Marks Entry (w/Breakup), when there are more than 2 components, select the first row of student marks instead (remember to exclude the last column which auto calculates in CLiC).`r`n
		Ensure there is nothing in the cell below the last mark. Students with no marks should be given a zero in Excel. For Exam Marks Entry and OBE Exam Mark Entry(w/Breakup), copy the special grades (W, R, U, I) from CLiC marks entry page into the cell for the student's mark.`r`n
		(2) Open Chrome, log into CLiC and navigate to the relevant marks entry page for your subject. Click on the component required, and place your cursor inside the box for the first student's mark.`r`n
		(3) Press Ctrl+[ for CW/Exam Marks Entry, Press Ctrl+] for CW Marks Sub Component Entry Page or Press Ctrl+\ for OBE Exam Marks Entry page. Do not touch the keyboard while the script runs.`r`n
		Repeat this with as many columns of marks as you need, selecting the correct start of columns in Excel and CLiC respectively. If you experience errors, you can try again without refreshing the page. This program does not click the save or submit buttons. You should be safe from errors, but remember to save your work.`r`n`r`n
		This program can also check if the order of Student IDs in Excel and CLiC matches:`r`n
		(1) Open the Excel file with your marks and place the Excel cell cursor at the top of the column of Student IDs (on the first student's ID).`r`n
		(2) Open Chrome, log into CLiC and navigate to the relevant marks entry page for your subject. Make sure the cursor is not in the input box. (If you just opened the page, you don't have to do anything. Or you can click randomly somewhere on the text in the page.)`r`n
		(3) Press Ctrl+; (semi-colon). Do not touch the keyboard while the script runs.`r`n`r`n
		While this script is running, a white & blue 'S' icon will be in your system tray. Shortcuts shown in buttons works from any program.`r
		This program was built by Willie Poh at Hackerspace MMU's Hackathon No. 23. Version 1.1.
	)"

	MsgBox welcomemsg, "Welcome to MarkyCopyPasta!", "iconi"
}

CopyExcelColumnToChrome(option) {
	WaitNoAltKey()
	Marks := CopyExcelColumn(option)

	if !Marks {
		MsgBox "Failed to copy only marks/grades (numbers, R, W, U, I) from Excel.",, "iconx"
		return false
	}

	PasteColumnInChrome(Marks, option)
}

CheckStudentIDOrder() {
	WaitNoAltKey()
	ExcelStudentIDs := GetStudentIDExcel()

	if !ExcelStudentIDs {
		MsgBox "Failed to copy only Student IDs (numbers) from Excel.",, "iconx"
		return false
	}

	ChromeStudentIDs := GetStudentIDChrome()

	if ChromeStudentIDs.Length < 1 {
		MsgBox "Failed to copy Student IDs from CLiC page.",, "iconx"
		return false
	}

	; MsgBox StrJoin(",", ExcelStudentIds)
	; MsgBox StrJoin(",", ChromeStudentIDs)

	if ExcelStudentIds.Length == ChromeStudentIDs.Length
		longer := false
	else if ExcelStudentIds.Length > ChromeStudentIDs.Length
		longer := "There are more Student IDs in Excel than in CLiC."
	else if ExcelStudentIds.Length < ChromeStudentIDs.Length
		longer := "There are more Student IDs in CLiC than in Excel."

	For index, esid in ExcelStudentIds {
		if ChromeStudentIDs.Has(index) {
			if esid != ChromeStudentIDs[index] {
				if !longer {
					MsgBox "Student ID lists do not match! Mismatch found at position number " index, "Mismatch", "iconx"
					Exit
				}
				else {
					MsgBox "Student ID lists do not match! Mismatch found at position number " index ". " longer, "Mismatch", "iconx"
					Exit
				}
			}
		}
		else {
			MsgBox "Student ID lists do not match! Mismatch found at position number " index ". " longer, "Mismatch", "iconx"
			Exit
		}
	}

	if !longer
		MsgBox "Student ID lists match!", "Match!", "iconi"
	else
		MsgBox "Student ID lists do not match! " longer, "Mismatch", "iconx"
}

CopyExcelColumn(option) {
	if !SwitchToExcelWindow()
		Exit

	Sleep 500

	A_Clipboard := "xyzblah"

	; Copy contents of a cell to clipboard
	Send "{Esc}"
	Sleep 60
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
	Marks.pop ; Remove the last blank element of the array

	if(option=="FinalOBE") {
		AllMarks := []
		for RowMarks in Marks {
			allmark := StrSplit(RowMarks, "`t") ; Last item in this array is a blank
			AllMarks.push(allmark)
		}
		Marks := AllMarks
	}

	; if(option=="FinalOBE") {
	; 	MsgBox Join2D(Marks)
	; }
	; else {
	; 	MsgBox StrJoin(", ", Marks)
	; }

	Send "{Esc}"
	Sleep 100
	Send "{Right}"
	Send "{Left}"
	Sleep 100

	if(option!="FinalOBE") {
		for mark in Marks
			if !IsNumber(mark) and mark != "R" and mark != "W" and mark != "I" and mark != "U"
				return false
	}
	else if(option=="FinalOBE") {
		for RowMarks in Marks {
			for mark in RowMarks {
				if !IsNumber(mark) and mark != "R" and mark != "W" and mark != "I" and mark != "U"
					return false
			}
		}
	}

	return Marks
}

PasteColumnInChrome(Marks, option) {
	WaitNoAltKey()
	if !SwitchToChromeWindow()
		Exit

	Sleep 500

	; Perform check - make sure it's in data entry mode - input fields usually have 0.00 inside to start with
	A_Clipboard := "xyzblah"
	Sleep 60
	Send "^a"
	Sleep 100  
	Send "^c"
	Sleep 250

	; Input fields usually have 0.00 inside to start with
	if !IsNumber(A_Clipboard) and !WinActive("OBE Exam Mark Entry(w/Breakup) - Google Chrome ahk_exe chrome.exe") {
		MsgBox "Cursor is not in input field. Please click / place the cursor into the first marks entry field.",, "iconx"
		Exit
	}
	; Account for OBE Exam Mark Entry(w/Breakup) page which is blank to start with, or may contain numbers (on other than 1st attempt)
	else if (A_Clipboard != "xyzblah") and !IsNumber(A_Clipboard) and WinActive("OBE Exam Mark Entry(w/Breakup) - Google Chrome ahk_exe chrome.exe") {
		MsgBox "Cursor is not in input field. Please click / place the cursor into the first marks entry field.",, "iconx"
		Exit
	}

	if(option!="FinalOBE") {
		For index, Mark in Marks {
			if (index == 1) and (Mark == 0) {
				Send "{Tab}"
				if Marks[2] == 0 {
					Send "{Tab}"
					if Marks[3] == 0 {
						Send "{Tab}"
						if Marks[4] ==0 {
							MsgBox "Detected that your first four marks are ZEROES (0). When there are more than three ZEROES (O) at the top of your marks table, please enter them manually and start your cursor in Excel and CLiC from the first non-zero mark.",, "iconx"
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
				if IsNumber(Mark) {
					Send Mark
					Send "{Tab}"
				}

				; Account for weird CLiC tabindex behavior jumping to page buttons with certain marks input fields
				if (index == 261) and (option=="Others") and ( WinActive("Course Work Marks - Google Chrome ahk_exe chrome.exe") or WinActive("CW Marks Entry - Google Chrome ahk_exe chrome.exe") )
					Send "{Tab 2}"
				else if (index == 191) and (option=="CWSub")
					Send "{Tab 4}"
			}
		}
	}
	else if(option=="FinalOBE") {

		if( Marks[1][1] == 0 ) {
			MsgBox "Detected that your first student's first mark is a ZERO (0). To avoid errors, please enter this manually and start your cursor in Excel and CLiC from the first non-zero mark.",, "iconx"
			Exit
		}

		For index, RowMarks in Marks {
			for index2, Mark in RowMarks {
				if IsNumber(Mark) {
					Send Mark
					Send "{Tab}"
				}
			}
		}
	}

	Sleep 12000
	MsgBox "Finished copying marks from Excel to CLiC! Please wait for the CLiC page to finish 'spinning.' Remember to check marks entered and click 'Save' once confirmed.`r`n`r`nIf you have entered zero marks, when you save or switch columns, you may have to click 'Ok' multiple times. This is normal.`r`n`r`nFor Exam Marks Entry Page, with a large number of students, copying may fail the first time. Please attempt marks copying a second time without refreshing the page. This usually solves the problem.",, "iconi"
}

GetStudentIDExcel() {
	if !SwitchToExcelWindow()
		Exit

	Sleep 500

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
	Sleep 100
	Send "{Right}"
	Send "{Left}"
	Sleep 100

	; Remove the last blank element of the array
	StudentIDs.pop

	for StudentID in StudentIDs
		if !IsNumber(StudentID)
			return false

	return StudentIDs
}

GetStudentIDChrome() {
	WaitNoAltKey()
	if !SwitchToChromeWindow()
		Exit

	; Make sure something new enters the clipboard
	A_Clipboard := "xyzblah"
	Sleep 200
	Send "^a"
	Sleep 100
	Send "^c"
	Sleep 1000
	if A_Clipboard == "xyzblah" {
		MsgBox "Failed to get Student IDs from page.",, "iconx"
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

UpdateReadiness() {
	global ExcelReadyColor
	global ExcelReadyText
	global ChromeReadyColor
	global ChromeReadyText

	excelReady := CheckExcelReadiness()
	if excelReady == 1 {
		ExcelReadyColor.Opt("+BackgroundGreen")
		ExcelReadyColor.Visible := False
		ExcelReadyColor.Visible := True
		ExcelReadyText.Text := "Excel window detected."
	}
	else if excelReady == 2 {
		ExcelReadyColor.Opt("+BackgroundRed")
		ExcelReadyColor.Visible := False
		ExcelReadyColor.Visible := True
		ExcelReadyText.Text := "Too many Excel windows detected."
	}
	else {
		ExcelReadyColor.Opt("+BackgroundRed")
		ExcelReadyColor.Visible := False
		ExcelReadyColor.Visible := True
		ExcelReadyText.Text := "Excel window not detected."
	}

	if CheckChromeReadiness() {
		ChromeReadyColor.Opt("+BackgroundGreen")
		ChromeReadyColor.Visible := False
		ChromeReadyColor.Visible := True
		ChromeReadyText.Text := "A marks entry page is detected."
	}
	else {
		ChromeReadyColor.Opt("+BackgroundRed")
		ChromeReadyColor.Visible := False
		ChromeReadyColor.Visible := True
		ChromeReadyText.Text := "No marks entry pages detected."
	}
}

SwitchToExcelWindow() {
	excelReady := CheckExcelReadiness()
	if !WinActive("ahk_class XLMAIN") {
		if excelReady == 1 {
			WinActivate("ahk_class XLMAIN")
			return true
		}
		else if excelReady == 2 {
			MsgBox "Too many Excel files open. Please open only the Excel file containing your student marks.",, "iconx"
			return false
		}
		else {
			MsgBox "No Excel file currently open. Please open an Excel file containing your student marks.",, "iconx"
			return false
		}
	}
	return true
}

CheckExcelReadiness() { ; Returns 1 if exactly one Excel window found. Else returns 2 for too many, and 0 for too few.
	if WinExist("ahk_class XLMAIN") {
		if CountExcelWindows() == 1 {
			return 1
		}
		else {
			return 2
		}
	}
	else {
		return 0
	}
}

CountExcelWindows() {
	WinHandles := []
	WinList := WinGetList("ahk_class XLMAIN")
	For Each, Win in WinList {
		WinHandles.Push(Win)
	}
	return WinHandles.Length
}

CheckChromeReadiness() {
	if (
		!WinExist("Course Work Marks - Google Chrome ahk_exe chrome.exe") and
		!WinExist("CW Marks Entry - Google Chrome ahk_exe chrome.exe") and
		!WinExist("Sub Component Data Entry - Google Chrome ahk_exe chrome.exe") and 
		!WinExist("Exam Marks Entry - Google Chrome ahk_exe chrome.exe") and
		!WinExist("Exam Marks - Google Chrome ahk_exe chrome.exe") and
		!WinExist("OBE Exam Mark Entry(w/Breakup) - Google Chrome ahk_exe chrome.exe")
		)
		return False
	else
		return True
}

SwitchToChromeWindow() {
	if !CheckChromeReadiness() {
		MsgBox "Your Google Chrome is not opened to a marks entry page. Please open Google Chrome to the correct page and place the cursor on the first value of the column of marks you wish to copy to.",, "iconx"
		return false
	}
	else {
		if WinExist("CW Marks Entry - Google Chrome ahk_exe chrome.exe")
			WinActivate("CW Marks Entry - Google Chrome ahk_exe chrome.exe")
		else if WinExist("Course Work Marks - Google Chrome ahk_exe chrome.exe")
			WinActivate("Course Work Marks - Google Chrome ahk_exe chrome.exe")
		else if WinExist("Sub Component Data Entry - Google Chrome ahk_exe chrome.exe")
			WinActivate("Sub Component Data Entry - Google Chrome ahk_exe chrome.exe")
		else if WinExist("Exam Marks Entry - Google Chrome ahk_exe chrome.exe")
			WinActivate("Exam Marks Entry - Google Chrome ahk_exe chrome.exe")
		else if WinExist("Exam Marks - Google Chrome ahk_exe chrome.exe")
			WinActivate("Exam Marks - Google Chrome ahk_exe chrome.exe")
		else if WinExist("OBE Exam Mark Entry(w/Breakup) - Google Chrome ahk_exe chrome.exe")
			WinActivate("OBE Exam Mark Entry(w/Breakup) - Google Chrome ahk_exe chrome.exe")
	}
	Sleep 60
	return true
}

StrJoin(sep, params) {
	for param in params
		str .= param . sep
	return SubStr(str, 1, -StrLen(sep))
}

Join2D( strArray2D ) {
  s := ""
  for i,array in strArray2D
    s .= ", [" . StrJoin(", ", array) . "]"
  return substr(s, 3)
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
