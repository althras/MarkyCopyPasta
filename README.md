# MarkyCopyPasta
Shortcuts (works from anywhere until you quit):
* Ctrl + [ - Copy for CW/Exam/OBE Exam Marks Entry pages
* Ctrl + ] - Copy for CW Marks Sub Component Entry page
* Ctrl + ; - Compare Student ID in Excel and CaMSys
* Alt + q - Quit the program
* Alt + h - See this help message

While this script is running, a white & blue 'S' icon will be in your system tray. This program copies a vertical column of marks from Excel into the coursework marks entry page in CaMSys.

How to use:
1. Open the Excel file containing your marks and place the Excel cell cursor at the top of the column of marks (on the first student's mark) you wish to copy. Make sure there is nothing in the cell below the last mark. Any students with no marks should be given a zero. For Exam Marks Entry and OBE Exam Mark Entry(w/Breakup), copy the special grades (W, R, U, I) from CaMSys marks entry page into the cell for the student's mark.
2. Open Chrome, log into CaMSys and navigate to the relevant marks entry page for your subject. Click on the component required, and place your cursor inside the box for the first student's mark.
3. Press Ctrl+[ for CW/Exam/OBE Exam Marks Entry or Press Ctrl+] for CW Marks Sub Component Entry Page. Do not touch the keyboard while the script runs.
Repeat this with as many columns of marks as you need, selecting the correct start of columns in Excel and CaMSys respectively. If you experience errors, you can try again without refreshing the page. This program does not click the save or submit buttons. You should be safe from errors, but remember to save your work.

This program can also check if the order of Student IDs in Excel and CaMSys matches:
1. Open the Excel file containing your marks and place the Excel cell cursor at the top of the column of Student IDs (on the first student's ID).
2. Open Chrome, log into CaMSys and navigate to the relevant marks entry page for your subject. Make sure the cursor is not in the input box. (If you just opened the page, you don't have to do anything. Or you can click randomly somewhere on the text in the page.)
3. Press Ctrl+; (semi-colon). Do not touch the keyboard while the script runs.

Press Alt+Q to Quit the script, or right-click the 'H' icon in your system tray and click Exit.

This program was built by Willie Poh at Hackerspace MMU's Hackathon No. 23. Version 0.3.2 (Beta Test)
