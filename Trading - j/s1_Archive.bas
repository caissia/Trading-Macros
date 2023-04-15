Attribute VB_Name = "s1_Archive"
Option Explicit
'contains 3 macros: Archive, Journal_New, Journal Switch

Sub Journal_Switch()
'ask to archive/create new journal

    Dim i
    Dim NewYr As Boolean
    NewYr = Year(Date) <> Year(Sheets("Range").Range("YearStart"))

    'if called from journal sheet
    If ActiveSheet.Name = "Journal" And tC = 4 Then

        i = MsgBox(Space(3) & "Archive the present journal or Create a new journal?" & Chr(10) & Chr(10) _
                & Space(24) & "[ archive = yes  |  create = no ]", 3, "Archive or Create")

        If i = 6 Then Call Archive
        If i = 7 Then Call Journal_New
        Exit Sub

    End If

    'first few days of new year check if new journal is required
    If Month(Now) = 1 And Day(Now) <= 5 And NewYr Then

        MsgBox "Remember to start a new trading journal every year." & Chr(10) _
                & "You should find the Investment journal under the trade folder.", 64, _
                "Remember"
 
 
        i = MsgBox(Space(5) & "Archive the present journal or Create a new journal?" & Chr(10) & Chr(10) _
                & Space(23) & "[ archive = yes  |  create = no ]", 3, "Archive or Create")

        If i = 6 Then Call Archive
        If i = 7 Then Call Journal_New

    End If

    If Month(Now) = 1 And Day(Now) <= 5 Then

        'check if journal is a template, if so ask to create a journal
        If Sheets("Range").Range("D23") = "Yes" Then Call Journal_New
    
        'check if current year matches the journal year, if not ask to update
        If CStr(Year(Now)) <> Right(Sheets("Range").Range("C21"), 4) Then
    
            i = MsgBox("The current year, " & Year(Now) & ", does not match" & Chr(10) _
                    & "     the year for this journal, " & Right(Sheets("Range").Range("C21"), 4) & "." & Chr(10) _
                    & "Would you like to update the year for this journal?", 36, _
                    "Update Year")
    
            If i = 6 Then
                Sheets("Range").Unprotect
                Sheets("Range").Range("C21") = "1/1/" & Year(Now)
                Sheets("Range").Range("G21") = "12/31/" & Year(Now)
                MsgBox "  The yearly range for this journal has been updated.", 0, "Current Year Update"
            End If
    
            'alert user where to change the date
            MsgBox Space(5) & "To manually change the dates for " & Chr(10) & "this journal go to Range C21 & G21.", 0, "Date Location"
    
        End If

    End If

End Sub

Sub Archive()
'save journal to archive it & start new journal year

    Dim i As Variant
    Dim yr As Variant
    Dim ext As String
    Dim cDir As String
    Dim nJrnl As String
    Dim svJrnl As String

    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    i = MsgBox("       Would you like to archive the present trade journal?" & Chr(10) & Chr(10) _
                & "     Once you specify the year the following will be created:" & Chr(10) _
                & "       - A new folder named after the selected year" & Chr(10) _
                & "       - This journal named after itself & selected year." & Chr(10) & Chr(10) _
                & "After this journal is archived it can be restored to start afresh.", 36, "Archive Journal?")
    If i <> 6 Then GoTo ex
                
retry:
        yr = MsgBox("             Save as this year or previous year?" & Chr(10) _
                      & "          This Year = ""Yes"", Previous Year = ""No""", 3, "Year?")
        If yr = 2 Then GoTo ex Else yr = yr - 6

        i = InStr(ThisWorkbook.Name, ".")
        yr = Format(Date, "yyyy") - yr
        ext = Right(ThisWorkbook.Name, i)
        cDir = Sheets("Range").Range("I16").Value & "\" & yr
        nJrnl = Left(ThisWorkbook.Name, i - 1) & " " & yr & ext
        svJrnl = cDir & "\" & nJrnl

        'check if new folder exists if not create new folder
        On Error GoTo err1
        If Len(Dir(cDir, vbDirectory)) = 0 Then
            MkDir cDir
        Else
err1:       i = MsgBox("The folder already exists or the following path does not exist." & Chr(10) & Chr(10) _
                    & "      " & cDir & Chr(10) & Chr(10) _
                    & "Do you want to skip creating a folder (yes) or retry (no)?", 35, "Create Folder Skip?")
            If i = 7 Then GoTo retry
            If i = 2 Then GoTo ex
        End If
        On Error GoTo 0

        'check if file exists if not save as
        On Error GoTo err2
        If Len(Dir(svJrnl, vbDirectory)) = 0 Then
            ThisWorkbook.SaveAs filename:=svJrnl
        Else
err2:       i = MsgBox("This journal already exists or the following path does not exist." & Chr(10) & Chr(10) _
                & " Journal:      " & nJrnl & Chr(10) _
                & " Directory:   " & cDir & Chr(10) & Chr(10) _
                & "Would you like to overwrite the existing file (yes) or retry (no)?", 35, "Overwrite File?")
            If i = 7 Then GoTo retry
            If i = 2 Then GoTo ex
            ThisWorkbook.SaveCopyAs svJrnl      'finally save workbook as archived file 'Workbooks.Open filename:=SavedWB 'in case file must be open
        End If

        'inform the user the new journal was created
        MsgBox "The current journal has been archived: " & Chr(10) & Chr(10) _
                & " Journal:      " & nJrnl & Chr(10) _
                & " Directory:   " & cDir, 64, "Success: Journal Archived"

        Application.Goto Sheets("Journal").Range("A1"), True
        Application.Calculation = xlAutomatic
        Application.ScreenUpdating = True

        'call macro to refresh journal, delete data
        Call Reset

        Exit Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ex:     MsgBox "Canceled " & "'""save as'' for this journal.", 64, "Canceled: Archive Journal"

        Application.Goto Sheets("Journal").Range("A1"), True
        Application.Calculation = xlAutomatic
        Application.ScreenUpdating = True

End Sub

Sub Journal_New()
'save as a new journal

    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    Dim aa As Byte   'answers to various inputboxes
    Dim pN As String 'partial name of new Journal
    Dim cF As String 'current folder location of old Journal
    Dim vN As String 'new name and directory of new journal
    Dim nN As String 'new name of new Journal
    Dim sN As String 'sample filename for new Journal


    pN = Sheets("Range").Range("C21").Text
    cF = Sheets("Range").Range("I16").Text
    vN = cF & "\" & "Trade Journal " & Year(pN) & ".xlsm"
    sN = "Trade Journal " & Year(pN)

    Windows(JournalTitle).Activate
    Application.Goto Sheets("Journal").Range("A1"), True
    Sheets("Journal").Unprotect
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        aa = MsgBox("Do you want to save this template or journal " & Chr(10) _
                    & "              as a new trade journal?", 4 + 32, "New Journal")
            If aa = 7 Then
                GoTo ex
            End If

        MsgBox "There are a several steps to creating a new Journal." & Chr(10) _
             & "  Step 1: Set directory " & Chr(10) _
             & "  Step 2: Set new filename" & Chr(10) _
             & "  Step 3: Set year covered" & Chr(10) _
             & "  Step 4: Set broker" & Chr(10) _
             & "  Step 5: Clear trade data (optional)", 64, "Steps: New Journal"

        aa = MsgBox("Do you want to create a new folder for the new Journal?", 4 + 32, "Step One: Create New Folder")

        If aa = 7 Then
            MsgBox "The default save location is" & Chr(10) & cF, , "Remember"
            GoTo nF
        End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

nD:     cF = InputBox("Please enter the directory and folder you would like to create:", _
                "Enter New Directory", cF)

        If cF = "" Then
            aa = MsgBox("Directory is missing or is the same as the example provided." & Chr(10) _
                    & "Try again?", 4 + 32, "Error")
                If aa = 6 Then
                    GoTo nD
                Else
                    GoTo fE
                End If
        End If

        On Error GoTo ER1
        'check if new folder exists 'create new folder
        If Len(Dir(cF, vbDirectory)) = 0 Then
            MkDir cF
        Else
ER1:        aa = MsgBox("The following folder already exists or the path does not exist:" & Chr(10) & cF _
                & Chr(10) & "Would you like to specify a new directory?", 4 + 32, "ERROR")
            If aa = 6 Then
                GoTo nD
            Else
fE:             aa = MsgBox("Are you sure you want to cancel setting up a new folder?", 4 + 48, "Cancel?")
                If aa = 7 Then GoTo nD
                MsgBox "Creating a new folder has been canceled.", 0 + 64, "New Journal Canceled"
                GoTo nF
            End If
        End If

        'Inform the user the new folder was created
        MsgBox "A new folder in the following directory has been created:" & Chr(10) _
                & cF, 64, "New Journal Success"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

nF:     nN = InputBox("Please enter a filename for the new Journal:", _
                "Step Two: New Filename", sN)

        If nN = sN Then
            aa = MsgBox("Filename is missing or is the same as the example provided." & Chr(10) _
                & "Try again?", 4 + 32, "Missing Filename")
            If aa = 6 Then
                GoTo nF
            Else
                GoTo nE
            End If
        End If

        If nN = "" Then GoTo nE

        'create string of directory plus new name of Journal
        vN = cF & "\" & nN & ".xlsm"

        On Error GoTo ER2
        'check if file exists 'saves the file with a new name
        If Len(Dir(nN, vbDirectory)) = 0 Then
            ThisWorkbook.SaveAs filename:=vN
        Else

ER2:    aa = MsgBox("The filename already exists or there is a problem" & Chr(10) _
                & "with the name or directory." & Chr(10) & Chr(10) & nN _
                & Chr(10) & Chr(10) & "Would you like to specify a new file name?", 4 + 32, "ERROR")
            If aa = 6 Then
                GoTo nF
            Else
nE:             aa = MsgBox("Are you sure you want to cancel creating a new Journal?", 4 + 48, "Cancel?")
                If aa = 7 Then GoTo nF
                GoTo ex
            End If
        End If

        'Inform the user the new Journal was created
        MsgBox "A new workbook named ''" & nN & "'' has been created.", 64, "New Journal Success"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Sheets("Range").Range("D23") = "No"
        Sheets("Range").Range("I16") = cF

        Call Set_Year
        Call Set_Broker
        Call Reset

        Application.Goto Sheets("Journal").Range("A1"), True
        Application.Calculation = xlAutomatic
        Application.ScreenUpdating = True

        Exit Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ex:     MsgBox "Creating a new Journal has been canceled.", 0 + 64, "New Journal Canceled"

        Application.Goto Sheets("Journal").Range("A1"), True
        Application.Calculation = xlAutomatic
        Application.ScreenUpdating = True

End Sub


