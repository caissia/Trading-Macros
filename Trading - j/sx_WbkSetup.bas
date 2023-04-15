Attribute VB_Name = "sx_WbkSetup"
Option Explicit
Public JournalTitle As String
Public JournalCaption As String

Sub SetupWorkbook()

    Dim s

    'avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'title of this journal
    JournalTitle = ThisWorkbook.Name

    'name caption of this journal
    JournalCaption = "Caissia"

    'change the excel caption
    Application.Caption = JournalCaption

    'replace excel icon on caption bar with wave icon
    Open_SetIcon "C:\Users\image\Documents\trade\fx\miscellanea\images\Queen_Transparent.ico", 0

    For s = 10 To 1 Step -1
        Sheets(s).Activate
        Sheets(s).Unprotect
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayGridlines = False
        Application.Goto Sheets(s).Range("A1"), True
        If Sheets(s).Name = "Note" Then ActiveSheet.ScrollArea = "A1:J41"
        If Sheets(s).Name = "-Note" Then ActiveSheet.ScrollArea = "A1:F11"
        If Sheets(s).Name = "Note-" Then ActiveSheet.ScrollArea = "A1:M14"
        If Sheets(s).Name = "-Note-" Then ActiveSheet.ScrollArea = "A1:X151"
        If Sheets(s).Name = "Rank" Then Range("L5").Select: ActiveSheet.ScrollArea = "A1:Y95"
        If Sheets(s).Name = "Range" Then Range("I21").Select: ActiveSheet.ScrollArea = "A1:V47"
        If Sheets(s).Name = "System" Then Range("I19").Select: ActiveSheet.ScrollArea = "A1:V29"
        If Sheets(s).Name = "Data" Then Range("C5").Select: ActiveSheet.ScrollArea = "A1:AA2058"
        If Sheets(s).Name = "Query" Then Range("C5").Select: ActiveSheet.ScrollArea = "A1:AA2011"
        If Sheets(s).Name = "Journal" Then Range("L19").Select: ActiveSheet.ScrollArea = "A1:ET1919"
        Sheets(s).Protect , UserInterfaceOnly:=True
        If Sheets(s).Name = "Range" Then Sheets(s).Unprotect
    Next

    'hide sheets only access is through vba
    Sheets("Range").Visible = xlSheetVeryHidden
    Sheets("Data").Visible = xlSheetVeryHidden

    Application.DisplayStatusBar = False
    Application.DisplayFormulaBar = False
    Application.WindowState = xlMaximized
    Application.Calculation = xlAutomatic
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"

    ActiveWindow.WindowState = xlMaximized
'    ActiveWindow.DisplayVerticalScrollBar = False
'    ActiveWindow.DisplayHorizontalScrollBar = False

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub FinalWorkbook()

    'avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'reset to standard excel caption
    Application.Caption = ""

    'close MT4 if open
    If IsProgramRunning("terminal.exe") Then
        TerminateProcess ("terminal.exe")
    End If

    'close ScreenHunter, screen capture app, if open
    If IsProgramRunning("ScreenHunter.exe") Then
        TerminateProcess ("ScreenHunter.exe")
    End If

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
