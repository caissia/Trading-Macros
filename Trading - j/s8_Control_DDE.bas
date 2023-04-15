Attribute VB_Name = "s8_Control_DDE"
Option Explicit
Public DDE As Boolean
Public MT4 As Boolean
'the following needed for ImageCapture
Private Const WM_CLOSE As Long = &H10
Private Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'contains 6 macros: Control_DDE, Toggle_DDE,Toggle_MT4, Connect_DDE, Pause_DDE, ImageCapture

Sub Control_DDE()
'pauses/connects DDE if market open

    If Not IsProgramRunning("terminal.exe") Then
        Shell ("C:\Program Files (x86)\FXCM MetaTrader 4\terminal.exe")
        Application.Wait Now + TimeSerial(0, 0, 4)
    End If

    Sheets("Range").EnableCalculation = True
    Call Connect_DDE

    Dim timenow As Date
    Dim timecls As Date
    Dim timeopn As Date

    timenow = TimeSerial(Hour(Now), Minute(Now), Second(Now))
    timecls = TimeSerial(13, 50, 0)
    timeopn = TimeSerial(14, 20, 0)

    If Weekday(Now()) = 6 And timenow < timecls Then

        Application.OnTime TimeSerial(13, 50, 0), "Pause_DDE"

    ElseIf Weekday(Now()) = 7 Or _
           Weekday(Now()) = 6 And timenow > timecls Or _
           Weekday(Now()) = 1 And timenow < timeopn Then

        Call Pause_DDE

    End If

End Sub

Sub Toggle_DDE()

    If DDE Then
        Sheets("Range").EnableCalculation = DDE
        Sheets("Range").Range("G23") = "Yes"
        DDE = Not DDE
    Else
        Sheets("Range").EnableCalculation = DDE
        Sheets("Range").Range("G23") = "No"
        DDE = Not DDE
    End If

End Sub

Sub Toggle_MT4()

    If MT4 Then

        If Not IsProgramRunning("terminal.exe") Then
            Shell ("C:\Program Files (x86)\FXCM MetaTrader 4\terminal.exe")
            Application.Wait Now + TimeSerial(0, 0, 4)
        End If

        'refocus this workbook ;can add 'DoEvents' after appactivate
        AppActivate Application.Caption

        Sheets("Range").EnableCalculation = True
        Call Connect_DDE
        MT4 = Not MT4

    Else

        If IsProgramRunning("terminal.exe") Then
            TerminateProcess ("terminal.exe")
        End If

        Sheets("Range").Range("G23") = "No"
        MT4 = Not MT4

    End If

End Sub

Sub Connect_DDE()
'this re-connects the DDE link feed to MT4

    Dim iCell As String
    Dim iSheet As String

    iCell = ActiveCell.Address
    iSheet = ActiveSheet.Name

    Sheets("Range").EnableCalculation = True
    Application.ScreenUpdating = False
    Sheets("Range").Visible = True
    Sheets("Range").Activate

    Sheets("Range").Range("L23").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(IF('MT4'|ASK!EURAUD=""N\A"","""",'MT4'|ASK!EURAUD),"""")"

    Sheets("Range").Range("M23").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(IF('MT4'|BID!EURAUD=""N\A"","""",'MT4'|BID!EURAUD),"""")"

    Sheets("Range").Range("G23") = "Yes"

    Sheets("Range").Range("I21").Select
    Sheets("Range").Visible = xlSheetVeryHidden

    Sheets(iSheet).Activate
    Sheets(iSheet).Range(iCell).Select

    Application.ScreenUpdating = True

End Sub

Sub Pause_DDE()

    Sheets("Range").Range("G23") = "No"
    Sheets("Range").EnableCalculation = False

End Sub

Sub ImageCapture()
'starts screenhunter app

    Dim ImageAppWin As String

    If Not IsProgramRunning("ScreenHunter.exe") Then
        Shell ("C:\Program Files (x86)\Wisdom-soft ScreenHunter\ScreenHunter.exe")  'open screenhunter
        AppActivate Application.Caption                                             'refocus on journal
        Application.Wait Now + TimeSerial(0, 0, 1)
    End If

    If IsProgramRunning("ScreenHunter.exe") Then                                    'wait until app can be found
        Do
        DoEvents                                                                    'gives vba time to process w/o freezing
        ImageAppWin = FindWindow(vbNullString, " ScreenHunter ")                    'acquire screenhunter's window handle number
        AppActivate Application.Caption
        Loop Until ImageAppWin > 0
    End If

    AppActivate Application.Caption
    Call SendMessage(ImageAppWin, WM_CLOSE, 0, 0)                                   'close but not exit screenhunter's initial window

End Sub
