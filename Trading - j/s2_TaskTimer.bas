Attribute VB_Name = "s2_TaskTimer"
Option Explicit
Public dF, rE, tC
Public Blue, Gray, Red, Tan
Public cF1, cF2, cF3, cF4, cF5, cF6, cF7
Public TT1, TT2, TT3, TT4, TT5, TT6, TT7, TT8
'contains 11 macros: SelectTask, TimeTask1 through TimeTask8, TimeDiff, ClearRoutine

Sub SelectTask()
's2, times each step in the trade routine

    Select Case ActiveCell.Row - 5

        Case 1: Call TimeTask1
        Case 2: Call TimeTask2
        Case 3: Call TimeTask3
        Case 4: Call TimeTask4
        Case 5: Call TimeTask5
        Case 6: Call TimeTask6
        Case 7: Call TimeTask7
        Case 8: Call TimeTask8

    End Select

End Sub

Sub TimeTask1()

    dF = ""

    Tan = RGB(221, 217, 196)
    Blue = RGB(0, 112, 192)
    Gray = RGB(77, 77, 77)
    Red = RGB(156, 0, 0)

    TT1 = datetime.Now

    cF1 = 1
    cF2 = 0
    cF3 = 0
    cF4 = 0
    cF5 = 0
    cF6 = 0
    cF7 = 0

    tC = ActiveCell.Address

    Range(tC).Offset(0, 0).Font.Color = Blue
    Range(tC).Offset(0, 1).Font.Color = Blue

    Range(Range(tC).Offset(1, 0).Address & ":" & Range(tC).Offset(7, 1).Address).Font.Color = Gray
    Range(tC).Offset(8, 0).Font.Color = Gray

    Application.ScreenUpdating = False
    Range("RoutineData").ClearContents
    Application.ScreenUpdating = True

    Range(tC).Offset(-3, 0) = "Market Routine"
    Range(tC).Offset(1, 0).Select

End Sub

Sub TimeTask2()

    dF = ""

    If cF1 <> 1 Or cF2 = 2 Then Exit Sub

    cF2 = 2

    TT2 = datetime.Now

    dF = TT2 - TT1

    Call TimeDiff

    tC = ActiveCell.Address

    Range(tC).Offset(-1, 0).Font.Color = Gray
    Range(tC).Offset(-1, 1).Font.Color = Gray
    Range(tC).Offset(0, 0).Font.Color = Blue
    Range(tC).Offset(0, 1).Font.Color = Blue

    Range(tC).Offset(-1, 6) = rE

    Range(tC).Offset(1, 0).Select

End Sub


Sub TimeTask3()

    dF = ""

    If cF2 <> 2 Or cF3 = 3 Then Exit Sub

    cF3 = 3

    TT3 = datetime.Now

    dF = TT3 - TT2

    Call TimeDiff

    tC = ActiveCell.Address

    Range(tC).Offset(-1, 0).Font.Color = Gray
    Range(tC).Offset(-1, 1).Font.Color = Gray
    Range(tC).Offset(0, 0).Font.Color = Blue
    Range(tC).Offset(0, 1).Font.Color = Blue

    Range(tC).Offset(-1, 6) = rE

    Range(tC).Offset(1, 0).Select

End Sub


Sub TimeTask4()

    dF = ""

    If cF3 <> 3 Or cF4 = 4 Then Exit Sub

    cF4 = 4

    TT4 = datetime.Now

    dF = TT4 - TT3

    Call TimeDiff

    tC = ActiveCell.Address

    Range(tC).Offset(-1, 0).Font.Color = Gray
    Range(tC).Offset(-1, 1).Font.Color = Gray
    Range(tC).Offset(0, 0).Font.Color = Blue
    Range(tC).Offset(0, 1).Font.Color = Blue

    Range(tC).Offset(-1, 6) = rE

    Range(tC).Offset(1, 0).Select

End Sub


Sub TimeTask5()

    dF = ""

    If cF4 <> 4 Or cF5 = 5 Then Exit Sub

    cF5 = 5

    TT5 = datetime.Now

    dF = TT5 - TT4

    Call TimeDiff

    tC = ActiveCell.Address

    Range(tC).Offset(-1, 0).Font.Color = Gray
    Range(tC).Offset(-1, 1).Font.Color = Gray
    Range(tC).Offset(0, 0).Font.Color = Blue
    Range(tC).Offset(0, 1).Font.Color = Blue

    Range(tC).Offset(-1, 6) = rE

    Range(tC).Offset(1, 0).Select

End Sub


Sub TimeTask6()

    dF = ""

    If cF5 <> 5 Or cF6 = 6 Then Exit Sub

    cF6 = 6

    TT6 = datetime.Now

    dF = TT6 - TT5

    Call TimeDiff

    tC = ActiveCell.Address

    Range(tC).Offset(-1, 0).Font.Color = Gray
    Range(tC).Offset(-1, 1).Font.Color = Gray
    Range(tC).Offset(0, 0).Font.Color = Blue
    Range(tC).Offset(0, 1).Font.Color = Blue

    Range(tC).Offset(-1, 6) = rE

    Range(tC).Offset(1, 0).Select

End Sub

Sub TimeTask7()

    dF = ""

    If cF6 <> 6 Or cF7 = 7 Then Exit Sub

    cF7 = 7

    TT7 = datetime.Now

    dF = TT7 - TT6

    Call TimeDiff

    tC = ActiveCell.Address

    Range(tC).Offset(-1, 0).Font.Color = Gray
    Range(tC).Offset(-1, 1).Font.Color = Gray
    Range(tC).Offset(0, 0).Font.Color = Blue
    Range(tC).Offset(0, 1).Font.Color = Blue

    Range(tC).Offset(-1, 6) = rE

    Range(tC).Offset(1, 0).Select

End Sub

Sub TimeTask8()

    dF = ""

    If cF7 <> 7 Then Exit Sub

    TT8 = datetime.Now

    dF = TT8 - TT7

    Call TimeDiff

    tC = ActiveCell.Address

    Range(tC).Offset(-1, 0).Font.Color = Gray
    Range(tC).Offset(-1, 1).Font.Color = Gray
    Range(tC).Offset(0, 0).Font.Color = Blue
    Range(tC).Offset(0, 1).Font.Color = Blue
    Range(tC).Offset(1, 0).Font.Color = Red

    Range(tC).Offset(-1, 6) = rE

    Range(tC).Offset(1, 0).Select

    'calculate total time
    dF = TT8 - TT1
    Call TimeDiff

    Range(tC).Offset(-10, 0) = "Market Routine - " & rE

End Sub

Sub TimeDiff()

    Dim d As Variant
    Dim h As Variant
    Dim m As Variant
    Dim s As Variant

    d = Int(dF)
    h = Hour(dF)
    m = Minute(dF)
    s = Second(dF)
    rE = ""

    If d > 1 Then d = d & " days  "
    If d = 1 Then d = d & " day  "
    If d = 0 Then d = ""

    If h > 1 Then h = h & " hours  "
    If h = 1 Then h = h & " hour  "
    If h = 0 Then h = ""

    If m > 1 Then m = m & " mins  "
    If m = 1 Then m = m & " min  "
    If m = 0 Then m = ""

    If s > 1 Then s = s & " secs  "
    If s = 1 Then s = s & " sec  "
    If s = 0 Then s = ""

    rE = d & h & m & s
    If rE = "" Then rE = "0 secs  "

End Sub

Sub ClearRoutine()
'resets tasktimers and market routine

    cF1 = 0
    cF2 = 0
    cF3 = 0
    cF4 = 0
    cF5 = 0
    cF6 = 0
    cF7 = 0

    tC = ActiveCell.Address

    Range("RoutineData").ClearContents

    Range(tC).Offset(-11, 0) = "Market Routine"

    Range(Range(tC).Offset(-8, 0).Address & ":" & Range(tC).Offset(-1, 1).Address).Font.Color = Gray
    Range(tC).Font.Color = Tan

    Range(tC).Offset(-8, 0).Select

End Sub
