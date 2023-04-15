Attribute VB_Name = "s9_OptTransfer"
Option Explicit

Sub OptTransfer()
's9, extract the optimal trade data from Journal & place it in Data sheet

    Dim OptTrades() As Variant
    Dim sCell As String

    Dim r As Variant
    Dim n As Variant
    Dim c As Variant
    Dim i As Variant
    Dim s As Variant
    Dim t As Variant

    'check if trade is to be entered
'    c = MsgBox("Transfer Optimal Trades from Journal to the Data table?", 4 + 32, "Data Transfer")
'        If c <> 6 Then Exit Sub

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Sheets("Data").Visible = True

    'first cell in Data sheet
    sCell = "C2007"

    'acquire all the optimal trades thus far
    OptTrades = Range("Journal_OptData")

    'clear contents of existing Data table of optimal wkly trades
    Range("Data_Opt_Table").ClearContents

    'transfer the optimal trade data to the Data sheet
    For r = 1 To UBound(OptTrades, 1)                       ' first array dimension is rows
        For c = 1 To UBound(OptTrades, 2)                   ' second array dimension is columns
            i = i + 1
            n = (i - 1) Mod 19                              ' identifies start of next trade
            If OptTrades(r, c) <> "" And n = 0 Then
                t = t + 1
                Sheets("Data").Range(sCell).Offset(t - 1, 0) = OptTrades(r + 0, c)      ' setup
                Sheets("Data").Range(sCell).Offset(t - 1, 1) = OptTrades(r + 1, c)      ' currency pair
                Sheets("Data").Range(sCell).Offset(t - 1, 2) = OptTrades(r + 2, c)      ' time frame
                Sheets("Data").Range(sCell).Offset(t - 1, 4) = OptTrades(r + 3, c)      ' week day
                Sheets("Data").Range(sCell).Offset(t - 1, 6) = OptTrades(r + 4, c)      ' time
                Sheets("Data").Range(sCell).Offset(t - 1, 8) = OptTrades(r + 5, c)      ' pip
                Sheets("Data").Range(sCell).Offset(t - 1, 13) = OptTrades(r + 6, c)     ' date
            End If
        Next c
    Next r

    'restore journal settings
    Application.Goto Sheets("Data").Range("A1"), True
    Sheets("Data").Visible = xlSheetVeryHidden
    Sheets("Journal").Unprotect
    Application.Goto Sheets("Journal").Range("A1"), True
    Sheets("Journal").Range("L19").Select
    Sheets("Journal").Protect

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    'data transfer from Journal sheet to Data sheet successfully completed
    n = Int(((Now - DateSerial(Year(Now), 1, 0)) + 6) / 7)
    If t = "" Then
        MsgBox Space(5) & "There are no Optimal trades to transfer to the Data table.", 0, "Optimal Trade Section Empty"

    ElseIf t <> n Then
        MsgBox Space(5) & "Succesfully transferred Optimal Trades from" & Chr(10) _
             & Space(12) & "the trade Journal to the Data table." & Chr(10) & Chr(10) _
             & Space(8) & "However, there are some trades missing." & Chr(10) _
             & Space(24) & "Optimal Trades:" & Space(8) & t & Chr(10) _
             & Space(24) & "Number of weeks:" & Space(4) & n, 0, "Optimal Transfer Complete (" & t & ")"
    Else
        MsgBox Space(5) & "Succesfully transferred Optimal Trades from" & Chr(10) _
             & Space(13) & "the trade Journal to the Data table.", 0, "Optimal Transfer Complete (" & t & ")"
    End If

End Sub
