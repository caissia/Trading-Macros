Attribute VB_Name = "s9_DataTransfer"
Option Explicit

Sub DataTransfer()
's9, extract trade data from the Journal & place it in Data sheet

    Dim StartJournal() As Variant       'start addresses of Setup entries on Journal sheet
    Dim TradeSetups() As Variant        'names of all the trade setups
    Dim firstaddress As String          'captures first address for loop to avoid duplicates
    Dim StartData() As Variant          'start address of Setup trade entries on Data sheet
    Dim StartPos() As Variant           'start position # of each setup in Data sheet
    Dim n_Setup() As Variant            'acronymns of trade setups
    Dim Complete As Boolean             'determines if each entered trade is complete
    Dim j_trade As String               'actual address of trade in Journal sheet to copy
    Dim d_trade As String               'actual address in Data sheet to paste the trade
    Dim i_trade As String               'ID # to be created after trade is completed
    Dim m_trade As String               'max number of possible trades in Journal sheet
    Dim o_date As Date                  'open date of trade
    Dim c_date As Date                  'close date of trade
    Dim test As String
    Dim wsd As String
    Dim wsj As String
    Dim d As Variant
    Dim h As Variant
    Dim m As Variant
    Dim a As Variant
    Dim t As Variant
    Dim r As Variant
    Dim p As Variant
    Dim n As Variant
    Dim i As Variant
    Dim e As Variant
    Dim c As Range

    'check if trade is to be entered
'    i = MsgBox("Transfer trades from Journal to the Data table?", 4 + 32, "Data Transfer")
'        If i <> 6 Then Exit Sub

    'setup worksheets
    wsj = "Journal"
    wsd = "Data"

    'setup variables
    m_trade = Application.Max(Sheets(wsj).Range("A:A"))

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Sheets(wsd).Visible = True

    'resize array to match # of setups
    n = Range("Setups").Rows.Count
    ReDim StartJournal(n)
    ReDim TradeSetups(n)
    ReDim StartData(n)
    ReDim StartPos(n)
    ReDim n_Setup(n)

    'erase the data table before continuuing
    Sheets(wsd).Range("DataTable").ClearContents

    i = 0

    'acquire the trade setup names
    For Each c In Range("Setups")
        If c.Value = "" Then Exit For
        i = i + 1
        TradeSetups(i) = c.Value
    Next c

    i = 0

    'acquire acronyms of trade setups
    For Each c In Range("Setup_ID")
        If c.Value = "" Then Exit For
        i = i + 1
        n_Setup(i) = c.Value
    Next c

    i = 0

    'acquire the first column addresses for each setup in Journal sheet
    With Range("Journal_Header")
        For i = 1 To n
            Set c = .Find(TradeSetups(i), LookIn:=xlValues, SearchOrder:=xlByColumns)
            If Not c Is Nothing Then
                firstaddress = c.Address
                Do
                    StartJournal(i) = c.Offset(2, 1).Address
                    Set c = .FindNext(c)
                Loop While Not c Is Nothing And c.Address <> firstaddress
            End If
        Next i
    End With

    i = 0

    'acquire the start position # of each setup in Data sheet
    For Each c In Range("StartPosition")
        i = i + 1
        StartPos(i) = Left(c.Value, InStr(c.Value, " ") - 1)
    Next c

    i = 0

    'acquire the start address for each setup in Data sheet
    With Range("DataTradeNum")
        For i = 1 To n
            Set c = .Find(StartPos(i), LookIn:=xlValues, SearchOrder:=xlByColumns, Lookat:=xlWhole)
            If Not c Is Nothing Then
                firstaddress = c.Address
                Do
                    StartData(i) = c.Address
                    Set c = .FindNext(c)
                Loop While Not c Is Nothing And c.Address <> firstaddress
            End If
        Next i
    End With

    i = 0

    'find trades in Journal sheet to enter in Data sheet
    For r = 1 To m_trade        'row offset
        For i = 1 To n          'column offset
            t = (r - 1) * 19
            test = Sheets(wsj).Range(StartJournal(i)).Offset(t)
            If test <> "" Then
                Complete = True
                'check if all trade data is present if not skip
                For e = 1 To 17
                    test = Sheets(wsj).Range(StartJournal(i)).Offset(t + e)
                    If test = "" Then Complete = False: Exit For
                Next e
                If Complete Then
                    p = Sheets(wsj).Range(StartJournal(i)).Offset(t, -2) - 1        'captures the number trade in Journal sheet
                    j_trade = Sheets(wsj).Range(StartJournal(i)).Offset(t).Address  'start of trade to copy from Journal sheet
                    d_trade = Range(StartData(i)).Offset(p, 1).Address              'start of trade to paste to Data sheet
                    GoSub Transfer                                                  'actually transfer the data to Data sheet
                    a = a + 1                                                       'keep track of number of trades transferred
                End If
            End If
        Next i
    Next r

    'restore to select cells
    Application.Goto Sheets(wsd).Range("A1"), True
    Sheets(wsd).Visible = xlSheetVeryHidden
    Sheets(wsj).Unprotect
    Application.Goto Sheets(wsj).Range("A1"), True
    Sheets(wsj).Range("L19").Select
    Sheets(wsj).Protect

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    'data transfer from Journal sheet to Data sheet successfully completed
    If a = "" Then
        MsgBox Space(5) & "There are no Journal trades to transfer to the Data table.", 0, "Trade Journal Empty"
    Else
        MsgBox Space(5) & "Succesfully transferred trade data from" & Chr(10) _
             & Space(9) & "the trade Journal to the Data table.", 0, "Data Transfer Complete (" & a & ")"
    End If

Exit Sub

Transfer:

    'create internal ID # for this trade
    t = Sheets(wsd).Range(d_trade).Offset(0, -1)
    If Len(t) = 1 Then t = "000" & t
    If Len(t) = 2 Then t = "00" & t
    If Len(t) = 3 Then t = "0" & t

    t = t & UCase(Left(Sheets(wsj).Range(j_trade).Offset(7), 1))
    t = t & Format(Sheets(wsj).Range(j_trade).Offset(3), "mmddyy")

    i_trade = n_Setup(i) & t    'setup acronymn, 4 digit trade number on Data sheet, date opened

    If Left(Sheets(wsj).Range(j_trade).Offset(5), 3) = "CHF" Then t = "F" Else t = Left(Sheets(wsj).Range(j_trade).Offset(5), 1)
    If Mid(Sheets(wsj).Range(j_trade).Offset(5), 4, 3) = "CHF" Then t = t & "F" Else t = t & Mid(Sheets(wsj).Range(j_trade).Offset(5), 4, 1)

    i_trade = i_trade & t       'setup acronymn, 4 digit trade number on Data sheet, date opened, currency pair shorthand

    t = Sheets(wsj).Range(j_trade).Offset(13)
    If t < 0 Then t = "-"
    If t > 0 Then t = "+"

    i_trade = i_trade & t       'setup acronymn, 4 digit trade number on Data sheet, date opened, currency pair shorthand, plus or minus based on profit

    'calculate the time elapsed between trade open and close
    o_date = Sheets(wsj).Range(j_trade).Offset(3)
    c_date = Sheets(wsj).Range(j_trade).Offset(16)

    t = c_date - o_date

    d = Int(t)
    h = Hour(t)
    m = Minute(t)
    's = Second(t)

    If d > 1 Then d = d & " days "
    If d = 1 Then d = d & " day "
    If d = 0 Then d = ""

    If h > 1 Then h = h & " hours "
    If h = 1 Then h = h & " hour "
    If h = 0 Then h = ""

    If m > 1 Then m = m & " mins "
    If m = 1 Then m = m & " min "
    If m = 0 Then m = ""

    'If s > 1 Then s = s & " secs "
    'If s = 1 Then s = s & " sec "
    'If s = 0 Then s = ""

    'finalized elapsed time: trade open to close
    t = d & h & m '& s

    'copy/paste without the clipboard from Journal sheet to Data sheet values only
    Sheets(wsd).Range(d_trade).Offset(0, 0).Value = i_trade                                        'composed       ~~> ID #
    Sheets(wsd).Range(d_trade).Offset(0, 1).Value = TradeSetups(i)                                 'Setup          ~~> Setup
    Sheets(wsd).Range(d_trade).Offset(0, 2).Value = Sheets(wsj).Range(j_trade).Offset(0).Value     'ID #           ~~> Order #
    Sheets(wsd).Range(d_trade).Offset(0, 3).Value = Sheets(wsj).Range(j_trade).Offset(1).Value     'Adj. Balance   ~~> Adjusted Balance
    Sheets(wsd).Range(d_trade).Offset(0, 4).Value = Sheets(wsj).Range(j_trade).Offset(2).Value     'Escore         ~~> Escore
    Sheets(wsd).Range(d_trade).Offset(0, 5).Value = Sheets(wsj).Range(j_trade).Offset(3).Value     'Date Open      ~~> Date Open
    Sheets(wsd).Range(d_trade).Offset(0, 6).Value = Sheets(wsj).Range(j_trade).Offset(3).Value     'Date Open      ~~> Day Open
    Sheets(wsd).Range(d_trade).Offset(0, 7).Value = Sheets(wsj).Range(j_trade).Offset(3).Value     'Date Open      ~~> T Open
    Sheets(wsd).Range(d_trade).Offset(0, 8).Value = Sheets(wsj).Range(j_trade).Offset(4).Value     'Time Frame     ~~> Time Frame
    Sheets(wsd).Range(d_trade).Offset(0, 9).Value = Sheets(wsj).Range(j_trade).Offset(5).Value     'Currency Pair  ~~> Pair
    Sheets(wsd).Range(d_trade).Offset(0, 10).Value = Sheets(wsj).Range(j_trade).Offset(6).Value    'Lots placed    ~~> Lots
    Sheets(wsd).Range(d_trade).Offset(0, 11).Value = Sheets(wsj).Range(j_trade).Offset(7).Value    'Direction      ~~> Direction
    Sheets(wsd).Range(d_trade).Offset(0, 12).Value = Sheets(wsj).Range(j_trade).Offset(8).Value    'Entry          ~~> Entry
    Sheets(wsd).Range(d_trade).Offset(0, 13).Value = Sheets(wsj).Range(j_trade).Offset(9).Value    'Exit           ~~> Exit
    Sheets(wsd).Range(d_trade).Offset(0, 14).Value = Sheets(wsj).Range(j_trade).Offset(10).Value   'Stoploss       ~~> S/L
    Sheets(wsd).Range(d_trade).Offset(0, 15).Value = Sheets(wsj).Range(j_trade).Offset(11).Value   'Target         ~~> TP
    Sheets(wsd).Range(d_trade).Offset(0, 16).Value = Sheets(wsj).Range(j_trade).Offset(12).Value   'Pips Earned    ~~> Pips
    Sheets(wsd).Range(d_trade).Offset(0, 17).Value = Sheets(wsj).Range(j_trade).Offset(13).Value   'Profit         ~~> Profit
    Sheets(wsd).Range(d_trade).Offset(0, 18).Value = Sheets(wsj).Range(j_trade).Offset(14).Value   'Swap Charge    ~~> Swap
    Sheets(wsd).Range(d_trade).Offset(0, 19).Value = Sheets(wsj).Range(j_trade).Offset(15).Value   'Commission     ~~> Comm
    Sheets(wsd).Range(d_trade).Offset(0, 20).Value = Sheets(wsj).Range(j_trade).Offset(16).Value   'Date Close     ~~> Date Close
    Sheets(wsd).Range(d_trade).Offset(0, 21).Value = Sheets(wsj).Range(j_trade).Offset(16).Value   'Date Close     ~~> Day Close
    Sheets(wsd).Range(d_trade).Offset(0, 22).Value = Sheets(wsj).Range(j_trade).Offset(16).Value   'Date Close     ~~> T Close
    Sheets(wsd).Range(d_trade).Offset(0, 23).Value = Sheets(wsj).Range(j_trade).Offset(17).Value   'Pscore         ~~> Pscore
    Sheets(wsd).Range(d_trade).Offset(0, 24).Value = t                                             'calculated     ~~> Total Time

Return

End Sub
