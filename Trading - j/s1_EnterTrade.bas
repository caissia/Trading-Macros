Attribute VB_Name = "s1_EnterTrade"
Option Explicit

Sub Trade_Entry()
Attribute Trade_Entry.VB_Description = "Paste trade in appropriate setup column"
Attribute Trade_Entry.VB_ProcData.VB_Invoke_Func = "T\n14"

    Dim a As Variant                                                    'temp variable
    Dim b As Variant                                                    'temp variable
    Dim c As Variant                                                    'loop variable
    Dim d As Variant                                                    'misc variable
    Dim i As Variant                                                    'misc variable
    Dim m As Integer                                                    'number of trades in a setup
    Dim n As Integer                                                    'number of setups in use

    Dim wsd As String                                                   'worksheet detailedstatement
    Dim wsj As String                                                   'worksheet journal
    Dim wsr As String                                                   'worksheet range
    Dim eQ As Workbook                                                  'workbook investment journal
    Dim sQ As Workbook                                                  'workbook detailedstatement

    Dim Ticket As String                                                'order#/Ticket#/ID# of trade
    Dim Balance As Variant                                              'starting balance before trade was placed
    Dim eScore As Integer                                               'self assessment of executing the trade
    Dim DateOpen As String                                              'date/time trade was opened
    Dim TimeFrame As String                                             'time frame of trade, i.e. daily
    Dim Direction As String                                             'direction: long or short
    Dim Lots As Double                                                  'lots controlled
    Dim CurrencyPair As String                                          'currency pair traded, i.e. EURNZD
    Dim PriceOpen As Double                                             'price of currency at outset of trade
    Dim Stoploss As Double                                              'points for stop loss
    Dim TargetPt As Double                                              'points targeted to gain from trade
    Dim DateClose As String                                             'date/time trade was closed
    Dim PriceClose As Double                                            'price of currency at end of trade
    Dim Commission As Variant                                           'commission charged by broker
    Dim Taxes As Variant                                                'taxes charged by broker
    Dim Swap As Variant                                                 'swap charged by broker (usually Wed night)
    Dim pips As Double                                                  'points in percentage from trade
    Dim Profit As Variant                                               'profit gained from trade
    Dim pScore As Integer                                               'self assessment of success of trade

    Dim tarray() As String                                              'array of all ticket numbers in journal
    Dim taddrs() As String                                              'array of all ticket numbers addresses
    Dim setup() As String                                               'array of trade setups

    Dim lCell As Variant                                                'row of last entry in detailed statement
    Dim tCell As String                                                 'cell address of trade to enter
    Dim jCell As String                                                 'static address of first trade in journal
    Dim fCell As String                                                 'static address of first trade in statement
    Dim pCell As String                                                 'cell address to paste trade
    Dim tCol As Variant                                                 'address of column number for new entry
    Dim pCol As Integer                                                 'column to paste in journal

    Dim StartRow As Integer                                             'first row where the data inputs start
    Dim LastRow As Integer                                              'for last row used in Journal
    Dim mFactor As Integer                                              'variable to calculate points
    Dim trades As Integer                                               'counter of trades entered
    Dim iSetup As String                                                'holds setups for inputbox
    Dim nSetup As String                                                'holds the entered setup

    StartRow = 20                                                       'first row of trades in journal
    jCell = "C20"                                                       'address of first trade in journal

    wsr = "Range"                                                       'sheet with semi-permanent data, i.e. setups
    wsj = "Journal"                                                     'journal to enter trades
    wsd = "DetailedStatement"                                           'detailed statement of completed trades

    Application.Goto Sheets(wsj).Range("L19")
    Application.ScreenUpdating = False
    Application.Calculation = xlManual


'~~> this section should run once if multiple trades need to be entered
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'check if trade is to be entered
    i = MsgBox(Space(5) & "Do you want to enter trade data?", 4, "Trade Entry")
        If i = 7 Then GoTo Cancel

    'check if workbook from trade terminal is open 'uses the function 'IsFileOpen'
    On Error Resume Next
    If IsFileOpen("C:\Users\image\Documents\trade\fx\statements\fxcm\DetailedStatement.xls") = False Then
        Set sQ = Workbooks.Open("C:\Users\image\Documents\trade\fx\statements\fxcm\DetailedStatement.xls")
    End If

    'journal investment workbook and detailed statement workbook
    Set eQ = Workbooks("Investment.xlsm")
    Set sQ = Workbooks("DetailedStatement.xls")

    'acquire the address of the first trade and the last entry in detailed statement
    With sQ.Sheets(wsd)
        For Each c In .Range("A1:A300")
            If c = "Ticket" Then tCell = c.Offset(1).Address: fCell = c.Offset(1, 13).Address
            If Left(c.Value, 4) = "Open" Then lCell = c.Offset(-1).Row - Range(tCell).Row: Exit For
        Next c
    End With

    'acquire max trades possible in journal and number of setups used
    m = Application.Max(eQ.Sheets(wsj).Range("A:A"))
    For Each c In eQ.Sheets(wsr).Range("Setups")
        If c <> "" Then n = n + 1
    Next c

    'redimension array to fit max number of trades per setup allowed, and number of setups
    ReDim tarray(m)
    ReDim taddrs(m)
    ReDim setup(n)

    'acquire ticket number of all trades already in the journal to avoid duplicate entries
    i = 0
    For a = 0 To m - 1               'row offset
        For b = 0 To n - 1           'column offset
            If eQ.Sheets(wsj).Range(jCell).Offset(a * 19, b * 12) <> "" Then
                i = i + 1
                tarray(i) = eQ.Sheets(wsj).Range(jCell).Offset(a * 19, b * 12)
                taddrs(i) = eQ.Sheets(wsj).Range(jCell).Offset(a * 19, b * 12).Address
            End If
        Next b
    Next a

    'redimension arrays to max number of trades currently in journal
    ReDim Preserve tarray(i)
    ReDim Preserve taddrs(i)

    'acquire setup entered & populate setup array for inputbox
        i = 1
        For Each c In Workbooks(JournalTitle).Sheets(wsr).Range("Setups")
            If c.Value = "" Then Exit For
            setup(i) = c.Value
            If i = 1 Then
                iSetup = "" & i & "        ~~      " & setup(i) & Chr(10)
            End If
            If i > 1 And Len(i) = 1 Then
                iSetup = iSetup & i & "        ~~      " & setup(i) & Chr(10)
            End If
            If i > 1 And Len(i) = 2 Then
                iSetup = iSetup & i & "      ~~      " & setup(i) & Chr(10)
            End If
            i = i + 1
        Next c


'~~> this section should loop in the case of entry of multiple trades
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

nTrade:

    'check if ticket# is empty and if so exit
    Ticket = sQ.Sheets(wsd).Range(tCell).Offset(0, 0).Value
    If Ticket = "" Then
        MsgBox "Ticket # cell is empty." & Chr(10) _
                & "Entry of this trade has been terminated.", 64, "Entry Canceled"
        GoTo Cancel
    Else
        GoSub cdata
    End If

    'check if trade was already added to journal
    For i = 1 To UBound(tarray)
        If Ticket = tarray(i) Then
            a = eQ.Sheets(wsj).Range("Journal_Header").Row
            b = eQ.Sheets(wsj).Range(taddrs(i)).Offset(0, -2).Column
            c = eQ.Sheets(wsj).Cells(a, b)
            d = eQ.Sheets(wsj).Range(taddrs(i)).Offset(0, -2)

            Application.Goto eQ.Sheets(wsj).Range(taddrs(i)).Offset(-1, -2), True
            eQ.Sheets(wsj).Range(taddrs(i)).Select

            Application.ScreenUpdating = True

            MsgBox "This trade, ticket #" & Ticket & ", already exists" & Chr(10) _
                & "under the following setup:  " & c & " #" & d & "." & Chr(10) & Chr(10) _
                & "The next trade will be examined.", 64, "Trade #" & Ticket & " Canceled"

            Application.ScreenUpdating = False

            Application.Goto eQ.Sheets(wsj).Range("A1"), True
            eQ.Sheets(wsj).Range("B2").Select
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False

            GoTo sTrade
        End If
    Next i

setup:

    'find column to paste in by way of the Setup
    i = InputBox(iSetup, "Trade Setup?", "Input answer here [1 - " & n & "]")
    If i = "" Then GoTo Cancel

    If IsNumeric(i) = False Or i < 1 Or i > n Then
        i = MsgBox("            A trade setup was not chosen." & Chr(10) _
                    & " Would you like to reenter the a trade setup?", 68, "Retry?")
        If i = 6 Then GoTo setup Else GoTo Cancel
    End If

    'determine column number to paste trade data
    If i = 1 Then: pCol = 3: Else: pCol = ((i - 1) * 12) + 3
    nSetup = i

    'determine cell to paste relative to
    LastRow = eQ.Sheets(wsj).Cells(Rows.Count, pCol).End(xlUp).Row
    If LastRow < StartRow Then: pCell = Cells(StartRow, pCol).Address(False, False)
    If LastRow >= StartRow Then: pCell = Cells(LastRow + 2, pCol).Address(False, False)

    'check if cell is correct
    tCol = Range(pCell).Offset(0, -2).Address(False, False)
    If IsEmpty(eQ.Sheets(wsj).Range(tCol).Value) Or IsNumeric(eQ.Sheets(wsj).Range(tCol).Value) = False Then
        MsgBox "Journal entry terminated." & Chr(10) _
        & "Missing column number at " & tCol & ".", 64, "Entry Canceled"
        GoTo Cancel
    End If

    'confirm cell location for trade entry
    Application.Goto eQ.Sheets(wsj).Range(pCell).Offset(-1, -2), True
    ActiveCell.Offset(1, 2).Activate

    Application.ScreenUpdating = True
    i = MsgBox("Cell location for trade data is selected." & Chr(10) _
             & "Continue trade entry?", 4 + 32, "Confirm Location")
    If i <> 6 Then
        GoTo Cancel
    End If
    Application.ScreenUpdating = False

Balance:

    'acquire starting balance from detailed statement
    For Each c In sQ.Sheets(wsd).Range("D1:D300")
        If c = "Deposit" Then
            Balance = c.Offset(0, 1)
            Balance = Replace(Balance, " ", "")
            Exit For
        End If
    Next c

    'sum profit before trade if any
    With sQ.Sheets(wsd)
        a = Range(tCell).Row
        b = Range(fCell).Row
        If a > b Then
            For Each c In sQ.Sheets(wsd).Range("N" & b & ":N" & a - 1)
                i = Replace(c, " ", "")
                Balance = Balance + (i * 1)
            Next c
        End If
    End With

    'manually enter balance if it is blank
    If Balance <> "" And Balance <> 0 Then
        GoTo TimeFrame
    Else
        'enter starting balance
        i = InputBox("Adjusted Balance?", "Starting Balance", "Input answer here ($)")
        If i = "" Then GoTo Cancel

        If IsNumeric(i) = False Or IsEmpty(i) Then
            MsgBox "Please enter the starting balance.", 64, "Balance"
            GoTo Balance
        End If

        If i < 275 Then
            MsgBox "Trade cannot be placed with that amount." & Chr(10) _
                 & "Please enter starting balance equal or greater than $275", 64, "Minimum Balance"
            GoTo Balance
        Else
            Balance = i
        End If
    End If

TimeFrame:

    'enter time frame
    i = InputBox("" _
            & "1 ~~ Monthly" & Chr(10) _
            & "2 ~~ Weekly" & Chr(10) _
            & "3 ~~ Daily" & Chr(10) _
            & "4 ~~ 4 Hour" & Chr(10) _
            & "5 ~~ Hourly" & Chr(10) _
            & "6 ~~ 30 min" & Chr(10) _
            , "Time Frame?", "Input answer here [1 - 6]")

    If i = "" Then GoTo Cancel
    If i = "" Or IsNumeric(i) = False Or i < 0 Or i > 6 Then
        MsgBox "Please enter a number between 1 and 6.", 64, "Time Frame Required"
        GoTo TimeFrame
    End If

    Select Case i
        Case Is = 1: TimeFrame = "Monthly"
        Case Is = 2: TimeFrame = "Weekly"
        Case Is = 3: TimeFrame = "Daily"
        Case Is = 4: TimeFrame = "4 Hour"
        Case Is = 5: TimeFrame = "Hourly"
        Case Is = 6: TimeFrame = "30 Min"
    End Select

eScore:

    'enter execution score
    i = InputBox("Execution Score?" & Chr(10) & Chr(10) _
            & "5   Followed trade as dictated in my plan" & Chr(10) _
            & "4   Followed trade entry but closed position before" & Chr(10) _
            & "      predetermined target was hit" & Chr(10) _
            & "3   Followed trade entry but removed or did not" & Chr(10) _
            & "      place stop/target and let position run" & Chr(10) _
            & "2   Entered setup late and/or did not set target" & Chr(10) _
            & "1   Impulse Trade" & Chr(10) _
            , "Score: Trade Execution", "Input answer here [1 - 5]")

    If i = "" Then GoTo Cancel
    If IsNumeric(i) = False Or IsEmpty(i) Or i < 1 Or i > 5 Then
        MsgBox "Please enter a number between 1 and 5.", 64, "Escore Required"
        GoTo eScore
    Else
        eScore = i
    End If

pScore:

    'enter performance score
    i = InputBox("Performance Score?" & Chr(10) & Chr(10) _
            & "5   Target hit" & Chr(10) _
            & "4   Profitable but at different exit" & Chr(10) _
            & "3   Out at even" & Chr(10) _
            & "2   Out at better price than stop..." & Chr(10) _
            & "      but negative trade" & Chr(10) _
            & "1   Stop hit or credit close (leverage hit)" & Chr(10) _
            , "Score: Trade Performance", "Input answer here [1 - 5]")

    If i = "" Then GoTo Cancel
    If IsNumeric(i) = False Or IsEmpty(i) Or i < 1 Or i > 5 Then
        MsgBox "Please enter a number between 1 and 5.", 64, "Pscore Required"
        GoTo pScore
    Else
        pScore = i
    End If

    'clean up dates
    DateOpen = Replace(DateOpen, ".", "-")
    DateOpen = Left(DateOpen, Len(DateOpen) - 3)
    DateClose = Replace(DateClose, ".", "-")
    DateClose = Left(DateClose, Len(DateClose) - 3)

    'clean up direction
    Select Case Direction
        Case Is = "sell": Direction = "short"
        Case Is = "buy": Direction = "long"
    End Select

    'calculate stoploss and target profit in points
    If InStr(PriceOpen, ".") = 2 Then: mFactor = 10000: Else mFactor = 100
    Stoploss = Round(Abs(PriceOpen - Stoploss) * mFactor, 0)
    TargetPt = Round(Abs(PriceOpen - TargetPt) * mFactor, 0)
    If Direction = "long" Then: pips = Round((PriceClose - PriceOpen) * mFactor, 1)
    If Direction = "short" Then: pips = Round((PriceOpen - PriceClose) * mFactor, 1)

    'format the profit to currency
    Profit = Trim(Profit) * 1
    Profit = Replace(Profit, " ", "")
    Profit = Format(Profit, "Currency")

    'find and sum additional commissions if any
    With sQ.Sheets(wsd)
        For i = 0 To lCell
            a = Left(sQ.Sheets(wsd).Range(tCell).Offset(i, 3), 10)
            b = Right(sQ.Sheets(wsd).Range(tCell).Offset(i, 3), 8)
            c = sQ.Sheets(wsd).Range(tCell).Offset(i, 13).Value
            If a = "Commission" And b = Ticket Then Commission = Commission + c
        Next i
    End With

    'check if taxes were taken since there is no journal entry for it
    If Taxes <> 0 Then
        Commission = Commission + Taxes
        MsgBox "Taxes were charged and added to the commission entry." & Chr(10) _
        & "Commission = " & Commission & Chr(10) _
        & "Taxes     = " & Taxes, 64, "Commission Plus Tax"
    End If

    'enter trade in journal
    GoSub eTrade

    'notify of succesful journal entry
    On Error Resume Next
    Application.Goto eQ.Sheets(wsj).Range(tCol).Offset(-1), True
    ActiveCell.Offset(1).Activate

    Application.ScreenUpdating = True

    a = eQ.Sheets(wsj).Range(pCell).Offset(0, -2)

    MsgBox setup(nSetup) & " #" & a & " trade has been entered succesfully!" & Chr(10) _
    & "Ticket #" & Ticket & ".", 64, "Success!"

    Application.ScreenUpdating = False

    'return to start of journal
    Application.Goto eQ.Sheets(wsj).Range("A1"), True
    eQ.Sheets(wsj).Range("B2").Select
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False

    'keep track of number of trades entered
    trades = trades + 1

sTrade:

    'check for more trades to enter
    a = sQ.Sheets(wsd).Range(tCell).Offset(1).Value
    b = sQ.Sheets(wsd).Range(tCell).Offset(1, 3).Value

    If IsNumeric(a) And IsNumeric(b) Then

        tCell = sQ.Sheets(wsd).Range(tCell).Offset(1).Address

        i = MsgBox("Another trade is listed," & Chr(10) _
            & " Ticket #" & a & "." & Chr(10) & Chr(10) _
            & "Continue trade entry?", 4 + 32, "Enter Next Trade?")

        If i = 6 Then GoTo nTrade

    End If

    'check for any Saturday dates
    Call c_TradeDate

    'provide total number of trades entered
    MsgBox Space(5) & "There are no more trades to enter." & Chr(10) _
         & Space(8) & "Total trades entered: " & trades, 0, "Trade Entry Complete"

Cancel:

    'close the workbook detailed statement
    If IsFileOpen("C:\Users\image\Documents\trade\fx\statements\fxcm\DetailedStatement.xls") = True Then
        sQ.Close SaveChanges:=False
    End If

    Application.Goto Sheets(wsj).Range("A1"), True
    Sheets(wsj).Range("L19").Select

    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

Exit Sub

cdata:

    'capture trade data from detailed statement
    Ticket = sQ.Sheets(wsd).Range(tCell).Offset(0, 0).Value
    DateOpen = sQ.Sheets(wsd).Range(tCell).Offset(0, 1).Value
    Direction = sQ.Sheets(wsd).Range(tCell).Offset(0, 2).Value
    Lots = sQ.Sheets(wsd).Range(tCell).Offset(0, 3).Value
    CurrencyPair = UCase(sQ.Sheets(wsd).Range(tCell).Offset(0, 4).Value)
    PriceOpen = sQ.Sheets(wsd).Range(tCell).Offset(0, 5).Value
    Stoploss = sQ.Sheets(wsd).Range(tCell).Offset(0, 6).Value
    TargetPt = sQ.Sheets(wsd).Range(tCell).Offset(0, 7).Value
    DateClose = sQ.Sheets(wsd).Range(tCell).Offset(0, 8).Value
    PriceClose = sQ.Sheets(wsd).Range(tCell).Offset(0, 9).Value
    Commission = sQ.Sheets(wsd).Range(tCell).Offset(0, 10).Value
    Taxes = sQ.Sheets(wsd).Range(tCell).Offset(0, 11).Value
    Swap = sQ.Sheets(wsd).Range(tCell).Offset(0, 12).Value
    Profit = sQ.Sheets(wsd).Range(tCell).Offset(0, 13).Value

Return

eTrade:

    'enter trade in journal
    eQ.Sheets(wsj).Range(pCell).Offset(0, 0) = Ticket
    eQ.Sheets(wsj).Range(pCell).Offset(1, 0) = Balance
    eQ.Sheets(wsj).Range(pCell).Offset(2, 0) = eScore
    eQ.Sheets(wsj).Range(pCell).Offset(3, 0) = DateOpen
    eQ.Sheets(wsj).Range(pCell).Offset(4, 0) = TimeFrame
    eQ.Sheets(wsj).Range(pCell).Offset(5, 0) = CurrencyPair
    eQ.Sheets(wsj).Range(pCell).Offset(6, 0) = Lots
    eQ.Sheets(wsj).Range(pCell).Offset(7, 0) = Direction
    eQ.Sheets(wsj).Range(pCell).Offset(8, 0) = PriceOpen
    eQ.Sheets(wsj).Range(pCell).Offset(9, 0) = PriceClose
    eQ.Sheets(wsj).Range(pCell).Offset(10, 0) = Stoploss
    eQ.Sheets(wsj).Range(pCell).Offset(11, 0) = TargetPt
    eQ.Sheets(wsj).Range(pCell).Offset(12, 0) = pips
    eQ.Sheets(wsj).Range(pCell).Offset(13, 0) = Profit
    eQ.Sheets(wsj).Range(pCell).Offset(14, 0) = Swap
    eQ.Sheets(wsj).Range(pCell).Offset(15, 0) = Commission
    eQ.Sheets(wsj).Range(pCell).Offset(16, 0) = DateClose
    eQ.Sheets(wsj).Range(pCell).Offset(17, 0) = pScore

Return

End Sub
