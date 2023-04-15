Attribute VB_Name = "s2_Projection"
Option Explicit
'contains 2 macros: Adjust, Projection

Sub Projection()
's2, projects trade results

    Dim Balance As Double
    Dim balanceA As Double
    Dim tradesperday As Byte
    Dim daysperweek As Byte
    Dim monthstraded As Integer
    Dim edge As Double                  'specifies the winning percentage
    Dim minloss As Integer
    Dim maxloss As Integer
    Dim minpips As Integer
    Dim maxpips As Integer
    Dim maxlotsize As Integer
    Dim maxleverage As Double
    Dim counter As Integer
    Dim tradestotal As Integer
    Dim LotSize As Double
    Dim Leverage As Double
    Dim pips As Double
    Dim loss As Double
    Dim Profit As Double
    Dim profitA As Double
    Dim profitN As Double
    Dim profitP As Double
    Dim neg_trades As Integer
    Dim pos_trades As Integer
    Dim pipArray As Variant
    Dim profitArray As Variant
    Dim profitNArray As Variant
    Dim profitPArray As Variant
    Dim answer As Integer
    Dim p As Variant
    Dim n As Variant
    Dim a As Variant

    Dim CurrencyPair As String
    Dim currencyprice As Double
    Dim lotcontract As Double
    Dim spread As Double
    Dim factor As Double
    Dim maxleveragedollars As Double
    Dim pipvalue As Double
    Dim maxsl As Double
    Dim maxpipsl As Double
    Dim maxdollarsl As Double
    Dim optleverage As Double
    Dim subcounter As Integer

    Dim c As Range
    Dim i As Integer
    Dim eQ As Integer
    Dim qPair As String
    Dim price1 As Double
    Dim price2 As Double
    Dim price3 As Double
    Dim price4 As Double
    Dim price5 As Double
    Dim price6 As Double
    Dim gainper As Double
    Dim leverage_pair As String
    Dim leverage_price As Double
    Dim all_pairs(24) As String
    Dim all_prices(24) As String
    Dim sumprofit As Double
    Dim sumpips As Double
    Dim tradegoal As Long
    Dim tradegoalck As Long

    Application.Calculation = xlManual

    'clear previous data if any
    Range("Projections").ClearContents

    'acquire start address to offset
    tC = ActiveCell.Address

    Balance = Range(tC).Offset(1, 2)
    balanceA = Range(tC).Offset(1, 2)
    tradesperday = Range(tC).Offset(2, 2)
    daysperweek = Range(tC).Offset(3, 2)
    monthstraded = Range(tC).Offset(4, 2)
    edge = Range(tC).Offset(5, 2)
    minloss = Range(tC).Offset(6, 2)
    maxloss = Range(tC).Offset(7, 2)
    minpips = Range(tC).Offset(8, 2)
    maxpips = Range(tC).Offset(9, 2)
    maxleverage = Range(tC).Offset(10, 2)
    maxlotsize = Range(tC).Offset(11, 2)

    'leverage data
    CurrencyPair = Range(tC).Offset(3, -13)         'pair (left)
    spread = Range(tC).Offset(3, -11)               'pair bid (right)
    currencyprice = Range(tC).Offset(4, -11)        'price, below pair bid (right)
    lotcontract = Range(tC).Offset(5, -11)          'lot size: 10000, 10000, or 1000 (right)
    maxsl = Range(tC).Offset(6, -11)                'stop loss (number on the right)

    a = Range(tC).Offset(12, 0)
    tradegoal = Mid(a, 10, Len(a) - 10)
    If tradegoal = 1 Then
        tradegoal = 1000000
    Else
        tradegoal = tradegoal * 1000
    End If

    price1 = Sheets("Range").Range("L44")   'AUDUSD
    price2 = Sheets("Range").Range("L28")   'USDCAD
    price3 = Sheets("Range").Range("L33")   'USDCHF
    price4 = Sheets("Range").Range("L46")   'GBPUSD
    price5 = Sheets("Range").Range("L41")   'USDJPY
    price6 = Sheets("Range").Range("L47")   'NZDUSD

    '~~> populate arrays
    With Sheets("Range")
        i = 0
        While i < 25
            For Each c In .Range("Pairs").Cells
                all_pairs(i) = c.Value
                i = i + 1
            Next c
        Wend
        i = 0
        While i < 25
            For Each c In .Range("Price").Cells
                all_prices(i) = c.Value
                i = i + 1
            Next c
        Wend
    End With

    '~~> acquire correct currency pair to calculate max leverage in dollars
    If Left(CurrencyPair, 3) = "USD" Then leverage_pair = CurrencyPair

    If Left(CurrencyPair, 3) = "AUD" Then leverage_pair = "AUDUSD"
    If Left(CurrencyPair, 3) = "EUR" Then leverage_pair = "EURUSD"
    If Left(CurrencyPair, 3) = "GBP" Then leverage_pair = "GBPUSD"
    If Left(CurrencyPair, 3) = "NZD" Then leverage_pair = "NZDUSD"

    If Left(CurrencyPair, 3) = "CAD" Then leverage_pair = "USDCAD"
    If Left(CurrencyPair, 3) = "CHF" Then leverage_pair = "USDCHF"

    '~~> acquire currency pair price to calculate correct leverage
    For i = 0 To 24
        If leverage_pair = all_pairs(i) Then
            Exit For
        End If
    Next

    neg_trades = 1 'must start at 1 to have the conditional work below
    pos_trades = 1

    tradestotal = monthstraded * daysperweek * tradesperday * 4.3333

    ReDim pipArray(1 To tradestotal) As Long
    ReDim profitArray(1 To tradestotal) As Long
    ReDim profitNArray(1 To tradestotal) As Long
    ReDim profitPArray(1 To tradestotal) As Long

    If minpips <= maxpips And minloss <= maxloss Then

        '~~> determine optimum leverage to take based on how much
        'cushion (pip s/l) for trade to go against trade direction
        For counter = 1 To tradestotal

            optleverage = maxleverage

            For subcounter = 1 To (maxleverage * 10000)

                LotSize = Application.RoundDown((Balance * 0.9) / ((lotcontract * optleverage) + (1 * maxsl)), 2)
                If LotSize > maxlotsize Then LotSize = maxlotsize

                If Left(leverage_pair, 3) = "USD" Then
                    maxleveragedollars = lotcontract * maxleverage * LotSize
                Else
                    maxleveragedollars = lotcontract * maxleverage * all_prices(i) * LotSize
                End If
                maxleveragedollars = Application.Round(maxleveragedollars, 2)

                '~~> calculate correct pip value
                qPair = Right(CurrencyPair, 3)
                If qPair = "AUD" Then eQ = 1
                If qPair = "CAD" Then eQ = 2
                If qPair = "CHF" Then eQ = 3
                If qPair = "GBP" Then eQ = 4
                If qPair = "JPY" Then eQ = 5
                If qPair = "NZD" Then eQ = 6
                If qPair = "USD" Then eQ = 7

                'determine decimal point based on yen vs non-yen
                If eQ = 5 Then factor = 0.01 Else: factor = 0.0001

                Select Case eQ
                    Case 1
                       pipvalue = (lotcontract * factor) * price1 * LotSize
                    Case 2
                       pipvalue = (lotcontract * factor) / price2 * LotSize
                    Case 3
                       pipvalue = (lotcontract * factor) / price3 * LotSize
                    Case 4
                       pipvalue = (lotcontract * factor) * price4 * LotSize
                    Case 5
                       pipvalue = (lotcontract * factor) / price5 * LotSize
                    Case 6
                       pipvalue = (lotcontract * factor) * price6 * LotSize
                    Case 7
                       pipvalue = (lotcontract * factor) * LotSize
                End Select

                pipvalue = Application.Round(pipvalue, 5)

                '~~> end calculate correct pip value
                ''''''''''''''''''''''''''''''''''''

                maxpipsl = Application.Round(((Balance - maxleveragedollars) / pipvalue), 1)

                If maxpipsl >= maxsl + spread Then
                    Exit For
                Else
                    optleverage = Application.RoundDown(optleverage + 0.001, 3)
                End If

            Next subcounter

            maxdollarsl = Application.Round(maxpipsl * pipvalue, 2)
            Leverage = Application.Round(maxleveragedollars / (lotcontract * all_prices(i)), 2)

            If Leverage < maxleverage Then
                counter = tradestotal
                answer = MsgBox("Leverage of last trade:  " & Leverage & "" & Chr(10) _
                        & " " & Chr(10) _
                        & "Would you like possible solutions?" & Chr(10), 1 + 64, "Leverage Exceeded")
                If answer = 1 Then
                    MsgBox "1. Increase the leverage," & Chr(10) _
                            & "2. Increase the maximum pip loss," & Chr(10) _
                            & "3. Increase the maximum lot size.", , "Solutions:"
                End If
            End If

                pips = Application.RandBetween(minpips, maxpips)
                loss = Application.RandBetween(-maxloss, -minloss)

            If pos_trades / (counter) <= edge Then

                Profit = pips * pipvalue
                profitP = pips * pipvalue

                pipArray(counter) = pips
                profitArray(counter) = Profit
                profitPArray(counter) = profitP

                profitA = pips * Application.RoundUp(pipvalue / 2, 0)
                profitA = profitA + (pips * Application.RoundUp(pipvalue / 2, 0))

                pos_trades = pos_trades + 1     'counts wins

                sumprofit = sumprofit + Profit  'sums all profit
                sumpips = sumpips + pips        'sums all pips

            Else

                Profit = loss * pipvalue
                profitN = loss * pipvalue

                pipArray(counter) = loss
                profitArray(counter) = Profit
                profitNArray(counter) = profitN

                profitA = loss * Application.RoundUp(pipvalue / 2, 0)
                profitA = profitA + (loss * Application.RoundUp(pipvalue / 2, 0))

                neg_trades = neg_trades + 1     'counts losses

            End If

            'determine how many trades until $ goal attained
            If tradegoalck = 0 And Balance >= tradegoal Then tradegoalck = counter

            'acquire new balance
            If Balance + Profit <= 30 Or balanceA + profitA <= 30 Then
                counter = tradestotal
                MsgBox "Next trade will result in a negative balance.", 16, "Projection Halted!"
                Exit Sub
            Else
                Balance = Balance + Profit
                balanceA = balanceA + profitA
            End If

                If Application.sum(profitPArray) = 0 Then
                    p = 1
                Else
                    p = Application.sum(profitPArray)
                End If

                If Abs(Application.sum(profitNArray)) = 0 Then
                    n = 1
                Else
                    n = Abs(Application.sum(profitNArray))
                End If

            'determine average profit per pip or trade
            If pos_trades > 1 Then
                a = Range(tC).Offset(12, 4)
                If Mid(a, 10, 1) = "P" Then
                    gainper = sumprofit / sumpips
                ElseIf Mid(a, 10, 1) = "T" Then
                    gainper = sumprofit / pos_trades
                End If
            End If

            '~~> paste values onto worksheet before moving to next trade
            Range(tC).Offset(1, 6) = Balance
            Range(tC).Offset(2, 6) = balanceA
            Range(tC).Offset(3, 6) = pos_trades - 1                    'Total wininning trades
            Range(tC).Offset(4, 6) = neg_trades - 1                    'Total losing trades
            Range(tC).Offset(5, 6) = pos_trades + neg_trades - 2       'Total Trades (counters started at 1 so subtract 2 for both counters
            Range(tC).Offset(6, 6) = p / n                             'Profit Factor

            Range(tC).Offset(8, 6) = Application.Max(profitPArray)     'largest upswing
            Range(tC).Offset(9, 6) = Application.Min(profitNArray)     'largest drawdown


            Range(tC).Offset(12, 2) = tradegoalck                      'average profit per pip/trade
            Range(tC).Offset(12, 6) = gainper                          'average profit per pip/trade

        Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim cell As Variant
        Dim J As Long
        Dim iNumCells As Long
        Dim iUVals As Long
        Dim sUCells As Variant

        iNumCells = Abs(Application.Count(pipArray))
        ReDim sUCells(iNumCells) As Long

        iUVals = 0
        For Each cell In pipArray                                   'counts unique elements
            If cell <> 0 Then
                For J = 1 To iUVals
                    If sUCells(J) = cell Then
                        Exit For
                    End If
                Next J
                If J > iUVals Then
                    iUVals = iUVals + 1
                    sUCells(iUVals) = cell
                End If
            End If
        Next cell
        Range(tC).Offset(7, 6) = iUVals            'counts unique elements

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim s As Variant
        Dim lCounter As Long
        Dim lMaxCount As Long
        Dim lCounterP As Long
        Dim lMaxCountP As Long

        lCounter = 0
        lMaxCount = 0

        On Error Resume Next 'counts contiguous losses
        For Each s In profitArray
            If s < 0 Then
                lCounter = lCounter + 1
                If lCounter > lMaxCount Then
                    lMaxCount = lCounter
                End If
            Else
                lCounter = 0
            End If
        Next s
        Range(tC).Offset(10, 6) = lMaxCount        'counts contiguous losses

        lCounterP = 0
        lMaxCountP = 0

        On Error Resume Next 'counts contiguous wins
        For Each s In profitArray
            If s > 0 Then
                lCounterP = lCounterP + 1
                If lCounterP > lMaxCountP Then
                    lMaxCountP = lCounterP
                End If
            Else
                lCounterP = 0
            End If
        Next s
        Range(tC).Offset(11, 6) = lMaxCountP       'counts contiguous wins

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Else
        MsgBox "Inputs need to be corrected before proceeding.", 16, "Projection Halted"
    End If

    Application.Calculation = xlAutomatic

End Sub

Sub Adjust()
's2, makes equity = starting balance if either changed

    Application.EnableEvents = False

    If Range(tC).Column = 5 Then Range("R18").Value = Range("E18").Value
    If Range(tC).Column = 18 Then Range("E18").Value = Range("R18").Value

    Application.EnableEvents = True

End Sub
