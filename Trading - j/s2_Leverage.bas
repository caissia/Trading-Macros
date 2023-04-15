Attribute VB_Name = "s2_Leverage"
Option Explicit
'contains 2 macros: ClearNTP, Leverage

Sub Leverage()
's2, calculates leverage, etc.

    Dim Equity As Double
    Dim Currency_Pair As String
    Dim Currency_Price As Double
    Dim Lot_Size As Double
    Dim Lot_Contract As Long
    Dim spread As Double
    Dim Max_SL As Integer
    Dim Margin_Equity_Percentage As Double
    Dim Opt_Leverage As Double
    Dim Max_Leverage As Double
    Dim Max_Leverage_Dollars As Double
    Dim Pip_Value As Double
    Dim Max_Pip_SL As Double
    Dim Max_Dollar_SL As Double
    Dim Percent_Gain As Double
    Dim Pips_Required As Double

    Dim factor As Double
    Dim subcounter As Integer

    Dim c As Range
    Dim b As Variant
    Dim a As Variant
    Dim i As Integer
    Dim eQ As Integer
    Dim qPair As String
    Dim price1 As Double
    Dim price2 As Double
    Dim price3 As Double
    Dim price4 As Double
    Dim price5 As Double
    Dim price6 As Double
    Dim leverage_pair As String
    Dim leverage_price As Double
    Dim all_pairs(24) As String
    Dim all_prices(24) As String

    tC = ActiveCell.Address
    tC = ActiveCell.Offset(1).Address       'in order to avoid issues with the merged cell

    'acquire inputs
    Equity = Range(tC).Offset(0, 1)
    Max_Leverage = Range(tC).Offset(1, 1)
    Currency_Pair = Range(tC).Offset(2, 0)
    spread = Range(tC).Offset(2, 1)
    Currency_Price = Range(tC).Offset(3, 1)
    Lot_Contract = Range(tC).Offset(4, 1)
    Max_SL = Range(tC).Offset(5, 1)

    a = Range(tC).Offset(11, 0).Address
    b = Mid(Range(a), 2, 1)
    If b = "%" Then
        Percent_Gain = Left(Range(a), 1) / 100
    Else
        Percent_Gain = Left(Range(a), 2) / 100
    End If

    price1 = Sheets("Range").Range("L44")   'AUDUSD
    price2 = Sheets("Range").Range("L28")   'USDCAD
    price3 = Sheets("Range").Range("L33")   'USDCHF
    price4 = Sheets("Range").Range("L46")   'GBPUSD
    price5 = Sheets("Range").Range("L41")   'USDJPY
    price6 = Sheets("Range").Range("L47")   'NZDUSD

    '~~> populate arrays
    With ThisWorkbook.Sheets("Range")
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
    If Left(Currency_Pair, 3) = "USD" Then leverage_pair = Currency_Pair

    If Left(Currency_Pair, 3) = "AUD" Then leverage_pair = "AUDUSD"
    If Left(Currency_Pair, 3) = "EUR" Then leverage_pair = "EURUSD"
    If Left(Currency_Pair, 3) = "GBP" Then leverage_pair = "GBPUSD"
    If Left(Currency_Pair, 3) = "NZD" Then leverage_pair = "NZDUSD"

    If Left(Currency_Pair, 3) = "CAD" Then leverage_pair = "USDCAD"
    If Left(Currency_Pair, 3) = "CHF" Then leverage_pair = "USDCHF"

    '~~> acquire currency pair price to calculate correct leverage
    For i = 0 To 24
        If leverage_pair = all_pairs(i) Then
            Exit For
        End If
    Next i

    'MsgBox i & ". " & Currency_Pair & " = " & all_pairs(i) & " = " & all_prices(i)

    Opt_Leverage = Max_Leverage

    '~~> determine optimum leverage to take based on how much
    'cushion (pip s/l) for trade to go against trade direction

    For subcounter = 1 To (Max_Leverage * 10000)

        Lot_Size = Application.RoundDown((Equity * 0.9) / ((Lot_Contract * Opt_Leverage) + (1 * Max_SL)), 1)

        If Left(leverage_pair, 3) = "USD" Then
            Max_Leverage_Dollars = Lot_Contract * Max_Leverage * Lot_Size
        Else
            Max_Leverage_Dollars = Lot_Contract * Max_Leverage * all_prices(i) * Lot_Size
        End If

        Max_Leverage_Dollars = Application.Round(Max_Leverage_Dollars, 2)

        '~~> begin to calculate correct pip value
        qPair = Right(Currency_Pair, 3)
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
               Pip_Value = (Lot_Contract * factor) * price1 * Lot_Size
            Case 2
               Pip_Value = (Lot_Contract * factor) / price2 * Lot_Size
            Case 3
               Pip_Value = (Lot_Contract * factor) / price3 * Lot_Size
            Case 4
               Pip_Value = (Lot_Contract * factor) * price4 * Lot_Size
            Case 5
               Pip_Value = (Lot_Contract * factor) / price5 * Lot_Size
            Case 6
               Pip_Value = (Lot_Contract * factor) * price6 * Lot_Size
            Case 7
               Pip_Value = (Lot_Contract * factor) * Lot_Size
        End Select
        '~~> end calculate correct pip value
        ''''''''''''''''''''''''''''''''

        Max_Pip_SL = Application.Round(((Equity - Max_Leverage_Dollars) / Pip_Value), 1)

        If Max_Pip_SL >= Max_SL Then
            Exit For
        Else
            Opt_Leverage = Application.RoundDown(Opt_Leverage + 0.001, 3)
        End If

    Next subcounter

    Max_Dollar_SL = Max_Pip_SL * Pip_Value
    Margin_Equity_Percentage = Equity / Max_Leverage_Dollars
    Pips_Required = Round(((Equity * Percent_Gain) / Pip_Value) + spread, 1)

    'paste results
    Range(tC).Offset(6, 1) = Lot_Size
    Range(tC).Offset(7, 1) = Pip_Value
    Range(tC).Offset(8, 1) = Max_Pip_SL
    Range(tC).Offset(9, 1) = Max_Dollar_SL
    Range(tC).Offset(10, 1) = Margin_Equity_Percentage
    Range(tC).Offset(11, 1) = Pips_Required & " pips"

End Sub

Sub ClearNTP()
's2
    Range("F29:J29").ClearContents
    Range("F29").Select

End Sub
