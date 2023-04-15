Attribute VB_Name = "s3_OptRanking"
Option Explicit

Sub OptRank()
's3, rank optimal wkly trade data

    Dim a As Variant                                    'misc, loop, count
    Dim b As Variant                                    'misc, loop, count
    Dim c As Variant                                    'misc, loop, count
    Dim i As Variant                                    'holds the ranked indexes
    Dim t As Integer                                    'loop through entire macro
    Dim ws As String                                    'holds worksheet name
    Dim Q As Collection                                 'holds unique values for comparison
    Dim sCell As String                                 'holds address of first entry
    Dim duplicate As Integer                            'counts duplicates

    Dim dupl() As Variant                               'array holds duplicates
    Dim orig() As Variant                               'array holds original range data
    Dim pips() As Variant                               'array holds range of pip data
    Dim psum() As Variant                               'array holds sums of all setups

    'avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'set up
    ws = "Rank"
    sCell = "R60"
    Set Q = New Collection

    'delete previous data
    Sheets(ws).Range("OptData").ClearContents

    'move through required ranges w/loop
    For t = 1 To 5

        If t = 1 Then orig = Range("Data_Opt_Setup")
        If t = 2 Then orig = Range("Data_Opt_cPair")
        If t = 3 Then orig = Range("Data_Opt_Chart")
        If t = 4 Then orig = Range("Data_Opt_wkDay")
        If t = 5 Then orig = Range("Data_Opt_cTime")
        If t = 6 Then orig = Range("Data_Opt_cDate")

        'create unique collection for comparison
        For Each a In orig
            On Error Resume Next
            If a <> "" Then Q.Add a, CStr(a)
        Next a

        'redimension arrays to number of unique items
        ReDim dupl(Q.Count)
        ReDim psum(Q.Count)

        'find number of occurences as duplicates
        For a = 1 To Q.Count

            For Each b In orig
                If b <> "" And b = Q(a) Then
                    duplicate = duplicate + 1
                End If
            Next b

            dupl(a) = duplicate
            duplicate = 0
'            Debug.Print dupl(a) & " = " & Q(a)
        Next a

        'rank index of duplicates by function
        i = RankArray(arr:=dupl)

        'debug.print check for accuracy
'        For a = 1 To UBound(i)
'            If i(a) <> "" Then
'                Debug.Print "Rank #" & a & " is " & dupl(i(a)) & " of " & Q(i(a))
'            End If
'        Next a

        'add pips if loop is on setup
        If t = 1 Then pips() = Range("Data_Opt_sPips")

        If t = 1 Then
            For a = 1 To Q.Count
                For c = 1 To UBound(orig)
                    If orig(a, 1) = "" Then Exit For
                    If Q(a) = orig(c, 1) Then
                       psum(a) = psum(a) + pips(c, 1)
                    End If
                Next c
            Next a
        End If

        'paste ranked values
        For a = 1 To UBound(i)
            b = Sheets(ws).Range(sCell).Offset(a - 1, -1)
            If a = b Then
                If t = 1 Then
                    Sheets(ws).Range(sCell).Offset(a - 1, 5) = dupl(i(a))     'paste # of trades
                    Sheets(ws).Range(sCell).Offset(a - 1, 6) = psum(i(a))     'paste pips
                End If
                If t < 6 Then
                    Sheets(ws).Range(sCell).Offset(a - 1, t - 1) = Q(i(a))
                End If
                If t = 6 Then
                    Sheets(ws).Range(sCell).Offset(a - 1, t + 1) = Q(i(a))
                End If
            End If
        Next a

        'clear collection for next iteration
        Set Q = Nothing
        Set Q = New Collection

    Next t

    'return journal settings
    Application.Goto Sheets(ws).Range("L5")
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
