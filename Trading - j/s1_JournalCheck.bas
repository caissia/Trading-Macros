Attribute VB_Name = "s1_JournalCheck"
Option Explicit
'contains 6 macros: Journal_Check, m_Data, m_Image, m_RawData, m_OptImage, c_TradeDate

Sub Journal_Check()

Dim i As Integer

    Application.Goto Range("L19")

    i = MsgBox(Space(5) & "Do you want to check the journal for the following?" & Chr(10) _
                & Space(36) & "missing data..." & Chr(10) _
                & Space(36) & "missing images..." & Chr(10) _
                & Space(36) & "missing image data..." & Chr(10) _
                & Space(36) & "incorrect dates...", 4, "Journal Check")

    If i = 6 Then

        Call m_Data
        Call m_Image
        Call m_RawData
        Call m_OptImage
        Call c_TradeDate

        MsgBox Space(5) & "Journal check completed.", 0, "Completed"

    End If

End Sub

Sub m_Data()
'checks if any trade data is missing for corresponding image

    Dim dSetup As String
    Dim setup(15) As String                                                 'Array of 16 trade setups
    Dim cSetup(15) As String
    Dim bSetup(15) As String
    Dim aSetup(15, 100) As String
    Dim ImageAddress(3200) As String                                        'Array of image addresses
    
    Dim DataRow As Integer                                                  'Row of trade data
    Dim StartRow As Integer                                                 'Start row of trade data
    Dim DataColumn As Integer                                               'Column of trade data
    Dim ImageColumn As Integer                                              'Column of image
    Dim TradeNumber As Integer                                              'Number listed off trade data
    
    Dim r As Range
    Dim s As Shape
    Dim i As Variant
    Dim b, c, m, n, t, y, z As Integer

    StartRow = 20

    Windows(JournalTitle).Activate
    Application.Goto Sheets("Journal").Range("A1"), True
    Sheets("Journal").Unprotect
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

'    '~~> Check if image is to be entered
'    i = MsgBox("Do you want to check for images missing trade data?", 4 + 32, "Check Trade Data")
'    If i = 7 Then: MsgBox "Check for missing trade data canceled.", 0 + 64, "Canceled": GoTo ex

    '~~> populate setup array
    With ThisWorkbook.Sheets("Range")
        i = 0
        For Each r In .Range("Setups").Cells
            If r.Value = "" Then Exit For
            setup(i) = r.Value
            i = i + 1
        Next r
    End With

    '~~> populate array with current image addresses
    With Sheets("Journal")
        On Error Resume Next
        For Each s In .Shapes
            If StartRow <= s.TopLeftCell.Row Then
                If s.Type = msoPicture Then
                    ImageAddress(z) = s.TopLeftCell.Address
                    z = z + 1
                End If
            End If
        Next s
    End With

    '~~> check for missing items
    For y = 0 To z

        'acquire address of image
        i = ImageAddress(y)

        'acquire column and row of image and data
        ImageColumn = Range(i).Column
        DataRow = Range(i).Row

        'acquire column of data
        If ImageColumn = 4 Or ImageColumn = 8 Then DataColumn = 3
        If (ImageColumn Mod 12) - 4 = 0 Then DataColumn = ImageColumn - 1
        If (ImageColumn Mod 12) - 8 = 0 Then DataColumn = ImageColumn - 5

        'check if data address was already checked, if so go to next address
        If t <> Cells(DataRow, DataColumn).Address Then

            t = Cells(DataRow, DataColumn).Address

            'acquire setup(i) number
            i = (DataColumn - 3) / 12

            'acquire the number of setup trade
            TradeNumber = Cells(DataRow, DataColumn).Offset(0, -2).Value

            'check for missing data in datacolumn
            c = 0: n = 0
            Cells(DataRow, DataColumn).Activate
            While n < 18
                If IsEmpty(ActiveCell) Then c = c + 1
                ActiveCell.Offset(1).Activate
                n = n + 1
            Wend

            'record all missing items to setup
            If c = 1 Then aSetup(i, TradeNumber) = "        #" & TradeNumber & " is missing " & c & " item"
            If c > 1 Then aSetup(i, TradeNumber) = "        #" & TradeNumber & " is missing " & c & " items"

        End If

    Next y

    '~~> tally all missing items
    For i = 0 To 15
        For n = 1 To 100
            If aSetup(i, n) <> "" Then
                If bSetup(i) = "" Then
                    bSetup(i) = aSetup(i, n)
                Else
                    bSetup(i) = bSetup(i) & Chr(10) & aSetup(i, n)
                End If
            End If
        Next n
    Next i

    For i = 0 To 15
        If bSetup(i) <> "" Then cSetup(i) = setup(i) & Chr(10) & bSetup(i)
    Next i

    For i = 0 To 15
        If cSetup(i) <> "" Then
            If dSetup = "" Then dSetup = cSetup(i) Else dSetup = dSetup & Chr(10) & cSetup(i)
        End If
    Next

    '~~> report result
    If z = 0 Then
        MsgBox "There are no images to check" & Chr(10) _
               & "corresponding trade data.", 0 + 64, "No Missing Items"
        GoTo ex
    End If

    If c > 0 Then
        MsgBox dSetup, 0, "Missing Trade Data of Image"
    Else
        MsgBox "Currently there are no missing trade" & Chr(10) _
             & " items for corresponding images.", 0 + 64, "Trade Data Complete"
    End If

ex:

    Application.Goto Sheets("Journal").Range("A1"), True
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

End Sub

Sub m_Image()
'checks if image is missing for corresponding trade data

    Dim r As Range
    Dim s As Shape
    Dim c As Variant
    Dim StartRow As Integer
    Dim a, i, Q, t, y, z As Integer

    Dim aImage As String
    Dim setup(15) As String
    Dim mAddress(1600) As String
    Dim ImageAddress(1600) As String
    Dim ImageMissing(1600) As String

    StartRow = 20

    Windows(JournalTitle).Activate
    Application.Goto Sheets("Journal").Range("A1"), True
    Sheets("Journal").Unprotect
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

'    '~~> verify that check for missing images
'    i = MsgBox("Do you want to check for missing images?", 4 + 32, "Check Images")
'    If i = 7 Then: MsgBox "Check for missing images canceled.", 0 + 64, "Canceled": GoTo ex

    '~~> populate setup array
    i = 0
    With Sheets("Range")
        For Each r In .Range("Setups").Cells
            If r.Value = "" Then Exit For
            setup(i) = r.Value
            i = i + 1
        Next r
    End With

    '~~> populate all images array addresses
    '1 for first image, 2 for second image
    i = 0
    With Sheets("Journal")
        For Each s In .Shapes
            ImageAddress(i) = s.TopLeftCell.Address
            i = i + 1
            On Error Resume Next
        Next s
    End With

    '~~> set range of  trade data
    'create array of missing addresses
    i = 0: y = 0: z = 0
    With Sheets("Journal")
        For Each r In .Range("Journal_Data").Cells
            If IsEmpty(r.Value) Then
            Else
                While y < 18
                    t = r.Offset(-y, -2)
                        If t <> 0 And c <> r.Offset(-y).Address And r.Offset(-y, -2).Row >= StartRow Then
                            mAddress(z) = r.Offset(-y).Offset(0, 1).Address 'acquire 1st image address
                            z = z + 1
                            mAddress(z) = r.Offset(-y).Offset(0, 5).Address 'acquire 2nd image address
                            z = z + 1
                            c = r.Offset(-y).Address
                            y = 18
                        End If
                    y = y + 1
                Wend
                y = 0
            End If
        Next r
    End With

    '~~> create list of all setups with missing images
    t = 0:
    If z > 0 Then
        For a = 0 To z
            Q = Filter(ImageAddress, mAddress(a))
            If UBound(Q) < 0 And mAddress(a) <> "" Then         'image not found so...
                t = t + 1                                       'tally missing images
                y = Range(mAddress(a)).Column                   'acquire column number
                If (y - 4) Mod 12 = 0 Then                      '1st image is missing
                    i = (y - 4) / 12                            'acquire Setup(i) number
                    c = Range(mAddress(a)).Offset(0, -3) + 0.1  'acquire listed setup number
                Else                                            '2nd image is missing
                    i = (y - 8) / 12                            'acquire Setup(i) number
                    c = Range(mAddress(a)).Offset(0, -7) + 0.2  'acquire listed setup number
                End If
                If ImageMissing(i) = "" Then                    'assemble missing images by setup
                    ImageMissing(i) = c
                Else
                    ImageMissing(i) = ImageMissing(i) & ", " & c
                End If
            End If
        Next a
    End If

    'final assembly of all missing images
    For i = 0 To 15
        If ImageMissing(i) <> "" Then
            If i = 0 Then aImage = "   " & setup(i) & ": " & ImageMissing(i)
            If i > 0 Then aImage = aImage & Chr(10) & "   " & setup(i) & ": " & ImageMissing(i)
        End If
    Next i

    'alert user of found entry
    If t = 0 Then
        MsgBox "There are no missing images " & Chr(10) _
             & " for corresponding trades.", 64, "Investment"
    ElseIf t = 1 Then
        MsgBox "Missing image found for corresponding trade..." & Chr(10) & aImage, 64, "Investment"
    ElseIf t > 1 Then
        MsgBox "There are " & t & " missing images " & Chr(10) _
             & "  for corresponding trades..." & Chr(10) & aImage, 64, "Investment"
    End If

ex:

    Application.Goto Sheets("Journal").Range("A1"), True
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

End Sub

Sub m_RawData()
'checks for incomplete trade data... if data is missing for any trade

    Dim test As String
    Dim dSetup As String
    Dim setup(15) As String                                                 'Array of 16 trade setups
    Dim cSetup(15) As String
    Dim bSetup(15) As String
    Dim aSetup(15, 100) As String
    Dim TradeAddress(3200) As String                                        'Array of image addresses

    Dim DataRow As Integer                                                  'Row of trade data
    Dim StartRow As Integer                                                 'Start row of trade data
    Dim DataColumn As Integer                                               'Column of trade data
    Dim TradeNumber As Integer                                              'Number listed off trade data

    Dim s As Shape
    Dim r As Range
    Dim i As Variant
    Dim a, b, c, m, n, t, y, z As Integer

    StartRow = 20

    Windows(JournalTitle).Activate
    Application.Goto Sheets("Journal").Range("A1"), True
    Sheets("Journal").Unprotect
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

'    '~~> Check if image is to be entered
'    i = MsgBox("Do you want to check for missing trade data?", 4 + 32, "Check Trade Data")
'    If i = 7 Then: MsgBox "Check for missing trade data canceled.", 0 + 64, "Canceled": GoTo ex

    '~~> populate setup array
    With ThisWorkbook.Sheets("Range")
        i = 0
        For Each r In .Range("Setups").Cells
            If r.Value = "" Then Exit For
            setup(i) = r.Value
            i = i + 1
        Next r
    End With

    '~~> create array of trade addresses
    i = 0: y = 0: z = 0
    With Sheets("Journal")
        For Each r In .Range("Journal_Data").Cells
            If Not IsEmpty(r.Value) Then
                While y < 18
                    t = r.Offset(-y, -2)
                        If t <> 0 And c <> r.Offset(-y).Address And r.Offset(-y, -2).Row >= StartRow Then
                            TradeAddress(z) = r.Offset(-y).Address
                            z = z + 1
                            c = r.Offset(-y).Address
                            y = 18
                        End If
                    y = y + 1
                Wend
                y = 0
            End If
        Next r
    End With

    '~~> check for missing items
    For y = 0 To z

        'acquire address of image
        If Len(TradeAddress(y)) > 0 Then
            a = TradeAddress(y)
        End If

        'if there are no images then exit
        If a = "" Then
            MsgBox "There are no trades to" & Chr(10) _
                 & " check missing items.", 0 + 64, "No Missing Items"
            GoTo ex
        End If

        'acquire column and row of data
        DataColumn = Range(a).Column
        DataRow = Range(a).Row

        'check if data address was already checked, if so go to next address
        If t = Cells(DataRow, DataColumn).Address Then

            t = Cells(DataRow, DataColumn).Address

            'acquire setup(i) number
            i = (DataColumn - 3) / 12

            'acquire the number of setup trade
            TradeNumber = Cells(DataRow, DataColumn).Offset(0, -2).Value

            'check for missing data in datacolumn
            c = 0: n = 0
            Cells(DataRow, DataColumn).Activate
            While n < 18
                If IsEmpty(ActiveCell) Then c = c + 1
                ActiveCell.Offset(1).Activate
                n = n + 1
            Wend

            'record all missing items to setup
            If c = 1 Then aSetup(i, TradeNumber) = "        #" & TradeNumber & " is missing " & c & " item"
            If c > 1 Then aSetup(i, TradeNumber) = "        #" & TradeNumber & " is missing " & c & " items"

        End If

    Next y

    '~~> tally all missing items
    For i = 0 To 15
        For n = 1 To 100
            If aSetup(i, n) <> "" Then
                If bSetup(i) = "" Then
                    bSetup(i) = aSetup(i, n)
                Else
                    bSetup(i) = bSetup(i) & Chr(10) & aSetup(i, n)
                End If
            End If
        Next n
    Next i

    For i = 0 To 15
        If bSetup(i) <> "" Then cSetup(i) = setup(i) & Chr(10) & bSetup(i)
    Next i

    For i = 0 To 15
        If cSetup(i) <> "" Then
            If dSetup = "" Then dSetup = cSetup(i) Else dSetup = dSetup & Chr(10) & cSetup(i)
        End If
    Next

    '~~> report result
    If c > 0 Then
        MsgBox dSetup, 0, "Missing Raw Trade Data"
    Else
        MsgBox "Currently there are no missing trade items." & Chr(10) _
             & "All trades are complete.", 0 + 64, "Trade Data Complete"
    End If

ex:

    Application.Goto Sheets("Journal").Range("A1"), True
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    
End Sub

Sub m_OptImage()
'checks if image is missing for corresponding optimum trade data

    Dim s As Shape
    Dim i, r, t As Integer

    Dim nDate As Date
    Dim aImage As String
    Dim tLabel As Integer
    Dim iDate(51) As Date
    Dim wDate(51) As String
    Dim StartDate As String
    Dim ImageColumn As Integer
    Dim mAddress(51) As String
    Dim DateMissing(51) As Date
    Dim LabelMissing(51) As Integer
    Dim ImageMissing(51) As String
    Dim ImageAddress(51) As String

    StartDate = "ET20"
    ImageColumn = 150

    Windows(JournalTitle).Activate
    Application.Goto Sheets("Journal").Range("A1"), True
    Sheets("Journal").Unprotect
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

'    '~~> verify that check for missing images
'    i = MsgBox("Do you want to check for optimal missing images?", 4 + 32, "Check Images")
'    If i = 7 Then: MsgBox "Check for missing images canceled.", 0 + 64, "Canceled": GoTo ex

    '~~> acquire date
    nDate = Date

    '~~> populate date array
    i = 0
    While i < 52
        iDate(i) = Range(StartDate).Offset(r, 0)
        i = i + 1
        r = r + 19
    Wend

    '~~> populate weekly date array
    i = 0: r = 0
    While i < 52
        wDate(i) = Range(StartDate).Offset(r - 1, -4)
        wDate(i) = Application.Substitute(wDate(i), " ", "", 3)
        wDate(i) = Application.Substitute(wDate(i), " ", "", 2)
        i = i + 1
        r = r + 19
    Wend

    '~~> populate array with addresses of all images
    i = 0: r = 0
    While i < 52
        mAddress(i) = Range(StartDate).Offset(r, -4).Address
        i = i + 1
        r = r + 19
    Wend

    '~~> populate array with weekly optimal trade images
    With Sheets("Journal")
        On Error Resume Next
        For Each s In .Shapes
            If ImageColumn = s.TopLeftCell.Column Then
                tLabel = s.TopLeftCell.Offset(0, -1)
                ImageAddress(tLabel - 1) = s.TopLeftCell.Address
            End If
        Next s
    End With

    '~~> check for missing images up to current date
    i = 0
    While i < 52
        If ImageAddress(i) <> mAddress(i) And iDate(i) < nDate Then
            t = t + 1
            DateMissing(i) = Range(mAddress(i)).Offset(-1, 4)
            LabelMissing(i) = Range(mAddress(i)).Offset(0, -1).Value
            ImageMissing(i) = "Image #" & LabelMissing(i) & " for " & wDate(i)
        End If
        i = i + 1
    Wend

    '~~> final assembly of all missing images
    i = 0: r = 0
    While i < 52
        If ImageMissing(i) <> "" Then
            If r = 0 Then aImage = "   " & ImageMissing(i)
            If r > 0 Then aImage = aImage & Chr(10) & "   " & ImageMissing(i)
            r = r + 1
        End If
        i = i + 1
    Wend

    'alert user of found entry
    If t = 0 Then
        MsgBox "There are no missing images for" & Chr(10) _
             & "  the optimal weekly trade.", 64, "Investment"
    ElseIf t = 1 Then
        MsgBox "Missing Image Found..." & Chr(10) & aImage, 64, "Optimal Weekly Trade"
    ElseIf t > 1 Then
        MsgBox "There are " & t & " missing images..." & Chr(10) & aImage, 64, "Optimal Weekly Trade"
    End If

ex:

    Application.Goto Sheets("Journal").Range("A1"), True
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

End Sub

Sub c_TradeDate()
'checks for any Saturday dates in trade data

    Dim c As Range
    Dim b As String
    Dim a, i, p, s As Integer

    Dim aDate As String
    Dim SatDate As String
    Dim setup(15) As String
    Dim SetupDate(1600) As String

    Sheets("Journal").Unprotect
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    'acquire setup entered & populate setup array
    With ThisWorkbook.Sheets("Range")
        For Each c In .Range("Setups").Cells
            If c.Value = "" Then Exit For
            setup(i) = c.Value
            i = i + 1
        Next c
    End With

    'font red if saturday dates
    For Each c In Range("Journal_Data")
        If IsDate(c) Then
            If Weekday(c) = 7 Then
                GoSub st
                If c.Font.Color <> 255 Then
                    With c.Font
                        .Color = -16777060
                    End With
                End If
            Else
                With c.Font
                    .ColorIndex = xlAutomatic
                End With
            End If
        End If
    Next

    'final tally of all dates found
    For i = 0 To 15
        If SetupDate(i) <> "" Then
            If i = 0 Then aDate = "   " & setup(i) & ": " & SetupDate(i)
            If i > 0 Then aDate = aDate & Chr(10) & "   " & setup(i) & ": " & SetupDate(i)
        End If
    Next i

    'alert user of found entry
    If s = 1 Then
        MsgBox "Saturday Date Found:" & Chr(10) & aDate, 64, "Investment"
    ElseIf s > 1 Then
        MsgBox "Saturday Dates Found:" & Chr(10) & aDate, 64, "Investment"
    End If

    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

Exit Sub

st:

    'acquire column number of date and thereby setup
    a = c.Column
    a = (a - 3) / 12

    'acquire number of setup
    b = c.Offset(-3, -2)
    If b = "" Then b = c.Offset(-16, -2)

    'keep tally of sat dates
    If s = 0 Or a > p Then
        SatDate = b
    ElseIf a = p Then
        SatDate = ", " & b
    End If
    If s = 0 Then
        SetupDate(a) = SatDate
    ElseIf SatDate <> "" Then
        SetupDate(a) = SetupDate(a) & SatDate
    End If

    'keep track of setup change
    p = a

    'keep track of sat dates
    s = s + 1

    'continue loop to find sat dates
    Return

End Sub
