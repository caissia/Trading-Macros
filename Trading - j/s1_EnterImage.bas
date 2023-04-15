Attribute VB_Name = "s1_EnterImage"
Option Explicit
'Required to check clipboard for image
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long

Sub Image_Entry()
Attribute Image_Entry.VB_ProcData.VB_Invoke_Func = "I\n14"

    Dim i As Variant                                                    'Misc variable
    Dim n As Variant                                                    'Setup number
    Dim O As Variant                                                    'Offset number
    Dim t As Variant                                                    'Misc variable
    Dim Q As Variant                                                    'Query variable
    Dim c As Range                                                      'Loop variable
    Dim s As Shape                                                      'Image variable

    Dim setup1 As String                                                'Setup of optimal trade
    Dim symbol As String                                                'Symbol of optimal trade
    Dim datetime1 As String                                             'Date of optimal trade
    Dim datetime2 As String                                             'Day of optimal trade
    Dim datetime3 As String                                             'Time of optimal trade
    Dim piprange As Variant                                             'Pip range of optimal trade
    Dim TimeFrame As String                                             'Time frame of optimal trade
    Dim Direction As String                                             'Direction of optimal trade

    Dim wDate As Date                                                   'First Sunday of current year
    Dim LastRow As Integer                                              'For last row used in Journal
    Dim StartRow As Integer                                             'First row of data inputs start
    Dim PasteCell As String                                             'Cell address of permanent paste
    Dim DataColumn As Integer                                           'Column to check trade data
    Dim ImageColumn As Integer                                          'Column of image paste
    Dim PreviousData As String                                          'Range of previous data
    Dim PreviousImage As String                                         'Address of previous image

    Dim wsp As String                                                   'Worksheet to paste variable
    Dim setup(15) As String                                             'Array of up to 16 trade setups
    Dim ShapeAddress(1600) As String                                    'Array of all shape addresses

    StartRow = 20                                                       'First row to enter image
    LastRow = 1901                                                      'Last row of image entry
    wsp = "Journal"                                                     'Sheet to log recent trade

    Application.Goto Range("A1"), True
    Application.Goto Sheets("Journal").Range("L19")
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    '~~> check if image is to be entered
    i = MsgBox(Space(5) & "Do you want to enter an image for the last trade placed?", 4, "Insert Image")
        If i = 7 Then GoTo ex

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '~~> check if the clipboard holds an image
    Const CF_BITMAP = 2
    Dim DataCB As Long
    Dim IsBmp As Long

    'open clipboard
    DataCB = OpenClipboard(0&)

    'check if we were successful
    If DataCB <> 0 Then
        'test if the data in Clipboard is an image by
        'trying to get a handle to the Bitmap
        IsBmp = GetClipboardData(CF_BITMAP)

        'if Bitmap not found
        If IsBmp = 0 Then
            MsgBox "Image not found on clipboard." & Chr(10) & "Image entry canceled.", 0, "Canceled"
            DataCB = CloseClipboard
            GoTo ex
        End If
    End If

    'close clipboard
    DataCB = CloseClipboard

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '~~> Acquire setup array for inputbox
        i = 1
        For Each c In Range("Setups")
            If c.Value = "" Then i = i - 1: Exit For
            setup(i) = c
            If i = 1 Then
                t = "" & i & "        ~~      " & setup(i) & Chr(10)
            End If
            If i > 1 And Len(i) = 1 Then
                t = t & i & "        ~~      " & setup(i) & Chr(10)
            End If
            If i > 1 And Len(i) = 2 Then
                t = t & i & "      ~~      " & setup(i) & Chr(10)
            End If
            i = i + 1
        Next c

    '~~> Find column to paste by way of Setup
    t = t & Chr(10) & "        or" & Chr(10) & Chr(10) & "mm/dd/yy      ~~      date of optimal weekly trade "
    i = InputBox(t, "Trade Setup of Image?", "Enter a number or date")

    '~~> determine if date was entered... if so it is the "optimal" trade of the week
    If IsDate(i) Then GoTo opt

    '~~> determine column number to paste trade data
    If IsNumeric(i) = False Or i < 0 Or i > 16 Or i = "" Then MsgBox "A trade setup or date was not entered." & Chr(10) & "Image entry will terminate.", 0, "Investment": GoTo ex
    If i = 1 Then: DataColumn = 3: Else: DataColumn = ((i - 1) * 12) + 3
    If i = 1 Then: ImageColumn = 4: Else: ImageColumn = ((i - 1) * 12) + 4
    n = i   'retain setup of image

    '~~> goto setup column
    Application.Goto Reference:=Sheets(wsp).Cells(StartRow, DataColumn).Offset(-1, -2), Scroll:=True
    ActiveCell.Offset(1).Select
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False

    '~~> check whether image is for the 1st or 2nd slot
    i = MsgBox("  Open Trade ~~ 'yes'" & Chr(10) _
             & "  Close Trade ~~ 'no'", 3, "Opened or Closed Trade?")

    If i <> 6 And i <> 7 Then: MsgBox "Entry canceled." & Chr(10) & "Image entry will terminate.", 0, "Investment": GoTo ex
    If i = 7 Then
        ImageColumn = ImageColumn + 4           'shifts over to 2nd slot
        O = 7                                   'offset number to get to column number of trade
    Else
        O = 3                                   'offset number to get to column number of trade
    End If

    '~~> populate all shapes array addresses
    i = 0
    With Sheets("Journal")
        For Each s In .Shapes
            ShapeAddress(i) = s.TopLeftCell.Address
            On Error Resume Next
            i = i + 1
        Next s
    End With

    '~~> check if image exists at PasteCell location
    '    if image exist goto next image entry row
    i = 0: t = StartRow
    For i = 0 To 100
        Q = Filter(ShapeAddress, Cells(t, ImageColumn).Address)
        If UBound(Q) >= 0 And Cells(t, ImageColumn).Address <> "" Then t = t + 19 Else: Exit For
    Next i
    PasteCell = Cells(t, ImageColumn).Address(False, False)

    '~~> check if PasteCell row is part of the start of image rows and not beyond image boundary
    i = (Range(PasteCell).Row Mod 19) - 5
    If Range(PasteCell).Row < StartRow Or Range(PasteCell).Row > LastRow Then
        MsgBox "Paste address is outside of the image area." & Chr(10) _
            & "Image entry canceled.", 0, "Canceled"
        PasteCell = Cells(StartRow, ImageColumn).Address(False)
        GoTo ex
    End If
    If i <> 0 Then
        MsgBox "Paste cell row is incorrect." & Chr(10) & "Image entry canceled.", 0, "Canceled"
        GoTo ex
    End If

    '~~> check if previous image is present... if so continue else abort
    If StartRow <> Range(PasteCell).Row Then
        If O = 3 Then PreviousImage = Range(PasteCell).Offset(-19, 4).Address
        If O = 7 Then PreviousImage = Range(PasteCell).Offset(-19, -4).Address

        Q = Filter(ShapeAddress, PreviousImage)
        If UBound(Q) < 0 And PreviousImage <> "" Then i = 0 Else i = 1

        If i = 0 Then
            If O = 3 Then O = 7 Else: If O = 7 Then O = 3
            i = Range(PreviousImage).Offset(0, -O)
            If O = 3 Then i = i + 0.1 Else i = i + 0.2
            Application.Goto Reference:=Sheets(wsp).Range(PreviousImage).Offset(-1, -O), Scroll:=True
            Range(PreviousImage).Offset(0, -O).Activate
            Application.ScreenUpdating = True
            MsgBox "Previous image #" & i & " is missing." & Chr(10) & "Image entry canceled.", 0, "Canceled"
            GoTo ex
        End If
    End If

    '~~> acquire range of previous trade
    If O = 3 Then PreviousData = Range(PasteCell).Offset(-19, -1).Address(False, False)
    If O = 7 Then PreviousData = Range(PasteCell).Offset(-19, -5).Address(False, False)
    PreviousData = PreviousData & ":" & Range(PreviousData).Offset(17).Address(False, False)

    '~~> check if any cells in range are empty
    i = Range(PasteCell).Row
    If i <> StartRow Then
        i = 0
        With ThisWorkbook.Sheets(wsp)
            For Each c In .Range(PreviousData).Cells
                If IsEmpty(c.Value) Then i = i + 1
            Next c
        End With
        Else: i = 0
    End If

    '~~> alert user if any empty cells found
    If i <> 0 And Range(PasteCell).Row <> StartRow Then

        t = Cells(Range(PreviousData).Row, DataColumn).Offset(0, -2).Address(False, False)
        Application.Goto Reference:=Sheets(wsp).Range(t).Offset(-1), Scroll:=True
        Range(t).Select
        Application.ScreenUpdating = True

        t = Range(PreviousData).Row
        t = Cells(t, DataColumn).Offset(0, -2)

        If i = 1 Then
            MsgBox setup(n) & " #" & t & " is missing " & i & " item." & Chr(10) & "Image entry canceled.", 0, "Canceled"
        End If
        If i > 1 Then
            MsgBox setup(n) & " #" & t & " is missing " & i & " items." & Chr(10) & "Image entry canceled.", 0, "Canceled"
        End If

        GoTo ex
    End If

    '~~> confirm paste location and paste the image
    t = Cells(Range(PasteCell).Row, DataColumn).Offset(-1, -2).Address(False, False)
    Application.Goto Reference:=Sheets(wsp).Range(t), Scroll:=True
    Range(PasteCell).Activate
    Application.ScreenUpdating = True

    i = MsgBox("Cell location for image is selected." & Chr(10) _
             & "Continue image entry?", 4, "Confirm Image Entry")
    If i <> 6 Then Application.Goto Sheets(wsp).Range("A1"), True: GoTo ex

    Application.ScreenUpdating = False

    Range(PasteCell).Activate
    On Error Resume Next
    Sheets(wsp).Paste
    If Err Then
        MsgBox "Image error." & Chr(10) & "Image entry canceled.", 0, "Canceled"
        Err.Clear
        GoTo ex
    End If

    '~~> move the image to fit
    Selection.ShapeRange.IncrementLeft 1  'move it left
    Selection.ShapeRange.IncrementTop 1   'move it down

    '~~> clear clipboard
    'Application.CutCopyMode = False

    '~~> acquire range of image and place border
    i = Range(PasteCell).Address & ":" & Range(PasteCell).Offset(17, 3).Address
    Sheets(wsp).Range(i).Select
    GoSub border

    '~~> shift focus off image 'first acquire column number of setup
    PasteCell = Range(PasteCell).Offset(0, -O).Address(False, False)
    Application.Goto Reference:=Sheets(wsp).Range(PasteCell).Offset(-1), Scroll:=True
    Range(PasteCell).Activate

    '~~> acquire slot number of pasted image
    i = Range(PasteCell)
    If O = 3 Then i = i + 0.1 Else i = i + 0.2

    '~~> notify of succesful image entry
    Application.ScreenUpdating = True       'in order to show the image
    MsgBox "The " & setup(n) & " image number " & i & "" & Chr(10) _
         & "  has been entered successfully!", 0, "Investment"
    GoTo dn

ex:

    'Application.Goto Sheets("Journal").Range("A1"), True
    'Application.Goto Sheets("Journal").Range("L19")
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

Exit Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

opt:

'    i = Weekday(DateSerial(Year(Date), 1, 1))
'    'acquire first sunday of current year in serial date format
'    wDate = DateSerial(Year(Date), 1, 1)
'    While i <> 1
'        i = Weekday(wDate)
'        wDate = wDate + 1
'    Wend

    'acquire first sunday of current year in serial date format
    wDate = Sheets("Journal").Range("EO19")

    '~~> check for the current year
    If Year(i) <> Year(Now) Then
        MsgBox Space(4) & "The date must be within the journal's year, " & Year(i) & "." & Chr(10) _
             & Space(26) & "Image entry canceled.", 0, "Canceled"
        GoTo ex
    End If

    '~~> acquire the setup of the image
    setup1 = InputBox("Enter the setup of this image.", "Setup of Image?")
    If WorksheetFunction.IsText(setup1) = False Then
        MsgBox "The setup range must be text." & Chr(10) _
             & "  User input: " & setup1 & "" & Chr(10) _
             & "Image entry canceled.", 0, "Canceled"
        GoTo ex
    End If

    '~~> acquire the currency pair of the image
    symbol = InputBox("Enter the currency pair of this image.", "Currency Pair of Image?")
    If WorksheetFunction.IsText(symbol) = False Then
        MsgBox "The currency pair must be text." & Chr(10) _
             & "  User input: " & symbol & "" & Chr(10) _
             & "Image entry canceled.", 0, "Canceled"
        GoTo ex
    End If

    '~~> acquire the time frame of the image
    TimeFrame = InputBox("Enter the time frame of this image.", "Timeframe of Image?")
    If WorksheetFunction.IsText(TimeFrame) = False Then
        MsgBox "The timeframe must be text." & Chr(10) _
             & "  User input: " & TimeFrame & "" & Chr(10) _
             & "Image entry canceled.", 0, "Canceled"
        GoTo ex
    End If

    '~~> acquire the time of the image
    datetime1 = InputBox("Enter the Date of this image.", "Date of Image?")
    If WorksheetFunction.IsText(datetime1) = False Then
        MsgBox "The time must be entered as text." & Chr(10) _
             & "  User input: " & datetime1 & "" & Chr(10) _
             & "Image entry canceled.", 0, "Canceled"
        GoTo ex
    End If

    '~~> acquire the weekday of the image
    datetime2 = InputBox("Enter the weekday of this image.", "Weekday of Image?")
    If WorksheetFunction.IsText(datetime2) = False Then
        MsgBox "The weekday must be text." & Chr(10) _
             & "  User input: " & datetime2 & "" & Chr(10) _
             & "Image entry canceled.", 0, "Canceled"
        GoTo ex
    End If

    '~~> acquire the time of the image
    datetime3 = InputBox("Enter the time of this image.", "Time of Image?")
    If WorksheetFunction.IsText(datetime3) = False Then
        MsgBox "The time must be entered as text." & Chr(10) _
             & "  User input: " & datetime3 & "" & Chr(10) _
             & "Image entry canceled.", 0, "Canceled"
        GoTo ex
    End If

    '~~> acquire the pip range of the image
    piprange = InputBox("Enter the pip range of this image.", "Piprange of Image?")
    If IsNumeric(piprange) = False Then
        MsgBox "The pip range must be a number." & Chr(10) _
             & "  User input: " & piprange & "" & Chr(10) _
             & "Image entry canceled.", 0, "Canceled"
        GoTo ex
    End If

    '~~> acquire the direction of the image
    Direction = InputBox("Enter the direction of this image.", "Direction of Image?")
    If Direction <> "long" And Direction <> "short" Then
        MsgBox "The direction must be either long or short." & Chr(10) _
             & "  User input: " & Direction & "" & Chr(10) _
             & "Image entry canceled.", 0, "Canceled"
        GoTo ex
    End If

    '~~> acquire number of image slot
        t = 0
        Do While t < 52
            If wDate - 7 <= i And i <= wDate Then Exit Do
            wDate = wDate + 7
            t = t + 1
        Loop

    '~~> determine cell to paste image
    If t = 1 Then t = 20 Else t = t * 19 + 1
    ImageColumn = 146

    '~~> goto weekly optimal trade column
    Application.Goto Reference:=Sheets(wsp).Cells(StartRow, ImageColumn).Offset(-1, -13), Scroll:=True
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False

    '~~> check cell paste address is correct
    i = (Cells(t, ImageColumn).Row Mod 19) - 1
    If t < StartRow Or t > LastRow Then
        MsgBox "Paste address is outside of the image area." & Chr(10) _
            & "Image entry canceled.", 0, "Canceled"
        GoTo ex
    End If
    If i <> 0 Then
        MsgBox "Paste cell row is incorrect." & Chr(10) _
        & "Image entry canceled.", 0, "Canceled"
        GoTo ex
    End If

    '~~> check if image exists at PasteCell location
    With Sheets(wsp)
        On Error Resume Next
        For Each s In .Shapes

            If Not Application.Intersect(s.TopLeftCell, .Cells(t, ImageColumn)) Is Nothing Then
                Application.Goto Sheets(wsp).Cells(t, ImageColumn).Offset(-1, -1), True
                ActiveCell.Offset(1).Activate
                ActiveWindow.SmallScroll ToLeft:=16         'move msgbox away from image
                Application.ScreenUpdating = True: Application.ScreenUpdating = False
                i = Cells(t, ImageColumn).Offset(-1).Value
                i = MsgBox(Space(8) & "An image exists for " & i & "." & Chr(10) _
                 & Space(5) & "Would you like to replace the image?", 4, "Replace or Cancel")
                If i = 6 Then
                    s.Delete
                    n = True    'records window scroll to avoid duplicating the move
                    Exit For
                Else
                    MsgBox "Image entry canceled.", 0, "Canceled"
                    GoTo ex
                End If
            End If
        Next s
    End With

    '~~> set cell format to string before pasting time
    Cells(t, ImageColumn).Offset(4, 4).NumberFormat = "@"

    '~~> paste data of optimal trade
    Cells(t, ImageColumn).Offset(0, 4) = setup1
    Cells(t, ImageColumn).Offset(1, 4) = symbol
    Cells(t, ImageColumn).Offset(2, 4) = TimeFrame
    Cells(t, ImageColumn).Offset(3, 4) = datetime2
    Cells(t, ImageColumn).Offset(4, 4) = datetime3
    Cells(t, ImageColumn).Offset(5, 4) = piprange
    Cells(t, ImageColumn).Offset(6, 4) = datetime1
    Cells(t, ImageColumn).Offset(7, 4) = Direction

    '~~> check if for blank cells in data
    i = Range(Cells(t, ImageColumn).Offset(0, 4).Address & ":" & Cells(t, ImageColumn).Offset(6, 4).Address).Address
    For Each c In Range(i)
        If c = "" Or c = " " Or c = "  " Or c = "   " Then
            MsgBox "Some image data is missing. Check and try again", 0, "Warning"
        End If
    Next c

    '~~> set cell format to invisible after pasting data
    For i = 0 To 6
        Cells(t, ImageColumn).Offset(i, 4).NumberFormat = ";;;"
    Next i

    '~~> paste image
    Cells(t, ImageColumn).Activate
    On Error Resume Next
    ActiveSheet.Paste
    If Err Then
        MsgBox "Image error." & Chr(10) & "Image entry canceled.", 0, "Canceled"
        Err.Clear
        GoTo ex
    End If

    '~~> move the image to fit
    Selection.ShapeRange.IncrementLeft 1  'move it left
    Selection.ShapeRange.IncrementTop 1   'move it down

    '~~> clear clipboard
    'Application.CutCopyMode = False

    '~~> acquire range of image and place border
    c = Cells(t, ImageColumn).Address & ":" & Cells(t, ImageColumn).Offset(17, 3).Address
    Range(c).Select
    GoSub border

    '~~> shift focus off image
    Application.Goto Sheets(wsp).Cells(t, ImageColumn).Offset(-1, -1), True
    ActiveCell.Offset(1).Activate
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False

    '~~> acquire slot number of pasted image
    i = Cells(t, ImageColumn).Offset(0, -1).Value

    '~~> notify of successful image entry
    If Not n Then ActiveWindow.SmallScroll ToLeft:=16   'move msgbox away from image
    Application.ScreenUpdating = True                   'in order to show the image
    MsgBox Space(5) & "The Optimal Weekly Trade image, #" & i & "," & Chr(10) _
         & Space(12) & "has been entered successfully!", 0, "Investment"

dn:

    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

Exit Sub

border:

    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.249946592608417
        .Weight = xlThin
    End With
    Return

End Sub
