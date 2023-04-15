Attribute VB_Name = "sx_Miscellanea"
Option Explicit
'contains 11 macros: modSetIcon, Open_SetIcon, Open_ResetIconToExcel, IsFileOpen, IsProgramRunning
'                    TerminateProcess, GetFormula, RankArray, Reminder, Toggle

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modSetIcon
' By Chip Pearson, chip@cpearson.com, www.cpearson.com/SetIcon.aspx
' This module contains code to change the icon of the Excel main
' window. The code is compatible with 64-bit Office.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If VBA7 And Win64 Then
'''''''''''''''''''''''''''''
' 64 bit Excel
'''''''''''''''''''''''''''''
Private Declare PtrSafe Function SendMessageA Lib "user32" _
      (ByVal hwnd As LongPtr, ByVal wMsg As LongLong, ByVal wParam As LongLong, _
      ByVal lParam As LongLong) As LongPtr

Private Declare PtrSafe Function ExtractIconA Lib "shell32.dll" _
      (ByVal hInst As LongPtr, ByVal lpszExeFileName As String, _
      ByVal nIconIndex As LongPtr) As Long

Private Const ICON_SMALL = 0&
Private Const ICON_BIG = 1&
Private Const WM_SETICON = &H80

#Else
'''''''''''''''''''''''''''''
' 32 bit Excel
'''''''''''''''''''''''''''''
Private Declare PtrSafe Function SendMessageA Lib "user32" _
      (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, _
      ByVal lParam As Long) As Long

Private Declare PtrSafe Function ExtractIconA Lib "shell32.dll" _
      (ByVal hInst As Long, ByVal lpszExeFileName As String, _
      ByVal nIconIndex As Long) As Long

Private Const ICON_SMALL As Long = 0&
Private Const ICON_BIG As Long = 1&
Private Const WM_SETICON As LongPtr = &H80
#End If

Sub Open_SetIcon(filename As String, Optional Index As Long = 0)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetIcon
' This procedure sets the icon in the upper left corner of
' the main Excel window. FileName is the name of the file
' containing the icon. It may be an .ico file, an .exe file,
' or a .dll file. If it is an .ico file, Index must be 0
' or omitted. If it is an .exe or .dll file, Index is the
' 0-based index to the icon resource.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If VBA7 And Win64 Then
    ' 64 bit Excel
    Dim hwnd As LongPtr
    Dim HIcon As LongPtr
#Else
    ' 32 bit Excel
    Dim hwnd As Long
    Dim HIcon As Long
#End If
    Dim n As Long
    Dim s As String
    If Dir(filename, vbNormal) = vbNullString Then
        ' file not found, get out
        Exit Sub
    End If
    ' get the extension of the file.
    n = InStrRev(filename, ".")
    s = LCase(Mid(filename, n + 1))
    ' ensure we have a valid file type
    Select Case s
        Case "exe", "ico", "dll"
            ' OK
        Case Else
            ' invalid file type
            Err.Raise 5
    End Select
    hwnd = Application.hwnd
    If hwnd = 0 Then
        Exit Sub
    End If
    HIcon = ExtractIconA(0, filename, Index)
    If HIcon <> 0 Then
        SendMessageA hwnd, WM_SETICON, ICON_SMALL, HIcon
    End If
End Sub

Sub Open_ResetIconToExcel()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ResetIconToExcel
' This resets the Excel window's icon. It is assumed to
' be the first icon resource in the Excel.exe file.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim FName As String
    FName = Application.Path & "\excel.exe"
    Open_SetIcon FName
End Sub

Function IsFileOpen(filename As String)
' This function checks to see if a file is open or not. If the file is
' already open, it returns True. If the file is not open, it returns
' False. Otherwise, a run-time error occurs because there is
' some other problem accessing the file.

    Dim filenum As Integer, errnum As Integer

    On Error Resume Next                    ' Turn error checking off.
    filenum = FreeFile()                    ' Get a free file number.

    ' Attempt to open the file and lock it.
    Open filename For Input Lock Read As #filenum
    Close filenum                           ' Close the file.
    errnum = Err                            ' Save the error number that occurred.
    On Error GoTo 0                         ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum
        Case 0: IsFileOpen = False          ' No error occurred. File is NOT already open by another user.
        Case 70: IsFileOpen = True          ' Error number for "Permission Denied." File is already opened by another user.
       'Case Else: Error errnum             ' Another error occurred.
    End Select

End Function

Function IsProgramRunning(Process As String)
'Returns true if program process is running: e.g., check = IsProgramRunning(excel.exe)

    Dim objList As Object

    Set objList = GetObject("winmgmts:").ExecQuery("select * from win32_process where name='" & Process & "'")

    If objList.Count > 0 Then IsProgramRunning = True
    If objList.Count < 1 Then IsProgramRunning = False

End Function

Sub TerminateProcess(app_exe As String)
'This kills any process running

    Dim Process As Object
    For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = '" & app_exe & "'")
        Process.Terminate
    Next
    
    'alternative this can be used if the program was called with "vPid Shell"...~~> Call Shell("TaskKill /F /PID " & CStr(vPID), 4)
    'a way to open a program with vPID ~~> vPID = Shell("C:\Program Files (x86)\FXCM MetaTrader 4\terminal.exe")

End Sub

Function GetFormula(cell As Range) As String
'This function shows the formula as text

   GetFormula = cell.Formula

End Function

Function RankArray(arr As Variant) As Variant
'accepts array and returns array of ranked indexes
'ranks based on elements from high to low
'received arr index must start with 1

    Dim d()                                             'receives the incoming array
    Dim orig()                                          'holds the first duplicates to add tag
    Dim rank()                                          'holds the rank of the indexes of array d
    Dim skip()                                          'holds the elements already checked for duplicates

    Dim a As Variant                                    'loop variable
    Dim b As Variant                                    'misc., loop, string
    Dim c As Variant                                    'misc., loop, string
    Dim tag As Variant                                  'decimal portion to add to elements in array to rank
    Dim pass As Boolean                                 'holds if items have already been checked/skipped
    Dim duplicate As Integer                            'counts duplicates for each element in array d

    'count elements in received array
    For Each a In arr
        If a <> "" Then c = c + 1
    Next a

    'redim dynamic variables
    ReDim skip(c)
    ReDim rank(c)
    ReDim orig(c)
    ReDim d(c)

    'pass array to new array
    For a = 1 To c
        d(a) = arr(a)
    Next a

'~~> make list to rank unique if duplicate values exist

    'add tag to duplicates of original
    For a = 1 To UBound(d)
        If IsEmpty(d(a)) Then Exit For

        For b = a To UBound(d)
            If IsEmpty(d(b)) Then Exit For

            'skip items already checked
            On Error Resume Next
            For Each c In skip
                If a = c Then pass = False
            Next c
            On Error GoTo 0

            If d(a) = d(b) And a <> b And pass Then
                duplicate = duplicate + 1

                If duplicate = 1 Then
                    orig(a) = a                         'array of first values that have duplicates
                End If

                If duplicate > 0 Then
                    skip(b) = b                         'array of values already evaluated to skip
                    tag = 10 ^ -(duplicate + 1)         'decimal portion to add, e.g. 0.01, 0.001, etc
                    d(b) = d(b) + tag                   'progressively & uniquely tag duplicates
                    'c = c + 1                           'debug.print check for accuracy
                    'Debug.Print "#" & c & " found: d(" & a & ")= d(" & b & ")" & ", duplicate = " & duplicate
                End If
            End If

            pass = True
        Next b

        duplicate = 0
    Next a

    'rank first duplicates higher by adding tag
    For a = 1 To UBound(d)
        If d(a) = "" Then Exit For
        If a = orig(a) Then d(a) = d(a) + 0.1
    Next a

    'debug.print check unique list
'    For a = 1 To UBound(d)
'        If IsEmpty(d(a)) Then Exit For
'        If Len(a) = 1 Then b = "   =  " Else b = "  =  "
'        Debug.Print "d(" & a & ")" & b & d(a)
'    Next a

'''''''''''''''''''' end of making list to rank unique ''''''''''''''''''''

'~~> rank the unique list now

    'wksheet function 'large' ranks unique items
    For a = 1 To UBound(d)
        If d(a) = "" Then Exit For
        On Error Resume Next
        For b = 1 To UBound(d)
            If d(b) = "" Then Exit For
            If d(a) = Application.Large(d, b) Then
                rank(b) = a
            End If
        Next b
    Next a

    'debug.print check that values are ranked properly
'    For a = 1 To UBound(rank)
'        If rank(a) = "" Then Exit For
'        If Len(a) = 1 Then b = "   =  " Else b = "  =  "
'        If Len(rank(a)) = 1 Then c = ")   =  " Else c = ")  =  "
'        Debug.Print "Rank #" & a & b & "d(" & rank(a) & c & d(rank(a))
'    Next a

    'return the array of indexes ranked from high to low element
    RankArray = rank()

'''''''''''''''''''''' end of ranking the unique list '''''''''''''''''''''

End Function

Sub Reminder()
'reminder to "save as" the detailed statement from MT4
'before the end of each month to avoid trade data loss

    Dim LastDayOfMonth As Date
    Dim Reminder(4) As Date
    Dim DaysLeft As Integer
    Dim cMonth As String
    Dim wkDay As Integer
    Dim i, t As Integer

    LastDayOfMonth = DateSerial(Year(Date), Month(Date) + 1, 0)
    DaysLeft = LastDayOfMonth - Date
    wkDay = Weekday(LastDayOfMonth)
    cMonth = MonthName(Month(Date))

    For i = 1 To 3

        If wkDay = 7 Then
            Reminder(i) = LastDayOfMonth - i

        ElseIf wkDay > 2 Then
            t = i - 1
            Reminder(i) = LastDayOfMonth - t

        ElseIf wkDay = 2 Then
            If i < 3 Then t = i - 1 Else t = i
            Reminder(i) = LastDayOfMonth - t

        ElseIf wkDay = 1 Then
            If i = 1 Then t = i - 1 Else t = i
            Reminder(i) = LastDayOfMonth - t

        End If

    Next i

    If Date = Reminder(1) Or Date = Reminder(2) Or Date = Reminder(3) Or Date = LastDayOfMonth Then
    MsgBox "Remember to 'save as' a detailed statement from MT4" & Chr(10) _
         & "    before the end of the month to avoid losing data." & Chr(10) & Chr(10) _
         & "                       ~ " & DaysLeft & " days left in " & cMonth & " ~", 0, "Save:  MT4 DetailedStatement"
    End If

End Sub

Sub Toggle()

    If Application.CommandBars("Ribbon").Visible = True Then

        Application.ScreenUpdating = False
        ActiveSheet.Protect

        Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
        Application.DisplayFormulaBar = False

'        ActiveWindow.DisplayHorizontalScrollBar = False
'        ActiveWindow.DisplayVerticalScrollBar = False

        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayGridlines = False
        'Application.DisplayFullScreen = True       'this does not allow the toggle to work

        ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
        Application.ScreenUpdating = True

        If ActiveSheet.Name = "Range" Then ActiveSheet.Unprotect

    Else

        Application.ScreenUpdating = False
        ActiveSheet.Unprotect

        Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",True)"
        'Application.DisplayFormulaBar = True       'simple preference

        ActiveWindow.DisplayHorizontalScrollBar = True
        ActiveWindow.DisplayVerticalScrollBar = True

        'ActiveWindow.DisplayHeadings = True        'simple preference
        'ActiveWindow.DisplayGridlines = True
        'Application.DisplayFullScreen = False

'        ActiveWindow.DisplayHorizontalScrollBar = True
'        ActiveWindow.DisplayVerticalScrollBar = True

'        ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
        Application.ScreenUpdating = True

    End If

End Sub
