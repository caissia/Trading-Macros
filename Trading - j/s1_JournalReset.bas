Attribute VB_Name = "s1_JournalReset"
Option Explicit
'contains 3 macros: Reset, Set_Broker, Set_Year

Sub Reset()
'clears all data and images

    Dim a, b, c As Byte
    Dim d As Variant
    Dim s As Shape

    Application.Goto Sheets("Journal").Range("A1"), True
    Application.Goto Sheets("Journal").Range("L19")
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    'counts how many cells are empty excluding formulas
    c = Application.Count(Range("Journal_Data"))
    If c = 0 Then c = Application.Count(Range("ET20:ET1020"))

    a = MsgBox(Space(5) & "Clear all the trade data and images in this Journal?", 4, "Permanent Data Wipe")

        'prevents error if range is empty
        If a = 6 And c <> 0 Then

            Range("Journal_Data").SpecialCells(xlCellTypeConstants).ClearContents           'deletes data but not formulas
            Range("Journal_OptData").ClearContents                                          'deletes optimal trade data

            On Error GoTo r1
            With Sheets("Journal")
               For Each s In .Shapes
                   If Not Application.Intersect(s.TopLeftCell, .Range("Journal_Images")) Is Nothing Then
                     If s.Type = msoPicture Then s.Delete
                   End If
                Next s
            End With

            MsgBox Space(5) & "Successfully cleared the trade data and images.", 0, "Success"

        End If

r2:

    Application.Goto Sheets("Journal").Range("A1"), True
    Application.Goto Sheets("Journal").Range("L19")
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

Exit Sub

r1:

    Sheets("Journal").Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
    MsgBox Space(5) & "There are no images to delete.", 0, "Image Deletion Canceled"
    GoTo r2

End Sub

Sub Set_Broker()

    Dim bb As Byte  'responses
    Dim cc As Byte  'counter
    Dim nB As Byte  'number of brokers
    Dim bC As Range 'current broker
    Dim bR As Range 'brokers
    Dim cB As Range 'cycles thru brokers

    Set bR = Sheets("Range").Range("Brokers")
    Set bC = Sheets("Range").Range("I3")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    nB = 0
    For Each cB In bR       'counts number of brokers
        If cB <> 0 Then
            nB = 1 + nB
        End If
    Next cB
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    bb = MsgBox(Space(5) & "The broker listed is " & bC & "." & Chr(10) _
              & Space(5) & "Would you like to keep the current broker?", 4, "Step Four: Keep Broker?")

    If bb = 6 Then
        MsgBox Space(5) & "The broker listed, " & bC & ", will not be changed." & Chr(10) _
             & Space(5) & "The broker is listed in the following location:" & Chr(10) _
             & Space(5) & "Range I3", 0, "Broker Unchanged"
        Exit Sub
    End If

    If bb = 7 Then

        cc = 1

        MsgBox Space(5) & "There are " & nB - 1 & " brokers to choose from.", 0, "Number of Brokers"

        For Each cB In bR.Cells

            If bC <> cB Then

                bb = MsgBox(Space(5) & "Would you like to use the following broker?" & Chr(10) _
                    & cB, 4, "Choose this Broker?")

                If bb = 6 Then
                    MsgBox Space(5) & "The broker " & cB & " has been saved to the following location:" & Chr(10) _
                         & Space(5) & "Range I3", 0, "Broker Saved"
                    Sheets("Range").Range("I3") = cB
                    Exit Sub
                End If

            End If

            cc = 1 + cc

            If cc > nB Then
                MsgBox Space(5) & "A broker has not been chosen." & Chr(10) _
                     & Space(5) & "The broker listed, " & bC & ", remains on file." & Chr(10) _
                     & Space(5) & "The broker is listed in the following location:" & Chr(10) _
                     & Space(5) & "Range I3", 0, "Broker Unchanged"
                Exit Sub
            End If

        Next

    End If

End Sub

Sub Set_Year()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    Dim a As Variant
    Dim Ck As Date
    Dim st As Date      'start date
    Dim eT As Date      'end date

        Ck = DateAdd("yyyy", 2, Now)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'1. get start date from user 2. check if cancel 3. check if proper date then format it
1:
    a = MsgBox(Space(5) & "Please enter the Start Date for this Journal", 1, "Step Three: Start Date")

        If a = 2 Or a = 7 Then
            GoTo eD1
        End If
    
        frmCalendar.Show 'calls the calendar
    
        If IsDate(DC) = False Then
            GoTo ER1
        Else
            st = DC
        End If
    
        If Year(st) < Year(Now) Or st > Ck Then
            a = MsgBox("The year entered should be current." & Chr(10) _
                    & "Try again?", 1 + 64, "ERROR")
            If a = 1 Then
                GoTo 1
            Else
                GoTo eD1
            End If
        End If
    
        If Weekday(st) = 7 Then
            a = MsgBox("The date entered falls on market close." & Chr(10) _
                    & "Try again?", 1 + 64, "ERROR")
            If a = 1 Then
                GoTo 1
            Else
                GoTo eD1
            End If
        End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'1. get end date from user 2. check if cancel 3. check if proper date then format it
2:
    a = MsgBox(Space(5) & "Please enter the End Date for this Journal", 1, "Step Three: End Date")

        If a = 2 Or a = 7 Then
            GoTo eD2
        End If

        frmCalendar.Show 'calls the calendar

        'MsgBox "The End Date chosen is: " & dC

        If IsDate(DC) = False Then
            GoTo ER2
        Else
            eT = DC
        End If

        If Year(eT) < Year(Now) Or eT > Ck Then
            a = MsgBox("The year entered should be current." & Chr(10) _
                    & "Try again?", 1 + 64, "ERROR")
            If a = 6 Then
                GoTo 2
            Else
                GoTo eD2
            End If
        End If

        If eT < st Or eT = st Then
            a = MsgBox("The end date must be later than the start date." & Chr(10) _
                    & "Try again?", 1 + 64, "ERROR")
            If a = 1 Then
                GoTo 2
            Else
                GoTo eD2
            End If
        End If

        If Weekday(eT) = 7 Then
            a = MsgBox("The date entered falls on market close." & Chr(10) _
                    & "Try again?", 1 + 64, "ERROR")
            If a = 1 Then
                GoTo 2
            Else
                GoTo eD2
            End If
        End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'1. get end date from user 2. check if cancel 3. check if proper date then format it
3:

    a = MsgBox(Space(5) & "Are these dates correct?" & Chr(10) _
            & "" & Chr(10) _
            & "Start Date: " & st & "" & Chr(10) _
            & "End Date:  " & eT, 3 + 32, "Trade Journal Year")

        If a = 7 Then GoTo 1
        If a = 2 Then GoTo eD1

        If a = 6 Then
            Sheets("Range").Unprotect
            Sheets("Range").Range("C21") = st
            Sheets("Range").Range("G21") = eT
            Sheets("Range").Protect
            MsgBox "The dates have been saved to the following address:" & Chr(10) _
                & "Range C21 and G21.", 64, "Dates Saved"
        End If

    Exit Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ER1:

    a = MsgBox(Space(5) & "The entered date resulted in an error." & Chr(10) _
            & "Try again?", 1, "ERROR")
    If a = 2 Then GoTo ex
    GoTo 1

ER2:

    a = MsgBox(Space(5) & "The entered date resulted in an error." & Chr(10) _
            & "Try again?", 1, "ERROR")
    If a = 2 Then GoTo ex
    GoTo 2

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

eD1:

    a = MsgBox(Space(5) & "Do you want to cancel entering the year for this trading Journal?", 4, "Cancel?")
    If a = 7 Then
        GoTo 1
    Else
        MsgBox Space(5) & "Setting the new trade year has been canceled.", 0, "Year Setting Canceled"
        GoTo ex
    End If

eD2:

    a = MsgBox(Space(5) & "Do you want to cancel entering the year for this trading Journal?", 4, "Cancel?")
    If a = 7 Then
        GoTo 2
    Else
        MsgBox Space(5) & "Setting the new trade year has been canceled.", 0, "Year Setting Canceled"
        GoTo ex
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ex:

    MsgBox Space(5) & "The year has not been set for this new Journal." & Chr(10) _
                & Space(8) & "The dates are located at the following address:" & Chr(10) _
                & Space(8) & "Range C21 and G21.", 0, "Warning"

End Sub
