Attribute VB_Name = "s4_SearchFilter"
Option Explicit
'contains 3 macros: ClearResult, ClearSearch, Query

Sub Query()
's4, applies an AutoFilter to the Data sheet based on the criteria on the Query sheet
'then copies/pastes the visible rows onto Query sheet result section

    Dim wsRaw As Worksheet
    Dim wsQ   As Worksheet
    Dim rng   As Range
    Dim cell  As Range

    'setup
    Application.ScreenUpdating = False
    ThisWorkbook.Sheets("Query").Unprotect
    ThisWorkbook.Sheets("Data").Unprotect

    Set wsRaw = ThisWorkbook.Sheets("Data")
    Set wsQ = ThisWorkbook.Sheets("Query")
    Set rng = wsQ.Range("B5:AA5")                                               'Place where the criteria are
    wsQ.Range("QueryResults").ClearContents                                     'Clears contents in Query worksheet

    'prepare Autofilter
    With wsRaw
        .AutoFilterMode = False
        .Rows(3).AutoFilter

        For Each cell In rng
            If cell <> "" Then _
                .Rows(3).AutoFilter Field:=cell.Column, Criteria1:=cell.Text    'The -0 shifts from the Query to Data wksht
        Next cell

        .Range("B4:AA2003").SpecialCells(xlVisible).Copy                        'Start of Data Table to search
        wsQ.Range("B12").PasteSpecial xlPasteValues                             'Place to paste filtered results

        .AutoFilterMode = False
    End With

    'for search response
    Range("A5") = "TRUE"

    ThisWorkbook.Sheets("Query").Protect
    ThisWorkbook.Sheets("Data").Protect
    Application.ScreenUpdating = True
    Range("A1").Select
    Range("B4").Select

End Sub

Sub ClearResult()

    Application.ScreenUpdating = False

    'delete previous results if any
    Range("QueryResults").ClearContents

    Application.ScreenUpdating = True

    'for search response
    Range("A5") = "FALSE"

    Range("A1").Select
    Range("C5").Select

End Sub

Sub ClearSearch()

    Application.ScreenUpdating = False

    'delete previous results/search terms
    Range("QueryResults").ClearContents
    Range("QuerySearch").ClearContents

    Application.ScreenUpdating = True

    'for search response
    Range("A5") = "FALSE"

    Range("A1").Select
    Range("C5").Select

End Sub

