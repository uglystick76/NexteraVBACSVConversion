Sub PullDistrictAndBuildingCodes()

Dim LastRow As Long
Dim i As Integer
Dim ws_count As Integer
Dim firstp As Integer
Dim lastp As Integer
Dim mytext As String

ws_count = ActiveWorkbook.Worksheets.Count
For i = 2 To ws_count
    Sheets(i).Activate
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    With Columns("B")
        .ColumnWidth = 15
        .NumberFormat = "0"
    End With
    
    With Columns("C")
        .ColumnWidth = 15
        .NumberFormat = "0"
    End With

    With Columns("G")
        .ColumnWidth = 15
        .NumberFormat = "0"

    End With
    Range("B4").Select
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' j is set to 4 because the first four rows are title rows.'
    For j = 4 To LastRow
        mytext = Range("A" & j).Value
               
               
        firstp = InStr(1, mytext, "(", 1)
        lastp = InStr(1, mytext, ")", 1)
        CopyText = Mid(mytext, firstp + 1, lastp - firstp - 1)
        CopyText2 = Right(mytext, 12)
        Range("B" & j).Value = CopyText
        Range("C" & j).Value = CopyText2
        CopyText = Empty
        CopyText2 = Empty
    Next j
    
    LastRow = 0
    
    
Next i

End Sub