Sub GetDistrictInfo()
'
'4/24/2018
'Get the District Info from B2 and filldown this must be run after KuTools of Importing spreadsheets is run.
' After combining with KuTools all the CSV files, Add the following Macros to the cobmined excel documenbt.

'
    Dim ws_count As Integer
    Dim i As Integer
    Dim title As String
    Dim length As Integer
    
        ws_count = ActiveWorkbook.Worksheets.Count
        For i = 2 To ws_count

            Sheets(i).Activate
            Columns("A:A").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("B2").Select
            length = Len(ActiveCell.Text) - 13
            title = Mid(ActiveCell.Text, 13, length)
            Range("A4").Select
            ActiveCell.Value = title
            Range("A4:A" & lRow).Select
            Selection.FillDown
        Next i
End Sub






