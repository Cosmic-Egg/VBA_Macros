Sub SortPCDataLeftToRight()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim sortRow As Long, firstCol As Long, lastCol As Long
    
    Set ws = ThisWorkbook.Sheets("PC_Data")
    sortRow = 1
    firstCol = 5 ' Column E
    lastCol = ws.Cells(sortRow, ws.Columns.Count).End(xlToLeft).Column
    
    If lastCol < firstCol Then
        MsgBox "No data to sort in columns E and beyond.", vbExclamation
        Exit Sub
    End If

    ' Define the full data range to sort (from row 1 to the last used row, columns E to last)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set dataRange = ws.Range(ws.Cells(1, firstCol), ws.Cells(lastRow, lastCol))

    ' Perform left-to-right sort by row 1
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(sortRow, firstCol), ws.Cells(sortRow, lastCol)), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        .SetRange dataRange
        .Header = xlNo
        .Orientation = xlLeftToRight
        .MatchCase = False
        .Apply
    End With
End Sub
