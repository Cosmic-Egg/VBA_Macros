Sub AccessRangeFromAnotherWorkbook()
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim targetRange As Range
    Dim valueFromRange As Variant
    
    ' Specify the name of the workbook and worksheet
    On Error Resume Next ' In case the workbook is already open
    Set sourceWorkbook = Workbooks("SourceWorkbook.xlsx")
    On Error GoTo 0 ' Turn error handling back on

    ' If the workbook is not open, open it
    If sourceWorkbook Is Nothing Then
        Set sourceWorkbook = Workbooks.Open("C:\Path\To\Your\SourceWorkbook.xlsx")
    End If

    ' Reference the specific worksheet
    Set sourceWorksheet = sourceWorkbook.Worksheets("Sheet1") ' Change as needed
    
    ' Access a specific range (e.g., A1)
    Set targetRange = sourceWorksheet.Range("A1")
    
    ' Get the value from the range
    valueFromRange = targetRange.Value
    
    ' Output the value to the Immediate Window (Ctrl + G to view)
    Debug.Print "Value from SourceWorkbook.xlsx, Sheet1, A1: " & valueFromRange

    ' Optionally close the workbook (if opened in this script)
    ' sourceWorkbook.Close SaveChanges:=False
End Sub

