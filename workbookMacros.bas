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

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim fileName As String
    Dim fileDialog As Object
    
    ' Check if the double-clicked cell is A1 on Sheet1
    If Not Intersect(Target, Me.Range("A1")) Is Nothing And Me.Name = "Sheet1" Then 'You can edit the sheetname and range as needed
        ' Open file dialog to select a file
        Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
        
        ' Show the dialog
        If fileDialog.Show = -1 Then
            ' Get the full path of the selected file
            fileName = fileDialog.SelectedItems(1)
            
            ' Insert the full file name into the selected cell (A1)
            Target.Value = fileName
            
            ' Prevent the default double-click action (such as editing the cell)
            Cancel = True
        End If
    End If
End Sub

Sub SearchFileByName()
    Dim searchText As String
    Dim searchDirectory As String
    Dim fileName As String
    Dim filePath As String
    Dim resultRange As Range
    Dim resultCell As Range
    
    ' Get the search text from cell B1
    searchText = ThisWorkbook.Sheets("Sheet1").Range("B1").Value
    
    ' Specify the directory to search in
    searchDirectory = "C:\Path\To\Your\Directory\" ' Change as needed
    
    ' Set the range to output the results (starting from cell C1)
    Set resultRange = ThisWorkbook.Sheets("Sheet1").Range("C1")
    Set resultCell = resultRange
    
    ' Check if the directory exists
    If Dir(searchDirectory, vbDirectory) = "" Then
        MsgBox "The specified directory does not exist.", vbExclamation
        Exit Sub
    End If
    
    ' Search for files in the directory
    fileName = Dir(searchDirectory & "*.*")
    Do While fileName <> ""
        ' Check if the file name contains the search text
        If InStr(1, fileName, searchText, vbTextCompare) > 0 Then
            ' Output the matching file name to the result range
            resultCell.Value = fileName
            ' Move to the next cell in the result range
            Set resultCell = resultCell.Offset(1, 0)
        End If
        ' Get the next file name
        fileName = Dir
    Loop
    
    ' Inform the user if no matching files were found
    If resultRange.Value = "" Then
        MsgBox "No files found containing the text: " & searchText, vbInformation
    End If
End Sub