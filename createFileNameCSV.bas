Sub RemoveNumbersAndSaveAsCSV()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim cleanedNames As Collection
    Dim cleanedName As String
    Dim filePath As String
    Dim fileNum As Integer
    Dim i As Integer
    
    ' Set the worksheet and range
    Set ws = ThisWorkbook.Sheets("Sheet1") 
    Set rng = ws.Range("A1:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
    
    ' Initialize the collection to store cleaned names
    Set cleanedNames = New Collection
    
    ' Loop through each cell in the range
    For Each cell In rng
        ' Remove numbers from the name
        cleanedName = RemoveNumbers(cell.Value)
        ' Add the cleaned name to the collection
        cleanedNames.Add cleanedName
    Next cell
    
    ' Set the file path for the CSV file
    filePath = Application.GetSaveAsFilename("CleanedFileNames.csv", "CSV Files (*.csv), *.csv")
    
    ' Check if the user canceled the save dialog
    If filePath = "False" Then Exit Sub
    
    ' Open the file for writing
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    ' Write the cleaned names to the CSV file
    For i = 1 To cleanedNames.Count
        Print #fileNum, cleanedNames(i)
    Next i
    
    ' Close the file
    Close #fileNum
    
    ' Inform the user
    MsgBox "Cleaned file names have been saved to " & filePath, vbInformation
End Sub

Function RemoveNumbers(ByVal text As String) As String
    Dim i As Integer
    Dim result As String
    Dim char As String
    
    ' Initialize the result string
    result = ""
    
    ' Loop through each character in the text
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        ' Check if the character is not a number
        If Not IsNumeric(char) Then
            result = result & char
        End If
    Next i
    
    ' Return the result
    RemoveNumbers = result
End Function