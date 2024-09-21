Attribute VB_Name = "arrayUtils"
Option Explicit

Enum arrayErrors
    errColumnDoesNotExist = vbObjectError + 513
    errParameterNotArray = vbObjectError + 514
    errIncorrectColumnNumber = vbObjectError + 515
    errIncorrectRowNumber = vbObjectError + 516
    errInvalidArrayPosition = vbObjectError + 517
    errIncorrectNumberofRows = vbObjectError + 518
    errIncorrectNumberofColumns = vbObjectError + 519
    errParameterArrayEmpty = vbObjectError + 520
    errNotArray = vbObjectError + 521
    errNot2DArray = vbObjectError + 522
    errArrayNotSet = vbObjectError + 523
End Enum

' Author: Paul Kelly - https://courses.excelmacromastery.com/
' Course: EffectiveExcelVBA
' Useful Utils for arrays


' Name: GetCurrentRegion
' Description: Gets the CurrentRegion but with the option to remove header rows.
'
' Parameters:
' dataRange: Any range in the data.
' headerSize: The size of the header to remove.
'
' Return Value:
' The CurrentRegion of data with the header rows removed.
Public Function GetCurrentRegion(ByVal dataRange As Range, Optional headerSize As Long = 1) As Range
    Set GetCurrentRegion = dataRange.CurrentRegion
    If headerSize > 0 Then
        With GetCurrentRegion
            ' Remove the header
            Set GetCurrentRegion = .Offset(headerSize).Resize(.rows.Count - headerSize)
        End With
    End If
End Function

' Name: ArrayToRange
' Description: Writes a given array to the worksheet range
'
' Parameters:
' data: The array to write.
' firstCellRange: The destination range where the array will be written.
' numberOfRows: Used to specify the number of rows. If not used then the
'                array row size will be used.
' numberOfRows: Used to specify the number of columns. If not used then the
'                array column size will be used.
' clearExistingData: When true it will remove any existing data in the current region
' clearExistingHeaderSize: Specifies the size of the header to remain when clearing
'                          existing data. Note if clearExistingData is false then this
'                          parameter is not used.

                
Public Sub arrayToRange(ByRef data As Variant _
                        , ByVal firstCellRange As Range _
                        , Optional ByVal numberOfRows As Long = -1 _
                        , Optional ByVal numberOfColumns As Long = -1 _
                        , Optional ByVal clearExistingData As Boolean = True _
                        , Optional ByVal clearExistingHeaderSize As Long = 1)
    If clearExistingData = True Then
        firstCellRange.CurrentRegion.Offset(clearExistingHeaderSize).ClearContents
    End If
    Dim rows As Long, columns As Long
    If numberOfRows = -1 Then
        rows = UBound(data, 1) - LBound(data, 1) + 1
    Else
        rows = numberOfRows
    End If
    
    If numberOfColumns = -1 Then
        columns = UBound(data, 2) - LBound(data, 2) + 1
    Else
        columns = numberOfColumns
    End If
    
    firstCellRange.Resize(rows, columns).Value = data
    
End Sub


' Name: arrayCopyRow
' Description: Copies a row from one array to another
'
' Parameters:
' destinationArray: The row is copied to this array.
' destinationRow: The row is copied to this row in the destination array.
' sourceArray: The row is copied from this array.
' sourceRow: The row is copied from this row in the source array.
'
Public Sub arrayCopyRow(ByRef destinationArray As Variant _
                        , ByVal destinationRow As Long _
                        , ByRef sourceArray As Variant _
                        , ByVal sourceRow As Long)
    
    If UBound(destinationArray, 2) <> UBound(sourceArray, 2) Then
        Err.Raise errIncorrectNumberofColumns, "arrayCopyRow" _
            , "The number of columns in the arrays do not match"
    End If
    Dim i As Long
    For i = LBound(sourceArray, 2) To UBound(sourceArray, 2)
        destinationArray(destinationRow, i) = sourceArray(sourceRow, i)
    Next i
    
End Sub


' Name: arraySetSize
' Description: Sets the size of the destinationArray to the size of the sourceArray
'
Public Sub arraySetSize(ByRef destinationArray As Variant _
                        , ByRef sourceArray As Variant)
                        
    ReDim destinationArray(LBound(sourceArray, 1) To UBound(sourceArray, 1) _
                            , LBound(sourceArray, 2) To UBound(sourceArray, 2))

End Sub


Function filter_array_by_text(ByRef originalArray As Variant, criteria As String, columnIndex As Long) As Variant
    Dim filteredArray() As Variant
    Dim rowCount As Long, colCount As Long
    Dim i As Long, filteredCount As Long

    ' Get the dimensions of the original array
    rowCount = UBound(originalArray, 1)
    colCount = UBound(originalArray, 2)

    ' Count how many rows match the criteria
    For i = 1 To rowCount
        If InStr(1, originalArray(i, columnIndex), criteria, vbTextCompare) > 0 Then
            filteredCount = filteredCount + 1
        End If
    Next i

    ' Resize the filtered array
    If filteredCount > 0 Then
        ReDim filteredArray(1 To filteredCount, 1 To colCount)

        ' Populate the filtered array with matching rows
        filteredCount = 0
        For i = 1 To rowCount
            If InStr(1, originalArray(i, columnIndex), criteria, vbTextCompare) > 0 Then
                filteredCount = filteredCount + 1
                For j = 1 To colCount
                    filteredArray(filteredCount, j) = originalArray(i, j)
                Next j
            End If
        Next i
    End If

    ' Return the filtered array
    filter_array_by_text = filteredArray
End Function


Function sum_and_group_data(ByRef dataArray As Variant, category_column as Long, value_column as Long) As Variant
    Dim dict As Object
    Dim category As Variant
    Dim value As Double
    Dim resultArray() As Variant
    Dim outputRow As Long
    Dim i As Long

    ' Create a new dictionary
    Set dict = CreateObject("Scripting.Dictionary")

    ' Loop through the input array to sum values by category
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        category = dataArray(i, category_column) ' Column 1 (Category)
        value = dataArray(i, value_column) ' Column 2 (Value)

        ' If the category is not in the dictionary, add it
        If Not dict.Exists(category) Then
            dict.Add category, value
        Else
            ' If it exists, sum the value
            dict(category) = dict(category) + value
        End If
    Next i

    ' Resize the result array to hold unique categories and their sums
    ReDim resultArray(1 To dict.Count, 1 To 2)
    
    ' Populate the result array with categories and their totals
    outputRow = 1
    For Each category In dict.Keys
        resultArray(outputRow, 1) = category
        resultArray(outputRow, 2) = dict(category)
        outputRow = outputRow + 1
    Next category

    ' Return the result array
    sum_and_group_data = resultArray
End Function

Function filter_array_by_text_with_multiple_criteria(originalArray As Variant, criteria As Variant, checkColumn As Long) As Variant
    Dim filteredArray() As Variant
    Dim i As Long
    Dim outputRow As Long
    Dim match As Boolean

    ' Initialize the filtered array with a maximum size
    ReDim filteredArray(1 To UBound(originalArray, 1), 1 To UBound(originalArray, 2))
    
    ' Loop through the original array
    outputRow = 0
    For i = LBound(originalArray, 1) To UBound(originalArray, 1)
        match = False ' Assume no match found initially

        ' Loop through the criteria to check for matches
        Dim j As Long
        For j = LBound(criteria) To UBound(criteria)
            If InStr(1, originalArray(i, checkColumn), criteria(j), vbTextCompare) > 0 Then
                match = True ' Match found
                Exit For
            End If
        Next j
        
        ' If a match was found, add the entire row to the filtered array
        If match Then
            outputRow = outputRow + 1
            Dim colIndex As Long
            For colIndex = LBound(originalArray, 2) To UBound(originalArray, 2)
                filteredArray(outputRow, colIndex) = originalArray(i, colIndex)
            Next colIndex
        End If
    Next i
    
    ' Resize the output array to remove unused rows
    If outputRow > 0 Then
        ReDim Preserve filteredArray(1 To outputRow, 1 To UBound(originalArray, 2))
    Else
        ' If no matches, return an empty array
        ReDim filteredArray(1 To 1, 1 To 1)
        filteredArray(1, 1) = "No matches found"
    End If

    ' Return the filtered array
    filter_array_by_text_with_multiple_criteria = filteredArray
End Function

