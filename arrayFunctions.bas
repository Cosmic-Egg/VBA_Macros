Function ExtractColumnsFromRange(rng As Range, ParamArray columns() As Variant) As Variant
    Dim result() As Variant
    Dim i As Long, j As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim colIndex As Long
    
    ' Count the number of rows in the range
    rowCount = rng.Rows.Count
    ' Count the number of columns specified in columns parameter
    colCount = UBound(columns) - LBound(columns) + 1
    
    ' Resize the result array to hold the data
    ReDim result(1 To rowCount, 1 To colCount)
    
    ' Loop through each row in the range
    For i = 1 To rowCount
        ' Loop through each specified column
        For j = LBound(columns) To UBound(columns)
            colIndex = columns(j)
            ' Check if the column index is valid
            If colIndex > 0 And colIndex <= rng.Columns.Count Then
                result(i, j - LBound(columns) + 1) = rng.Cells(i, colIndex).Value
            Else
                result(i, j - LBound(columns) + 1) = "N/A" ' or handle as needed
            End If
        Next j
    Next i
    
    ExtractColumnsFromRange = result
End Function

Sub get_current_region(ByRef rng as range, optional headerLength as long = 1) 
    set rng = rng.CurrentRegion
    set rng = rng.Offset(headerLength).Resize(rng.rows.count - headerLength)
End Sub


Function GetCurrentRegion(ByVal dataRange As Range, Optional headerSize As Long = 1) As Range
    Set GetCurrentRegion = dataRange.CurrentRegion
    If headerSize > 0 Then
        With GetCurrentRegion
            ' Remove the header
            Set GetCurrentRegion = .Offset(headerSize).Resize(.rows.Count - headerSize)
        End With
    End If
End Function

Sub arrayToRange(ByRef data As Variant _
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

Sub arraySetSize(ByRef destinationArray As Variant _
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
        filteredArray = application.transpose(filteredArray)
        'ReDim Preserve filteredArray(1 To outputRow, 1 To UBound(originalArray, 2))
        ReDim Preserve filteredArray(1 To UBound(originalArray, 2), 1 To outputRowput)
        filteredArray = application.transpose(filteredArray)
    Else
        ' If no matches, return an empty array
        ReDim filteredArray(1 To 1, 1 To 1)
        filteredArray(1, 1) = "No matches found"
    End If

    ' Return the filtered array
    filter_array_by_text_with_multiple_criteria = filteredArray
End Function

Function filter_array_by_text_with_multiple_criteria_improved(originalArray As Variant, criteria As Variant, checkColumn As Long) As Variant
    Dim filteredArray() As Variant
    Dim tempArray() As Variant
    Dim i As Long
    Dim outputRow As Long
    Dim match As Boolean

    ' Check if originalArray has any rows
    If IsEmpty(originalArray) Or (TypeName(originalArray) <> "Variant()" And UBound(originalArray, 1) < 1) Then
        filter_array_by_text_with_multiple_criteria = Array()
        Exit Function
    End If

    ' Initialize temporary storage for matches
    ReDim tempArray(1 To UBound(originalArray, 1), 1 To UBound(originalArray, 2))
    
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
        
        ' If a match was found, add the entire row to the temporary array
        If match Then
            outputRow = outputRow + 1
            Dim colIndex As Long
            For colIndex = LBound(originalArray, 2) To UBound(originalArray, 2)
                tempArray(outputRow, colIndex) = originalArray(i, colIndex)
            Next colIndex
        End If
    Next i

    ' Resize the output array only once if outputRow > 0
    If outputRow > 0 Then
        ReDim filteredArray(1 To outputRow, 1 To UBound(originalArray, 2))
        Dim k As Long
        For k = 1 To outputRow
            For colIndex = LBound(originalArray, 2) To UBound(originalArray, 2)
                filteredArray(k, colIndex) = tempArray(k, colIndex)
            Next colIndex
        Next k
    Else
        ' Return an empty array if no matches were found
        filter_array_by_text_with_multiple_criteria = Array()
        Exit Function
    End If

    ' Return the filtered array
    filter_array_by_text_with_multiple_criteria = filteredArray
End Function

Function join_arrays(ByRef arr1 As Variant,ByRef arr2 As Variant, commonColIndex1 As Long, commonColIndex2 As Long) As Variant
    Dim joinedArray() As Variant
    Dim dict As Object
    Dim totalRows As Long
    Dim totalCols As Long
    Dim i As Long, j As Long
    Dim key As Variant
    
    ' Create a dictionary to hold the rows from the second array
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Populate the dictionary with keys from the second array based on the common column
    For i = LBound(arr2, 1) To UBound(arr2, 1)
        key = arr2(i, commonColIndex2)
        If Not dict.Exists(key) Then
            dict.Add key, i ' Store the row index
        End If
    Next i

    ' Calculate the total number of rows for the joined array
    totalRows = 0
    For i = LBound(arr1, 1) To UBound(arr1, 1)
        key = arr1(i, commonColIndex1)
        If dict.Exists(key) Then
            totalRows = totalRows + 1
        End If
    Next i
    
    ' Calculate the total number of columns for the joined array
    totalCols = UBound(arr1, 2) - LBound(arr1, 2) + 1 + UBound(arr2, 2) - LBound(arr2, 2) + 1
    
    ' Resize the joined array
    ReDim joinedArray(0 To totalRows - 1, 0 To totalCols - 1)

    ' Populate the joined array
    Dim rowIndex As Long
    rowIndex = 0
    For i = LBound(arr1, 1) To UBound(arr1, 1)
        key = arr1(i, commonColIndex1)
        If dict.Exists(key) Then
            ' Copy data from the first array
            For j = LBound(arr1, 2) To UBound(arr1, 2)
                joinedArray(rowIndex, j - LBound(arr1, 2)) = arr1(i, j)
            Next j
            
            ' Copy data from the second array
            Dim arr2Row As Long
            arr2Row = dict(key) ' Get the row index from the dictionary
            For j = LBound(arr2, 2) To UBound(arr2, 2)
                joinedArray(rowIndex, j - LBound(arr1, 2) + UBound(arr1, 2) - LBound(arr1, 2) + 1) = arr2(arr2Row, j)
            Next j
            
            rowIndex = rowIndex + 1
        End If
    Next i

    ' Return the joined array
    join_arrays = joinedArray
End Function
