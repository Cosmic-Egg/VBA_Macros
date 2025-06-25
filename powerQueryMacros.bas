Sub RefreshMultipleQueriesAndWait()
    Dim queryNames As Variant
    Dim conn As WorkbookConnection
    Dim name As Variant
    Dim allDone As Boolean

    ' List of Power Query connection names (as seen in Data → Queries & Connections)
    queryNames = Array("Sales Data", "Customer Info", "Orders")  ' <-- change to your actual query names

    ' Start refreshing all listed queries
    For Each name In queryNames
        On Error Resume Next  ' In case a query name is missing
        Set conn = ThisWorkbook.Connections(name)
        If Not conn Is Nothing Then
            conn.Refresh
        End If
    Next name

    ' Wait until all are done refreshing
    Do
        allDone = True
        For Each name In queryNames
            Set conn = ThisWorkbook.Connections(name)
            If Not conn Is Nothing Then
                If conn.Refreshing Then
                    allDone = False
                    Exit For  ' No need to check further — still waiting
                End If
            End If
        Next name
        DoEvents  ' Allow Excel to continue working during the wait
    Loop Until allDone

    MsgBox "All queries have been refreshed.", vbInformation
End Sub
