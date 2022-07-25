Attribute VB_Name = "Tools"
'InCollection function due to Net Framework not loading correct libraries for ArrayList
Public Function InCollection(col As Collection, key As String) As Boolean
    Dim i As Variant
    InCollection = False
    For Each i In col
        If StrComp(i, key) = 0 Then
            InCollection = True
        End If
    Next i
End Function

'Simple helper function to hide rows on a selected worksheet
'   Arg - 'beginRow': Integer reference to the beginning row
'   Arg - 'chkCol': Integer reference to row being checked
'   Arg - 'lastRow': Integer reference to the last row in the sheet
'   Arg - 'workSHT': Worksheet reference to sheet being operated on
Public Function hideRows(ByVal beginRow As Integer, ByVal chkCol As Integer, ByVal lastRow, ByRef workSHT As Worksheet)
    Dim cellRef As Range
    For i = beginRow To lastRow
        Set cellRef = workSHT.Cells(i, chkCol)
        If isEmpty(cellRef.Value) Or cellRef.Value = 0 Then
            workSHT.Cells(i, chkCol).EntireRow.Hidden = True
        End If
    Next i
End Function

'Function to print Dictionary key and value pairs to console
Function printDict(ByVal dict As Dictionary)
    Dim key
    For Each key In dict.Keys
        Debug.Print (key + ": " + CStr(dict(key)))
    Next key
End Function

