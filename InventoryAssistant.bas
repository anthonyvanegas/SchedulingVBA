Attribute VB_Name = "InventoryAssistant"
'Enum for getDict()
Enum countDictSelector
    getReqCntDict = 3       'Dictionary for requested count values on the count sheet
    getNetCntDict = 4       'Dictionary for net count values on the count sheet
End Enum

'Notes: Add some sort of way to adjust page breaks on the variance sheet

'Subroutine to set updated counts from the schedule to our scheduling sheet. This is usually done in the morning for the previous days production day. For instance, on Tuesday,
'       this would be done for Mondays production. This subroutine also finds the variances between these counts.
'
'Instructions:  Please select the julian number for the production schedule inside
'                   this sheet.
'Notes:         Any file path changes can be made in the variables below.
Sub InventoryBalance()
    'Turn off all display alerts
    Application.DisplayAlerts = False
    
    'Define variables for our scheduling WB
    Dim curWS As Worksheet: Set curWS = _
                                ThisWorkbook.Worksheets(ThisWorkbook.ActiveSheet.Name)              'Reference to our current WS (our schedule)
    Dim julianNum As Integer: julianNum = ActiveCell.Value
    Dim adjClm As Integer: adjClm = ActiveCell.Column + 6
    
    'Define variables for the production WB
    Dim productionWB As Workbook: Set productionWB = _
                                    Workbooks.Open("U:\REMOVED" + _
                                        CStr(Year(Date)) + _
                                        getJulianString(julianNum))                                 'Reference to the production schedule
    Dim countWS As Worksheet: Set countWS = _
                                productionWB.Worksheets(4)                                          'Reference to 'Anthony's Count Sheet' inside of ProductionWB
                                
    'Copy and paste the counts into our current worksheet
    countWS.Range(countWS.Cells(2, 4).Address, _
                    countWS.Cells(countWS.Cells(Rows.Count, 1).End(xlUp).row, 4).Address).Copy
    curWS.Range(Cells(ActiveCell.row + 2, ActiveCell.Column + 11).Address).PasteSpecial xlPasteValues
    
    'Get variances between counts
    MsgBox ("Loading Variance Report")
    Dim varianceDict As Dictionary: Set varianceDict = getVarianceDict( _
                                        getCountDict(countWS, getReqCntDict), _
                                        getCountDict(countWS, getNetCntDict))                       'Reference to our variance dictionary that holds the differences between req. and net. counts
    productionWB.Close SaveChanges:=False
    
    'Go through scheduling sheet and apply variances for non-forecastable items
    Call setVariances(varianceDict, curWS, adjClm)
    
    'Create report to show what items were adjusted and for how much
    Call genReport(varianceDict, julianNum)
    
    'Reset display alerts
    Application.DisplayAlerts = True
    
End Sub

'Function to format our julian number in accordance to our production schedule file naming schema
'   Arg - 'julianNum': Our julian number for the production schedule being manipulated
Public Function getJulianString(ByVal julianNum As Integer) As String
    'Logic to make sure we follow sheet naming logic
    Select Case julianNum
        Case 0 To 9
            getJulianString = (" sched00" + CStr(julianNum) + ".xls")
        Case 10 To 99
            getJulianString = (" sched0" + CStr(julianNum) + ".xls")
        Case 100 To 366
            getJulianString = (" sched" + CStr(julianNum) + ".xls")
        Case Else   'Case where there is no clear julian number
            err.Raise Number:=vbObjectError + 513, _
                        Description:="Incorrect cell selected, please select Julian Number"
    End Select
End Function
 
'Function to create dictionary data sets
'   Arg - 'countWS': Reference to "Anthony's Count Sheet (SBT)" in the Production Schedule
'   Arg - 'dictSelect': Reference to ENUM selector for either the requested or net counts
Private Function getCountDict(ByRef countWS As Worksheet, ByVal dictSelect As countDictSelector) As Dictionary
    Set getCountDict = New Dictionary    'Init the Dictionary
    
    Dim startRow As Integer: startRow = 2                                               'Row where counts start
    Dim lastRow As Integer: lastRow = countWS.Cells(Rows.Count, 1).End(xlUp).row        'Row where counts end
    Dim row As Variant                                                                  'For loop iterator
    
    'For each row, get the count associated with the requested column
    For row = 2 To lastRow
        If countWS.Cells(row, dictSelect) <> "" Then
            getCountDict.Add countWS.Cells(row, 1).Value, countWS.Cells(row, dictSelect).Value
        End If
    Next row
End Function

'Function to create variance dictionary set
'   Arg - 'reqDict': Reference to our requested counts from the count sheet
'   Arg - 'netDict': Reference to our net counts from the count sheet
Private Function getVarianceDict(ByRef reqDict As Dictionary, ByRef netDict As Dictionary) As Dictionary
    Set getVarianceDict = New Dictionary                    'Init the Dictionary
    Dim item As Variant
    Dim reqCnt As Integer
    Dim netCnt As Integer
    For Each item In reqDict
        reqCnt = reqDict(item)
        If netDict.Exists(item) = True Then                 'If the item exists in the net dictionary
            'Compare the netCnt to the reqCnt
            netCnt = netDict(item)
            If reqCnt <> netCnt Then                        'Request count does not equal net count so we find the variance
                getVarianceDict.Add CStr(item), (reqCnt - netCnt)
            End If
        Else                                                'If the requested item is not on the net dictionary then the item is short
            reqCnt = -reqCnt
            getVarianceDict.Add CStr(item), reqCnt
        End If
    Next item
    'Loop through net dict to check if their are any new items
    For Each item In netDict
        If reqDict.Exists(item) <> True Then
            getVarianceDict.Add item, netDict(item)
        End If
    Next item
End Function

'Function to set our variances inside of our production schedule (in our requested production dates adjustment column)
'   Arg - 'varianceDict': Reference to our variance dictionary
'   Arg - 'adjClm': Reference to our adjustment column for the given production day
Private Function setVariances(ByRef varianceDict As Dictionary, ByRef curWS As Worksheet, ByRef adjClm As Integer)
    'Go through sheet and clean up anything that is not forecastable
    Dim locationDict As Dictionary: Set locationDict = _
                                        ProjectionAssistant.getDict(ActiveCell, getNonForecastLocationDict) 'Reference to non-forecastable item row locations
    Dim itemNum As Variant                                                                                  'Our dictionary iterator, contains the item number
    Dim orgAdjAmt As Integer                                                                                'The original amount inside of the adjustment cell
    
    'Iterate through each variance
    For Each itemNum In varianceDict
        'Check to see if the variance exists in the non-forecastable item location dictionary
        If locationDict.Exists(itemNum) Then
            'Add the adjustment to the adjustment column
            orgAdjAmt = curWS.Cells(locationDict(itemNum), adjClm).Value
            curWS.Cells(locationDict(itemNum), adjClm).Value = orgAdjAmt - varianceDict(itemNum)
        End If
    Next itemNum
End Function

'Function generates variance report sheet
'   Arg - 'varianceDict': Reference to our variance dictionary
'   Arg - 'curWB': Reference to our current workbook
Private Function genReport(ByRef varianceDict As Dictionary, ByRef julianNum As Integer)
    Dim rowCnt As Integer: rowCnt = 3
    Dim varianceWS As Worksheet: Set varianceWS = _
                            ThisWorkbook.Worksheets("REMOVED")                       'Reference to variance WS
    
    'Clear report first
    varianceWS.Range("A" + CStr(rowCnt) + _
                        ":B" + CStr(varianceWS.Cells(Rows.Count, 1).End(xlUp).row)).ClearContents
    
    'Set julian number
    varianceWS.Cells(1, 2).Value = julianNum
    
    'For each variance, place it into the sheet
    For Each itemNum In varianceDict
        varianceWS.Cells(rowCnt, 1).Value = itemNum
        varianceWS.Cells(rowCnt, 2).Value = -varianceDict(itemNum)
        rowCnt = rowCnt + 1
    Next itemNum
    
    'Unhide sheet if it is hidden
    If varianceWS.Visible = xlSheetHidden Then
        varianceWS.Visible = xlSheetVisible
    End If
    
    'Activate Sheet
    varianceWS.Activate
End Function
