Attribute VB_Name = "SchedulingAssistant"
'Subroutine for scheduling orders on the worksheet
Sub CombineOrders()
Attribute CombineOrders.VB_ProcData.VB_Invoke_Func = "C\n14"
    Dim ans As Integer: ans = MsgBox("Combine Orders?", vbYesNoCancel, "Scheduling Assistant")        'Get answer
    If ans = 6 Then                                 'Combine orders
        Call scheduleOrder(ActiveCell, 0, True)
    ElseIf ans = 7 Then                             'Do not combine orders
        Call scheduleOrder(ActiveCell, 0, False)
    End If
End Sub

'Function for the entire process of scheduling an order on a single row
'   Arg - 'orderCell': Range referencing a order cell
'   Arg - 'safetyStock': Integer referencing added safety stock (used by Projection Assistant Module)
'   Arg - 'combine': Boolean referencing whether or not we are combining the order or not
Private Function scheduleOrder(ByVal orderCell As Range, ByVal safetyStock As Integer, ByVal combine As Boolean)
    'Variable Declorations
    Dim produceQty As Integer: produceQty = Abs(Cells(orderCell.row, orderCell.Column + 4).Value) + safetyStock             'Total to be produced
    Dim nexOrderCell As Range: Set nexOrderCell = Cells(orderCell.row, orderCell.Column + 13)                               'The next first order delivery cell
    
    If orderCell.Locked = False Then                'Check to make sure we are not refrencing the wrong cell
        If combine = True Then
            Call scheduleLine(orderCell, produceQty)
            Call scheduleOrder(nexOrderCell, 0, True)
        ElseIf combine = False Then
            produceQty = produceQty - Abs(Cells(orderCell.row, orderCell.Column + 1).Value)
            Call scheduleLine(orderCell, produceQty)
            Call scheduleOrder(nexOrderCell, 0, True)
        End If
    End If
End Function

'Function for the single line scheduling process of a single production day
'   Arg - 'orderCell': Range referencing a order cell
'   Arg - 'produceQty': Integer referencing the quantity to be produced in cases
Private Function scheduleLine(ByVal orderCell As Range, ByVal produceQty As Integer)
    Dim packSize As Double: packSize = Cells(orderCell.row, 4).Value                                    'Pack Size
    Dim batchSize As Double: batchSize = Cells(orderCell.row, 5).Value                                  'Batch Size
    Dim poundageCell As Range: Set poundageCell = Cells(orderCell.row, orderCell.Column + 5)            'Poundage cell on selected date
    Dim qtyCell As Range: Set qtyCell = Cells(orderCell.row, orderCell.Column + 8)                      'Units cell on selected date
    
    'Schedule based on whether or not the item is a batch or not
    If batchSize = 0 Then
        poundageCell.Value = produceQty * packSize
        qtyCell.Value = produceQty
    Else
        poundageCell.Value = (produceQty * packSize) / batchSize
        qtyCell.Value = produceQty
    End If
    
    'Clear contents of line if we are not scheduling anything
    If poundageCell.Value = 0 Or qtyCell.Value = 0 Then
        poundageCell.ClearContents
        qtyCell.ClearContents
    End If
End Function

'Subroutine to remove IKI extras from a recipe on the schedule
Sub UseExtra()
Attribute UseExtra.VB_ProcData.VB_Invoke_Func = "V\n14"
    If Cells(ActiveCell.row, 5).Value = 0 Then 'Check to see if the item is a batch
        ActiveCell.Offset(0, 7).Value = ActiveCell.Offset(0, 7).Value - (ActiveCell.Value * Cells(ActiveCell.row, 4).Value) ' Get new #
        ActiveCell.Offset(0, 5).Value = ActiveCell.Offset(0, 5).Value + ActiveCell.Value 'Balance inventory
    Else
        ActiveCell.Offset(0, 7).Value = ActiveCell.Offset(0, 7).Value - ((ActiveCell.Value * Cells(ActiveCell.row, 4).Value) / Cells(ActiveCell.row, 5).Value)
        ActiveCell.Offset(0, 5).Value = ActiveCell.Offset(0, 5).Value + ActiveCell.Value
    End If
End Sub

'Subroutine to add an order to an order cell
Sub AddOrder()
Attribute AddOrder.VB_ProcData.VB_Invoke_Func = "X\n14"
    Dim addON As String
    addON = InputBox("Enter order qty to add on:")
    ActiveCell.Value = -(Abs(ActiveCell.Value) + val(addON))
End Sub

'Subroutine to schedule out production after IKI extras have been included into the recipes
'
'Instructions:  Please select the julian number for the production schedule inside
'                   this sheet. You must also have your production sheet open.
'Notes:         Any file path changes can be made in the variables below.
Sub ScheduleProduction()
    Dim julianCell As Range: Set julianCell = Range(ActiveCell.Address)
    Dim scheduleWS As Worksheet: Set scheduleWS = ThisWorkbook.Worksheets(ActiveSheet.Name)                     'Reference to the current scheduling worksheet
    Dim yieldWS As Worksheet: Set yieldWS = ThisWorkbook.Worksheets("REMOVED")                                  'Reference to the yield sheet, which contains the yield and recipe calculators
    Dim outputWS As Worksheet: Set outputWS = ThisWorkbook.Worksheets("REMOVED")                                'Reference to the output sheet, which formats the recipes
    Dim stockWS As Worksheet: Set stockWS = ThisWorkbook.Worksheets("REMOVED")                                  'Reference to stock sheet, to be given to SBT warehouse
    Dim productionSch As Workbook: Set productionSch = _
                                    Workbooks("Production" + CStr(Year(Date)) + _
                                    InventoryAssistant.getJulianString(julianCell.Value))                       'Reference to the production schedule
    
    'Get Raw Recipes
        Application.DisplayAlerts = False                               'Configs for windows
        
    'Set raw recipes and check for minimum batch sizes
        Call setRawRecipes(julianCell, scheduleWS, yieldWS)
        
        If (minBatchCheck(yieldWS)) Then
        'Set converted Recipes
            Call setConvertedRecipes(yieldWS, outputWS, productionSch)
        'Get Counts
            Call setCounts(julianCell, scheduleWS, productionSch)
        'Set stock counts
            Call createStockList(julianCell, scheduleWS, stockWS)
        'Protect production day
            Call SecurityAssistant.protectProductionDay(scheduleWS, julianCell, True)
        Else
            MsgBox ("Re-run macro after changes are complete")
        End If
                
        Application.DisplayAlerts = True                                'Configs for windows
End Sub

'Function to get converted recipes from raw recipes
'   Arg - 'julianCell': Range reference to the julian number cell associated with the production day
'   Arg - 'scheduleWS': Worksheet reference to the selected week schedule worksheet
'   Arg - 'yieldWS': Worksheet reference to the yield sheet, which contains the yield and recipe calculators
Private Function setRawRecipes(ByVal julianCell As Range, ByVal scheduleWS As Worksheet, ByRef yieldWS As Worksheet)
    'Copy column that contains recipes
    Dim lastRow As Integer: lastRow = scheduleWS.Cells(Rows.Count, 1).End(xlUp).row
    scheduleWS.Range(Cells(julianCell.row + 2, julianCell.Column + 8).Address, _
                        Cells(lastRow, julianCell.Column + 8).Address).Copy
                        
    'Paste over to poundage/yield consolidator
    yieldWS.Cells(2, 3).PasteSpecial xlPasteValues
    
    'Copy column that contains poundages
    scheduleWS.Range(scheduleWS.Cells(julianCell.row + 2, julianCell.Column + 10).Address, _
                        scheduleWS.Cells(lastRow, julianCell.Column + 10).Address).Copy
                        
    'Paste over to raw poundage section
    yieldWS.Cells(2, 6).PasteSpecial xlPasteValues
    
End Function

'Function runs through the yieldWS's poundages to check if we have any recipes under 25#
'   Arg - 'yieldWS': Worksheet reference to the yield sheet, which contains the yield and recipe calculators
Private Function minBatchCheck(ByRef yieldWS As Worksheet) As Boolean
    minBatchCheck = True
    Dim minBatchDict As Dictionary: Set minBatchDict = New Dictionary
    
    Dim stRow As Integer: stRow = 2
    Dim lastRow As Integer: lastRow = yieldWS.Cells(Rows.Count, 1).End(xlUp).row
    Dim chkColumn As Integer: chkColumn = 7
    
    Dim chkNum As Integer
    For i = stRow To lastRow
        chkNum = yieldWS.Cells(i, chkColumn).Value
        If (chkNum > 0 And chkNum < 25) Then
            minBatchDict.Add CStr(yieldWS.Cells(i, 1).Value), chkNum
            minBatchCheck = False
        End If
    Next i
    
    
    If (minBatchCheck = False) Then
        Dim alertString As String
        Dim key
        For Each key In minBatchDict.Keys
            alertString = alertString & " -" & CStr(key) & ": #" & minBatchDict(key) & Chr(10)
        Next key
        
        MsgBox ("Check these items for minimum batch sizes: " & Chr(10) & alertString)
    End If
    
End Function
    

'Function to set converted recipes into the production schedule
'   Arg - 'yieldWS': Worksheet reference to the yield sheet, which contains the yield and recipe calculators
'   Arg - 'outputWS': Worksheet reference to the output sheet, which formats the recipes
'   Arg - 'productionSch': Worksheet reference to the production schedule
Private Function setConvertedRecipes(ByVal yieldWS As Worksheet, ByRef outputWS As Worksheet, ByVal productionSch As Workbook)
    'Copy consolidated poundage
        Dim lastRow As Integer: lastRow = outputWS.Cells(Rows.Count, 1).End(xlUp).row
        outputWS.Range(outputWS.Cells(2, 4).Address, outputWS.Cells(lastRow, 4).Address).Copy
        
    'Paste over to schedule
        Dim productionSheet As Worksheet: Set productionSheet = productionSch.Worksheets("REMOVED")
        productionSheet.Cells(5, 9).PasteSpecial xlPasteValues
        
    'Clean up the recipes
        Dim beginRow As Integer: beginRow = 5
        lastRow = productionSheet.Cells(Rows.Count, 1).End(xlUp).row
        Dim chkCol As Integer: chkCol = 9
        Call Tools.hideRows(beginRow, chkCol, lastRow, productionSheet)
End Function

'Function sets the counts into the Production Schedule from the Scheduleing worksheets. These are the final counts that IKI needs to send to SBT.
'   Arg - 'scheduleWS': Worksheet reference to the selected week schedule worksheet
'   Arg - 'productionSch': Worksheet reference to the production schedule
Private Function setCounts(ByVal julianCell As Range, ByVal scheduleWS As Worksheet, ByVal productionSch As Workbook)
    'Copy counts
        Dim lastRow As Integer: lastRow = scheduleWS.Cells(Rows.Count, 1).End(xlUp).row
        scheduleWS.Range(Cells(julianCell.row + 2, julianCell.Column + 11).Address, Cells(lastRow, julianCell.Column + 11).Address).Copy
    'Paste counts
        Dim countSheet As Worksheet: Set countSheet = productionSch.Worksheets("REVMOED")
        countSheet.Cells(2, 3).PasteSpecial xlPasteValues
        countSheet.Cells(2, 4).PasteSpecial xlPasteValues
        
    'Clean up counts
        Dim beginRow As Integer: beginRow = 2
        lastRow = countSheet.Cells(Rows.Count, 1).End(xlUp).row
        Dim chkCol As Integer: chkCol = 3
        Call Tools.hideRows(beginRow, chkCol, lastRow, countSheet)
End Function

'Function to create stock list for SBT Warehouse based off of what is being projected for the selected day's production
'   Arg - 'julianCell': Range reference to the julian number cell associated with the production day
'   Arg - 'scheduleWS': Worksheet reference to the selected week schedule worksheet
'   Arg - 'stockWS': Worksheet reference to the stock sheet inside this workbook
Private Function createStockList(ByVal julianCell As Range, ByRef scheduleWS As Worksheet, ByRef stockWS As Worksheet)

    'Copy and Paste values into Stock List
    Dim lastRow As Integer: lastRow = Cells(Rows.Count, 1).End(xlUp).row
    scheduleWS.Range(Cells(julianCell.row + 2, julianCell.Column + 2), Cells(lastRow, julianCell.Column + 2)).Copy
    stockWS.Cells(5, 4).PasteSpecial xlPasteValues
    
    'Get date into sheet
    Dim productionDate As Date: productionDate = DateAdd("d", (julianCell.Value - scheduleWS.Cells(3, 2).Value), CDate(scheduleWS.Cells(1, 2).Value2))
    stockWS.Cells(1, 2).Value = productionDate
    
    'Check to see if item is being produced
    Dim forecastLocationDict As Dictionary: Set forecastLocationDict = ProjectionAssistant.getDict(julianCell, getForecastLocationDict)
    Dim key1 As Variant
    For Each key1 In forecastLocationDict.Keys()
        Dim location As Integer: location = forecastLocationDict(key1)
        Set cellRef = scheduleWS.Cells(location, julianCell.Column + 8)
        If isEmpty(cellRef.Value) Then
            stockWS.Cells(location, 5).Interior.ColorIndex = 1
        Else
            stockWS.Cells(location, 5).Interior.ColorIndex = 2
        End If
    Next key1
    
    'Unhide sheet
    stockWS.Visible = xlSheetVisible
        
End Function
    
