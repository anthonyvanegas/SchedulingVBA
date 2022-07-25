Attribute VB_Name = "ProjectionAssistant"
Option Explicit

'Created by Anthony Vanegas 1/26/22
'Projection Assistant for scheduling worksheets

'Enum for getRange()
Enum rangeSelector
    projectionCopy = 0
    projectionPaste = 1
    safetyStock = 2
    applySafeStock = 3
End Enum

'Enum for getDict()
Enum dictSelector
    getProjectionDict = 0
    getForecastLocationDict = 1
    getNonForecastLocationDict = 3
    getSafeStockDict = 4
End Enum

'Subroutine to schedule safety stock on a single order line
Sub SingleAddSafetyStock()
    Dim addON As String: addON = InputBox("Enter safety stock qty to add on:")
    Dim forecastList As Collection: Set forecastList = getForecastList()
    
    If Tools.InCollection(forecastList, CStr(Cells(ActiveCell.row, 1).Value)) Then                           'Checks to see if item is forecastable
        Call SchedulingAssistant.scheduleOrder(ActiveCell, CStr(addON), True)
    End If
End Sub

'Subroutine to scheduele out given projections from projection sheets.
Sub SetProjection()
    'Copy row data and add to dictionary
    Dim projectionDict As Dictionary: Set projectionDict = getDict(getRange(projectionCopy), getProjectionDict)

    'Get references to rows and add to other dictionary
    Dim rng As Range: Set rng = getRange(projectionPaste)
    Dim pasteDict As Dictionary: Set pasteDict = getDict(rng, getForecastLocationDict)
    
    'Copy projection from projection sheet
    Dim key As Variant
    For Each key In projectionDict.Keys()
        Cells(pasteDict(key), rng.Column).Value = -projectionDict(key)
    Next
End Sub

'Get range based on selected prompt given by Enum
'   Arg - 'rangeSelect': Enum prompt decision
Function getRange(ByVal rangeSelect As rangeSelector) As Range
    Dim prompt As String 'Prompt that shows up on the input box
    
    'Select prompt based on enum
    Select Case rangeSelect
        Case projectionCopy
            prompt = "Select the column you'd like to pull the projection from:"
        Case projectionPaste
            prompt = "Select the column you'd like to paste the projection to:"
    End Select
    
    'Get range from user
    On Error Resume Next
        Set getRange = Application.InputBox( _
            Title:="Projection Assistant", _
            prompt:=prompt, _
            Type:=8)
    On Error GoTo 0
End Function

'Function to create dictionary data sets
'   Arg - 'rng': Range object
Function getDict(ByVal rng As Range, ByVal dictSelect As dictSelector) As Dictionary
    Set getDict = New Dictionary    'Init the Dictionary
    
    Dim forecastList As Collection  'Reference to forecast list containing item numbers we can forecast with
    
    'Activate the worksheet
    rng.Worksheet.Activate
    'Get the last row
    Dim lastRow As Integer: lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
    
    'Set dictionary values
    Dim i As Integer
    Dim itemNum As String
    Select Case dictSelect
        Case getProjectionDict  'Case where we need dict to have values of projection
            For i = 4 To lastRow
                If Cells(i, rng.Column).Value <> 0 Then
                    getDict.Add CStr(Cells(i, 1).Value), Cells(i, ActiveCell.Column).Value
                End If
            Next i
        Case getForecastLocationDict    'Case where we need dict to have position values of forecastable items on schedule
            Set forecastList = getForecastList()
            For i = 5 To lastRow
                itemNum = Cells(i, 1).Value
                'Check to see if the itemNum number is not blank and that it is a forecastable item
                    If Tools.InCollection(forecastList, itemNum) Then
                        getDict.Add itemNum, i
                    End If
            Next i
        Case getNonForecastLocationDict
            Set forecastList = getForecastList()
            For i = 5 To lastRow
                itemNum = Cells(i, 1).Value
                'Check to see if the itemNum number is not blank and that it is not a forecastable item
                    If Tools.InCollection(forecastList, itemNum) <> True And itemNum <> "" Then
                        getDict.Add itemNum, i
                    End If
            Next i
    End Select
End Function

'Makes a list of the forecastable items
Function getForecastList() As Collection
    Dim sheet As Worksheet: Set sheet = ThisWorkbook.Worksheets("REMOVED")   'We check the STOCK sheet for rows that are unhiden.
    Set getForecastList = New Collection
    
    'Fill in ArrayList
    Dim lastRow As Integer: lastRow = sheet.Cells(Rows.Count, 1).End(xlUp).row
    Dim i As Integer
    
    For i = 3 To lastRow
        If (sheet.Cells(i, 1).EntireRow.Hidden = False) Then
            getForecastList.Add (sheet.Cells(i, 1).Value)
        End If
    Next i
    
End Function


Function printForecastList()
    Dim forecastList As Collection: Set forecastList = getForecastList()

    Dim i As Integer
    For i = 1 To forecastList.Count
        Debug.Print (forecastList.item(i))
    Next i
    
End Function


