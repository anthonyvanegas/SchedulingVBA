Attribute VB_Name = "SecurityAssistant"
Option Explicit
'Sub override for the protectProductionDay function, normally this gets activated from ScheduleProduction sub
Sub OverrideProductionDayProtection()
    Dim julianCell As Range: Set julianCell = Range(ActiveCell.Address)
    Dim scheduleWS As Worksheet: Set scheduleWS = ThisWorkbook.Worksheets(ActiveSheet.Name)                     'Reference to the current scheduling worksheet
    
    Dim ans As Integer: ans = MsgBox("Select Yes to Protect or No to Unprotect:", vbYesNoCancel, "Scheduling Assistant")        'Get answer
    If ans = 6 Then
        Call protectProductionDay(scheduleWS, julianCell, True)
    ElseIf ans = 7 Then
        Call protectProductionDay(scheduleWS, julianCell, False)
    End If

End Sub

'Function locks/unlocks the order section of a production day
'   Arg - 'julianCell': Range reference to the julian number cell associated with the production day
'   Arg - 'scheduleWS': Worksheet reference to the selected week schedule worksheet
'   Arg - 'protect': Boolean reference to protection state
Function protectProductionDay(ByRef scheduleWS As Worksheet, ByRef julianCell As Range, ByRef protect As Boolean)
    'Unprotect the entire sheet
        scheduleWS.Unprotect ("")
    'Find the range of cells for the production day to be locked/unlocked (We are only going to use the order cells)
        Dim lastRow As Integer: lastRow = scheduleWS.Cells(Rows.Count, 1).End(xlUp).row
        Dim protectionRange As Range: Set protectionRange = Range(Cells(5, julianCell.Column + 1), _
                                                                Cells(lastRow, julianCell.Column + 5))
        'If an order cell is white change it to a darker color to signify that it is locked.
        Dim cell As Variant
        For Each cell In protectionRange
            If cell.Interior.ColorIndex = 2 Or cell.Interior.ColorIndex = -4142 And protect Then
                cell.Interior.ColorIndex = 15
                cell.Locked = True
            ElseIf cell.Interior.ColorIndex = 15 Then
                cell.Interior.ColorIndex = 2
                cell.Locked = False
            End If
        Next cell
        
        scheduleWS.protect ("")
End Function
        
