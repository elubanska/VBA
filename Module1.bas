Attribute VB_Name = "Module1"
Sub TickerCalc()
'by ELO

'Dim i As Double
Dim Rng As Range
Dim Counter As Double
Dim Counter2 As Double
Dim Ticker As String
Dim Volume As Double
Dim Open_val As Double
Dim Close_val As Double
Dim SummaryRow As Double
Dim StartRow As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim aSheet As Object

For Each aSheet In ActiveWorkbook.Sheets
'MsgBox (aSheet.Name)
aSheet.Activate
'Calculate how many rows do we have filled
Set Rng = Range("A1", Range("A1").End(xlDown))
Counter = Rng.Count

'MsgBox Counter
StartRow = 2
SummaryRow = 2
Volume = 0

'Let's sort to make life easier
'First ticker, then date

'Range("A1:H" & Counter).Select
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("A1"), Order:=xlAscending
     .SortFields.Add Key:=Range("B1"), Order:=xlAscending
     .SetRange Range("A1:H" & Counter)
     .Header = xlYes
     .Apply
End With

'Sheet.Sort Key:=Range("A2:A" & Counter), SortOn:=xlSortOnValues, Order:=xlAscending
'Sort Key:=Range("A2:A" & Counter), SortOn:=xlSortOnValues, Order:=xlAscending

'If opening value = 0
If Cells(StartRow, 3) = 0 Then
        If Cells(StartRow + 1, 3).Value = 0 Then
            Open_val = Cells(StartRow + 2, 3).Value
            Cells(SummaryRow, 14).Value = Open_val
        Else
        Open_val = Cells(StartRow + 1, 3).Value
        Cells(SummaryRow, 14).Value = Open_val
        End If
Else
        Open_val = Cells(StartRow, 3).Value
        Cells(SummaryRow, 14).Value = Open_val
End If

For i = 2 To Counter

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        Ticker = Cells(i, 1).Value
        Volume = Volume + Cells(i, 7).Value
        Close_val = Cells(i, 6).Value
        
        If Cells(i + 1, 3).Value = 0 Then
            If Cells(i + 2, 3).Value = 0 Then
                Open_val = Cells(i + 3, 3).Value
            Else
                Open_val = Cells(i + 2, 3).Value
            End If
        Else
            Open_val = Cells(i + 1, 3).Value
        End If
        
        
        Cells(SummaryRow, 10).Value = Ticker
        Cells(SummaryRow, 13).Value = Volume
        Cells(SummaryRow, 15).Value = Close_val
        Cells(SummaryRow + 1, 14).Value = Open_val
        SummaryRow = SummaryRow + 1
        Volume = 0
        Open_val = 0
     Else
        Volume = Volume + Cells(i, 7).Value
     End If
Next i

'Time for additional calculations:
Set Rng = Range("j1", Range("j1").End(xlDown))
Counter2 = Rng.Count

For j = 2 To Counter2
    Yearly_Change = Cells(j, 15).Value - Cells(j, 14).Value
    If Cells(j, 14).Value = 0 Then
        Percent_Change = 0
    Else
        Percent_Change = (Cells(j, 15).Value / Cells(j, 14).Value) - 1
    End If
    
    Cells(j, 11).Value = Yearly_Change
    Cells(j, 12).Value = Percent_Change
    Cells(j, 12).NumberFormat = "0.0%"
    If Cells(j, 11).Value > 0 Then
        Cells(j, 11).Interior.Color = vbGreen
    Else
        Cells(j, 11).Interior.Color = vbRed
    End If
Next j
'MsgBox (Counter2)
'Calculate the greatest
    Cells(2, 19).Value = WorksheetFunction.Max(Range("l2:l" & Counter2))
    Cells(2, 19).NumberFormat = "0.0%"
    'Cells(2, 18).Value = WorksheetFunction.Match(Range("s2"), Range("l1:l" & Counter2), 0)
    Cells(2, 18).Value = Cells(WorksheetFunction.Match(Range("s2"), Range("l1:l" & Counter2), 0), 10)
    Cells(3, 19).Value = WorksheetFunction.Min(Range("l2:l" & Counter2))
    Cells(3, 19).NumberFormat = "0.0%"
    Cells(3, 18).Value = Cells(WorksheetFunction.Match(Range("s3"), Range("l1:l" & Counter2), 0), 10)
    Cells(4, 19).Value = WorksheetFunction.Max(Range("m2:m" & Counter2))
    'Cells(4, 18).Value = Cells(WorksheetFunction.Match(Range("s4"), Range("l2:l" & Counter2), 0), 10)


Next aSheet

End Sub



