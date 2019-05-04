Attribute VB_Name = "Module3"
Sub StockData_Hard()

Dim WS As Worksheet
For Each WS In Worksheets
WS.Activate

'Determine the Last Row
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row


Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Ticker As String
Dim Percent_Change As Double
Dim Volume As Double


WS.Range("I1").Value = "Ticker"
WS.Range("J1").Value = "Yearly Change"
WS.Range("K1").Value = "Percent Change"
WS.Range("L1").Value = "Total_Stock_Volume"
Volume = 0

Dim Row As Double
Row = 2
Dim Column As Integer
Column = 1
Dim i As Long
Open_Price = Cells(2, Column + 2).Value

'Loop through all Ticker symbol
For i = 2 To LastRow
If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
Ticker = Cells(i, Column).Value
Cells(Row, Column + 8).Value = Ticker
Close_Price = Cells(i, Column + 5).Value

Yearly_Change = Close_Price - Open_Price
Cells(Row, Column + 9).Value = Yearly_Change
'Add Percent_Change
If (Open_Price = 0 And Close_Price = 0) Then
Percent_Change = 0
ElseIf (Open_Price = 0 And Close_Price <> 0) Then
Percent_Change = 1
Else
Percent_Change = Yearly_Change / Open_Price
Cells(Row, Column + 10).Value = Percent_Change
Cells(Row, Column + 10).NumberFormat = "0.00%"
End If
Volume = Volume + Cells(i, Column + 6).Value
Cells(Row, Column + 11).Value = Volume
'Add One to the summary table Row
Row = Row + 1
Open_Price = Cells(i + 1, Column + 2)
Volume = 0
'If cells are the same ticker
Else
Volume = Volume + Cells(i, Column + 6).Value
End If
Next i

Yearly_ChangeLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row

For j = 2 To Yearly_ChangeLastRow
If Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0 Then
Cells(j, Column + 9).Interior.ColorIndex = 10
ElseIf Cells(j, Column + 9).Value < 0 Then
Cells(j, Column + 9).Interior.ColorIndex = 3
End If
Next j

 'Set Greatest % Increase, %Decrease, and Total Volume
 Cells(2, Column + 14).Value = "Greatest % Increase"
 
 Cells(3, Column + 14).Value = "Greatest % Decrease"
 Cells(4, Column + 14).Value = "Greatest Total Volume"
 Cells(1, Column + 15).Value = "Ticker"
 Cells(1, Column + 16).Value = "Value"
  'Look through each rows to find the greatest value and its associate ticker
  For Z = 2 To Yearly_ChangeLastRow
  If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & Yearly_ChangeLastRow)) Then
  Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
  Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
  Cells(2, Column + 16).NumberFormat = "0.00%"
  ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & Yearly_ChangeLastRow)) Then
  
  Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
  Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
  Cells(3, Column + 16).NumberFormat = "0.00%"
  ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & Yearly_ChangeLastRow)) Then
 
 Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
 Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
  End If
Next Z

Next WS

End Sub
 
