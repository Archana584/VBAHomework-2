Attribute VB_Name = "Module1"
Sub Stock_Analysis_Easy()
Dim WS As Worksheet
Dim Ticker As String
Dim Total_Stock_Volume As Double
Dim Summary_Table_Row As Integer

For Each WS In Worksheets
WS.Activate

Total_Stock_Volume = 0

'Keep track of the location for each row/line in tne summary table
Summary_Table_Row = 2
WS.Range("I1").Value = "Ticker"
WS.Range("J1").Value = "Total Stock Volume"

'Determine the Last row
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Set the Ticker Value
Ticker = Cells(i, 1).Value

'Add to Total_Stock_Volume
Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
WS.Range("I" & Summary_Table_Row).Value = Ticker

'Print the Total_Stock_Volume to the summary Table
WS.Range("J" & Summary_Table_Row).Value = Total_Stock_Volume

'Add one to the Summary table row
Summary_Table_Row = Summary_Table_Row + 1

'Reset Total_Stock_Volume
Total_Stock_Volume = 0


Else
Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

End If

Next i

Next WS

End Sub

