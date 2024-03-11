VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_analysis()

'define ranges and set variables for the worksheets

Dim works As Worksheet
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_Change As Double
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As Double
Dim ticker_greatest_increase As String
Dim ticker_greatest_decrease As String
Dim ticker_greatest_volume As String

'set initial values for greatest increase, decrease, and volume
greatest_increase = 0
greatest_decrease = 0
greatest_volume = 0

'error defer

On Error Resume Next

For Each works In ThisWorkbook.Worksheets
'adding headers to columns in all worksheets

works.Cells(1, 9).Value = "Ticker"
works.Cells(1, 10).Value = "Yearly Change"
works.Cells(1, 11).Value = "Percent Change"
works.Cells(1, 12).Value = "Total Stock Volume"

'loop through all rows in worksheets

Summary_Table_Row = 2

For i = 2 To works.UsedRange.Rows.Count

If works.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'finding values in the columns

ticker = works.Cells(i, 1).Value

vol = works.Cells(i, 7).Value

year_open = works.Cells(i, 3).Value

year_close = works.Cells(i, 6).Value

yearly_change = year_close - year_open

percent_Change = yearly_change / year_open

works.Cells(Summary_Table_Row, 9).Value = ticker
works.Cells(Summary_Table_Row, 10).Value = yearly_change
works.Cells(Summary_Table_Row, 11).Value = percent_Change
works.Cells(Summary_Table_Row, 12).Value = vol

 ' Find the greatest % increase, % decrease, and total volume and hold those variables for the summary page
             If percent_Change > greatest_increase Then
                greatest_increase = percent_Change
                ticker_greatest_increase = ticker
            ElseIf percent_Change < greatest_decrease Then
                greatest_decrease = percent_Change
                ticker_greatest_decrease = ticker
            End If
            
            If vol > greatest_volume Then
                greatest_volume = vol
                ticker_greatest_volume = ticker
            End If
        

Summary_Table_Row = Summary_Table_Row + 1




'reset volume to zero

vol = 0

End If
'format colors in yearly change columns, all worksheets

If yearly_change > 0 Then
                    works.Cells(i, 10).Interior.ColorIndex = 4
                ElseIf yearly_change < 0 Then
                    works.Cells(i, 10).Interior.ColorIndex = 3
                End If


Next i



works.Columns("K").NumberFormat = "0.00%"

'Format cells to fit data
works.Columns("I:L").AutoFit

Next
'create new  Summary worksheet for totals, and name it

Sheets.Add.Name = "Summary"

'move to first sheet
Sheets("Summary").Move Before:=Sheets(1)

'Specify location of summary sheet, and its rows/columns

Set Summary = Worksheets("Summary")
Summary.Cells(2, 1).Value = "Greatest % Increase"
Summary.Cells(3, 1).Value = "Greatest % Decrease"
Summary.Cells(4, 1).Value = "Greatest Total Volume"

Summary.Cells(1, 2).Value = "Ticker"
Summary.Cells(1, 3).Value = "Value"

'add in values for greatest increase, decrease,and volume calculated above

Summary.Cells(2, 2).Value = ticker_greatest_increase
Summary.Cells(3, 2).Value = ticker_greatest_decrease
Summary.Cells(4, 2).Value = ticker_greatest_volume
Summary.Cells(2, 3).Value = greatest_increase
Summary.Cells(3, 3).Value = greatest_decrease
Summary.Cells(4, 3).Value = greatest_volume

'format cells to clean up the summary sheet

Summary.Columns("A:C").AutoFit
Summary.Cells(2, 3).NumberFormat = "0.00%"
Summary.Cells(3, 3).NumberFormat = "0.00%"




End Sub

