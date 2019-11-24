VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub TickerLoop()
'Set the initial variable for holding the ticker symbol
Dim Ticker_Symbol As String

'Set an initial variale for holding the total stock volume per symbol
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

'Set the First row of the ticker symbol
Dim Open_Row As Long

'Set the Last value of the ticker symbol
Dim Close_Row As Long

'Set the open value
Dim Year_Open_Value As Double

'Set the close value
Dim Year_Close_Value As Double

Dim lastrow As Long

Dim YearlyChange As Double

Dim PercentChange As Double


'add unique ticker symbols to a summary table

Dim Ticker_Summary_Row As Long
Ticker_Summary_Row = 2

'reset number of rows variable
    lastrow = 0

'count the number of rows in active sheet.
   lastrow = Cells(Rows.Count, 1).End(xlUp).Row


'Loop through all Ticker Symbols
For i = 2 To lastrow

'Check if the ticker symbol is the same
 
If Cells(i + 1, 1).Value <> Cells(i, 1) Then

'Set the Ticker Symbol
Ticker_Symbol = Cells(i, 1).Value

'Add to total stock vol
Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

'Print the TickerSymbol in the summary table
Range("I" & Ticker_Summary_Row).Value = Ticker_Symbol

'Print the total stock volume in the summary table
Range("l" & Ticker_Summary_Row).Value = Total_Stock_Volume

'Add one to the summary table row
Ticker_Summary_Row = Ticker_Summary_Row + 1

'Reset the stock vol total
Total_Stock_Volume = 0

Else
'Add to the Total stock volume total
Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value


End If

Next i

'Set header row for the 4 columns
Range("I1").Value = "Ticker Symbol"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"


For i = 2 To 850000

Open_Row = Range("A:A").Find(What:=Cells(i, 9).Value, After:=Range("A1")).Row
Close_Row = Range("A:A").Find(What:=Cells(i, 9).Value, After:=Range("A1"), lookat:=xlWhole, searchdirection:=xlPrevious).Row

'Set the open and close values
Year_Open_Value = Cells(Open_Row, 3).Value
Year_Close_Value = Cells(Close_Row, 6).Value

'Find the difference

YearlyChange = Year_Close_Value - Year_Open_Value

'Populate values for change and %

Cells(i, 10).Value = YearlyChange
Cells(i, 11).Value = Round(YearlyChange / Year_Open_Value, 4)


Next i

End Sub

