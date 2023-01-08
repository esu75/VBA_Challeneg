VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Author: Esubalew Adal
'This is for sheet1 or 2018 data
'This VBA code runs only on sheet1 (2018)

Sub VBA_Challenge()

'Variable diclaration

Dim Summary_Table_Row As Integer
Dim ClosePrice As Double
Dim OpenPrice As Double
Dim Percent_Change As Double
Dim Yearly_Change As Double
Dim Ticker_Total_Volume As Double
Dim Greatest_Percent_Increase As Double
Dim Greatest_Percent_Decrease As Double
Dim Greatest_Percent_Decrease_Ticker As String
Dim Geatest_Volume_Ticker As String
Dim Greatest_Total_Volume As Double
Dim TickerName As String
Dim Greatest_Percent_Increase_Ticker As String

'Initializing variables

Greatest_Percent_Increase = 0
Greatest_Percent_Decrease = 0
Greatest_Total_Volume = 0
Ticker_Total_Volume = 0
Summary_Table_Row = 2
Greatest_Percent_Increase_Ticker = ""
Greatest_Percent_Decrease_Ticker = ""
Geatest_Volume_Ticker = ""
TickerName = ""

' Get the last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Create Summary Table Headers

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("I1:Q1").Columns.AutoFit
Range("O2:O4").Columns.AutoFit
Range("A1:Q1").Font.Bold = True

OpenPrice = Cells(2, 3).Value   'For capturing the first ticker (AAB)open price just for this mehtod only.
                                'If using the array method,or other methods, explicitly assigning the openPrice value
                                'is not neccessary


'Iteration for the current worksheet rows

For i = 2 To lastrow

' Comparing the adjacent rows values

If Cells(i, 1) <> Cells(i + 1, 1) Then  ' conditon to compare adjacent cells

'If this condition is true, do this

TickerName = Cells(i, 1).Value   ' The ticker name at that i value

ClosePrice = Cells(i, 6).Value   ' Closeprice at that i value

Yearly_Change = ClosePrice - OpenPrice

' calculating the percent change between the yearly change and openprice

Percent_Change = ((ClosePrice / OpenPrice) - 1) * 100

'volume when the if condition is true

Ticker_Total_Volume = Ticker_Total_Volume + Cells(i, 7).Value

'Populate the Summary Table Values

Range("I" & Summary_Table_Row).Value = TickerName

Range("J" & Summary_Table_Row).Value = Yearly_Change

Range("K" & Summary_Table_Row).Value = (Percent_Change) & "%"

Range("L" & Summary_Table_Row).Value = Ticker_Total_Volume

'Adding color to the cells to indicate good (green)and bad (red)performing stocks

If (Yearly_Change > 0) Then

Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

Else

Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

End If
'Increment the Summary Table Rows by 1

Summary_Table_Row = Summary_Table_Row + 1

'Reset variables to capture the next ticker OpenPrice and ClosePrice where it happens @ (i+1)

Yearly_Change = 0            'Resets the value every loop to capture the next ticker value
ClosePrice = 0               'Resets the value every loop to capture the next ticker value
OpenPrice = Cells(i + 1, 3)  'capture the next ticker openPrice

'Comparison to determine the maximum and minimum percent changes

If Percent_Change > Greatest_Percent_Increase Then

Greatest_Percent_Increase = Percent_Change

Greatest_Percent_Increase_Ticker = TickerName

ElseIf Percent_Change < Greatest_Percent_Decrease Then

Greatest_Percent_Decrease = Percent_Change

Greatest_Percent_Decrease_Ticker = TickerName

End If
 'Determine the maximum volume and its ticker value
 
If Ticker_Total_Volume > Greatest_Total_Volume Then

Greatest_Total_Volume = Ticker_Total_Volume

Geatest_Volume_Ticker = TickerName

End If
'Reset to to capture the next ticker values

Percent_Change = 0
Ticker_Total_Volume = 0

Else
'Otherwise, if the above condition is not True, do this repeatdely untilt the condition becomes true. The volumes values will be added row-by-row

Ticker_Total_Volume = Ticker_Total_Volume + Cells(i, 7).Value

End If
Next i

'Populating the Summary_Table after the iteration

Range("P2").Value = Greatest_Percent_Increase_Ticker
Range("P3").Value = Greatest_Percent_Decrease_Ticker
Range("P4").Value = Geatest_Volume_Ticker

Range("Q2").Value = Greatest_Percent_Increase & "%"
Range("Q3").Value = Greatest_Percent_Decrease & "%"
Range("Q4").Value = Greatest_Total_Volume

End Sub

