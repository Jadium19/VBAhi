Attribute VB_Name = "Module1"
Sub StockTickerLoop()

'Set MainWs as a worksheet object
Dim headers() As Variant
Dim MainWS As Worksheet
Dim WB As Workbook

Set WB = ActiveWorkbook

'Set header info
headers() = Array("Ticker ", "Date ", "Open", "High", "Low", "Close", "Volume", " ", "Ticker", "Yearly_Change", "%Change", "StockVol", " ", " ", " ", "Ticker", "Value")

For Each MainWS In WB.Sheets
With MainWS
.Rows(1).Value = ""
For i = LBound(headers()) To UBound(headers())
.Cells(1, 1 + i).Value = headers(i)

Next i
.Rows(1).Font.Bold = True
.Rows(1).VerticalAlignment = xlCenter
End With

Next MainWS

'Loop through worksheets in workbook
For Each MainWS In Worksheets

'Set initial variables for calcs.

Dim TickerName As String
TickerName = " "
Dim TotalTickerVol As Double
TotalTickerVol = 0
Dim BegPrice As Double
BegPrice = 0
Dim EndPrice As Double
EndPrice = 0
Dim YearlyPriceChg As Double
YearlyPriceChg = 0
Dim YearlyPriceChgPercent As Double
YearlyPriceChgPercent = 0
Dim MaxTickerName As String
MaxTickerName = " "
Dim MinTickerName As String
MinTickerName = " "
Dim MaxPercent As Double
MaxPercent = 0
Dim MinPercent As Double
MinPercent = 0
Dim MaxVolTickerName As String
MaxVolTickerName = " "
Dim MaxVol As Double
MaxVol = 0

'Set location variables
Dim SummaryTableRow As Long
SummaryTableRow = 2

'Set Row count for workbook
Dim Lastrow As Long
 
'loop through all worksheets
Lastrow = MainWS.Cells(Rows.Count, 1).End(xlUp).Row

'Set inital value
BegPrice = MainWS.Cells(2, 3).Value

'Loop from the beginning to the end of worksheet
For i = 2 To Lastrow

'check ticker name
If MainWS.Cells(i + 1, 1).Value <> MainWS.Cells(i, 1).Value Then

'Set the ticker name
TickerName = MainWS.Cells(i, 1).Value

'Calculate
EndPrice = MainWS.Cells(i, 6).Value
YearlyPriceChg = EndPrice - BegPrice

'set condition for a zero value
If BegPrice <> 0 Then
YearlyPriceChgPercent = (YearlyPriceChg / BegPrice) * 100

End If

'add to ticker total vol
TotalTickerVol = TotalTickerVol + MainWS.Cells(i, 7).Value

'print ticker name in summary table
MainWS.Range("I" & SummaryTableRow).Value = TickerName

'Print yearly price change
MainWS.Range("J" & SummaryTableRow).Value = YearlyPriceChg

'Color fill

If (YearlyPriceChg > 0) Then
MainWS.Range("J" & SummaryTableRow).Interior.ColorIndex = 4

ElseIf (YearlyPriceChg <= 0) Then
MainWS.Range("J" & SummaryTableRow).Interior.ColorIndex = 3

End If

'print yearly price change as percent in summary table
MainWS.Range("K" & SummaryTableRow).Value = CStr(YearlyPriceChgPercent)

'print total stock vol in summary table
MainWS.Range("L" & SummaryTableRow).Value = TotalTickerVol

'Add 1 to the sumary table count
SummaryTableRow = SummaryTableRow + 1

'Get next beg price
BegPrice = MainWS.Cells(i + 1, 3).Value

'Do Calculations
If (YearlyPriceChgPercent > MaxPercent) Then
MaxPercent = YearlyPriceChgPercent
MaxTickerName = TickerName

ElseIf (YearlyPriceChgPercent < MinPercent) Then
MinPercent = YearlyPriceChgPercent
MinTickerName = TickerName

End If

If (TotalTickerVol > MaxVol) Then
MaxVol = TotalTickerVol
MaxVolTickerName = TickerName

End If

'Reset Values
YearlyPriceChgPercent = 0
TotalTickerVol = 0

Else
TotalTickerVol = TotalTickerVol + MainWS.Cells(i, 7).Value

End If

Next i

'print values
MainWS.Range("Q2").Value = (CStr(MaxPercent) & "%")
MainWS.Range("Q3").Value = (CStr(MinPercent) & "%")
MainWS.Range("P2").Value = MaxTickerName
MainWS.Range("P3").Value = MinTickerName
MainWS.Range("Q4").Value = MaxVol
MainWS.Range("O2").Value = "Greatest % Increase"
MainWS.Range("O3").Value = "Greatest % Decrease"
MainWS.Range("O4").Value = "Greatest Total Vol"

Next MainWS

End Sub
