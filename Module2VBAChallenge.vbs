Attribute VB_Name = "Module1"
Sub stocksoutput()

Dim x As Long
Dim z As Long
x = Application.Worksheets.Count

For z = 1 To x

Worksheets(z).Activate

'''''''''''''''''''''''''''''''''''''''''''''''
Dim a As Long
Dim b As Long
Dim yearlychange As Double
Dim percentchange As Double
Dim Tsvcells As Double
Dim firstdateamount As Double
Dim lastdateamount As Double
Dim valuerow As Long
Dim subvaluerow As Long

lastrowA = Cells(Rows.Count, 1).End(xlUp).Row
lastrowI = Cells(Rows.Count, 9).End(xlUp).Row
valuerow = 2
subvaluerow = 1
Tsvcells = 0
tsv = 0

Range("A2:A" & lastrowA).Copy
Cells(2, 9).Select
ActiveSheet.Paste
Range("I2:I" & lastrowA).RemoveDuplicates Columns:=1, Header:=xlNo


'''

For a = 2 To lastrowA

If Right(Cells(a, 2), 4) = "0102" Then
    firstdateamount = Cells(a, 3).Value
    
ElseIf Right(Cells(a, 2), 4) = "1231" Then
    lastdateamount = Cells(a, 6).Value
    
   yearlychange = lastdateamount - firstdateamount
   Range("J" & valuerow).Value = yearlychange
   
   If yearlychange > 0 Then
    Range("J" & valuerow).Interior.ColorIndex = 4
    
    ElseIf yearlychange < 0 Then
        Range("J" & valuerow).Interior.ColorIndex = 3
        
    End If
    
    percentchange = yearlychange / firstdateamount
    Range("K" & valuerow).Value = percentchange
    Range("K" & valuerow).NumberFormat = "0.00%"
    
  valuerow = valuerow + 1
  
End If

Next a

'''

For b = 2 To lastrowA



    If Cells(b, 1).Value <> Cells(b + 1, 1).Value Then
    
    Tsvcells = Tsvcells + Cells(b, 7).Value
    subvaluerow = subvaluerow + 1
    Range("L" & subvaluerow).Value = Tsvcells
    Tsvcells = 0
    
    Else
    
    Tsvcells = Tsvcells + Cells(b, 7).Value
    
    End If
Next b

'''

Dim c As Long
Dim lastrowK As Long
Dim PercentMax As Double
Dim PercentMaxTicker As String
Dim PercentMin As Double
Dim PercentMinTicker As String
Dim Percentrange As Range
Dim StockVMax As Double
Dim StockVTicker As String
Dim StockVRange As Range


Set Percentrange = Range("K:K")
Set StockVRange = Range("L:L")
lastrowK = Cells(Rows.Count, 11).End(xlUp).Row


Range("Q2").Value = Application.WorksheetFunction.max(Range("K2:K" & lastrowK))
PercentMax = Range("Q2").Value
PercentMaxTicker = Application.Match(PercentMax, Percentrange, 0)
Range("Q2").NumberFormat = "0.00%"
Range("P2").Value = Cells(PercentMaxTicker, "I").Value


Range("Q3").Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrowK))
PercentMin = Range("Q3").Value
PercentMinTicker = Application.Match(PercentMin, Percentrange, 0)
Range("Q3").NumberFormat = "0.00%"
Range("P3").Value = Cells(PercentMinTicker, "I").Value

Range("Q4").Value = Application.WorksheetFunction.max(Range("L2:L" & lastrowK))
StockVMax = Range("Q4").Value
StockVTicker = Application.Match(StockVMax, StockVRange, 0)
Range("P4").Value = Cells(StockVTicker, "I").Value

'''''''''''''''

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"


Columns("I:Q").AutoFit



Next z

End Sub
