Attribute VB_Name = "Module1"
Sub stockdata()
    Dim summaryrow As Double
    Dim totalvolume As Double
    Dim closingprice As Double
    Dim openingprice As Double
    Dim percentchange As Double
    Dim lastrow As Long
    
    summaryrow = 2
    totalvolume = 0
    closingprice = 0
    openingprice = Cells(2, 3).Value
    percentchange = 0
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Headers
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"

For i = 2 To lastrow
    '7 is vol column
    totalvolume = totalvolume + Cells(i, 7).Value
          
    'if current cell ticker diff from next row ticker then
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'get current rows closing price
        closingprice = Cells(i, 6).Value
        
        yearlychange = closingprice - openingprice
        'add yearly change to summary table
        Cells(summaryrow, 10).Value = yearlychange
        'add color index code green if positive red if negative
        If Cells(summaryrow, 10).Value > 0 Then
            Cells(summaryrow, 10).Interior.ColorIndex = 4
        Else
            Cells(summaryrow, 10).Interior.ColorIndex = 3
        End If
        
        'set percent change and protect
        If openingprice > 0 Then
            percentchange = yearlychange / openingprice
        Else
            percentchange = 0
        End If
        'put percent change in summary table
        Cells(summaryrow, 11).Value = percentchange
        
        openingprice = Cells(i + 1, 3).Value
        
        'add ticker to summary table
        Cells(summaryrow, 9).Value = Cells(i, 1).Value
        'add totalvolume to summary table
        Cells(summaryrow, 12).Value = totalvolume
        
        
        summaryrow = summaryrow + 1
        totalvolume = 0
        
        
               
    End If
 

Next i

End Sub

