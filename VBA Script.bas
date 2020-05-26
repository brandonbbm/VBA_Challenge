Attribute VB_Name = "Module1"
Sub StockCharts()

Dim i As Long

' Set initial variable for ticker
Dim ticker As String

' Set initial variable for total stock volume
Dim volume As Variant
volume = 0

' Set final row value
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
'Set string values for list titles
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"
Range("N2") = "Greatest Percent Increase"
Range("N3") = "Greatest Percent Decrease"
Range("N4") = "Greatest Total Volume"
Range("O1") = "Ticker"
Range("P1") = "Value"

'Set variable for ticker2, greatest % increase, decrease, and total volume
Dim Max_Inc, Max_Dec As Double
Dim Max_Vol As Double

'Set value for year open price
Dim yearopen As Double

'Set value for year closing price
Dim yearclose As Double

'Set value for yearly change
Dim yearchange As Double

'Set value for percent change
Dim percentchange As Variant

'Set location for title values
Dim types As Long
types = 2

'Set initial value for open price for first sotkc
yearopen = Cells(2, 3).Value

For i = 2 To lastrow
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Set ticker value
        ticker = Cells(i, 1).Value
        
        Range("I" & types).Value = ticker
            
        'Set volume value
        volume = Cells(i, 7).Value + volume
            
        Range("L" & types).Value = volume
                
        'Set year close and year change
        yearclose = Cells(i, 6).Value
        
        yearchange = yearclose - yearopen
        
        Range("J" & types).Value = yearchange
        
        'Calculate percent change and input as percent, remove divisible by 0 error
        If yearopen = 0 Then
                percentchange = 0
            
        Else
            percentchange = (yearchange / yearopen)
            
        End If
        
        Range("K" & types).Value = percentchange
        Range("K" & types).NumberFormat = "0.00%"
        
        'Reset year open to next stock
        yearopen = Cells(i + 1, 3).Value
           
        'Continue type
        types = types + 1
            
        'Reset volume
        volume = 0
        
        Else
        
            'Add to volume for same stock
            volume = Cells(i, 7).Value + volume
        
            Range("L" & types).Value = volume
                
        End If
        
    Next i

'Set initial input for max increase, max decrease, max volume
Max_Inc = Cells(2, 11).Value
Min_Inc = Cells(3, 11).Value
Max_Vol = Cells(2, 12).Value

For i = 2 To lastrow
    
    'Set color index for negative percentage change
    If Cells(i, 11).Value < 0 Then
    
        Cells(i, 11).Interior.ColorIndex = 3
        
        'Set color index for positive percentage change
        ElseIf Cells(i, 11).Value > 0 Then
        
            Cells(i, 11).Interior.ColorIndex = 4
    
    End If
    
    'Set value and ticker for maximum increase
    If Cells(i, 11).Value > Max_Inc Then
    
        Max_Inc = Cells(i, 11).Value
        Inc_Ticker = Cells(i, 9).Value
    
    'Set value and ticker for maximum decrease
    ElseIf Cells(i, 11).Value < Max_Dec Then
    
        Max_Dec = Cells(i, 11).Value
        Dec_Ticker = Cells(i, 9).Value
    
    End If
    
    'Set value and ticker for maximum volume
    If Cells(i, 12).Value > Max_Vol Then
    
        Max_Vol = Cells(i, 12).Value
        Vol_Ticker = Cells(i, 9).Value
    
    End If
    
    'Removes additional prints/color index beyond available stocks
    If Cells(i, 9).Value = "" Then
    
        Cells(i, 10).Value = ""
        
        Cells(i, 12).Value = ""
        
        Cells(i, 11).Interior.ColorIndex = xlNone
    
    End If

    
Next i

'Print value for max increase, decrease, volume, and tickers, and set percent formatting
Cells(2, 15).Value = Inc_Ticker
Cells(2, 16).Value = Max_Inc
Cells(2, 16).NumberFormat = "0.00%"
Cells(3, 15).Value = Dec_Ticker
Cells(3, 16).Value = Max_Dec
Cells(3, 16).NumberFormat = "0.00%"
Cells(4, 15).Value = Vol_Ticker
Cells(4, 16).Value = Max_Vol

'Autofit Column Width
Columns("A:O").EntireColumn.AutoFit

End Sub

