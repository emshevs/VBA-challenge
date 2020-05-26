Attribute VB_Name = "Module1"
Sub StockChallenge()

For Each ws In Worksheets

'Headers

    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
'Variables

    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Volume As Double
    Volume = 0
    
    Dim open_price As Double
    open_price = 0
    Dim close_price As Double
    close_price = 0
    
    Dim lastrow As Double
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
    
'Loop through tickers

For i = 2 To lastrow


    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        Ticker = ws.Cells(i, 1).Value
        Volume = Volume + ws.Cells(i, 7).Value

          ws.Range("I" & Summary_Table_Row).Value = Ticker
          ws.Range("L" & Summary_Table_Row).Value = Volume

        Volume = 0

        close_price = ws.Cells(i, 6)
        
       ' To avoid any errors
        If open_price = 0 Then
            YearlyChange = 0
            PercentChange = 0
        'Calculate year change and percentage change
        Else:
            YearlyChange = close_price - open_price
            PercentChange = (close_price - open_price) / open_price
            
        End If

        ' Inserting values to summary table
            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            ws.Range("K" & Summary_Table_Row).Value = PercentChange
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

            Summary_Table_Row = Summary_Table_Row + 1

    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
         open_price = ws.Cells(i, 3)


    Else: Volume = Volume + ws.Cells(i, 7).Value

    End If

    Next i

    'Conditional Formatting

    For j = 2 To lastrow

        If ws.Range("J" & j).Value > 0 Then
            ws.Range("J" & j).Interior.ColorIndex = 4

        ElseIf ws.Range("J" & j).Value < 0 Then
            ws.Range("J" & j).Interior.ColorIndex = 3
        
    End If

    Next j
    
    ' Headers for Challenge

        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
    'Define and assign variables

        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As Double

        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
    
    ' Loop to find greatest increase and input value + ticker

        For k = 2 To lastrow


        If ws.Cells(k, 11).Value > GreatestIncrease Then
            GreatestIncrease = ws.Cells(k, 11).Value
            ws.Range("Q2").Value = GreatestIncrease
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("P2").Value = ws.Cells(k, 9).Value
        
        End If

        Next k
    
    ' Loop to find greatest decrease and input value + ticker

    For l = 2 To lastrow
    
        If ws.Cells(l, 11).Value < GreatestDecrease Then
            GreatestDecrease = ws.Cells(l, 11).Value
            ws.Range("Q3").Value = GreatestDecrease
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("P3").Value = ws.Cells(l, 9).Value
        End If
    
   Next l
   
    ' Loop to find greatest volume and input value + ticker

    For m = 2 To lastrow
    
        If ws.Cells(m, 12).Value > GreatestVolume Then
            GreatestVolume = ws.Cells(m, 12).Value
            ws.Range("Q4").Value = GreatestVolume
            ws.Range("Q4").NumberFormat = "0.00E+00"
            ws.Range("P4").Value = ws.Cells(m, 9).Value
            
        End If
  
    Next m

'Apply to all worksheets
    
Next ws

End Sub
