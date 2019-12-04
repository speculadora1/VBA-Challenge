Attribute VB_Name = "Module2"
Sub stockMarket():

' Declare all variables that will be used in the sub-routine
    Dim lastRow As Long
    Dim sumTblRow As Integer
    Dim tickerSymbol As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim incTicker As String
    Dim decTicker As String
    Dim volTicker As String
    Dim greatInc As Double
    Dim greatDec As Double
    Dim greatVol As Double
    
' Begin the outer for loop
    For Each ws In Worksheets
    
    ' Input summary table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
    ' Initialize necessary variables so they reset before the next worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        sumTblRow = 2
        tickerSymbol = "x"
        openPrice = 0
        closePrice = 0
        totalVolume = 0
        greatInc = 0
        greatDec = 0
        greatVol = 0
        
    ' Open the inner for loop
        For i = 2 To lastRow
        
        ' Check if the current ticker is the same as the previous
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                tickerSymbol = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
        ' Check if the current ticker is the same as the next
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                closePrice = ws.Cells(i, 6).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
            ' Add data to summary table
                ws.Range("I" & sumTblRow).Value = tickerSymbol
                ' Set value for yearly change & color conditionally
                    If (closePrice - openPrice) > 0 Then
                        ws.Range("J" & sumTblRow).Interior.ColorIndex = 4
                        ws.Range("J" & sumTblRow).Value = closePrice - openPrice
                    ElseIf (closePrice - openPrice) < 0 Then
                        ws.Range("J" & sumTblRow).Interior.ColorIndex = 3
                        ws.Range("J" & sumTblRow).Value = closePrice - openPrice
                    Else
                        ws.Range("J" & sumTblRow).Value = closePrice - openPrice
                    End If
                ' Conditionally set value for percentage change
                    If openPrice = 0 Then
                        ws.Range("K" & sumTblRow).Value = 0
                    Else
                        ws.Range("K" & sumTblRow).Value = (closePrice - openPrice) / openPrice
                    End If
                ws.Range("L" & sumTblRow).Value = totalVolume
                
            ' Check for any changes to greatest increase / decrease / volume
                If openPrice = 0 Then
                ElseIf ((closePrice - openPrice) / openPrice) > greatInc Then
                    incTicker = tickerSymbol
                    greatInc = (closePrice - openPrice) / openPrice
                ElseIf ((closePrice - openPrice) / openPrice) < greatDec Then
                    decTicker = tickerSymbol
                    greatDec = (closePrice - openPrice) / openPrice
                ElseIf totalVolume > greatVol Then
                    volTicker = tickerSymbol
                    greatVol = totalVolume
                End If
                
            ' Reset totalVolume and increment sumTblRow
                totalVolume = 0
                sumTblRow = sumTblRow + 1
        
        ' Handle else case => current ticker is the same as next and previous
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value

        ' End conditionals
            End If
            
    ' Go to next row
        Next i
        
    ' Format summary table
        ws.Columns("J").NumberFormat = "$0.00"
        ws.Columns("K").NumberFormat = "0.00%"
    
    ' Add greatest increase / decrease / volume data to summary table
        ws.Range("P2").Value = incTicker
        ws.Range("P3").Value = decTicker
        ws.Range("P4").Value = volTicker
        ws.Range("Q2").Value = greatInc
        ws.Range("Q3").Value = greatDec
        ws.Range("Q4").Value = greatVol
   ' Format above data
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("O6").NumberFormat = "0.00%"
        
' Go to next worksheet
    Next ws

End Sub
