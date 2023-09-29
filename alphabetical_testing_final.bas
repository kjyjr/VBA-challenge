Attribute VB_Name = "Module1"
Sub alphabetical_testing():

    For Each ws In Worksheets
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Columns("J:L").AutoFit
        ws.Columns("O").AutoFit
    
        
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
      
        Dim Ticker As String
        
        Dim TotalStockVolume As Double
            TotalStockVolume = 0
          
        Dim ClosingPrice As Double
        Dim OpeningPrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim GreatestPercentageIncrease As Double
        Dim GreatestPercentageDecrease As Double
        Dim GreatestTotalVolume As Double
    
        Dim summaryTableRows As Integer
            summaryTableRows = 2
        
        For Row = 2 To lastRow
       
            If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
                Ticker = Cells(Row, 1).Value
            
                Cells(summaryTableRows, 9).Value = Ticker
                       
                TotalStockVolume = TotalStockVolume + Cells(Row, 7).Value
                     
                Cells(summaryTableRows, 12).Value = TotalStockVolume
                                                
                summaryTableRows = summaryTableRows + 1
            
                ClosingPrice = Cells(Row, 6).Value
            
                    If Cells(Row - 1, 1).Value <> Cells(Row, 1).Value Then
                        OpeningPrice = Cells(Row, 3).Value
                
                        YearlyChange = ClosingPrice - OpeningPrice
                
                        Cells(summaryTableRows, 10).Value = YearlyChange
                
                        PercentChange = YearlyChange / ClosingPrice
                
                        Cells(summaryTableRows, 11).Value = PercentChange
                
                    Else
                        OpeningPrice = Cells(Row - 1, 3).Value
                
                    End If
                
            Else
            
                TotalStockVolume = TotalStockVolume + Cells(Row, 7).Value
          
            End If

        Next Row
      
        ' maxIncrease = WorksheetFunction.Max(Columns(11).EntireColumn)
        ' maxIndex = WorksheetFunction.Match(maxIncrease, Columns(11).EntireColumn, 0)
        ' Range("P2").Value = Range("I" & maxIncreaseIndex + 1).Value
        ' Range("Q2").Value = maxIncrease
    
        ' maxDecrease = WorksheetFunction.Max(Columns(11).EntireColumn)
        ' maxIndex = WorksheetFunction.Match(maxDecrease, Columns(11).EntireColumn, 0)
        ' Range("P2").Value = Range("I" & maxDecreaseIndex + 1).Value
        ' Range("Q2").Value = maxDecrease
    
        maxVolume = WorksheetFunction.Max(Columns(12).EntireColumn)
        maxIndex = WorksheetFunction.Match(maxVolume, Columns(12).EntireColumn, 0)
        Range("P4").Value = Range("I" & maxVolumeIndex + 1).Value
        Range("Q4").Value = maxVolume

    Next ws

End Sub
