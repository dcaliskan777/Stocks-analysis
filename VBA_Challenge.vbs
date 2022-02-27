Sub yearValueAnalysisRefactored()
yearValue = InputBox("What is the year would you like to analysis on?")
Dim startTime As Single
Dim endTime As Single
startTime = Timer

 'Format the output sheet on All Stocks Analysis worksheet
 
   Sheets("All Stocks Analysis WRF").Activate
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   
   'Create a header row
   
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

 
   ' Define variables
   Dim currentTicker As String
   Dim currentTickerStartingPrice As Single
   Dim currentTickerEndingPrice As Single
   Dim currentTickerTotalVolume As Double
   Dim tickerIndex As Integer
   Dim RowCount As Integer
   
   'Activate data worksheet
   
   Sheets(yearValue).Activate
   
   'Get the number of rows to loop over
   
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row
   
   'Insert a new ticker,  close price and which is not in the list
   
   Cells(RowCount + 1, 1).Value = "Dursun"
   Cells(RowCount + 1, 6).Value = 1
   Cells(RowCount + 1, 8).Value = 1
   
    'Initilize currentTicker, currentTickerStartingPrice, currentTickerTotalVolume and tickerIndex
    
    currentTicker = Range("A2").Value
    currentTickerStartingPrice = Cells(2, 6)
    currentTickerTotalVolume = 0
    tickerIndex = 4
    
       
       'Loop through rows in the data
       
       For j = 2 To RowCount + 1
       
           'Get total volume for current ticker
           
           If Cells(j, 1).Value = currentTicker Then

               currentTickerTotalVolume = currentTickerTotalVolume + Cells(j, 8).Value
               
               Else
               
               'Assign currenTickerEndingPrice
               
               currentTickerEndingPrice = Cells(j - 1, 6).Value
               
               'Activate Sheet All Stocks Analysis WRF and import currentTicker,currentTickerTotalVolume and return of currentTicker on it
               
               Worksheets("All Stocks Analysis WRF").Activate
               
               Cells(tickerIndex, 1).Value = currentTicker
       Cells(tickerIndex, 2).Value = currentTickerTotalVolume
       Cells(tickerIndex, 3).Value = currentTickerEndingPrice / currentTickerStartingPrice - 1
       
        'Activate data worksheet(yearValue)
   
   Sheets(yearValue).Activate
   
       'Initilize currentTicker,currentTickerStartingPrice,currentTickerTotalVolume and increase tickerIndex
       
       currentTicker = Cells(j, 1).Value
    currentTickerStartingPrice = Cells(j, 6)
    currentTickerTotalVolume = Cells(j, 8)
    tickerIndex = tickerIndex + 1

           End If

        
       Next j
     
   
   'Formatting
   
   Worksheets("All Stocks Analysis WRF").Activate
Range("A3:C3").Font.Bold = True
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("B4:B15").NumberFormat = "#,##0"
Range("C4:C15").NumberFormat = "0.0%"
Columns("B").AutoFit

dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If
        
        Next i
        
         'Clear aded ticker,  close price and which is not in the list
         
         Sheets(yearValue).Activate
   
   Cells(RowCount + 1, 1).Value = ""
   Cells(RowCount + 1, 6).Value = ""
   Cells(RowCount + 1, 8).Value = ""
        
endTime = Timer
MsgBox ("This code ran in " & (endTime - startTime) & " seconds for the year " & yearValue)
End Sub
