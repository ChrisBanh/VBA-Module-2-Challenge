
Option Explicit


Sub stock()

'==============================
'Loop through all of the sheets
'==============================
Dim ws As Worksheet

For Each ws In Worksheets

    '---------------
    'Insert the Year
    '---------------
    
    'Create a variable to hold the ticker symbol, total volume variable to determine the number of rows in each sheet. 
    

   
   Dim TickerName As String
   TickerName = " "
   
   
   Dim TickerTotal As Double
   TickerTotal = 0
   
   Dim Lastrow As Long
   Dim i As Long
   

   
   ' Determine the Last Row
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
 
   
   
   
   ' Add a Column for Ticker
        ws.Range("I1").EntireColumn.Insert
        
   ' Add a Column for Yearly Change
        ws.Range("J1").EntireColumn.Insert
        
   ' Add a Column for Percent Change
        ws.Range("K1").EntireColumn.Insert
    
   ' Add a Column for Total Stock Volume
        ws.Range("L1").EntireColumn.Insert
        
        
   ' Add the word Ticker to the First Column Header
        ws.Cells(1, 9).Value = "Ticker"
        
   ' Add the words Yearly Change to the Second Column Header
        ws.Cells(1, 10).Value = "Yearly Change"
        
   ' Add the words Percent Change to the Third Column Header
        ws.Cells(1, 11).Value = "Percent Change"
        
        
   ' Add the words Total Stock Volume to the Fourth Column Header
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
   'Add headers to the right of the original extrapolated data that will help indicate the largest metrics within the given dataset

     ws.Range("P1").Value = "Ticker"
     ws.Range("Q1").Value = "Value"
     ws.Range("O2").Value = "Greatest % Increase"
     ws.Range("O3").Value = "Greatest % Decrease"
     ws.Range("O4").Value = "Greatest Total Volume"
     
     
     ' Set Variables for the prices, changes, percentages and range of values. Also set variables for the bonus assessment.
     
      Dim openprice As Double
        openprice = 0
      Dim closeprice As Double
        closeprice = 0
      Dim pricechange As Double
        pricechange = 0
      Dim percentpricechange As Double
        percentpricechange = 0
      Dim Formatchange As Double
      Dim incTicker As String
      Dim incVal As Double
      Dim decTicker as String
      Dim decVal as Double
      incTicker = ""
      incVal = 0
      decTicker = ""
      decVal = 0
      Dim greatestTicker as String
      Dim greatestVolume as Double
      greatestTicker = ""
      greatestVolume = 0

    Dim TickerRow As Long: TickerRow = 1

        
   ' Loop through all of the stocks
      For i = 2 To Lastrow

   If openprice = 0 Then

          openprice = ws.Cells(i, 3).Value
      End If
      
   ' Check if we are still within the same Ticker Name, if not...
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     
     ' Add a row to the TickerRow
      TickerRow = TickerRow + 1
     
    ' Set the Ticker Name
     TickerName = ws.Cells(i, 1).Value
     
     ws.Cells(TickerRow, "I").Value = TickerName
     
     'Determine the change in ticker prices

        
          closeprice = ws.Cells(i, 6).Value
     
     
    pricechange = closeprice - openprice
    
    'Print out the change in price and fill the cells this colum depending on not profit or loss
    ws.Cells(TickerRow, "J").Value = pricechange

    If pricechange < 0 Then
            ws.Cells(TickerRow, "J").Interior.ColorIndex = 3
        ElseIf pricechange > 0 Then
            ws.Cells(TickerRow, "J").Interior.ColorIndex = 4


     End If
    'Highlight yearly change green or re depending on net profit or loss.

    

    'Determine and fill out the change
     
     percentpricechange = (pricechange / openprice)


     ws.Cells(TickerRow, "K").Value = percentpricechange
     
     ws.Cells(TickerRow, "K").NumberFormat = ".##%"
     
     If incVal < percentpricechange Then

     incVal = percentpricechange

     incTicker = TickerName

     End If

     If decVal >percentpricechange Then

     decVal = percentpricechange

     decTicker = TickerName

     End if
      
    'Reset openprice for every different ticker
    openprice = 0


    ' Add to the Ticker Total
    
    TickerTotal = TickerTotal + ws.Cells(i, 7).Value
    
    ws.Cells(TickerRow, "L").Value = TickerTotal
    ws.Cells(TickerRow, "L").NumberFormat = "0"

   If greatestVolume < TickerTotal Then

   greatestVolume = TickerTotal

   greatestTicker = TickerName

   End if

    ' Reset the TickerTotal
    TickerTotal = 0
    
    
    

    'If the the next row that follows up is of the same ticker name
    Else

    'Add to the TickerTotal
    TickerTotal = TickerTotal + ws.Cells(i, 7).Value

    End If
     
    Next i
    

   'Determine the biggeest gain and loss in terms of percentage and also determine the greatest volume for each sheet 

    ws.Cells(2, "P") = incTicker
    ws.Cells(2, "Q") = incVal

    ws.Range("Q1:Q3").NumberFormat = ".##%"

    ws.Cells(3, "P") = decTicker
    ws.cells(3, "Q") = decVal

    ws.Cells(4, "P") = greatestTicker
    ws.cells(4, "Q") = greatestVolume

    Next ws
    
End Sub




