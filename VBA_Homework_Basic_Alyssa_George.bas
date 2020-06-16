Attribute VB_Name = "Module1"
'Create a script that will loop through all the stocks for one year and output the following information.



'The ticker symbol.
Sub ticker()
For Each ws In Worksheets
  
  ' Set variables as strings or doubles
  Dim tickername As String
  Dim openprice As Double
  Dim closeprice As Double
  Dim yearlychange As Double
  Dim percentchange As Double
  Dim lastrow As Double
  Dim totalstockvolume As Double
  
  totalstockvolume = 0
  
  ' Determine last row
   lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Keep track of the location for each ticker in the summary table
  Dim stockvolumerow As Integer
  stockvolumerow = 2

  ' Loop through all tickers
    For i = 2 To lastrow
 ' Set titles of table
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 12).Value = "Total Stock Volume"
 
 'Compare tickers
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker name for checking
      tickername = ws.Cells(i, 1).Value
      
      ' Add to the total stock volume
      totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value

      ' Put ticker name in the table
      ws.Range("I" & stockvolumerow).Value = tickername
      
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

      'Put yearly change in table
        'set open price location in table
        openprice = ws.Cells(2, 3).Value
        'set close price as end of cells
        closeprice = ws.Cells(i, 6).Value
        'subtract close price from open price
        yearlychange = closeprice - openprice
        'put in table
        ws.Range("J" & stockvolumerow).Value = yearlychange
        'set conditional formatting for colors of table
        If yearlychange < 0 Then
        ws.Range("J" & stockvolumerow).Interior.ColorIndex = 3
        ElseIf yearlychange >= 0 Then
        ws.Range("J" & stockvolumerow).Interior.ColorIndex = 4
        End If
        
        'reset open price
        openprice = ws.Cells(i + 1, 3)

'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
        If openprice = 0 Then
        percentchange = 0
        ElseIf yearlychange <> 0 Then
        percentchange = yearlychange / openprice
        End If
       ' print percent change to summary table as a percent
       
        ws.Range("K" & stockvolumerow).Value = percentchange
        ws.Range("K" & stockvolumerow).NumberFormat = "0.00%"
      ' Print the total stock volume to the Summary Table
      ws.Range("L" & stockvolumerow).Value = totalstockvolume

      ' Add one to the summary table row
      stockvolumerow = stockvolumerow + 1
      
      ' Reset
      totalstockvolume = 0
      

    Else
      'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
      
      ' Add to total stock volume
      totalstockvolume = totalstockvolume + ws.Cells(i, 3).Value

    End If


  Next i


Next ws
MsgBox ("Complete Macro")

End Sub


'You should also have conditional formatting that will highlight positive change in green and negative change in red.
