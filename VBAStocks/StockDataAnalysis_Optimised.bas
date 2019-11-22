Attribute VB_Name = "Module1"
Sub StockDataAnalysis_Optimised()

'optimised the time complexity, O(n) runtime
'removed an extra For loop
'changed the last row calculation specific to a particular column

     Dim ws As Worksheet
    
    ' loop through all the sheets of the workbook
     For Each ws In ActiveWorkbook.Worksheets
       ws.Activate
       
        ' Add heading
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
       
        'Variables to hold values
        Dim Stock_Volume As Double
        Stock_Volume = 0
        Dim Ticker_Name As String
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
      
        
        'Variables for looping and counter
        Dim Row As Integer
        Row = 2
        Dim i As Long
       
        ' get the last row
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'Initial open price
        Open_Price = Cells(2, 3).Value
         
         ' Loop through all ticker symbol
        For i = 2 To lastRow
         ' Check if we are still with the same ticker symbol
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ' Set Ticker name
                Ticker_Name = Cells(i, 1).Value
                Cells(Row, 9).Value = Ticker_Name
                
                ' Add Total Volumn
                Stock_Volume = Stock_Volume + Cells(i, 7).Value
                Cells(Row, 12).Value = Stock_Volume
                
                ' Set close price
                Close_Price = Cells(i, 6).Value
                ' Add yearly change
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, 10).Value = Yearly_Change
                
                ' Add Percent Change
                'Error Handling if either close or Open price is 0 or for divide by 0 error
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf Open_Price = 0 Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, 11).Value = Percent_Change
                    Cells(Row, 11).NumberFormat = "0.00%"
                End If

                ' increment the row
                Row = Row + 1
                ' reset the Open Price
                Open_Price = Cells(i + 1, 3)
                ' reset the Volumn Total
                Stock_Volume = 0
            'if cells are the same ticker
            Else
                Stock_Volume = Stock_Volume + Cells(i, 7).Value
            End If
        Next i
        
        ' Determine the Last Row of Yearly Change per ws
        lastRow_YearlyChange = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
         ' Variables For Greatest % Increase, Greatest % Decrease, and Total Volume
        Dim Greatest_Increase As Double
        Greatest_Increase = 0
          Dim Greatest_Decrease As Double
        Greatest_Decrease = 0
          Dim Greatest_Total_Value As Double
        Greatest_Total_Value = 0
        
        ' Set the Cell Colors
        ' Set Values For Greatest % Increase, Greatest % Decrease, and Total Volume
        For j = 2 To lastRow_YearlyChange
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
                
           If Cells(j, 11).Value > Greatest_Increase Then
                Greatest_Increase = Cells(j, 11).Value
                Cells(2, 16).Value = Cells(j, 9).Value
                Cells(2, 17).Value = Cells(j, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(j, 11).Value < Greatest_Decrease Then
               Greatest_Decrease = Cells(j, 11).Value
                Cells(3, 16).Value = Cells(j, 9).Value
                Cells(3, 17).Value = Cells(j, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(j, 12).Value > Greatest_Total_Value Then
                Greatest_Total_Value = Cells(j, 12).Value
                Cells(4, 16).Value = Cells(j, 9).Value
                Cells(4, 17).Value = Cells(j, 12).Value
            End If
        Next j
        
        
    Next ws
        
End Sub








