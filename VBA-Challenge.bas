Attribute VB_Name = "Module1"
Sub VBA_Challenge()

       'Setting the dimensions to be used in the Sub

Dim Ticker As String
Dim Summary_Table As Long
Dim Sold_Stock As Double
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim i As Long
Dim Lastrow As Long
Dim TickerCount As Double
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total As Double
Dim Greatest_Increase_Name As String
Dim Greatest_Decrease_Name As String
Dim Greatest_Total_Name As String


       'Setting starting values

         TickerCount = 0
         Close_Price = 0
         Yearly_Change = 0
         Percent_Change = 0
         Greatest_Increase = 0
         Greatest_Total = 0
         Greatest_Decrease = 0
         
       ' Aligning summary tables
    
        Summary_Table = 2
        Greatest_Table = 2
        
   For Each ws In Worksheets
     
       'Detmining initial Opening price
       
        Open_Price = ws.Cells(2, 3).Value
        Greatest_Increase = 0
        Greatest_Total = 0
        Greatest_Decrease = 0
       'Setting headers for new tables
      
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        
             
       ' Defining last row
       
        Lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
       'Opening loop
        
        For i = 2 To Lastrow
           
       ' Determining if ticker name is the same as the previous value
           
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
       'if it doesn't then grab the ticker name, opening price and closing price
           
                Ticker = ws.Cells(i, 1).Value
                 
                Close_Price = ws.Cells(i, 6).Value
                
       'Calculating the change over the period to determine the yearly change
            
                Yearly_Change = Close_Price - Open_Price
                
       'determining the percentage change over the period of time and populating Greatest % increase & % Decrease
            
                                  
                    
                    If Open_Price <> 0 Then Percent_Change = (Yearly_Change / Open_Price)
                    
                
                        If (Percent_Change > Greatest_Increase) Then
                            Greatest_Increase = Percent_Change
                            Greatest_Increase_Name = Ticker
                            ws.Range("P2").Value = Greatest_Increase_Name
                            ws.Range("Q2").Value = Greatest_Increase
                            
                        End If
                        
                        If (Percent_Change < Greatest_Decrease) Then
                            Greatest_Decrease = Percent_Change
                            Greatest_Decrease_Name = Ticker
                            ws.Range("P3").Value = Greatest_Decrease_Name
                            ws.Range("Q3").Value = Greatest_Decrease
                                
                        End If
                   
                        If (Sold_Stock > Greatest_Total) Then
                            Greatest_Total = Sold_Stock + ws.Cells(i, 7).Value
                            Greatest_Total_Name = Ticker
                            ws.Range("P4").Value = Greatest_Total_Name
                            ws.Range("Q4").Value = Greatest_Total
                        End If
                    
                                             
                
       ' Populating the table
             
                Sold_Stock = Sold_Stock + ws.Cells(i, 7).Value
                
                ws.Range("M" & Summary_Table).Value = Sold_Stock
            
                ws.Range("J" & Summary_Table).Value = Ticker
                
                ws.Range("K" & Summary_Table).Value = Yearly_Change
                
                ws.Range("L" & Summary_Table).Value = Percent_Change
                
                
                                               
                                
       'Conditional formatting based on positive or negitive result
              
                    If (Yearly_Change > 0) Then
                
                         ws.Range("K" & Summary_Table).Interior.ColorIndex = 4
                         
                    ElseIf (Yearly_Change <= 0) Then
                    
                         ws.Range("K" & Summary_Table).Interior.ColorIndex = 3
                    
                    End If
                    
       'moving the displayed data to the next row
              
                Summary_Table = Summary_Table + 1
                
       'Resetting the table to move onto the next ticker
         
              Sold_Stock = 0
              Open_Price = ws.Cells(i + 1, 3).Value
              Close_Price = 0
              Yearly_Change = 0
              Percent_Change = 0
              
       'If the previous ticker value is the same as current i then add the value to cell
            Else
                Sold_Stock = Sold_Stock + ws.Cells(i, 7).Value
            
            End If
            
          
       'Reseting nested loop
       
        Next i
       
       'Resetting the table position before moving onto the next worksheet
           
            Summary_Table = 2
            Greatest_Table = 2
            
       'Correcting cell formatting with new data entered
       
             ws.Columns.AutoFit
             ws.Range("L" & Summary_Table).NumberFormat = "0.0%"
             ws.Range("Q2", "Q3").NumberFormat = "0.0%"
             ws.Range("Q4").NumberFormat = "0.00E+00"
             
       'Moving to the next worksheet
       
    Next ws
    
    
End Sub
