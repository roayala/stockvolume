Sub TickerStock()

    ' Declaración de variables
    
    Dim ticker, tickervol, tickerhighpercent, tickerlowpercent As String
    Dim tvolume, last_row, tvolume_mayor As Long
    Dim yfinal, yinital, ypercent, ypercenthigh, ypercentlow As Double
    Dim i, x As Integer
            
            
     For Each ws In Worksheets
     
                'Counts how many rows are until the end
                last_row = Cells(Rows.Count, 1).End(xlUp).Row
            
                x = 2                                   'This variable will help to print on the correspondant Row
                ticker = Cells(2, 1).Value2    'Establish the first value of ticker
                yinitial = Cells(2, 3).Value2  'Takes the first value of the year
                tvolume_mayor = 0              'Assigns the value for the greatest volume
                ypercenthigh = 0                  'Assigns the value for highest percentage of change
                ypercentlow = 0                    'Assigns the value for lowest percentage of change
                
                '---------Separate by ticker
                
                
                For i = 2 To last_row + 1     'The For loop starts to check the values of the whole table and print them
                   
                   If ticker = Cells(i, 1) Then 'This If helps to check if the ticker has changed
                      tvolume = tvolume + Cells(i, 7).Value2  'Adds the value to volume
                      ticker = Cells(i, 1).Value2                       'Changes the value of ticker
                    
                   ElseIf ticker <> Cells(i, 1) Then                 'If the ticker is different,
                     Cells(x, 10).Value = ticker                       'Prints the las value of the ticker
                     Cells(x, 13).Value = tvolume                   'Prints the value of the volume
                     yfinal = Cells(i - 1, 6).Value2                   'takes the last value of the list
                     Cells(x, 11).Value2 = yfinal - yinitial       'Prints the value of the difference in the year
                     On Error Resume Next
                     ypercent = (yfinal - yinitial) / yinitial      'Assings the percent value of change
                     Cells(x, 12).Value2 = ypercent               'prints the percent
                     
                     
                        If ypercent > ypercenthigh Then        'This If, checks if the percent of change is the greatest
                        ypercenthigh = ypercent
                        tickerhighpercent = ticker                 'I used differnt variables to take the value of the ticker for each checker
                        End If
                     
                     If ypercent < ypercentlow Then             'This If, checks if the percent of change is the lowest
                     ypercentlow = ypercent
                     tickerlowpercent = ticker
                     End If
                     
                     If tvolumemayor < tvolume Then          'This if checks the greatest Volume
                     tvolumemayor = tvolume
                     tickervol = ticker
                     End If
                     
                     yinitial = Cells(i, 3).Value2                          'Takes the new value of the start of the period
                     Cells(x, 12).NumberFormat = "0.00%"          'Adds format
                     
                     If Cells(x, 11) < 0 Then
                      Cells(x, 11).Interior.ColorIndex = 3
                      Cells(x, 11).Font.ColorIndex = 30
                     Else
                        Cells(x, 11).Interior.ColorIndex = 43
                        Cells(x, 11).Font.ColorIndex = 10
                     End If
                      
                      x = x + 1
                     ticker = Cells(i, 1).Value2
                     tvolume = Cells(i, 7).Value2
            
            
                     
                    
                   End If
                
                
                Next i
                '-Add up by ticker for each year
               
               Range("J1").Value = "Ticker"
               Range("K1").Value = "Yearly Change"
               Range("L1").Value = "Percent Change"
                Range("M1").Value = "Total Stock Volume"
                Range("Q1").Value = "Ticker"
                Range("R1").Value = "Value"
                
                Range("R4").Value = tvolumemayor
                Range("Q4").Value = tickervol
                Range("P4").Value = "Greatest Total Volume"
               
                Range("R2").Value = ypercenthigh
                Range("Q2").Value = tickerhighpercent
                Range("P2").Value = "Greatest % Increase"
                Range("R2").NumberFormat = "0.00%"
               
                  Range("R3").Value = ypercentlow
                Range("Q3").Value = tickerlowpercent
                Range("P3").Value = "Greatest % Decrease"
                Range("R3").NumberFormat = "0.00%"
        Next ws
        
        
    End Sub
    