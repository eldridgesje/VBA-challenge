Attribute VB_Name = "Module1"
Sub stock_Looper()

    '***LOOP THROUGH WORKBOOKS***

    For Each ws In Worksheets

    
        'Count the rows in the workbook
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

    '***CREATE THE SUMMARY FIELDS***

        'Name the fields
        ws.Range("i1").Value = "Ticker"
        
        ws.Range("j1").Value = "Yearly Change"
    
        ws.Range("k1").Value = "Percent Change"

        ws.Range("l1").Value = "Total Stock Volume"
        
        
    '***DEFINE VARIABLES***
        
        'Establish a variable to track summary (Ticker) row
        Dim TickerRow As Double
    
        TickerRow = 1
        
        'Establish a variable to track Ticker start value
        Dim StartValue As Double
        
        'Establish a variable to track Ticker end value
        Dim EndValue As Double
        
        'Establish a variable to track Ticker volume
        Dim Volume As Double
        
        
    '***USE LOOPS TO SUMMARIZE DATA***
        
        'Loop through to establish Ticker values
    
        For i = 2 To LastRow
        
            If ws.Cells(i, 1).Value <> ws.Cells((i - 1), 1).Value Then
            
                'Add ticker rows
                TickerRow = TickerRow + 1
                
                'Add Ticker symbol to summary
                ws.Cells(TickerRow, 9).Value = ws.Cells(i, 1).Value
                
                
            End If
            
            Next i


                
        'Reset ticker
        TickerRow = 1
                     
                     
                     
        'Loop through yearly change and volume
        For i = 2 To (LastRow + 1)
        
            If ws.Cells(i, 1).Value <> ws.Cells((i - 1), 1).Value Then
            
            'Distinguish first ticker from later tickers
                If i > 2 Then
                
                'Later ticker logic: Establish Values and Volumes,then record summary
                
                    TickerRow = TickerRow + 1
                    
                    
                    'Set EndValue for ending Ticker
                    EndValue = ws.Cells((i - 1), 6).Value
                                    
                    'Summarize Yearly Change for ending Ticker
                    ws.Cells(TickerRow, 10) = EndValue - StartValue
                    
                   
                        'Control for StartValues of 0
                   
                        If StartValue = 0 Then
                   
                            ws.Cells(TickerRow, 11) = Format(0, "Percent")
                            
                            ws.Cells(TickerRow, 11).Interior.ColorIndex = 6
                                       
                        
                        'Calculate and assign Percentage Change
                        
                        Else
                   
                            ws.Cells(TickerRow, 11).Value = Format(((EndValue - StartValue) / StartValue), "Percent")
                            
                            'Conditional formatting
                            
                            If ws.Cells(TickerRow, 11).Value > 0 Then
                                
                                ws.Cells(TickerRow, 11).Interior.ColorIndex = 4
                            
                            ElseIf ws.Cells(TickerRow, 11).Value < 0 Then
                                
                                ws.Cells(TickerRow, 11).Interior.ColorIndex = 3
                                
                            Else
                            
                                ws.Cells(TickerRow, 11).Interior.ColorIndex = 6

                            End If
                    
                        End If
                    
                    
                    'Set StarTValue for new Ticker
                    StartValue = ws.Cells(i, 3).Value
                               
                    
                    'Record total Volume for old Ticker
                    ws.Cells(TickerRow, 12) = Volume
                    
                    'Set initial volume for new Ticker
                    Volume = ws.Cells(i, 7).Value
                    
                                        
                Else
                
                'First ticker logic: Begin Value and Volume tracking
                
                    
                    StartValue = ws.Cells(i, 3).Value
                    
                    Volume = ws.Cells(i, 7).Value
                
                
                End If
            
            'For intermediary dates, add to total Volume
            
            Else
            
                Volume = Volume + ws.Cells(i, 7).Value
            
            End If
        
            Next i



'   ***CREATE TABLE DISPLAYING GREATEST PERCENT INCREASE, LOWEST PERCENT DECREASE, AND HIGHEST TOTAL VOLUME***

    'Add labels for table
    ws.Range("p1").Value = "Ticker"
    ws.Range("q1").Value = "Value"
    ws.Range("o2").Value = "Greatest % increase"
    ws.Range("o3").Value = "Greatest % decrease"
    ws.Range("o4").Value = "Greatest total volume"


    'Loop to find the greatest increase
    
    
    'Declare variables
    
    Dim Increase As Double
    
    Dim HighTicker As String
    
    Increase = 0
    
    
    'Begin loop
    
    For i = 2 To TickerRow
    
        If ws.Cells(i, 11).Value > Increase Then
        
            Increase = ws.Cells(i, 11).Value
            
            HighTicker = ws.Cells(i, 9).Value
            
            
        End If
        
        Next i
        
    
    'Write results
    
    ws.Range("q2").Value = Format(Increase, "Percent")
    ws.Range("p2").Value = HighTicker



    'Loop to find the greatest decrease
    
    
    'Declare variables
    
    Dim Decrease As Double
    
    Dim LowTicker As String
    
    Decrease = 0
    
    'Begin loop
    
    For i = 2 To TickerRow
    
        If ws.Cells(i, 11).Value < Decrease Then
        
            Decrease = ws.Cells(i, 11).Value
            
            LowTicker = ws.Cells(i, 9).Value
                       
            
        End If
        
        Next i
    
    'Write results
        
    ws.Range("q3").Value = Format(Decrease, "Percent")
    ws.Range("p3").Value = LowTicker


    
    'Loop to find greatest Volume
    
    
    'Declare variables
    
    Dim HighVolume As Double
    
    Dim VolumeTicker As String
    
    HighVolume = 0
    
    'Begin loop
    
    For i = 2 To TickerRow
    
        If ws.Cells(i, 12).Value > HighVolume Then
        
            HighVolume = ws.Cells(i, 12).Value
            
            VolumeTicker = ws.Cells(i, 9).Value
                       
            
        End If
        
        Next i
        
    'Write results
    
    ws.Range("q4").Value = HighVolume
    ws.Range("p4").Value = VolumeTicker
    
    

    Next ws


End Sub

