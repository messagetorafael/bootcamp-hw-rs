Sub allworksheet()
    
    '---------------------------------------------------------------
    ' In case of execution iterations, make sure to first clean up any previous/residual
    ' data + formating in the range before running the code
    '---------------------------------------------------------------
        
    For Each ws In Worksheets
    
        ws.Range("I:Z").Clear

    Next ws

    '#################################################################
    '#  BEGIN OF THE MAIN CODE
    '#################################################################


'TA: THIS IS THE "CHALLENGE" PART OF THE ASSIGMENT
    
    '#---------------------------------------------------------------
    '#LOOPING THRU ALL THE WORKSHEETS
    '#---------------------------------------------------------------
    For Each ws In Worksheets
        
        '#*-----------------------------------------------------------*
        '#* Declaring variables and defining initial values
        '#*-----------------------------------------------------------*

        '#*Rows counter/index
        '#*---------------------------------------------------------------
            Dim lastrow As Long                 'used to check where the last row with data in the column is
                
                lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row '>>> Count rows within column with contigous data
            
            Dim r As Long                       'store row position for loops and cell reference commands
        
        '#*Ticker codes
        '#*---------------------------------------------------------------
            Dim ticker As String                   ' temporary holds ticker code/name:
                                                   '   1) to be copied to the summary table.
                                                   '   2) for reference to calculate other info in the summary table.
            
            Dim tickerCodeMaxVol As String         ' stores ticker code for greatest total volume
            Dim tickerGreatestIncrease As String   ' ticker ticker code for greatest percent increase
            Dim tickerGreatestDecrease As String   ' ticker ticker code for greatest percent decrease
        
        
        '#*Ticker Counters for listing
        '#*---------------------------------------------------------------
            Dim tickercount As Double              'used for incrementing row position/list each ticker in the summary table
                tickercount = 2                    'off-setting so the list in the summary table starts from row 2
        
        
        '#*Volumes
        '#*---------------------------------------------------------------
            Dim tickerTotalVolSum As Double        'stores the total volume per ticker
                tickerTotalVolSum = 0              'Purposedly positioned to reset value in every loop cycle.
            
            Dim tickerGreatestVolume As Double     'stores the greatest total volume
                tickerGreatestVolume = 0           'Purposedly positioned to reset value in every loop cycle.
        
                        
        '#*Variables for Date, StockPrice, Volume, Percentage comparisons
        '#*---------------------------------------------------------------
            Dim YearBegin As Double                'used to compare and help determine earliest date in the data range
                YearBegin = 99999999               'Purposedly positioned to reset value in every loop cycle.
            
            Dim YearEnd As Double                  'used to check analysis was correct
                YearEnd = 0                        'Purposedly positioned to reset value in every loop cycle.
            
            Dim YearBeginStockOpen As Double       '(+) used to capture stock price
            Dim YearEndStockClose As Double        '(-) used to capture stock price
            Dim YearChange As Double               '(=) holds difference from two variables above
            
            Dim YearChangePercent As Double        'holds the numeric change value for convertion to percentual
            Dim YearChangePercentMin As Double     'holds the value for the Greatest Yearly Decrease
                YearChangePercentMin = 100         'Purposedly positioned to reset value in every loop cycle.
                
            Dim YearChangePercentMax As Double     ' holds the value for the Greatest Yearly Increase
                YearChangePercentMax = 0           'Purposedly positioned to reset value in every loop cycle.
            
                
        '##---------------------------------------------------------------
        '## LOOPING THRU THE LIST (DATA IN THE SHEET)
        '##--------------------------------------------------------------
            '## Summarizing Volumes, then grouping/listing ticker, Yearly Change,
            '## Percentage Change, Total Stock, ...
            '##---------------------------------------------------------------
            
            For r = 2 To lastrow
                        
            '##*---------------------------------------------------------------
            '## [MAIN IF] IF Ticker Code changes from one row to another, then copy and print
            '## the INFO on the summary table.
            '## ELSE, ticker does not change on next row then just sum volume up.
            '##*---------------------------------------------------------------

'TA: THIS IS THE EASY PART OF THE ASSIGMENT

                If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then

                    ticker = ws.Cells(r, 1).Value
                    tickerTotalVolSum = tickerTotalVolSum + ws.Cells(r, 7).Value
                    
                    ws.Range("I" & tickercount).Value = ticker                  '>>> print [EASY]
                    ws.Range("L" & tickercount).Value = tickerTotalVolSum       '>>> print [EASY]
                    ws.Range("L" & tickercount).NumberFormat = "$ #,##0.00" '>>> currency format

'TA: THIS IS THE MODERATE PART OF THE ASSIGMENT
                    '##*>---------------------------------------------------------------
                    '##*> 1) Determine Start/End dates to determine Open/Close stock values
                    '##*>---------------------------------------------------------------
                        If ws.Cells(r, 2).Value <= YearBegin Then
                                               
                                YearBeginStockOpen = ws.Cells(r, 3).Value  '>>>Open Stock value
                             
                            ElseIf ws.Cells(r, 2).Value > YearEnd Then
                            
                                YearEndStockClose = ws.Cells(r, 6).Value   '>>>Close Stock Value
                            
                        End If
                        
                    '##*> 1) THEN calculate and print: Yearly Change on Stock value with conditional formating
                    '##*>---------------------------------------------------------------
                        YearChange = (YearBeginStockOpen - YearEndStockClose)
                                                
                        With ws.Range("J" & tickercount)
                        
                            .Value = YearChange                                 '>>> print [MODERATE]
                            .NumberFormat = "$ #,##0.00"   '>>> currency number format
                            
                            If YearChange > 0 Then         '>>> conditional formatting
                            
                                    .Interior.ColorIndex = "4" 'Green
                                
                                ElseIf YearChange < 0 Then
                                
                                    .Interior.ColorIndex = "3" 'Red
                                
                            End If
                        
                        End With

                    '##*>---------------------------------------------------------------
                    '##*> 1) Calculate and print Yearly Change %
                    '##*>---------------------------------------------------------------
                   
                        If (YearChange = 0 Or YearBeginStockOpen = 0) Then  '<This IF is a fix to the OVERFLOW bug.
                                                                            'Removed all attempts to divide any
                                YearChangePercent = 0                       'possible value equals to *zero* or divided by 0>
                                                                
                                ws.Range("K" & tickercount).Value = FormatPercent(YearChangePercent) '>>> print [MODERATE]
                                
                            Else

                                YearChangePercent = (YearChange) / (YearBeginStockOpen)
                                ws.Range("K" & tickercount).Value = FormatPercent(YearChangePercent) '>>> print [MODERATE]
                                
                        
'TA: THIS IS THE HARD PART OF THE ASSIGMENT
                        
                            '##*>>---------------------------------------------------------------
                            '##*>> 2) Determine Tickers/Value for Greatest % Increase/Decrease
                            '##*>>---------------------------------------------------------------
                                  
                                '##*>>Storing Ticker for greatest Year Percentage Increase
                                If YearChangePercent > YearChangePercentMax Then
                                             
                                        YearChangePercentMax = YearChangePercent        '>>> stores during the list loop
                                        tickerGreatestIncrease = ws.Cells(r, 1).Value   '>>> stores during the list loop
                                                 
                                '##*>>'Copying Ticker for greatest Year Percentage decrease
                                    ElseIf YearChangePercent < YearChangePercentMin Then
                                                 
                                        YearChangePercentMin = YearChangePercent        '>>> stores during the list loop
                                        tickerGreatestDecrease = ws.Cells(r, 1).Value   '>>> stores during the list loop
                                             
                                End If
                                                    
                            '##*>>---------------------------------------------------------------
                            '##*>> 3) Determine Tickers/Value for Greatest Total Volume
                            '##*>>---------------------------------------------------------------
                                If tickerTotalVolSum > tickerGreatestVolume Then
                                                
                                           tickerGreatestVolume = tickerTotalVolSum '>>> stores during the list loop
                                           tickerCodeMaxVol = ticker                '>>> stores during the list loop
                                        
                                End If

                        
                        End If
                        
                    
                    '##*> Reseting variables used as reference counters before starting next cycle (list loop)
                    '##*>---------------------------------------------------------------
                        YearBegin = 99999999
                        YearEnd = 0
                        tickerTotalVolSum = 0
                        test = 0
                    
                        
                    '##*> 'Increment so that next ticker prints on next row
                    '##*>---------------------------------------------------------------
                        tickercount = tickercount + 1
                    
                    
            '##* [MAIN IF] ELSE, ticker does not change on next row then:
            '##*---------------------------------------------------------------
                    Else
                
'TA: THIS IS THE EASY PART OF THE ASSIGMENT
                      
                    '##*> Add volume to the total of the ticker
                    '##*>---------------------------------------------------------------
                        tickerTotalVolSum = tickerTotalVolSum + ws.Cells(r, 7).Value
                        
'TA: THIS IS THE MODERATE PART OF THE ASSIGMENT
                    '##*> 1) Check Start/End dates to determine Open/Close stock values
                    '##*>---------------------------------------------------------------
                    
                        If ws.Cells(r, 2).Value <= YearBegin Then
                        
                            YearBegin = ws.Cells(r, 2).Value            'this is not needed, used just to confirm correct dates were taken
                            YearBeginStockOpen = ws.Cells(r, 3).Value   '>>> stores during the list loop
                             
                        ElseIf ws.Cells(r, 2).Value > YearEnd Then
                            
                            YearEnd = ws.Cells(r, 2).Value              'idem
                            YearEndStockClose = ws.Cells(r, 6).Value    '>>> stores during the list loop
                            
                        End If
                                     
            '##* [MAIN IF] END
            '##*---------------------------------------------------------------
            
                End If
                 
        '## END LOOPING THRU THE LIST (DATA IN THE SHEET)
        '##--------------------------------------------------------------
            Next r
            
            
'TA: THIS IS THE HARD PART OF THE ASSIGMENT
        
        '#*-----------------------------------------------------------*
        '#* Create the table for Summary of the Greatest
        '#*-----------------------------------------------------------*
                   
            '->>>>>>>>>>>>>>>>>>>> '"Greatest % Increase"
            ws.Cells(2, 15).Value = tickerGreatestIncrease              '>>> Print [HARD]
            ws.Cells(2, 16).Value = FormatPercent(YearChangePercentMax) '>>> Print [HARD]
            
            '->>>>>>>>>>>>>>>>>>>> "Greatest % Decrease"
            ws.Cells(3, 15).Value = tickerGreatestDecrease              '>>> Print [HARD]
            ws.Cells(3, 16).Value = FormatPercent(YearChangePercentMin) '>>> Print [HARD]
            
            '->>>>>>>>>>>>>>>>>>>> "Greatest Total volume"
            ws.Cells(4, 15).Value = tickerCodeMaxVol                    '>>> PRINT [HARD]
            ws.Cells(4, 16).Value = tickerGreatestVolume                '>>> PRINT [HARD]
            ws.Cells(4, 16).NumberFormat = "$ #,##0.00" '>>> currency format
     
        '#* Create headers for the Ticker summary table
        '#*-----------------------------------------------------------*
        
            ' <formatting header cells>
            With ws.Range("I1:P1")
                .Font.Size = 14
                .Font.FontStyle = "Bold"
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                          
            End With
                
            '->>>>>>>>>>>>>>>>>>>> Header names
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
            ' <formatting header cells>
            With ws.Range("N2:N4")
                .Font.Size = 12
                .Font.FontStyle = "Bold"
                .HorizontalAlignment = xlLeft
                
            End With
        
        
        '#* Adjust headers and formatting of the "Greatest summary table
        '#*-----------------------------------------------------------*
        
            '->>>>>>>>>>>>>>>>>>>> Header names
            ws.Range("N2").Value = "Greatest % Increase"
            ws.Range("N3").Value = "Greatest % Decrease"
            ws.Range("N4").Value = "Greatest Total Volume"
            
            ' <formatting header cells>
            With ws.Range("O1:P1")
                .Font.Size = 14
                .Font.FontStyle = "Bold"
                .HorizontalAlignment = xlLeft
                            
            End With
            
            '->>>>>>>>>>>>>>>>>>>> Header names
            ws.Range("O1").Value = "Ticker"
            ws.Range("P1").Value = "Value"
                
              
        '#* Automatically adjust the columns' width to content
        '#*-----------------------------------------------------------*
            ws.Range("I1:P1").EntireColumn.AutoFit
    
    '# GO TO THE NEXT WORKSHEET
    '#---------------------------------------------------------------
                    
    Next ws
    
    '# END OF LOOPING THRU ALL THE WORKSHEETS
    '#---------------------------------------------------------------
       
    '#################################################################
    '#  END OF THE MAIN CODE
    '#################################################################
       
End Sub
