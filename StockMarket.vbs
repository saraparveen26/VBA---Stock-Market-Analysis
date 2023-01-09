Attribute VB_Name = "StockMarket"
Sub StockMarket()

'Loop through all worksheets in the workbook

    Dim ws As Worksheet

        For Each ws In ThisWorkbook.Worksheets



'Clear all previous formatting
    
        ws.Range("A:Q").FormatConditions.Delete
    
    


'Count number of data entries by determining number of last filled row in column A and storing it in LastRowA

    Dim LastRowA As Double

        LastRowA = ws.Cells(Rows.Count, "A").End(xlUp).Row





'=====> Add required column headers for data analysis

        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

            'Make the headers Bold

                With ws.Range("I1:L1")
                    .Font.Bold = True
                End With
    
       
       

'Declare variables

    'Input variables
        'Declare the variable counterA to loop through column A for Total Stock data
        'Declare the variable counterC to move down column C for Stock Open Price
        'Declare the variable counterF to move down column F for Stock Close Price
        'Declare the variable counterG to move down column G for Stock Volume
        
    'Output variables
        'Declare the variable counterI to move down column I for Unique Ticker/Stock values)
        'Declare the variable counterJ to move down column J for Yearly Change
        'Declare the variable counterK to move down column K for Percent Change
        'Declare the variable counterL to move down column L for Stock Volume
    
        Dim counterA As Double
        Dim counterC As Double
        Dim counterF As Double
        Dim counterG As Double
        Dim counterI As Double
        Dim counterJ As Double
        Dim counterK As Double
        Dim counterL As Double
        
    'Declare variables to store opening price, closing pricem yearly change, percent change, and total stock volume
        
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        
        
        
        
'=====> Retrieve UNIQUE TICKER/STOCK values to Column I with header "Ticker"

        'Assign starting value to start at 2nd row for output into column I so that header row is excluded

            counterI = 2
            
        'Loop through the original stock data to retrieve unique values
            
            For counterA = 2 To LastRowA
            
                'Copy unique Ticker values to column I "Ticker"
                
                    If ws.Cells(counterA + 1, "A").Value <> ws.Cells(counterA, "A").Value Then
                    
                        ws.Cells(counterI, "I").Value = ws.Cells(counterA, "A").Value
                
                            'Move down to next row in output column I "Ticker" for next output
                    
                            counterI = counterI + 1
                
                    End If
            
            'Move to next value in the loop
            
            Next counterA
        
    
  
'=====> Calculate YEARLY CHANGE, PERCENT CHANGE and TOTAL STOCK VOLUME
    
        'Assign starting value to start at 2nd row for columns C through L so that header row is excluded
 
            counterC = 2
            counterF = 2
            counterG = 2
            counterI = 2
            counterJ = 2
            counterK = 2
            counterL = 2
         
         'Assign starting value of zero to the variable TotalVolume
         
            TotalVolume = 0
         
        
        'Loop through original stock data to retrieve values for opening price, closing price and calculating total stock volume
    
            For counterA = 2 To LastRowA
                
    
                    'Add stock volume from column G and store it in the variable TotalVolume
                
                        TotalVolume = TotalVolume + ws.Cells(counterG, "G").Value
    
    
                'Determine Opening Stock Price
     
                    If ws.Cells(counterA, "A").Value = ws.Cells(counterI, "I").Value Then
                
                        'Retrieve opening price and store it in the variable OpenPrice
                    
                         OpenPrice = ws.Cells(counterC, "C").Value
           
                            'Move to next row in the column I for next Ticker
                  
                            counterI = counterI + 1
   
                    End If
  
  
  
                'Determine Closing Stock Price, Yealy Change, Percent Change and Total Stock Volume
    
                    If ws.Cells(counterA + 1, "A").Value <> ws.Cells(counterA, "A").Value Then
                        
                        'Retrieve closing price and store it in the variable ClosePrice
                            
                            ClosePrice = ws.Cells(counterF, "F").Value
    
                        'Calculate Yearly Change and output it in the column J "Yearly Change"
                            
                            YearlyChange = ClosePrice - OpenPrice
                            
                            ws.Cells(counterJ, "J").Value = YearlyChange
            
                        'Calculate Percent Change and output it in the column K "Percent Change"
                         
                            PercentChange = (YearlyChange / OpenPrice)
                            
                            ws.Cells(counterK, "K").Value = FormatPercent(PercentChange)
            
                        'Retrieve Total Stock Volume to output it in column L "Total Stock Volume"
                            
                            ws.Cells(counterL, "L").Value = TotalVolume
                                
                                'Reset Total Stock Volume to zero to restart for the next unique ticker/ stock
                                
                                TotalVolume = 0
            
                                     'Move down to next row in output column J, K and L for next output
                                     
                                        counterJ = counterJ + 1
                                        counterK = counterK + 1
                                        counterL = counterL + 1
    
                    End If
    
        
                'Move to next row in the column C, F and G for next Open and Close Price and Volume
                 
                    counterC = counterC + 1
                    counterF = counterF + 1
                    counterG = counterG + 1
      
      
            'Move to next value in the loop
            
            Next counterA
   
    
    
    
'=====> Apply CONDITIONAL FORMATTING

        'Declare variables to identify the target data range and formatting conditions

            Dim ChangeRange As Range
            Dim Positive As FormatCondition
            Dim Negative As FormatCondition
            Dim NoChange As FormatCondition

        'Identify target data range as filled rows under columns "Yearly Change" and "Percent Change"

            Set ChangeRange = ws.Range(ws.Range("J2"), ws.Range("K2").End(xlDown))

        'Delete any previous formatting for identified range

            ChangeRange.FormatConditions.Delete

        'Assign three criteria for formatting: Greater than 0, Less than 0, Equal to 0

            Set Positive = ChangeRange.FormatConditions.Add(xlCellValue, xlGreater, "=0.00")
            Set Negative = ChangeRange.FormatConditions.Add(xlCellValue, xlLess, "=0.00")
            Set NoChange = ChangeRange.FormatConditions.Add(xlCellValue, xlEqual, "=0.00")


        'Change cell color to Green if value is positive/greater than 0

            With Positive
                .Interior.Color = vbGreen
            End With

        'Change cell color to Red if value is negative/less than 0

            With Negative
                .Interior.Color = vbRed
            End With

        'Change cell color to Yellow if value has no changes/equal to 0

            With NoChange
                .Interior.Color = vbYellow
            End With

                
                
    
    
    
'=====> Calculate the GREATEST numbers for % INCREASE, % DECREASE, and TOTAL VOLUME

        'Add headers for columns and rows

            ws.Range("P1:Q1").Value = Array("Ticker", "Value")
            ws.Range("O2:O4").Value = ws.Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))

                'Make the headers Bold

                    With ws.Range("P1:Q1")
                        .Font.Bold = True
                    End With

                    With ws.Range("O2:O4")
                        .Font.Bold = True
                    End With

        'Count number of unique ticker value by determining last filled row in column I and storing it in LastRowI

            Dim LastRowI As Double

                LastRowI = ws.Cells(Rows.Count, "I").End(xlUp).Row
            

        'Declare variables to store greatest Increase, Decrease and StockVolume
            
            Dim Increase As Double
            Dim Decrease As Double
            Dim StockVolume As Double

        'Assign starting value to start at 2nd row for columns I, K and L so that header row is excluded
 
            counterI = 2
            counterK = 2
            counterL = 2
        
        'Assign starting value of zero to the variables Increase, Decrease and StockVolume
         
            Increase = 0
            Decrease = 0
            StockVolume = 0


        'Loop through retrieved data for unique values to calculate greatest increase, greatest decrease and greatest total volume
    
            For counterI = 2 To LastRowI
            
                    'Retrieve the greatest increase, store it in the variable "Increase"

                        If ws.Cells(counterI, "K").Value > Increase Then
                        
                        Increase = ws.Cells(counterI, "K").Value
                        
                            'Display the respective ticker symbol in column P

                                ws.Cells(2, "P").Value = ws.Cells(counterI, "I").Value
                        End If


                    
                    'Retrieve the greatest decrease, store it in the variable "Decrease"

                        If ws.Cells(counterI, "K").Value < Decrease Then

                        Decrease = ws.Cells(counterI, "K").Value
                        
                            'Display the respective ticker symbol in column P

                                ws.Cells(3, "P").Value = ws.Cells(counterI, "I").Value
                        End If
                        


                    'Retrieve the greatest volume, store it in the variable "StockVolume"

                        If ws.Cells(counterI, "L").Value > StockVolume Then
                
                        StockVolume = ws.Cells(counterI, "L").Value
                            
                            'Display the respective ticker symbol in column P

                                ws.Cells(4, "P").Value = ws.Cells(counterI, "I").Value

                        End If

            Next counterI

    
    'Display the retrieved values for greatest increase, decrease and volume in column Q

        ws.Cells(2, "Q").Value = FormatPercent(Increase)
        ws.Cells(3, "Q").Value = FormatPercent(Decrease)
        ws.Cells(4, "Q").Value = StockVolume




    'Adjust column width to display full values properly

        ws.Columns("A:Q").EntireColumn.AutoFit
 
    
    
    
    'Move to next worksheet
    
        Next ws
    
                


End Sub

