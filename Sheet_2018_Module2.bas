Attribute VB_Name = "Module2"
Sub ThreeYearStockData2018():

        Dim Worksheet As Worksheet: Set Worksheet = ThisWorkbook.Worksheets("2018")
        Dim nameOfWorksheet As String
        
        'Defining my current row
        Dim i As Long
        
        'Defining for Ticker
        Dim j As Long
        
        'Index Counter for Ticker row
        Dim tickerCount As Long
        
        'Defining the last row in column A
        Dim lastRowInA As Long
        
        'Defining the last row in column I
        Dim lastRowInI As Long
        
        'Defining percentageChange variable
        Dim percentageChange As Double
        
        'Defining percentageIncrease variable
        Dim greatestIncrease As Double
        
        'Defining greatestDecrease variable
        Dim greatestDecrease As Double
        
        'Defining the greatestVolume
        Dim greatestVolume As Double
        
        
        'assigning the name of the workSheet
        nameOfWorksheet = Worksheet.Name
        
        'Creating the column headers based on their Range
        Worksheet.Range("i1").Value = "Ticker"
        Worksheet.Range("j1").Value = "Yearly Change"
        Worksheet.Range("k1").Value = "Percent Change"
        Worksheet.Range("l1").Value = "Total Stock Volume"
        Worksheet.Range("P1").Value = "Ticker"
        Worksheet.Range("Q1").Value = "Value"
        Worksheet.Range("O2").Value = "Greatest % Increase"
        Worksheet.Range("O3").Value = "Greatest % Decrease"
        Worksheet.Range("O4").Value = "Greatest Total Volume"
        
        'setting the ticket count to avoid the header
        tickerCount = 2
        
        'starting row ignoring the header row
        j = 2
        
        'Find the last non-blank cell in column A
        lastRowInA = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox ("Last row in column A is " & lastRowInA)
        
            'using for loop to loop through the rows
            For i = 2 To lastRowInA
            
                
                If Worksheet.Cells(i + 1, 1).Value <> Worksheet.Cells(i, 1).Value Then
                
                'write and create ticker colum
                Worksheet.Cells(tickerCount, 9).Value = Worksheet.Cells(i, 1).Value
                
                'Calculating and writing year change
                Worksheet.Cells(tickerCount, 10).Value = Worksheet.Cells(i, 6).Value - Worksheet.Cells(j, 3).Value
                
                    
                    If Worksheet.Cells(tickerCount, 10).Value < 0 Then
                
                        'Using colour index to set cell colour to red
                        Worksheet.Cells(tickerCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                        'Using colour index to set cell colour to green
                        Worksheet.Cells(tickerCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculating and writing the percentage change
                    If Worksheet.Cells(j, 3).Value <> 0 Then
                    percentageChange = ((Worksheet.Cells(i, 6).Value - Worksheet.Cells(j, 3).Value) / Worksheet.Cells(j, 3).Value)
                    
                        'formatting to percentage
                        Worksheet.Cells(tickerCount, 11).Value = Format(percentageChange, "Percent")
                    
                    Else
                    
                        Worksheet.Cells(tickerCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'calculating and writing for volume column
                Worksheet.Cells(tickerCount, 12).Value = WorksheetFunction.Sum(Range(Worksheet.Cells(j, 7), Worksheet.Cells(i, 7)))
                
                'now we can increase the ticker count
                tickerCount = tickerCount + 1
                
                'increment the ticker
                j = i + 1
                
                End If
            
            Next i
            
        'allowing us to loop until finding the last empty row
        lastRowInI = Worksheet.Cells(Rows.Count, 9).End(xlUp).Row
        
        
        'declaring for the summary section
        greatestVolume = Worksheet.Cells(2, 12).Value
        greatestIncrease = Worksheet.Cells(2, 11).Value
        greatestDecrease = Worksheet.Cells(2, 11).Value
        
            'using for loops to loop and create the summary section
            For i = 2 To lastRowInI
            
                'if volume is greater than value therefore repopulate the cell
                If Worksheet.Cells(i, 12).Value > greatestVolume Then
                    greatestVolume = Worksheet.Cells(i, 12).Value
                    Worksheet.Cells(4, 16).Value = Worksheet.Cells(i, 9).Value
                
                Else
                
                    greatestVolume = greatestVolume
                
                End If
                
                'if greatestIncrease is greater than value therefore repopulate the cell
                If Worksheet.Cells(i, 11).Value > greatestIncrease Then
                    greatestIncrease = Worksheet.Cells(i, 11).Value
                    Worksheet.Cells(2, 16).Value = Worksheet.Cells(i, 9).Value
                
                Else
                
                    greatestIncrease = greatestIncrease
                
                End If
                
                'if greatestDecrease is lesser than value therefore repopulate the cell
                If Worksheet.Cells(i, 11).Value < greatestDecrease Then
                    greatestDecrease = Worksheet.Cells(i, 11).Value
                    Worksheet.Cells(3, 16).Value = Worksheet.Cells(i, 9).Value
                
                Else
                
                    greatestDecrease = greatestDecrease
                
                End If
                
            'writing the summary section to the worksheet
            Worksheet.Cells(2, 17).Value = Format(greatestIncrease, "Percent")
            Worksheet.Cells(3, 17).Value = Format(greatestDecrease, "Percent")
            Worksheet.Cells(4, 17).Value = Format(greatestVolume, "Scientific")
            
            Next i
            
        'function to automatically adjust spreadsheet
        Worksheets(nameOfWorksheet).Columns("A:Z").AutoFit
            
        
End Sub


