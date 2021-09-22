Attribute VB_Name = "Module1"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine summarizes the stocks with the following information:
'1. Ticker symbol
'2. Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'3.The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'4.The total stock volume of the stock.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub stockTickerSummary()

    '=======Variable declarations=======
    Dim aWorksheet As Worksheet
    Dim rowCnt As Long
    Dim workSheetName As String

    
    'Summary table variables
    Dim tickerSymbol As String
    Dim nextTickerSymbol As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVol As LongLong
    Dim yearlyChange As Double
    Dim percentageYearlyChange As Double
    Dim summaryTableRow As Integer
    
    
    'Additional information variables
    Dim greatestPercentageIncrease As Double
    Dim greatestPercentageDecrease As Double
    Dim greatestTotalVolume As Long
    
    
    '=======Variable initialization=======
     totalVol = 0
     summaryTableRow = 0
    
    'loop through all the worksheet in the workbook
    For Each aWorksheet In Worksheets
    
         aWorksheet.Activate  'Activate the worksheet
         summaryTableRow = 0
         
         rowCnt = aWorksheet.Cells(Rows.Count, 1).End(xlUp).Row 'Get the row count for the sheet
         
         'Loop through the rows to find the ticket symbol
         Dim rowNum As Long
         For rowNum = 2 To rowCnt
         
               'Capture the information for the first ticker symbol in the data set
               If (rowNum = 2) Then
                   openPrice = Cells(rowNum, 3).Value
               End If
            
               tickerSymbol = Cells(rowNum, 1).Value 'Get the Ticker symbol
              
               nextTickerSymbol = Cells(rowNum + 1, 1).Value  'Get the Ticker Symbol in the next row
               
                'Determine when the ticker symbol changes
                If tickerSymbol <> nextTickerSymbol Then
                  
                     closePrice = Cells(rowNum, 6).Value 'Capture the current ticker's closing price
                      
                     totalVol = totalVol + Cells(rowNum, 7).Value  'Add the volume to the total one last time for the current ticker
                     
                     ' ======= Update the Summary table =======
                      summaryTableRow = summaryTableRow + 1
                      Call updateSummary(tickerSymbol, openPrice, closePrice, totalVol, summaryTableRow)
                    
                      openPrice = Cells(rowNum + 1, 3).Value  'Capture the next ticker symbol's starting price
                             
                      totalVol = 0  'Reset the total volume
               Else
                    totalVol = totalVol + Cells(rowNum, 7).Value  'Add the volume to total
               End If
          
         Next rowNum
        
    Call updateBonusStats

    Next aWorksheet
    
End Sub


'Sub routine to update the summary table

Sub updateSummary(tickerSymbol As String, openPrice As Double, closePrice As Double, totalVolume As LongLong, summaryTableRow As Integer)

            Dim yearlyChange As Double
            Dim percentageYearlyChange As Double
     
            'if the summary table has not yet been created, add the header
            If summaryTableRow = 1 Then
                Cells(summaryTableRow, 13).Value = "Ticker"
                Cells(summaryTableRow, 14).Value = "Yearly Change"
                Cells(summaryTableRow, 15).Value = "Percentage Yearly Change"
                Cells(summaryTableRow, 16).Value = "Total Stock Volume"
                summaryTableRow = summaryTableRow + 1
            End If
       
            Cells(summaryTableRow, 13).Value = tickerSymbol 'Set the ticker name
          
            yearlyChange = closePrice - openPrice  'Calculate and Update the yearly change(openPrice - closePrice)
            Cells(summaryTableRow, 14).Value = yearlyChange  'format the yearly change cell based on the value
           
            If yearlyChange < 0 Then
                Range("n" & summaryTableRow).Interior.Color = VBA.ColorConstants.vbRed
            Else
                Range("n" & summaryTableRow).Interior.Color = VBA.ColorConstants.vbGreen
            End If
            
            'Calculate and update the percentage yearly change
            If openPrice <> 0 Then
                percentageYearlyChange = (yearlyChange / openPrice)
            Else
                percentageYearlyChange = 0
            End If

            Cells(summaryTableRow, 15).Value = percentageYearlyChange
            Range("o" & summaryTableRow).Style = "Percent"  'format the yearly change cell to show percentages
           
            Cells(summaryTableRow, 16).Value = totalVolume  'Update the total stock volume


End Sub


Sub updateBonusStats()

    Dim percentageChangeRange As Range
    Dim volRange As Range
    Dim greatestPercentageIncrease As Double
    Dim greatestPercentageDecrease As Double
    Dim greatestTotalVolume As LongLong
    
    Dim greatestPercentageIncreaseTicker As String
    Dim greatestPercentagedecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestPercentageRow As Integer
    Dim volumeRow As Integer
    Dim lastRow As Integer
    
    lastRow = Cells(Rows.Count, 13).End(xlUp).Row 'Get the row count for the summary sheet
    
    Set percentageChangeRange = Range("o2:o" & lastRow)
    Set volRange = Range("p2:p" & lastRow)
    
    greatestPercentageIncrease = Application.WorksheetFunction.Max(percentageChangeRange)
    greatestPercentageRow = WorksheetFunction.Match(greatestPercentageIncrease, percentageChangeRange, 0) + percentageChangeRange.Row - 1
    greatestPercentageIncreaseTicker = Range("m" & greatestPercentageRow).Value
     
    greatestPercentageDecrease = Application.WorksheetFunction.Min(percentageChangeRange)
    greatestPercentageRow = WorksheetFunction.Match(greatestPercentageDecrease, percentageChangeRange, 0) + percentageChangeRange.Row - 1
    greatestPercentagedecreaseTicker = Range("m" & greatestPercentageRow).Value
    
    greatestTotalVolume = Application.WorksheetFunction.Max(volRange)
    volumeRow = WorksheetFunction.Match(greatestTotalVolume, volRange, 0) + volRange.Row - 1
    greatestVolumeTicker = Range("m" & volumeRow).Value
   
    If Range(x1).Value = "" Then
        Range(x1).Value = "Ticker"
        Range(y1).Value = "Value"
        Range(w2).Value = "Greatest PercentageI ncrease"
        Range(w3).Value = "Greatest Percentage Decrease"
        Range(w4).Value = "Greatest Total Volume"
    End If
   
    Cells(2, 24).Value = greatestPercentageIncreaseTicker
    Cells(2, 25).Value = greatestPercentageIncrease
    Range("y2").Style = "percent"
    
    Cells(3, 24).Value = greatestPercentagedecreaseTicker
    Cells(3, 25).Value = greatestPercentageDecrease
    Range("y3").Style = "percent"
    
    Cells(4, 24).Value = greatestVolumeTicker
    Cells(4, 25).Value = greatestTotalVolume
    
End Sub



