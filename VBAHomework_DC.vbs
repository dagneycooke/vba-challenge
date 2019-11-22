Attribute VB_Name = "Module1"
' This macro will calculate yearly change, percent increase, and total stock volume for every year
' It will then calculate which stock had the greatest percent increase, the greatest percent decrease and the highest stock volume
' It will loop through all sheets in the workbook

Sub homework()

    ' initialize workbook
    Dim book As Workbook
    Set book = ActiveWorkbook
    
    ' determine total number of sheets in workbook
    totalSheets = book.Sheets.Count
     
    ' initialize all variables
    Dim rowTotal As Long ' total number of rows in sheet
    Dim columnTotal As Integer ' total number of columns in sheet
    Dim tickerCount As Integer ' number of different stocks in sheet
    Dim tickerName As String ' name of current stock
    Dim opening As Double ' open price at start of year
    Dim closing As Double ' closing price at end of year
    Dim numOfStocks As Integer ' number of stocks for current ticker

    
    ' determine total number of columns in the sheet
    columnTotal = Sheets(1).Cells(1, Columns.Count).End(xlToLeft).Column

' loop through all sheets using variable i
For i = 1 To totalSheets

    ' print labels into the necessary cells
    Sheets(i).Range("J1") = "Ticker"
    Sheets(i).Range("K1") = "Yearly Change"
    Sheets(i).Range("L1") = "Percent Change"
    Sheets(i).Range("M1") = "Total Stock Volume"
    Sheets(i).Range("Q1").Value = "Ticker"
    Sheets(i).Range("R1").Value = "Value"
    Sheets(i).Range("P2").Value = "Greatest % Increase"
    Sheets(i).Range("P3").Value = "Greatest % Decrease"
    Sheets(i).Range("P4").Value = "Greatest Total Volume"
    Sheets(i).Range("R2").NumberFormat = "0.00%" ' set this cell format to percentages
    Sheets(i).Range("R3").NumberFormat = "0.00%"  ' set this cell format to percentages
    
    ' reset variables for each sheet
    numOfStocks = 0
    opening = 0
    closing = 0
    tickerCount = 0
    
    'determine total number of rows in the sheet
    rowTotal = Sheets(i).Cells(Rows.Count, 2).End(xlUp).Row
        
    ' label all new columns

    
        ' loop through all rows
        For j = 2 To rowTotal
        
            ' sum total volume of stocks as you loop through the rows
            Sheets(i).Cells(tickerCount + 2, columnTotal + 6) = Sheets(i).Cells(j, 7) + Sheets(i).Cells(tickerCount + 2, columnTotal + 6)
            
            ' count the number of stocks per ticker
            numOfStocks = numOfStocks + 1
                
            ' if ticker name changes, do this
            If Sheets(i).Cells(j + 1, 1).Value <> Sheets(i).Cells(j, 1).Value Then
                
                ' increase ticker count by 1
                tickerCount = tickerCount + 1
                
                ' print ticker name in column 9
                Sheets(i).Cells(tickerCount + 1, columnTotal + 3) = Sheets(i).Cells(j, 1).Value
                 
                ' pull open value data from first row of the ticker
                opening = Sheets(i).Cells(j - (numOfStocks - 1), 3)
                
                ' pull close value data from last row of the ticker
                closing = Sheets(i).Cells(j, 6).Value
                 
                ' calculate yearly change
                Sheets(i).Cells(tickerCount + 1, columnTotal + 4) = closing - opening
                
                ' calculate percent change and then format it
                ' if one of the values is zero, assign percent change to 0
                If (closing = 0 Or opening = 0) Then
                    Sheets(i).Cells(tickerCount + 1, columnTotal + 5) = 0
                Else
                    Sheets(i).Cells(tickerCount + 1, columnTotal + 5) = ((closing / opening) - 1)
                End If
                Sheets(i).Cells(tickerCount + 1, columnTotal + 5).NumberFormat = "0.00%"
                 
                ' format the yearly change column depending on if the value is positive or negative
                If Sheets(i).Cells(tickerCount + 1, columnTotal + 4) > 0 Then
                    Sheets(i).Cells(tickerCount + 1, columnTotal + 4).Interior.ColorIndex = 4 ' if percent change is positive, make cell green
                ElseIf Sheets(i).Cells(tickerCount + 1, columnTotal + 4) < 0 Then
                    Sheets(i).Cells(tickerCount + 1, columnTotal + 4).Interior.ColorIndex = 3 ' if percent change is negative, make cell red
                End If
                
                ' reset the number of stocks for the next ticker iteration
                numOfStocks = 0
                    
            End If
        Next j
    Next i
    
' iterate through all summarize tickers in a sheet to find greatest percent increase, greatest percent decrease, and greatest total stock volume

' iterate through all sheets
For i = 1 To totalSheets
    
    ' calculate number of tickers
    rowTotal = Sheets(i).Cells(Rows.Count, 10).End(xlUp).Row
    
    ' iterate through tickers
    For j = 2 To rowTotal
    
        ' if the percent change value is higher than the one in R2, replace the value of R2 and then replace the ticker name
        If Sheets(i).Cells(j, 12) > Sheets(i).Range("R2") Then
            Sheets(i).Range("R2") = Sheets(i).Cells(j, 12).Value
            Sheets(i).Range("Q2") = Sheets(i).Cells(j, 10).Value
        End If
       ' if the percent change value is lower than the one in R3, replace the value of R3 and then replace the ticker name
        If Sheets(i).Cells(j, 12) < Sheets(i).Range("R3") Then
            Sheets(i).Range("R3") = Sheets(i).Cells(j, 12).Value
            Sheets(i).Range("Q3") = Sheets(i).Cells(j, 10).Value
        End If
        ' if the stock volume value is greater than the one in R4, replace the value of R4 and then replace the ticker name
        If Sheets(i).Cells(j, 13) > Sheets(i).Range("R4") Then
            Sheets(i).Range("R4") = Sheets(i).Cells(j, 13).Value
            Sheets(i).Range("Q4") = Sheets(i).Cells(j, 10).Value
        End If
    Next j
Next i


End Sub



