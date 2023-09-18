Attribute VB_Name = "Module1"
Sub stock_analysis()

'Variable to store ticker
Dim ticker As String

'Variable to store opening value for a ticker
Dim openValue As Double
openValue = 0

'Variable to store calculated yearly change for a ticker
Dim yearlyChange As Double
yearlyChange = 0

'Variable to store calculated percent change for a ticker
Dim percentChange As Double
percentChange = 0

'Variable for accumulating total stock volume for a ticker
Dim totalStockVolume As LongLong
totalStockVolume = 0

'Row counter for tracking next row for writing ticker values
'Initial value is 2, since this will be first row written to.
Dim rowCounter As Integer
rowCounter = 2

'Write column headers and set percentage style
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Columns("K").NumberFormat = "0.00%"
Range("L1").Value = "Total Stock Volume"
Range("N2").Value = "Greatest % Increase"
Range("P2").NumberFormat = "0.00%"
Range("N3").Value = "Greatest % Decrease"
Range("P3").NumberFormat = "0.00%"
Range("N4").Value = "Greatest Total Volume"

'Find last row in the worksheet that has data
'Source: https://stackoverflow.com/questions/18088729/row-count-where-data-exists
With ActiveSheet
lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
'MsgBox lastRow
End With

'Loop through all data rows of worksheet.
For i = 2 To lastRow
    
    'If it is first row of ticker, then save ticker name and opening value
    'and add volume to stock volume running total
    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        ticker = Cells(i, 1).Value
        openValue = Cells(i, 3).Value
        totalStockVolume = totalStockVolume + Cells(i, 7).Value
    
    'If it is last row of ticker, then add volume to runnning total,
    'calculate yearly change and percentage change,
    'write values in columns I to L,
    'check if current values should update greatest values on columns O and P
    'and add 1 to row counter and reset totalStockValue to 0
    ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        totalStockVolume = totalStockVolume + Cells(i, 7).Value
        
        'Yearly change = close value - open value
        yearlyChange = Cells(i, 6).Value - openValue
        'Percent change = yearly change / open value
        percentChange = yearlyChange / openValue
        
        Cells(rowCounter, 9).Value = ticker
        Cells(rowCounter, 10).Value = yearlyChange
        Cells(rowCounter, 11).Value = percentChange
        Cells(rowCounter, 12).Value = totalStockVolume
        
        'Format background color:
        'green if positive change
        'red if negative change
        'no color fill if 0 change
        If yearlyChange > 0 Then
            Cells(rowCounter, 10).Interior.ColorIndex = 4
            Cells(rowCounter, 11).Interior.ColorIndex = 4
        ElseIf yearlyChange < 0 Then
            Cells(rowCounter, 10).Interior.ColorIndex = 3
            Cells(rowCounter, 11).Interior.ColorIndex = 3
        End If
                      
        'Check if current percent change is either greatest increase or decrease and if so update spreadsheet
        If percentChange > Range("P2").Value Then
            Range("O2").Value = ticker
            Range("P2").Value = percentChange
        ElseIf percentChange < Range("P3").Value Then
            Range("O3").Value = ticker
            Range("P3").Value = percentChange
        End If
        
        'Check if current total stock volume is greatest and if so update spreadsheet
        If totalStockVolume > Range("P4").Value Then
            Range("O4").Value = ticker
            Range("P4").Value = totalStockVolume
        End If
        
        rowCounter = rowCounter + 1
        totalStockVolume = 0
        
    'If neither first or last row of a ticker, then add volume to runnning total.
    Else
        totalStockVolume = totalStockVolume + Cells(i, 7).Value
    End If
    
Next i

End Sub

