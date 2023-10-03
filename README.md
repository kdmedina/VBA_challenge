# VBA_challenge
Module 2 Challenge


    Sub alphabetical_testing()

    Dim currentName As String
    Dim nextName As String
    Dim totalSV As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim i As Long
    Dim lastRow As Long
    Dim curSheet As Worksheet
    
    ' Loop through all worksheets
For Each curSheet In ThisWorkbook.Worksheets
    curSheet.Cells(1, 9).Value = "Ticker"
    curSheet.Cells(1, 10).Value = "Yearly Change "
    curSheet.Cells(1, 11).Value = "Percent Change "
    curSheet.Cells(1, 12).Value = "Total Stock Volume"
    
    totalSV = 0
    groupNo = 1
        
    openPrice = curSheet.Cells(2, 3).Value
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        
        currentName = curSheet.Cells(i, 1).Value
        nextName = curSheet.Cells(i + 1, 1).Value

    If nextName = currentName Then
        totalSV = totalSV + curSheet.Cells(i, 7).Value
    Else
        totalSV = totalSV + Cells(i, 7).Value
        closePrice = Cells(i, 6).Value
        YrChange = closePrice - openPrice
        PctChange = (YrChange / openPrice)
        
        curSheet.Cells(groupNo + 1, 9).Value = currentName
        curSheet.Cells(groupNo + 1, 10).Value = YrChange
        curSheet.Cells(groupNo + 1, 11).Value = PctChange
        curSheet.Cells(groupNo + 1, 12).Value = totalSV
        
        totalSV = 0
        openPrice = Cells(i + 1, 3).Value
        groupNo = groupNo + 1
        
    End If

Next i

 ' Apply formatting to Yearly Change column
    For j = 2 To lastRow
        If curSheet.Cells(j, 10).Value > 0 Then
            curSheet.Cells(j, 10).Interior.ColorIndex = 4
        ElseIf curSheet.Cells(j, 10).Value < 0 Then
            curSheet.Cells(j, 10).Interior.ColorIndex = 3
        End If
    Next j

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

    ' Find and display greatest increase, decrease, and total volume
    greatest_increase = Application.WorksheetFunction.Max(curSheet.Range("K2:K" & lastRow))
    greatest_decrease = Application.WorksheetFunction.Min(curSheet.Range("K2:K" & lastRow))
    greatest_TotalVolume = Application.WorksheetFunction.Max(curSheet.Range("L2:L" & lastRow))

    curSheet.Cells(2, 17).Value = greatest_increase
    curSheet.Cells(2, 16).Value = curSheet.Cells(Application.WorksheetFunction.Match(greatest_increase, curSheet.Range("K2:K" & lastRow), 0) + 1, 9).Value

    curSheet.Cells(3, 17).Value = greatest_decrease
    curSheet.Cells(3, 16).Value = curSheet.Cells(Application.WorksheetFunction.Match(greatest_decrease, curSheet.Range("K2:K" & lastRow), 0) + 1, 9).Value

    curSheet.Cells(4, 17).Value = greatest_TotalVolume
    curSheet.Cells(4, 16).Value = curSheet.Cells(Application.WorksheetFunction.Match(greatest_TotalVolume, curSheet.Range("L2:L" & lastRow), 0) + 1, 9).Value
Next curSheet
        
    

    End Sub
