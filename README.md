# VBA_challenge
Module 2 Challenge


    Sub ()
    Dim currentname As String
    Dim nextname As String
    Dim totalSV As Double
    Dim openYear As Double
    Dim closeYear As Double
    Dim lastrow As Long
    Dim curSheet As Worksheet
    Dim groupnumber As Long ' Added declaration
    Dim openprice As Double ' Added declaration

    ' Loop through all worksheets
    For Each curSheet In ThisWorkbook.Worksheets
        curSheet.Cells(1, 9).Value = "Ticker"
        curSheet.Cells(1, 10).Value = "Yearly Change"
        curSheet.Cells(1, 11).Value = "Percent Change"
        curSheet.Cells(1, 12).Value = "Total Stock Volume"
        totalSV = 0
        groupnumber = 1
        openprice = curSheet.Cells(2, 3).Value
        lastrow = curSheet.Cells(curSheet.Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow
            currentname = curSheet.Cells(i, 1).Value
            nextname = curSheet.Cells(i + 1, 1).Value

            If nextname = currentname Then
                totalSV = totalSV + curSheet.Cells(i, 7).Value
            Else
                totalSV = totalSV + curSheet.Cells(i, 7).Value
                closeprice = curSheet.Cells(i, 6).Value
                YrChange = closeprice - openprice
                PctChange = YrChange / openprice
                curSheet.Cells(groupnumber + 1, 9).Value = currentname
                curSheet.Cells(groupnumber + 1, 10).Value = YrChange
                curSheet.Cells(groupnumber + 1, 11).Value = PctChange
                curSheet.Cells(groupnumber + 1, 12).Value = totalSV
                totalSV = 0
                openprice = curSheet.Cells(i + 1, 3).Value
                groupnumber = groupnumber + 1
            End If
        Next i

        ' Apply formatting to Yearly Change column
        For j = 2 To lastrow
            If curSheet.Cells(j, 10).Value > 0 Then
                curSheet.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf curSheet.Cells(j, 10).Value < 0 Then
                curSheet.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j

        ' Find and display greatest increase, decrease, and total volume
        greatest_increase = Application.WorksheetFunction.Max(curSheet.Range("K2:K" & lastrow))
        greatest_decrease = Application.WorksheetFunction.Min(curSheet.Range("K2:K" & lastrow))
        greatest_TotalVolume = Application.WorksheetFunction.Max(curSheet.Range("L2:L" & lastrow))

        curSheet.Cells(2, 16).Value = greatest_increase
        curSheet.Cells(2, 17).Value = curSheet.Cells(Application.WorksheetFunction.Match(greatest_increase, curSheet.Range("K2:K" & lastrow), 0) + 1, 9).Value

        curSheet.Cells(3, 16).Value = greatest_decrease
        curSheet.Cells(3, 17).Value = curSheet.Cells(Application.WorksheetFunction.Match(greatest_decrease, curSheet.Range("K2:K" & lastrow), 0) + 1, 9).Value

        curSheet.Cells(4, 16).Value = greatest_TotalVolume
        curSheet.Cells(4, 17).Value = curSheet.Cells(Application.WorksheetFunction.Match(greatest_TotalVolume, curSheet.Range("L2:L" & lastrow), 0) + 1, 9).Value
    Next curSheet
    End Sub
