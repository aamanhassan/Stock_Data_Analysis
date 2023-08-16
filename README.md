# Stock_Data_Analysis
Vba Scripting for Stocks
Sub Stock_Data()

    Dim Change As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim StartRow As Long
    Dim RowIndex As Long
    Dim RowCount As Long
    Dim ColumnIndex As Long
    Dim ws As Worksheet
    
    ' Loop through each worksheet
    For Each ws In Worksheets
        
        TotalVolume = 0
        Change = 0
        ColumnIndex = 0
        StartRow = 2  ' Start counting from the second row
        
    
        ' Find the last used row number 
        RowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
    
        For RowIndex = 2 To RowCount
            ' Check if the ticker symbol changes from the previous row
            If ws.Cells(RowIndex + 1, 1).Value <> ws.Cells(RowIndex, 1).Value Then
                ' Calculate the total volume for the current ticker
                TotalVolume = TotalVolume + ws.Cells(RowIndex, 7).Value
                
                ' Calculate results for the previous ticker
                If TotalVolume <> 0 Then
                    ' Calculate change and percentage change
                    Change = ws.Cells(RowIndex, 6).Value - ws.Cells(StartRow, 3).Value
                    If ws.Cells(StartRow, 3).Value <> 0 Then
                        PercentChange = Change / ws.Cells(StartRow, 3).Value
                    Else
                        PercentChange = 0
                    End If
                    
                    ' Display results in appropriate columns
                    ws.Range("I" & 2 + ColumnIndex).Value = ws.Cells(StartRow, 1).Value
                    ws.Range("J" & 2 + ColumnIndex).Value = Change
                    ws.Range("J" & 2 + ColumnIndex).NumberFormat = "0.00"
                    ws.Range("K" & 2 + ColumnIndex).Value = PercentChange
                    ws.Range("K" & 2 + ColumnIndex).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + ColumnIndex).Value = TotalVolume
                    
                    ' Apply formatting to show positive changes in green and negative changes in red
                    If Change > 0 Then
                        ws.Range("J" & 2 + ColumnIndex).Interior.ColorIndex = 4 ' Green
                    ElseIf Change < 0 Then
                        ws.Range("J" & 2 + ColumnIndex).Interior.ColorIndex = 3 ' Red
                    Else
                        ws.Range("J" & 2 + ColumnIndex).Interior.ColorIndex = 0
                    End If
                    
                    ' Move to the next set of data
                    StartRow = RowIndex + 1
                    ColumnIndex = ColumnIndex + 1
                End If
                
                ' Reset variables for the new ticker
                TotalVolume = 0
                Change = 0
            Else
                ' Accumulate volume for the same ticker
                TotalVolume = TotalVolume + ws.Cells(RowIndex, 7).Value
            End If
        Next RowIndex
        
        'Finding Greatest Percentage Increase and Decrease and Greatest Total Volume
        
        ' Initialize RowIndex before using it
        RowIndex = 2
        
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))

        Increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowIndex)), ws.Range("K2:K" & RowIndex), 0)
        Decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowIndex)), ws.Range("K2:K" & RowIndex), 0)
        Volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowIndex)), ws.Range("L2:L" & RowIndex), 0)

        ws.Range("P2") = ws.Cells(Increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(Decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(Volume_number + 1, 9)
        
    Next ws

End Sub
