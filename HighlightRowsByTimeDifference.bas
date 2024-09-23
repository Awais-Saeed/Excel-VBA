Attribute VB_Name = "Module1"
Sub HighlightRowsByTimeDifference()
    
    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim timeDiff As Double
    Dim threshold As Double
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name
    
    ' Get the last row with data in column B
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Threshold in minutes
    threshold = 4
    
    ' Loop through rows starting from the second row
    For i = 2 To lastRow
        ' Calculate the time difference in minutes. Time data is in Column B
        If IsDate(ws.Cells(i, 2).Value) And IsDate(ws.Cells(i - 1, 2).Value) Then
            timeDiff = DateDiff("n", ws.Cells(i - 1, 2).Value, ws.Cells(i, 2).Value)

            Debug.Print timeDiff
            
            ' If the time difference is greater than the threshold
            If timeDiff >= threshold Then
                ' Highlight the entire row
                ws.Rows(i).Interior.Color = RGB(255, 0, 0) ' Change to your preferred color
            End If
        End If
    Next i
End Sub

