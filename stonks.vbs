Sub stonks()
'Declaring worksheet variable
Dim ws As Worksheet
'Declaring row count, summary row and total volume as integers
Dim row, sumrow As Integer
'Declaring open and closing quantities to store prices.
Dim opn, cls, tvol, lrow  As Double
'Declaring ticker name.
Dim ticker As String

'Cycling through worksheets.
For Each ws In Worksheets
    ws.Activate
    'Headings for summary
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    'Starting summary under the headings and finding last row in our data
    sumrow = 2
    'Finding the last row in our raw data.
    lrow = Cells(Rows.Count, 1).End(xlUp).row
    'Checking each row and surrounding rows to determine unique tickers and other data.
    For i = 2 To lrow
        'Checking to see the end of ticker streaker to do final calculations.
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            'This if statement is to avoid division by zero.
            'Both internal processes summarize the data in desired format.
            If opn = 0 Then
                Cells(sumrow, 9).Value = Cells(i, 1).Value
                Cells(sumrow, 10).Value = Cells(i, 6).Value - opn
                Cells(sumrow, 10).NumberFormat = "$#,##0.00"
                Cells(sumrow, 11).Value = 0
                Cells(sumrow, 11).NumberFormat = "0.00%"
                Cells(sumrow, 12).Value = tvol + Cells(i, 7).Value
            Else
                Cells(sumrow, 9).Value = Cells(i, 1).Value
                Cells(sumrow, 10).Value = Cells(i, 6).Value - opn
                Cells(sumrow, 10).NumberFormat = "$#,##0.00"
                Cells(sumrow, 11).Value = Cells(sumrow, 10).Value / opn
                Cells(sumrow, 11).NumberFormat = "0.00%"
                Cells(sumrow, 12).Value = tvol + Cells(i, 7).Value
            End If
            'Conditional color formatting in the summative data.
            If Cells(sumrow, 10).Value > 0 Then
                Cells(sumrow, 10).Interior.ColorIndex = 4
                Cells(sumrow, 11).Interior.ColorIndex = 4
            ElseIf Cells(sumrow, 10).Value < 0 Then
                Cells(sumrow, 10).Interior.ColorIndex = 3
                Cells(sumrow, 11).Interior.ColorIndex = 3
            Else
                Cells(sumrow, 10).Interior.ColorIndex = 6
                Cells(sumrow, 11).Interior.ColorIndex = 6
            End If
            'Since is this done at the end we use the end of this code block to move to the next summative row.
            'Additionally, we reset the total volume to zero.
            sumrow = sumrow + 1
            tvol = 0
        'This is going to be triggered when we start a new ticker and thus we collect opening price.
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            opn = Cells(i, 3).Value
            tvol = tvol + Cells(i, 7).Value
        'When we aren't at the beginning or end of ticker count we just count the volume total.
        Else
            tvol = tvol + Cells(i, 7).Value
        End If
    Next i
Next ws
End Sub