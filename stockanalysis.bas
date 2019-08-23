Attribute VB_Name = "Module1"
Sub stock()

For Each WS In Worksheets
WS.Activate

    'set last row variable
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox (Str(lastRow))
    
    'ticker variable to hold ticker name
    Dim currentTicker As String
    'volume variable
    Dim Volume As Double
    Volume = 0
    'row variable to increment
    Dim Row As Integer
    Row = 2
    
    'yearly change variables
    Dim opening As Double
    Dim closing As Double
    Dim yearlyChange As Double
    
    'percent change variable
    Dim percentChange As Double
    
    'label columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Volume"
    
    'initial opening value
    opening = Cells(Row, 3).Value
    
    'loop through ticker column
    For i = 2 To LastRow
    
        'check if ticker name changes
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            '~~easy~~
            currentTicker = Cells(i, 1).Value
            'set ticker name
            Cells(Row, 9) = currentTicker
            'add to volume
            Volume = Volume + Cells(i, 7).Value
            'set volume value
            Cells(Row, 12) = Volume
            'reset volume
            Volume = 0
            
            '~~medium~~
            closing = Cells(i, 6).Value
            yearlyChange = closing - opening
            Cells(Row, 10) = yearlyChange
            'calculate percent change
            If opening = 0 Then
                percentChange = 0
            ElseIf (opening = 0 And closing = 1) Then
                percentChange = 1
            Else
                percentChange = yearlyChange / opening
                Cells(Row, 11) = percentChange
                Cells(Row, 11).NumberFormat = "0.00%"
            End If
            
            'reset opening price
            opening = Cells(i + 1, 3).Value
            
            'increment the row
               Row = Row + 1
        Else
            'ticker name the same, add to volume
            Volume = Volume + Cells(i, 7).Value
        End If
    
    Next i
    
    'last row of the yearly change column
    yearlyChangeLast = WS.Cells(Rows.Count, 10).End(xlUp).Row
    
    'conditionally format cells
    For j = 2 To yearlyChangeLast
        If (Cells(j, 10).Value >= 0) Then
            Cells(j, 10).Interior.ColorIndex = 10
        Else
            Cells(j, 10).Interior.ColorIndex = 3
        End If
    Next j
    
    '~~hard~~
    'label table
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'loop through rows
    For k = 2 To yearlyChangeLast
        'check for greatest % increase and assign
        If Cells(k, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & yearlyChangeLast)) Then
            Cells(2, 16).Value = Cells(k, 9).Value
            Cells(2, 17).Value = Cells(k, 11).Value
            Cells(2, 17).NumberFormat = "0.00%"
        End If
        'check for greatest % decrease and assign
        If Cells(k, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & yearlyChangeLast)) Then
            Cells(3, 16).Value = Cells(k, 9).Value
            Cells(3, 17).Value = Cells(k, 11).Value
            Cells(3, 17).NumberFormat = "0.00%"
        End If
        'check for greatest total volume and assign
        If Cells(k, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & yearlyChangeLast)) Then
            Cells(4, 16).Value = Cells(k, 9).Value
            Cells(4, 17).Value = Cells(k, 12).Value
        End If
    Next k

'autofit columns
Range("I:Q").EntireColumn.AutoFit

Next WS

End Sub
