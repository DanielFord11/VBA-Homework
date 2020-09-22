Attribute VB_Name = "Module1"
Sub Stocks()

Dim SavedStocks() As String
Dim ColumnSize As Long
Dim BonusColumn As Long
Dim TablePlace As Integer
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double
Dim AmIncrease As Double
Dim TickersArray() As String
Dim ColorIndex As Integer
Dim ws As Worksheet

Application.ScreenUpdating = False

                                            'Runs main code on each worksheet
For Each ws In Worksheets
    ws.Activate
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
        


                                            'Counter for place on output table and color index
    TablePlace = 2
    ColorIndex = 3


                                            'Measures the row count to set the array size and iteration count
    ColumnSize = ws.Cells(Rows.Count, "A").End(xlUp).Row
    ReDim TickersArray(ColumnSize)
    ReDim SavedStocks(7)

                                            'Creates an Array to hold all the Ticker Symbols for each entry
    For Tickers = 0 To ColumnSize
        TickersArray(Tickers) = ws.Cells(Tickers + 1, 1).Value
    Next Tickers

    '------Main loop to extract open/close values and calulate total volume/change per stock
    For i = 1 To ColumnSize
                                            'Compares ticker sybols within the array to determine if it's an opening stock
       If TickersArray(i - 1) <> TickersArray(i) Then
                                            'Adds opening day volume to table
              ws.Cells(TablePlace, 12) = ws.Cells(TablePlace, 12) + ws.Cells(i + 1, 7).Value
                                            'Adds opening stock values to array for later reference
            For j = 0 To 6
                SavedStocks(j) = ws.Cells(i + 1, j + 1).Value
            Next j

                                            'Compares ticker symbols within the array to determine if it's a closing stock
        ElseIf TickersArray(i) <> TickersArray(i + 1) Or i = ColumnSize Then

                                            'Handles divide by zero error for stocks that open as zero (sets to 1)
                On Error GoTo ZeroHandle
                                            'Determines which color to apply for conditional formating
                If ws.Cells(i + 1, 6).Value - SavedStocks(2) > 0 Then
                        ColorIndex = 4
                End If

                                            'Calculates/enters change/volume stats to table and clears the opening stock array
                  ws.Cells(TablePlace, 9) = SavedStocks(0)
                  ws.Cells(TablePlace, 10) = ws.Cells(i + 1, 6).Value - SavedStocks(2)
                  ws.Cells(TablePlace, 10).Interior.ColorIndex = ColorIndex
                  ws.Cells(TablePlace, 11) = Round((ws.Cells(i + 1, 6).Value - SavedStocks(2)) / SavedStocks(2), 4)
                  ws.Cells(TablePlace, 11).Interior.ColorIndex = ColorIndex
                  ws.Cells(TablePlace, 11).NumberFormat = "0.00%"
                  ws.Cells(TablePlace, 12) = ws.Cells(TablePlace, 12) + ws.Cells(i + 1, 7).Value
                TablePlace = TablePlace + 1
                ReDim SavedStocks(7)
                ColorIndex = 3

        Else
                                             'Increments the total volume for the non-opening/closing stocks
              ws.Cells(TablePlace, 12) = ws.Cells(TablePlace, 12) + ws.Cells(i + 1, 7).Value
        End If
    Next i

    '-------Bonus: Greatest Increase, Decrease, and Volume

    BonusColumn = ws.Cells(Rows.Count, "I").End(xlUp).Row

    'Increase
    For i = 2 To BonusColumn
        If GreatestIncrease < ws.Cells(i, 11).Value Then
            GreatestIncrease = ws.Cells(i, 11).Value
            ws.Range("N3").Value = ws.Cells(i, 9)
            ws.Range("O3") = GreatestIncrease
        End If
    Next i

    ws.Range("O3").NumberFormat = "0.00%"
    ws.Range("M3") = "Greatest Increase"


    'Decrease
    For i = 2 To BonusColumn
        If ws.Cells(i, 11).Value < GreatestDecrease Then
            GreatestDecrease = ws.Cells(i, 11).Value
            ws.Range("N4").Value = ws.Cells(i, 9)
            ws.Range("O4") = GreatestDecrease
        End If
    Next i

    ws.Range("O4").NumberFormat = "0.00%"
    ws.Range("M4") = "Greatest Decrease"

    'Greatest Volume
    For i = 2 To BonusColumn
        If GreatestVolume < ws.Cells(i, 12).Value Then
            GreatestVolume = ws.Cells(i, 12).Value
            ws.Range("N5").Value = ws.Cells(i, 9)
            ws.Range("O5") = GreatestVolume
        End If
    Next i

    ws.Range("M5") = "Greatest Total Volume"
    GreatestIncrease = 0
    GreatestDecrease = 100
    GreatestVolume = 0
    Application.ScreenUpdating = True

Next

Exit Sub

ZeroHandle:
            SavedStocks(2) = 1
            Resume

End Sub


