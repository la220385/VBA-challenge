Sub StockMarket()

    'Declared constants
    Const COLOR_RED As Integer = 3
    Const COLOR_GREEN As Integer = 4

    'Declared variables
    Dim Last_Row As Double
    Dim Total_Stock_Volume As LongLong
    Dim Table_Row As Double
    Dim Quarter_Begin_Value As Double
    Dim Quarter_End_Value As Double
    Dim Ticker As String
    Dim Date_Begin As Long
    Dim Date_Input As Long
    Dim Quarterly_Change_Value As Double
    Dim Percent_Change_Fraction As Double
    Dim Greatest_Increase_Value As Double
    Dim Greatest_Increase_Ticker As String
    Dim Greatest_Decrease_Value As Double
    Dim Greatest_Decrease_Ticker As String
    Dim Greatest_Volume_Value As LongLong
    Dim Greatest_Volume_Ticker As String
    Dim Input_Row As Long
    Dim ws As Worksheet

    'Loop through all worksheets
    For Each ws In Worksheets

        'Generating and formatting table labels
        ws.Range("I1, P1").Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Range("I1:L1, P1:Q1, O2:O4").Font.Bold = True

        'Determined the last row for each sheet
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Before loop begins, determine placement for the first row of output
        Table_Row = 2

        'Set starting values
        Greatest_Volume_Value = -1
        Greatest_Increase_Value = -99999
        Greatest_Decrease_Value = 99999
        Total_Stock_Volume = 0
        Date_Begin = 99999999

        'Nested For Loop that goes through all data on the worksheet
        For Input_Row = 2 To Last_Row

            Ticker = ws.Cells(Input_Row, 1).Value
            Date_Input = ws.Cells(Input_Row, 2).Value
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(Input_Row, 7).Value

            'First If statement
            If Ticker = ws.Cells(Input_Row + 1, 1).Value Then
                'Determine Quarter_Begin_Value
                If Date_Input < Date_Begin Then
                    Date_Begin = Date_Input
                    Quarter_Begin_Value = ws.Cells(Input_Row, 3).Value
                End If

            Else
                'Determine Quarter_End_Value
                Quarter_End_Value = ws.Cells(Input_Row, 6).Value
                'Calculation for Quarterly Change
                Quarterly_Change_Value = Quarter_End_Value - Quarter_Begin_Value
                'Calculation for Percent Change
                If Quarter_Begin_Value <> 0 Then
                    Percent_Change_Fraction = (Quarterly_Change_Value / Quarter_Begin_Value)
                Else
                    Percent_Change_Fraction = 0
                End If

                'Calculation for Greatest Total Volume
                If Total_Stock_Volume > Greatest_Volume_Value Then
                    Greatest_Volume_Value = Total_Stock_Volume
                    Greatest_Volume_Ticker = Ticker
                End If

                'Calculations for Greatest % increase and decrease
                If Percent_Change_Fraction > Greatest_Increase_Value Then
                    Greatest_Increase_Value = Percent_Change_Fraction
                    Greatest_Increase_Ticker = Ticker
                End If

                If Percent_Change_Fraction < Greatest_Decrease_Value Then
                    Greatest_Decrease_Value = Percent_Change_Fraction
                    Greatest_Decrease_Ticker = Ticker
                End If

                'Populate the table with output values
                ws.Cells(Table_Row, 9).Value = Ticker
                ws.Cells(Table_Row, 10).Value = Quarterly_Change_Value
                'Format Quarterly Change with colors
                If Quarterly_Change_Value >= 0 Then
                    ws.Cells(Table_Row, 10).Interior.ColorIndex = COLOR_GREEN
                Else
                    ws.Cells(Table_Row, 10).Interior.ColorIndex = COLOR_RED
                End If

                'Populate Percent Change without color
                ws.Cells(Table_Row, 11).Value = FormatPercent(Percent_Change_Fraction)

                'Populate Total Stock Volume
                ws.Cells(Table_Row, 12).Value = Total_Stock_Volume

                'Reset Total_Stock_Volume and Date_Begin for next loop
                Total_Stock_Volume = 0
                Date_Begin = 99999999
                'Determine where the next row of data output will start for the next loop
                Table_Row = Table_Row + 1

            End If

        Next Input_Row

        'Populate second table with output values
        ws.Range("P4").Value = Greatest_Volume_Ticker
        ws.Range("Q4").Value = Greatest_Volume_Value
        ws.Range("P2").Value = Greatest_Increase_Ticker
        ws.Range("Q2").Value = FormatPercent(Greatest_Increase_Value)
        ws.Range("P3").Value = Greatest_Decrease_Ticker
        ws.Range("Q3").Value = FormatPercent(Greatest_Decrease_Value)

        'Final formatting of tables
        ws.Columns("I:Q").AutoFit
        ws.Columns("I").HorizontalAlignment = xlLeft
        ws.Columns("O:P").HorizontalAlignment = xlLeft

    Next ws

End Sub

