Sub MainLoop()
' Set Screen updating to FALSE after the code
Application.ScreenUpdating = False
Application.EnableEvents = False
' Define Worksheet and Workbook
Dim wb As Workbook
Dim Currentws As Worksheet
Dim wsHeaders() As Variant
' Set the active workbook
Set wb = ActiveWorkbook
' Define the variant array for managing header information for each sheet
wsHeaders() = Array("Ticker ", "Date ", "Open", "High", "Low", "Close", "Volume", " ", "Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", " ", " ", " ", "Ticker", "Value")
' Populate the Header Row on each sheet
For Each Currentws In wb.Sheets
' Initialize the values of the first row
    With Currentws
    .Rows(1).Value = " "
' Loop through the variant array to populate the column headers across the first row
    For i = LBound(wsHeaders()) To UBound(wsHeaders())
    .Cells(1, 1 + i).Value = wsHeaders(i)
    Next i
' Format the populated values consistently and Autofit the Header Row
    .Rows(1).Font.Bold = True
    .Rows(1).VerticalAlignment = xlCenter
    .Rows(1).HorizontalAlignment = xlCenter
    .Rows(1).Interior.ColorIndex = 40
    End With
' Repeat for other worksheets
Next Currentws
' Loop through all sheets in the Workbook for variable declarations and subsequent computations
For Each Currentws In Worksheets
    ' Define the variables to be used for calculations
    Dim current_ticker As String
    current_ticker = " "
    Dim Total_Ticker_Annual_Volume As Double
    Total_Ticker_Annual_Volume = 0
    Dim Ticker_Price1 As Double
    Ticker_Price1 = 0
    Dim Ticker_End_Price As Double
    Ticker_End_Price = 0
    ' Define all derived variables based on calculations
    Dim Yearly_Price_Change As Double
    Yearly_Price_Change = 0
    Dim YearlyPercentPriceChange As Double
    YearlyPercentPriceChange = 0
    Dim Max_current_ticker As String
    Max_current_ticker = " "
    Dim Min_current_ticker As String
    Min_current_ticker = " "
    Dim Max_Percent As Double
    Max_Percent = 0
    Dim Min_Percent As Double
    Min_Percent = 0
    Dim Max_Volume_current_ticker As String
    Max_Volume_current_ticker = " "
    Dim Max_Volume As Double
    Max_Volume = 0
    ' Define variable for Summary Table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    ' Define variable for the total row count within the worksheet
    Dim Lastrow_Sheet As Long
    ' Compute the last row in each sheet using the Range End Method
    Lastrow_Sheet = Currentws.Cells(Rows.Count, 1).End(xlUp).Row
    ' Initialize the beginning stock price for the first ticker in a worksheet
    Ticker_Price1 = Currentws.Cells(2, 3).Value
    ' --------------------------------------------------------------------------
    'Loop each sheet from the first row (i.e Row 2) until the end of sheet
    For i = 2 To Lastrow_Sheet
        ' Check difference in Stock Ticker between successive rows
        If Currentws.Cells(i + 1, 1).Value <> Currentws.Cells(i, 1).Value Then
            ' Set the Stock Ticker Value
            current_ticker = Currentws.Cells(i, 1).Value
            ' Calculate Year End Price and Yearly Price Change
            Ticker_End_Price = Currentws.Cells(i, 6).Value
            Yearly_Price_Change = Ticker_End_Price - Ticker_Price1
            ' Calculate Yearly Price Change (in %)
            If Ticker_Price1 <> 0 Then
                YearlyPercentPriceChange = (Yearly_Price_Change / Ticker_Price1) * 100
            End If
            ' Calculate the Annual Volume for a particular stock
            Total_Ticker_Annual_Volume = Total_Ticker_Annual_Volume + Currentws.Cells(i, 7).Value
            ' Print the Stock Ticker Value in the corresponding column
            Currentws.Range("I" & Summary_Table_Row).Value = current_ticker
            ' Print the Annual Price change in the corresponding column
            Currentws.Range("J" & Summary_Table_Row).Value = Yearly_Price_Change
            ' Color fill the Annual Price change column based on the given requirements: RED for Negative, GREEN for Positive
            If (Yearly_Price_Change > 0) Then
                Currentws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf (Yearly_Price_Change <= 0) Then
                Currentws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            ' Print the percent change in the Annual Price of the Ticker in the corresponding column
            Currentws.Range("K" & Summary_Table_Row).Value = (CStr(YearlyPercentPriceChange) & "%")
            ' Lastly, print the Annual Volume of the Ticker in the corresponding column
            Currentws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Annual_Volume
            ' Increment the row count for the summary table
            Summary_Table_Row = Summary_Table_Row + 1
            ' Set the value for the next beginning ticket price in the loop
            Ticker_Price1 = Currentws.Cells(i + 1, 3).Value
            ' Calculate the MAX, MIN percent change in price
            If (YearlyPercentPriceChange > Max_Percent) Then
                Max_Percent = YearlyPercentPriceChange
                Max_current_ticker = current_ticker
            ElseIf (YearlyPercentPriceChange < Min_Percent) Then
                Min_Percent = YearlyPercentPriceChange
                Min_current_ticker = current_ticker
            End If
            ' Calculate the greatest annual volume across all ticker symbols
            If (Total_Ticker_Annual_Volume > Max_Volume) Then
                Max_Volume = Total_Ticker_Annual_Volume
                Max_Volume_current_ticker = current_ticker
            End If
            ' Reset values before the next iteration of the loop
            YearlyPercentPriceChange = 0
            Total_Ticker_Annual_Volume = 0
        ' Add Ticker Volume cumulatively for the same Ticker
        Else
            Total_Ticker_Annual_Volume = Total_Ticker_Annual_Volume + Currentws.Cells(i, 7).Value
        ' Main IF loop ends
        End If
    ' FOR loop ends - iterate through the next row of the sheet
    Next i

    Currentws.Range("O2").Value = "Greatest % Increase"
   
    Currentws.Range("O3").Value = "Greatest % Decrease"
    
    Currentws.Range("O4").Value = "Greatest Total Volume"
    
    Currentws.Range("O2:O4").Rows.AutoFit
    ' Assign values to each corresponding cell based on prior calculations
    Currentws.Range("P2").Value = Max_current_ticker
    Currentws.Range("P3").Value = Min_current_ticker
    Currentws.Range("P4").Value = Max_Volume_current_ticker
    Currentws.Range("P2:P4").Rows.AutoFit
    Currentws.Range("Q2").Value = (CStr(Max_Percent) & "%")
    Currentws.Range("Q3").Value = (CStr(Min_Percent) & "%")
    Currentws.Range("Q4").Value = Max_Volume
    Currentws.Range("Q2:Q4").Rows.AutoFit
    'Main Loop Ends
Next Currentws
' Move to the next worksheet instance and iterate the same set of steps
' Set Screen updating to TRUE before the end of the procedure
Application.ScreenUpdating = True
Application.EnableEvents = True
End Sub