Sub StockData()
' Declare Current as a worksheet object variable.
    Dim Current As Worksheet

' Loop through all of the worksheets in the active workbook.
     For Each Current In Worksheets
      
' Determine the Last Row
    LastRow = Current.Cells(Rows.Count, 1).End(xlUp).Row

' Set an initial variable for holding the Stock ticker Name
    Dim StockTicker As String

' Set an initial variable for holding the total volume per stock ticker
    Dim TotalVolume As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentYearlyChange As String
    TotalVolume = 0
    OpenPrice = 0
    ClosePrice = 0
    YearlyChange = 0
'Enter the column headers
    Current.Cells(1, "H").Value = "Stock Ticker"
    Current.Cells(1, "I").Value = "Total Stock Vol"
    Current.Cells(1, "J").Value = "Yearly Change"
    Current.Cells(1, "K").Value = "PercentYearly Change"

    

' Keep track of the location for each stock ticker in the summary table
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2

' Loop through all stock data
    For I = 2 To LastRow
    
' Check if we are still within the same Stock Ticker, if yes...
    If Current.Cells(I + 1, 1).Value = Current.Cells(I, 1).Value Then

'Add to the Total Volume
      TotalVolume = TotalVolume + Current.Cells(I, 7).Value
      If (OpenPrice = 0) Then
      OpenPrice = Current.Cells(I, 3)
      End If

    
' If the cell immediately following a row is the different Stock Ticker...
    Else
' Set the Stock ticker name
      StockTicker = Current.Cells(I, 1).Value

' Add to the TotalVolume
      TotalVolume = TotalVolume + Current.Cells(I, 7).Value
      
' Print the Stock Ticker in the Summary Table
      Current.Range("H" & SummaryTableRow).Value = StockTicker

' Print the Total Volume to the Summary Table
      Current.Range("I" & SummaryTableRow).Value = TotalVolume
      ClosePrice = Current.Cells(I, 6).Value
    
      
' Only if open price and close price for the ticker exist,

      If (Current.Cells(I, 3) <> 0) Then
      
        YearlyChange = ClosePrice - OpenPrice
        PercentYearlyChange = Str(Round((YearlyChange / OpenPrice) * 100)) & "%"
'Print the Yearly Change in the summary Table
        Current.Range("J" & SummaryTableRow).Value = YearlyChange
        Current.Range("K" & SummaryTableRow).Value = PercentYearlyChange
        
' If the open price and close price are 0, then enter 0 for yearly change and percent change
        Else
        Current.Range("J" & SummaryTableRow).Value = 0
        Current.Range("K" & SummaryTableRow).Value = "0%"
        End If
' if yearly change is positive, color it green

    If (Current.Cells(SummaryTableRow, 10).Value > 0) Then

        Current.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4

' Otherwise color it red
      Else

        Current.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3

    End If

' Move to the next row in the summary table
        SummaryTableRow = SummaryTableRow + 1

' Reset the TotalVolume,nStart,nEnd,OpenPrice,ClosePrice,YearlyChange

      TotalVolume = 0
      OpenPrice = 0
      ClosePrice = 0
      YearlyChange = 0
    
      
    End If
    Next I
    
' This line displays the worksheet name in a message box.
        MsgBox Current.Name
    Next Current

    End Sub


