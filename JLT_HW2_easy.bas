Attribute VB_Name = "Module1"
'JLT HW2 solution

Sub easy()

'Loop through all tabs
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

' Set an initial variable for holding the ticker
  Dim Ticker As String

  ' Set an initial variable for holding the total
  Dim TotalStockVolume As Double
  
  ' Set initial value
  TotalStockVolume = 0

  ' Make summary table headers
  Range("I1") = "Ticker"
  Range("J1") = "Total Stock Volume"

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all tickers
  Dim i As Long
  
  ' Determine the Last Row
  Dim LastRow As Long

    LastRow = WS.Cells(Rows.Count, "A").End(xlUp).Row
     
    For i = 2 To LastRow
    'The maximum number of rows in the final wkbk is 797711

    ' Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker
        Ticker = Cells(i, 1).Value

      ' Add to the total
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value

      ' Print the ticker in the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the total to the Summary Table
         Range("J" & Summary_Table_Row).Value = TotalStockVolume

      ' Add one to the summary table row
         Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the total
        TotalStockVolume = 0

    ' If the cell immediately following a row is the same brand...
        Else

      ' Add to the total
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value

        End If

    Next i

    'Fix the width of the columns
    ActiveSheet.Columns("A:P").AutoFit

Next WS

End Sub




