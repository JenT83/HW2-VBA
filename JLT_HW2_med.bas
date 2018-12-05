Attribute VB_Name = "Module2"
'JLT HW2 solution

Sub medium()

'Loop through all tabs
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

' Set an initial variable for holding the ticker
  Dim Ticker As String
  
  'Set variables
  Dim OpenPrice As Double
  Dim ClosePrice As Double
  Dim YearChange As Double
  Dim PercentChange As Double

  ' Set an initial variable for holding the total
  Dim TotalStockVolume As Double
  
  ' Set initial value
  TotalStockVolume = 0
  OpenPrice = Cells(2, 3).Value
  ClosedPrice = 0
  YearlyChange = 0
  PercentChange = 0

  ' Make summary table headers
  Range("I1") = "Ticker"
  Range("J1") = "Yearly Change"
  Range("K1") = "Percent Change"
  Range("L1") = "Total Stock Volume"

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  ' Determine the Last Row
  Dim LastRow As Long

    LastRow = WS.Cells(Rows.Count, "A").End(xlUp).Row

    For j = 2 To LastRow

    ' Check if we are still within the same ticker, if it is not...
        If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then

        'Set close price
        ClosedPrice = Cells(j, 6).Value


      ' Set the ticker
        Ticker = Cells(j, 1).Value

        ' Add to the total
        TotalStockVolume = TotalStockVolume + Cells(j, 7).Value

      'Determine yearly and percent change
        YearlyChange = ClosedPrice - OpenPrice

        If OpenProce = 0 And ClosedPrice = 0 Then
            PercentChange = 0
        Else
            PercentChange = ((ClosedPrice / OpenPrice) - 1)
        End If
        
      ' Print the ticker in the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker
        
      ' Print the yearly and percent change in the Summary Table
        Range("J" & Summary_Table_Row).Value = YearlyChange
        Range("K" & Summary_Table_Row).Value = PercentChange

      ' Print the total to the Summary Table
         Range("L" & Summary_Table_Row).Value = TotalStockVolume

      ' Add one to the summary table row
         Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the total
        TotalStockVolume = 0
        
        'Set open price
        OpenPrice = Cells(j + 1, 3).Value
      

    ' If the cell immediately following a row is the same brand...
        Else

      ' Add to the total
        TotalStockVolume = TotalStockVolume + Cells(j, 7).Value
        
        ' Stock open
        If OpenPrice = 0 Then
            OpenPrice = Cells(j, 3).Value
        End If
        
        End If

    Next j

  ' Conditional format cells
  Dim k As Long
  Dim LastRow2 As Long
    LastRow2 = WS.Cells(Rows.Count, "J").End(xlUp).Row
     
    For k = 2 To LastRow2
    
        If Cells(k, 10).Value >= 0 Then
            Cells(k, 10).Interior.ColorIndex = 4
        Else
            Cells(k, 10).Interior.ColorIndex = 3
        End If
    Next k

  Columns("K:K").Select
  Selection.Style = "Percent"
  Selection.NumberFormat = "0.00%"

'Fix the width of the column
ActiveSheet.Columns("A:P").AutoFit
       
Next WS

End Sub





